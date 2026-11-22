import React, { useEffect, useState } from 'react';
import { saveAs } from 'file-saver';
import { ListTree, Upload, FileDown, FileSpreadsheet, Trash2, Loader2, AlertTriangle } from 'lucide-react';
import { buildDeliverableTocIndex, formatPageLabelRanges } from '../utils/ssotValidator';
import { extractTocFromDeliverables, analyzeTocSample, buildTocPptx } from '../utils/pptTocBuilder';
import { buildTocWorkbook } from '../utils/tocExcelExporter';
import { updateTocDeck } from '../utils/pptTocUpdater';
import tocDefaultUrl from '../assets/templates/toc_default.pptx?url';

const DEFAULT_NAME = '목차_기본형(공공의료 AI ISP).pptx';
const KIND_LABEL = { cards: '카드형', table: '표형(테이블)', list: '텍스트박스형(행반복)' };

const panel = { background: 'rgba(0,0,0,0.2)', border: '1px solid var(--panel-border)', borderRadius: '12px', padding: '18px', display: 'flex', flexDirection: 'column', gap: '12px' };
const h3 = { margin: 0, fontSize: '15px', fontWeight: 700, color: 'var(--text-primary)' };
const muted = { fontSize: '12.5px', color: 'var(--text-secondary)', lineHeight: 1.6, margin: 0 };
const chip = (ok) => ({ fontSize: '12px', fontWeight: 600, padding: '3px 9px', borderRadius: '999px', background: ok ? 'rgba(16,185,129,0.14)' : 'rgba(148,163,184,0.14)', color: ok ? '#34d399' : 'var(--text-secondary)' });
const input = { background: 'rgba(0,0,0,0.3)', border: '1px solid var(--panel-border)', borderRadius: '6px', color: 'var(--text-primary)', padding: '6px 8px', fontSize: '13px', fontFamily: 'inherit' };
// 이 프로젝트에는 btn-primary 전역 스타일이 없어 버튼 모양·커서를 직접 지정
const actionBtn = (enabled, from = '#f97316', to = '#ea580c') => ({
    padding: '10px 16px', fontSize: '13.5px', fontWeight: 700, display: 'inline-flex', alignItems: 'center', gap: '8px', justifyContent: 'center',
    borderRadius: '8px', border: 'none', fontFamily: 'inherit', transition: 'filter .15s, transform .15s',
    cursor: enabled ? 'pointer' : 'not-allowed',
    color: enabled ? '#fff' : 'var(--text-secondary)',
    background: enabled ? `linear-gradient(135deg, ${from}, ${to})` : 'rgba(148,163,184,0.18)',
    opacity: enabled ? 1 : 0.7
});
const hoverOn = (e) => { if (!e.currentTarget.disabled) { e.currentTarget.style.filter = 'brightness(1.1)'; e.currentTarget.style.transform = 'translateY(-1px)'; } };
const hoverOff = (e) => { e.currentTarget.style.filter = ''; e.currentTarget.style.transform = ''; };
const dropZone = (active) => ({
    border: `2px dashed ${active ? '#f97316' : 'var(--panel-border)'}`, borderRadius: '10px', padding: '14px',
    background: active ? 'rgba(249,115,22,0.10)' : 'rgba(255,255,255,0.02)', display: 'flex', flexDirection: 'column', gap: '8px',
    alignItems: 'flex-start', transition: 'background .15s, border-color .15s'
});
const isPptx = (f) => /\.pptx$/i.test(f.name) && !f.name.startsWith('~$');

// 끌어서 놓기 영역 (파일 선택 버튼과 함께 사용)
function DropArea({ onFiles, children, testId }) {
    const [over, setOver] = useState(false);
    return (
        <div
            data-testid={testId}
            onDragOver={(e) => { e.preventDefault(); e.dataTransfer.dropEffect = 'copy'; if (!over) setOver(true); }}
            onDragEnter={(e) => { e.preventDefault(); setOver(true); }}
            onDragLeave={(e) => { if (!e.currentTarget.contains(e.relatedTarget)) setOver(false); }}
            onDrop={(e) => { e.preventDefault(); setOver(false); onFiles(Array.from(e.dataTransfer.files || [])); }}
            style={dropZone(over)}
        >
            {children}
        </div>
    );
}

const fileBtn = { display: 'inline-flex', alignItems: 'center', gap: '8px', padding: '9px 16px', borderRadius: '8px', border: '1px solid var(--panel-border)', cursor: 'pointer', fontSize: '13.5px', fontWeight: 600, color: 'var(--text-primary)', background: 'rgba(255,255,255,0.04)' };

// 샘플 목차 깊이 설명 (예: 2단계 · 절 › 항)
function depthInfo(slide) {
    if (!slide || !slide.levels) return null;
    if (slide.kind === 'cards') return { label: '장 › 절 (장마다 카드)', overview: '장 › 절', perChapter: '장 › 절', usesSubsections: false };
    const names = slide.topIsChapter ? ['장', '절', '항'] : ['절', '항', '세부 항목'];
    const chain = names.slice(0, slide.levels).join(' › ');
    const overviewChain = ['장', '절', '항'].slice(0, slide.levels).join(' › ');
    return { label: `${slide.levels}단계 · ${chain}`, overview: overviewChain, perChapter: chain, usesSubsections: slide.topIsChapter ? slide.levels >= 3 : slide.levels >= 2 };
}

export default function PptTocGeneratorTab({ setErrorMsg, setSuccessMsg }) {
    const [deliverables, setDeliverables] = useState([]);
    const [chapters, setChapters] = useState([]);
    const [isExtracting, setIsExtracting] = useState(false);
    const [sampleSource, setSampleSource] = useState('default');
    const [sampleFile, setSampleFile] = useState(null);
    const [sampleInfo, setSampleInfo] = useState(null);
    const [sel, setSel] = useState({ overviewIndex: null, detailIndex: null, chapterIndex: null });
    const [isAnalyzingSample, setIsAnalyzingSample] = useState(false);
    const [options, setOptions] = useState({ showPage: true, overviewTitle: '목차' });
    const [isBuilding, setIsBuilding] = useState(false);
    const [isExporting, setIsExporting] = useState(false);
    // 직접 올린 디자인: keep = 디자인 유지·목차와 페이지만 갱신, generate = 이 디자인으로 새로 생성
    const [customMode, setCustomMode] = useState('keep');
    const [updateResult, setUpdateResult] = useState(null);
    const [warnings, setWarnings] = useState([]);

    const resetMsg = () => { setErrorMsg?.(null); setSuccessMsg?.(null); };

    const addDeliverables = (list) => {
        const files = list.filter(isPptx);
        if (!files.length) { setErrorMsg?.('PPTX 파일만 올릴 수 있습니다.'); return; }
        resetMsg();
        setDeliverables(prev => {
            const names = new Set(prev.map(f => f.name));
            return [...prev, ...files.filter(f => !names.has(f.name))].sort((a, b) => a.name.localeCompare(b.name, 'ko'));
        });
        setChapters([]);
    };
    const onPickDeliverables = (e) => {
        const list = Array.from(e.target.files || []);
        e.target.value = '';
        if (list.length) addDeliverables(list);
    };

    const extract = async () => {
        resetMsg();
        setIsExtracting(true);
        try {
            const toc = await extractTocFromDeliverables(deliverables, buildDeliverableTocIndex, formatPageLabelRanges);
            if (!toc.some(c => c.sections.length)) {
                setErrorMsg?.('장·절 구조를 찾지 못했습니다. 슬라이드 상단 경로(예: "1. 환경분석 개요")나 목차 슬라이드가 있는 PPTX인지 확인해 주세요.');
            }
            setChapters(toc);
        } catch (err) {
            console.error(err);
            setErrorMsg?.(`목차 추출 중 오류가 발생했습니다: ${err.message}`);
        } finally {
            setIsExtracting(false);
        }
    };

    const loadSample = async (file, { quiet = false } = {}) => {
        setSampleFile(file);
        setSampleInfo(null);
        setUpdateResult(null);
        setIsAnalyzingSample(true);
        try {
            const info = await analyzeTocSample(file);
            setSampleInfo(info);
            setSel({ ...info.defaults });
            if (!quiet && info.defaults.overviewIndex == null && info.defaults.chapterIndex == null && info.defaults.detailIndex == null) {
                setErrorMsg?.('샘플에서 목차 항목 칸(같은 모양으로 반복되는 제목 칸·목차 표·장 카드)을 찾지 못했습니다. 목차 슬라이드에 {{제목}} 표시를 넣어 다시 올려 주세요.');
            }
        } catch (err) {
            console.error(err);
            setErrorMsg?.(`목차 샘플 분석 중 오류가 발생했습니다: ${err.message}`);
        } finally {
            setIsAnalyzingSample(false);
        }
    };

    // 기본형 샘플 (공공의료 AI ISP 표준 목차: 전체 목차 카드형 · 상세 목차 표형 · 장별 목차)
    const loadDefault = async () => {
        try {
            const res = await fetch(tocDefaultUrl);
            if (!res.ok) throw new Error(`HTTP ${res.status}`);
            const buf = await res.arrayBuffer();
            await loadSample(new File([buf], DEFAULT_NAME, { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }));
        } catch (err) {
            console.error(err);
            setErrorMsg?.(`기본형 목차 샘플을 불러오지 못했습니다: ${err.message}`);
        }
    };
    useEffect(() => { loadDefault(); }, []); // eslint-disable-line react-hooks/exhaustive-deps

    const chooseSource = (src) => {
        resetMsg();
        setSampleSource(src);
        if (src === 'default') loadDefault();
        else { setSampleFile(null); setSampleInfo(null); setSel({ overviewIndex: null, detailIndex: null, chapterIndex: null }); }
    };

    const applySampleFile = async (file) => {
        if (!file || !isPptx(file)) { setErrorMsg?.('목차 샘플은 PPTX 파일이어야 합니다.'); return; }
        resetMsg();
        setSampleSource('custom');
        await loadSample(file, { quiet: true });
    };
    const onPickSample = async (e) => {
        const file = e.target.files?.[0];
        e.target.value = '';
        if (file) await applySampleFile(file);
    };

    const updateChapter = (ci, patch) => setChapters(prev => prev.map((c, i) => (i === ci ? { ...c, ...patch } : c)));
    const updateSection = (ci, si, patch) => setChapters(prev => prev.map((c, i) => (i === ci
        ? { ...c, sections: c.sections.map((s, j) => (j === si ? { ...s, ...patch } : s)) }
        : c)));

    const keepMode = sampleSource === 'custom' && customMode === 'keep';
    const build = async () => {
        resetMsg();
        setWarnings([]);
        setUpdateResult(null);
        setIsBuilding(true);
        try {
            if (keepMode) {
                // 올린 디자인 그대로, 목차 글자와 시작 쪽만 점검·갱신
                const r = await updateTocDeck({ file: sampleFile, chapters, options: { showPage: options.showPage } });
                const d = new Date();
                const stamp = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
                saveAs(r.blob, `${sampleFile.name.replace(/\.pptx$/i, '')}_갱신_${stamp}.pptx`);
                setWarnings(r.warnings);
                setUpdateResult({ ...r, fileName: sampleFile.name });
                setSuccessMsg?.(`올린 디자인을 그대로 두고 ${r.updatedSlides.length}개 슬라이드의 목차를 점검했습니다. 바뀐 곳 ${r.changes.length}건 (아래 점검 결과 참고).`);
                return;
            }
            const { blob, warnings: w, slideCount, counts } = await buildTocPptx({ sampleFile, chapters, options: { ...options, ...sel } });
            const d = new Date();
            const stamp = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
            saveAs(blob, `목차_${stamp}.pptx`);
            setWarnings(w);
            setSuccessMsg?.(`목차 슬라이드 ${slideCount}장을 만들었습니다 (앞부분 목차 ${counts.front}장 · 장별 목차 ${counts.chapter}장). PowerPoint 구역으로 나뉘어 있어 장 앞에 옮겨 넣기 쉽습니다.`);
        } catch (err) {
            console.error(err);
            setErrorMsg?.(`목차 PPT 생성 중 오류가 발생했습니다: ${err.message}`);
        } finally {
            setIsBuilding(false);
        }
    };

    // 목차 엑셀 내려받기 (체크 해제 항목 제외, 고친 제목 반영)
    const exportExcel = async () => {
        resetMsg();
        setIsExporting(true);
        try {
            const { blob, counts } = await buildTocWorkbook(chapters);
            if (!counts.chapters) { setErrorMsg?.('내보낼 장이 없습니다. 포함할 장을 체크해 주세요.'); return; }
            const d = new Date();
            const stamp = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
            saveAs(blob, `목차_${stamp}.xlsx`);
            setSuccessMsg?.(`목차를 엑셀로 내려받았습니다 (${counts.chapters}개 장 · ${counts.sections}개 절 · ${counts.subsections}개 항).`);
        } catch (err) {
            console.error(err);
            setErrorMsg?.(`목차 엑셀 생성 중 오류가 발생했습니다: ${err.message}`);
        } finally {
            setIsExporting(false);
        }
    };
    const canExport = chapters.length > 0 && !isExporting;
    const excelBtn = (testId, compact = false) => (
        <button type="button" data-testid={testId} disabled={!canExport} onClick={exportExcel}
            onMouseEnter={hoverOn} onMouseLeave={hoverOff}
            style={{ ...actionBtn(canExport, '#10b981', '#059669'), ...(compact ? { padding: '7px 12px', fontSize: '12.5px' } : {}) }}>
            {isExporting ? <><Loader2 size={14} className="animate-spin" /> 엑셀 만드는 중…</> : <><FileSpreadsheet size={compact ? 14 : 16} /> 목차 엑셀로 내려받기</>}
        </button>
    );

    const slideAt = (i) => (i == null ? null : sampleInfo?.slides.find(s => s.index === i) || null);
    const overviewSlide = slideAt(sel.overviewIndex);
    const detailSlide = slideAt(sel.detailIndex);
    const chapterSlide = slideAt(sel.chapterIndex);
    const depth = depthInfo(chapterSlide || detailSlide || overviewSlide);
    const detailDepth = depthInfo(detailSlide);
    const usesSubsections = !!(depthInfo(chapterSlide)?.usesSubsections || detailDepth?.usesSubsections || (detailSlide?.kind === 'table' && detailSlide.levels >= 3));
    const updateSub = (ci, si, xi, patch) => setChapters(prev => prev.map((c, i) => (i === ci
        ? { ...c, sections: c.sections.map((s, j) => (j === si ? { ...s, subsections: s.subsections.map((x, k) => (k === xi ? { ...x, ...patch } : x)) } : s)) }
        : c)));
    const includedSections = chapters.filter(c => c.include).reduce((n, c) => n + c.sections.filter(s => s.include).length, 0);
    const includedSubs = chapters.filter(c => c.include).reduce((n, c) => n + c.sections.filter(s => s.include).reduce((m, s) => m + (s.subsections || []).filter(x => x.include).length, 0), 0);
    const canBuild = chapters.length > 0 && sampleFile && !isBuilding
        && (keepMode || sel.overviewIndex != null || sel.detailIndex != null || sel.chapterIndex != null);
    const slideSelect = (key, label, testId, allow) => (
        <label style={{ display: 'flex', flexDirection: 'column', gap: '4px', fontSize: '12.5px', color: 'var(--text-secondary)' }}>
            {label}
            <select data-testid={testId} value={sel[key] ?? ''} onChange={(e) => setSel(v => ({ ...v, [key]: e.target.value === '' ? null : Number(e.target.value) }))} style={{ ...input, height: '36px' }}>
                <option value="">만들지 않음</option>
                {sampleInfo.slides.map(sl => (
                    <option key={sl.index} value={sl.index} disabled={!sl.kind || !allow.includes(sl.kind) || sl.rowCount < 1}>
                        {`${sl.index + 1}번 · ${KIND_LABEL[sl.kind] || '인식 안 됨'}${sl.kind ? ` · 항목 ${sl.rowCount}` : ''} · ${sl.preview || '(텍스트 없음)'}`}
                    </option>
                ))}
            </select>
        </label>
    );

    return (
        <div className="animate-fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
            <div style={{ display: 'flex', alignItems: 'flex-start', gap: '10px' }}>
                <ListTree size={20} color="#f97316" style={{ marginTop: 2 }} />
                <p style={muted}>
                    최종 산출물(PPTX)의 장·절·항 구조와 시작 페이지(슬라이드 하단 표기 기준)를 읽어 <strong>표지 앞부분 목차</strong>(전체 목차 + 상세 목차)와 <strong>장별 세부목차</strong>(장마다 1장 이상)를 표(테이블) 서식으로 한 파일에 만듭니다.
                    목차 샘플(기본형 또는 직접 올린 파일)의 디자인(마스터·서식·도형·표)을 그대로 복제해 채우고, 칸이 모자라면 같은 모양으로 늘리며 넘치면 '(계속)'으로 이어 갑니다.
                </p>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(320px, 1fr))', gap: '16px' }}>
                {/* ① 최종 산출물 */}
                <section style={panel}>
                    <h3 style={h3}>① 최종 산출물 (PPTX, 여러 개)</h3>
                    <DropArea testId="toc-deliverables-drop" onFiles={addDeliverables}>
                        <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap', alignItems: 'center' }}>
                            <label style={{ ...fileBtn, cursor: 'pointer' }}>
                                <Upload size={15} /> 파일 선택
                                <input data-testid="toc-deliverables" type="file" accept=".pptx" multiple onChange={onPickDeliverables} style={{ display: 'none' }} />
                            </label>
                            {deliverables.length > 0 && (
                                <button type="button" onClick={() => { setDeliverables([]); setChapters([]); }}
                                    style={{ ...fileBtn, cursor: 'pointer', padding: '8px 12px', fontSize: '12.5px', fontFamily: 'inherit' }}>
                                    <Trash2 size={13} /> 비우기
                                </button>
                            )}
                        </div>
                        <span style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>또는 PPTX 파일을 여기로 끌어다 놓으세요 (여러 개 가능)</span>
                        {deliverables.length > 0 ? (
                            <ul style={{ margin: 0, paddingLeft: '18px', fontSize: '13px', color: 'var(--text-secondary)', display: 'flex', flexDirection: 'column', gap: '3px' }}>
                                {deliverables.map(f => <li key={f.name}>{f.name}</li>)}
                            </ul>
                        ) : <p style={muted}>장별 산출물 PPT(예: 1장. 환경분석.pptx … 4장. 이행계획.pptx)를 올려 주세요.</p>}
                    </DropArea>
                    <button type="button" disabled={!deliverables.length || isExtracting} onClick={extract}
                        onMouseEnter={hoverOn} onMouseLeave={hoverOff}
                        style={actionBtn(deliverables.length > 0 && !isExtracting)}>
                        {isExtracting ? <><Loader2 size={15} className="animate-spin" /> 장·절 구조 분석 중…</> : '목차 추출'}
                    </button>
                </section>

                {/* ② 목차 샘플 */}
                <section style={panel}>
                    <h3 style={h3}>② 목차 샘플 (디자인)</h3>
                    <div style={{ display: 'flex', gap: '6px', flexWrap: 'wrap' }}>
                        {[['default', '기본형 (공공의료 AI ISP)'], ['custom', '직접 올리기']].map(([k, label]) => (
                            <button key={k} type="button" data-testid={`toc-source-${k}`} onClick={() => chooseSource(k)}
                                style={{ fontFamily: 'inherit', padding: '7px 12px', borderRadius: '8px', fontSize: '12.5px', fontWeight: 700, cursor: 'pointer', border: `1px solid ${sampleSource === k ? '#f97316' : 'var(--panel-border)'}`, background: sampleSource === k ? 'rgba(249,115,22,0.16)' : 'transparent', color: sampleSource === k ? '#fb923c' : 'var(--text-secondary)' }}>
                                {label}
                            </button>
                        ))}
                    </div>
                    {sampleSource === 'custom' && (
                        <DropArea testId="toc-sample-drop" onFiles={(list) => applySampleFile(list.find(isPptx) || list[0])}>
                            <label style={{ ...fileBtn, cursor: 'pointer' }}>
                                <Upload size={15} /> {sampleFile ? '샘플 바꾸기' : '샘플 파일 선택'}
                                <input data-testid="toc-sample" type="file" accept=".pptx" onChange={onPickSample} style={{ display: 'none' }} />
                            </label>
                            <span style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>또는 목차 샘플 PPTX를 여기로 끌어다 놓으세요</span>
                        </DropArea>
                    )}
                    {isAnalyzingSample && <p style={muted}><Loader2 size={13} className="animate-spin" /> 샘플 디자인 분석 중…</p>}
                    {sampleFile && sampleInfo && (
                        <>
                            <span style={{ fontSize: '13px', color: 'var(--text-primary)', display: 'flex', gap: '8px', alignItems: 'center', flexWrap: 'wrap' }}>
                                {sampleFile.name} · {sampleInfo.slideCount}장
                                {sampleSource === 'default' && (
                                    <a href={tocDefaultUrl} download={DEFAULT_NAME} style={{ fontSize: '12px', color: '#fb923c' }}>기본형 내려받기(수정해 올릴 수 있음)</a>
                                )}
                            </span>
                            {sampleSource === 'custom' && (
                                <div data-testid="toc-custom-mode" style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                                    {[['keep', '디자인 유지 · 목차와 페이지만 갱신 (권장)', '올린 파일의 슬라이드·표·도형·서식은 그대로 두고, 목차 글자(장 제목·절/항 번호·제목)와 시작 쪽만 최종 산출물 기준으로 점검해 바꿉니다.'],
                                        ['generate', '이 디자인으로 새로 생성', '올린 슬라이드를 견본으로 목차 슬라이드를 새로 만듭니다(칸 수·배치가 내용에 맞게 바뀜).']].map(([k, label, desc]) => (
                                        <label key={k} style={{ display: 'flex', gap: '8px', alignItems: 'flex-start', cursor: 'pointer', padding: '8px 10px', borderRadius: '8px', border: `1px solid ${customMode === k ? '#f97316' : 'var(--panel-border)'}`, background: customMode === k ? 'rgba(249,115,22,0.08)' : 'transparent' }}>
                                            <input type="radio" name="toc-custom-mode" value={k} checked={customMode === k} onChange={() => { setCustomMode(k); setUpdateResult(null); }} style={{ marginTop: 3 }} />
                                            <span style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                                                <span style={{ fontSize: '13px', fontWeight: 700, color: 'var(--text-primary)' }}>{label}</span>
                                                <span style={{ fontSize: '12px', color: 'var(--text-secondary)', lineHeight: 1.5 }}>{desc}</span>
                                            </span>
                                        </label>
                                    ))}
                                </div>
                            )}
                            {!keepMode && (<>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                                <span style={{ fontSize: '12.5px', fontWeight: 700, color: 'var(--text-primary)' }}>표지 앞부분에 넣을 목차</span>
                                {slideSelect('overviewIndex', '전체 목차 (장 › 절)', 'toc-sel-overview', ['cards', 'list', 'table'])}
                                {slideSelect('detailIndex', '상세 목차 (장 › 절 › 항, 장 순서대로 이어서)', 'toc-sel-detail', ['table', 'list'])}
                                <span style={{ fontSize: '12.5px', fontWeight: 700, color: 'var(--text-primary)', marginTop: '4px' }}>장마다 넣을 목차 (장별 세부목차)</span>
                                {slideSelect('chapterIndex', '장별 목차 (장 표지 다음에 넣는 절 › 항 세부목차, 표형 테이블 권장)', 'toc-sel-chapter', ['table', 'list'])}
                            </div>
                            {chapterSlide && (
                                <div data-testid="toc-sample-detect" style={{ display: 'flex', flexWrap: 'wrap', gap: '6px' }}>
                                    <span style={chip(true)}>장별 목차: {chapterSlide.mode === 'token' ? '{{토큰}} 지정 방식' : `${KIND_LABEL[chapterSlide.kind]} 자동 인식`}</span>
                                    {depth && <span data-testid="toc-sample-depth" style={{ ...chip(true), background: 'rgba(249,115,22,0.16)', color: '#fb923c' }}>목차 깊이 {depth.label}</span>}
                                    <span style={chip(true)}>항목 칸 {chapterSlide.rowCount}개{chapterSlide.columns > 1 ? ` · ${chapterSlide.columns}단` : ''}</span>
                                    <span style={chip(chapterSlide.hasPage)}>페이지 칸 {chapterSlide.hasPage ? '✓' : '없음'}</span>
                                    <span style={chip(chapterSlide.hasChapterRoman)}>장 번호 칸 {chapterSlide.hasChapterRoman ? '✓' : '없음'}</span>
                                </div>
                            )}
                            </>)}
                        </>
                    )}
                    <p style={muted}>
                        직접 올릴 때는 한 파일에 전체 목차·상세 목차·장별 목차 슬라이드를 함께 넣으면 종류(카드형·표형·행 반복형)를 알아서 골라 둡니다.
                        인식이 어긋나면 목차 슬라이드의 글자를 <code>{'{{장번호}}'}</code> <code>{'{{장제목}}'}</code> <code>{'{{번호}}'}</code> <code>{'{{제목}}'}</code> <code>{'{{페이지}}'}</code>로 바꿔 올려 주세요.
                    </p>
                </section>

                {/* ③ 생성 옵션 */}
                <section style={panel}>
                    <h3 style={h3}>③ 생성</h3>
                    {keepMode ? (
                        <ul style={{ margin: 0, paddingLeft: '18px', fontSize: '12.5px', color: 'var(--text-secondary)', display: 'flex', flexDirection: 'column', gap: '3px' }}>
                            <li>올린 파일 그대로: 슬라이드·표·도형·서식 변경 없음</li>
                            <li>점검·갱신: 장 제목, 절·항 번호와 제목, 시작 쪽, 'n개 절', 머리글 장 표기</li>
                            <li>칸이 남으면 비우고, 모자라면 경고로 알림</li>
                        </ul>
                    ) : (
                    <ul style={{ margin: 0, paddingLeft: '18px', fontSize: '12.5px', color: 'var(--text-secondary)', display: 'flex', flexDirection: 'column', gap: '3px' }}>
                        <li>앞부분 목차: {[overviewSlide && `전체 목차(${KIND_LABEL[overviewSlide.kind]})`, detailSlide && `상세 목차(${KIND_LABEL[detailSlide.kind]})`].filter(Boolean).join(' + ') || '만들지 않음'}</li>
                        <li>장별 목차: {chapterSlide ? `장마다 ${KIND_LABEL[chapterSlide.kind]} 1장 이상` : '만들지 않음'}</li>
                    </ul>
                    )}
                    {!keepMode && overviewSlide?.kind === 'list' && (
                        <label style={{ display: 'flex', gap: '8px', alignItems: 'center', fontSize: '12.5px', color: 'var(--text-secondary)' }}>
                            전체 목차 장 제목 칸 문구
                            <input value={options.overviewTitle} onChange={e => setOptions(o => ({ ...o, overviewTitle: e.target.value }))} style={{ ...input, width: '110px' }} />
                        </label>
                    )}
                    <label style={{ display: 'flex', gap: '8px', alignItems: 'center', fontSize: '13.5px', color: 'var(--text-primary)' }}>
                        <input type="checkbox" checked={options.showPage} onChange={e => setOptions(o => ({ ...o, showPage: e.target.checked }))} style={{ width: 16, height: 16 }} />
                        시작 페이지 표기 (샘플에 페이지 칸이 있을 때)
                    </label>
                    <button type="button" disabled={!canBuild} onClick={build}
                        onMouseEnter={hoverOn} onMouseLeave={hoverOff}
                        style={{ ...actionBtn(canBuild), padding: '11px 16px', fontSize: '14px', fontWeight: 800 }}>
                        {isBuilding
                            ? <><Loader2 size={15} className="animate-spin" /> {keepMode ? '점검·갱신 중…' : '생성 중…'}</>
                            : <><FileDown size={16} /> {keepMode ? '올린 디자인 그대로 목차·페이지 갱신' : '샘플 디자인으로 목차 PPT 생성'}</>}
                    </button>
                    {excelBtn('toc-excel')}
                    <p style={muted}>
                        {chapters.length ? `포함: ${chapters.filter(c => c.include).length}개 장 · ${includedSections}개 절${usesSubsections ? ` · ${includedSubs}개 항` : ''}` : '① 목차 추출과 ② 샘플 선택을 마치면 생성할 수 있습니다.'}
                    </p>
                    {warnings.length > 0 && (
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                            {warnings.map((w, i) => (
                                <span key={i} style={{ fontSize: '12px', color: '#fbbf24', display: 'flex', gap: '6px' }}><AlertTriangle size={13} style={{ flex: 'none', marginTop: 2 }} />{w}</span>
                            ))}
                        </div>
                    )}
                </section>
            </div>

            {/* 디자인 유지 갱신 점검 결과 */}
            {updateResult && (
                <section style={{ ...panel, padding: 0, overflow: 'hidden' }} data-testid="toc-update-result">
                    <div style={{ padding: '14px 18px', borderBottom: '1px solid var(--panel-border)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '8px' }}>
                        <h3 style={h3}>점검 결과 · {updateResult.fileName}</h3>
                        <span style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>
                            갱신한 슬라이드 {updateResult.updatedSlides.join(', ')}번 · 바뀐 곳 {updateResult.changes.length}건{updateResult.warnings.length ? ` · 확인 필요 ${updateResult.warnings.length}건` : ''}
                        </span>
                    </div>
                    {updateResult.changes.length === 0 ? (
                        <p style={{ ...muted, padding: '14px 18px' }}>올린 목차가 최종 산출물과 모두 같습니다. 바꾼 곳이 없습니다.</p>
                    ) : (
                        <div style={{ maxHeight: '420px', overflow: 'auto' }}>
                            <table style={{ width: '100%', minWidth: '640px', borderCollapse: 'collapse', fontSize: '12.5px' }}>
                                <thead style={{ position: 'sticky', top: 0, background: 'var(--panel-bg, #111827)' }}>
                                    <tr style={{ textAlign: 'left', color: 'var(--text-secondary)' }}>
                                        <th style={{ padding: '7px 12px', width: '60px' }}>슬라이드</th>
                                        <th style={{ padding: '7px 12px', width: '120px' }}>구분</th>
                                        <th style={{ padding: '7px 12px', width: '110px' }}>위치</th>
                                        <th style={{ padding: '7px 12px', width: '70px' }}>항목</th>
                                        <th style={{ padding: '7px 12px' }}>이전</th>
                                        <th style={{ padding: '7px 12px' }}>변경</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {updateResult.changes.slice(0, 1000).map((c, i) => (
                                        <tr key={i} style={{ borderTop: '1px solid rgba(255,255,255,0.05)' }}>
                                            <td style={{ padding: '5px 12px', color: 'var(--text-secondary)' }}>{c.slide}</td>
                                            <td style={{ padding: '5px 12px', color: 'var(--text-secondary)' }}>{c.area}</td>
                                            <td style={{ padding: '5px 12px', color: 'var(--text-primary)' }}>{c.where}</td>
                                            <td style={{ padding: '5px 12px', color: 'var(--text-secondary)' }}>{c.field}</td>
                                            <td style={{ padding: '5px 12px', color: '#f87171', textDecoration: c.before ? 'line-through' : 'none' }}>{c.before || '(빈 칸)'}</td>
                                            <td style={{ padding: '5px 12px', color: '#34d399' }}>{c.after || '(비움)'}</td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                            {updateResult.changes.length > 1000 && <p style={{ ...muted, padding: '8px 12px' }}>처음 1,000건만 표시했습니다.</p>}
                        </div>
                    )}
                </section>
            )}

            {/* 목차 미리보기·편집 */}
            {chapters.length > 0 && (
                <section style={{ ...panel, padding: 0, overflow: 'hidden' }} data-testid="toc-preview">
                    <div style={{ padding: '14px 18px', borderBottom: '1px solid var(--panel-border)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '8px' }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                            <h3 style={h3}>목차 미리보기 · 편집</h3>
                            {excelBtn('toc-excel-preview', true)}
                        </div>
                        <span style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>제목을 고치거나 체크를 해제하면 생성 결과에 반영됩니다.{sampleInfo ? (usesSubsections ? ' 샘플 깊이에 맞춰 항(1.1 …)까지 표시합니다.' : ' 샘플 깊이에 맞춰 항은 제외합니다.') : ''}</span>
                    </div>
                    <div style={{ overflowX: 'auto' }}>
                        <table style={{ width: '100%', minWidth: '640px', borderCollapse: 'collapse', fontSize: '13px' }}>
                            <thead>
                                <tr style={{ background: 'rgba(255,255,255,0.04)', textAlign: 'left', color: 'var(--text-secondary)' }}>
                                    <th style={{ padding: '8px 12px', width: '48px' }}>포함</th>
                                    <th style={{ padding: '8px 12px', width: '70px' }}>번호</th>
                                    <th style={{ padding: '8px 12px' }}>제목</th>
                                    <th style={{ padding: '8px 12px', width: '120px' }}>시작 페이지</th>
                                    <th style={{ padding: '8px 12px', width: '220px' }}>출처 파일 · 페이지 범위</th>
                                </tr>
                            </thead>
                            <tbody>
                                {chapters.map((c, ci) => (
                                    <React.Fragment key={c.key}>
                                        <tr style={{ background: 'rgba(249,115,22,0.08)', borderTop: '1px solid var(--panel-border)' }}>
                                            <td style={{ padding: '8px 12px' }}><input type="checkbox" aria-label={`${c.title} 장 포함`} checked={c.include} onChange={e => updateChapter(ci, { include: e.target.checked })} /></td>
                                            <td style={{ padding: '8px 12px', fontWeight: 800, color: '#fb923c' }}>{c.roman || '-'}</td>
                                            <td style={{ padding: '6px 12px' }}><input aria-label="장 제목" value={c.title} onChange={e => updateChapter(ci, { title: e.target.value })} style={{ ...input, width: '100%', fontWeight: 700 }} /></td>
                                            <td style={{ padding: '8px 12px', color: 'var(--text-secondary)' }}>{c.startLabel || '-'}</td>
                                            <td style={{ padding: '8px 12px', color: 'var(--text-secondary)', fontSize: '12px' }}>{c.fileName}</td>
                                        </tr>
                                        {c.sections.map((s, si) => (
                                            <React.Fragment key={`${c.key}-${si}`}>
                                                <tr style={{ borderTop: '1px solid rgba(255,255,255,0.04)', opacity: c.include && s.include ? 1 : 0.45 }}>
                                                    <td style={{ padding: '6px 12px' }}><input type="checkbox" aria-label={`${s.title} 절 포함`} checked={s.include} onChange={e => updateSection(ci, si, { include: e.target.checked })} /></td>
                                                    <td style={{ padding: '6px 12px', color: 'var(--text-secondary)' }}>{c.roman ? `${c.roman}.${s.num}` : s.num}</td>
                                                    <td style={{ padding: '4px 12px' }}><input aria-label="절 제목" value={s.title} onChange={e => updateSection(ci, si, { title: e.target.value })} style={{ ...input, width: '100%' }} /></td>
                                                    <td style={{ padding: '6px 12px', color: s.startLabel ? 'var(--text-secondary)' : '#fbbf24' }}>{s.startLabel || '페이지 미확인'}</td>
                                                    <td style={{ padding: '6px 12px', color: 'var(--text-secondary)', fontSize: '12px' }}>{s.pageRange || ''}</td>
                                                </tr>
                                                {usesSubsections && (s.subsections || []).map((x, xi) => (
                                                    <tr key={`${c.key}-${si}-${xi}`} style={{ borderTop: '1px solid rgba(255,255,255,0.03)', opacity: c.include && s.include && x.include ? 1 : 0.45 }}>
                                                        <td style={{ padding: '4px 12px' }}><input type="checkbox" aria-label={`${x.title} 항 포함`} checked={x.include} onChange={e => updateSub(ci, si, xi, { include: e.target.checked })} /></td>
                                                        <td style={{ padding: '4px 12px 4px 24px', color: 'var(--text-secondary)', fontSize: '12px' }}>{`${s.num}.${x.num}`}</td>
                                                        <td style={{ padding: '3px 12px 3px 28px' }}><input aria-label="항 제목" value={x.title} onChange={e => updateSub(ci, si, xi, { title: e.target.value })} style={{ ...input, width: '100%', fontSize: '12.5px' }} /></td>
                                                        <td style={{ padding: '4px 12px', color: 'var(--text-secondary)', fontSize: '12px' }}>{x.startLabel || '-'}</td>
                                                        <td></td>
                                                    </tr>
                                                ))}
                                            </React.Fragment>
                                        ))}
                                        {c.sections.length === 0 && (
                                            <tr><td colSpan={5} style={{ padding: '8px 12px 8px 60px', color: '#fbbf24', fontSize: '12.5px' }}><AlertTriangle size={13} /> 이 파일에서 절 구조를 찾지 못했습니다.</td></tr>
                                        )}
                                    </React.Fragment>
                                ))}
                            </tbody>
                        </table>
                    </div>
                </section>
            )}
        </div>
    );
}
