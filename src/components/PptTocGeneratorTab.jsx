import React, { useState } from 'react';
import { saveAs } from 'file-saver';
import { ListTree, Upload, FileDown, Trash2, Loader2, AlertTriangle } from 'lucide-react';
import { buildDeliverableTocIndex, formatPageLabelRanges } from '../utils/ssotValidator';
import { extractTocFromDeliverables, analyzeTocSample, buildTocPptx } from '../utils/pptTocBuilder';

const panel = { background: 'rgba(0,0,0,0.2)', border: '1px solid var(--panel-border)', borderRadius: '12px', padding: '18px', display: 'flex', flexDirection: 'column', gap: '12px' };
const h3 = { margin: 0, fontSize: '15px', fontWeight: 700, color: 'var(--text-primary)' };
const muted = { fontSize: '12.5px', color: 'var(--text-secondary)', lineHeight: 1.6, margin: 0 };
const chip = (ok) => ({ fontSize: '12px', fontWeight: 600, padding: '3px 9px', borderRadius: '999px', background: ok ? 'rgba(16,185,129,0.14)' : 'rgba(148,163,184,0.14)', color: ok ? '#34d399' : 'var(--text-secondary)' });
const input = { background: 'rgba(0,0,0,0.3)', border: '1px solid var(--panel-border)', borderRadius: '6px', color: 'var(--text-primary)', padding: '6px 8px', fontSize: '13px', fontFamily: 'inherit' };
const fileBtn = { display: 'inline-flex', alignItems: 'center', gap: '8px', padding: '9px 16px', borderRadius: '8px', border: '1px solid var(--panel-border)', cursor: 'pointer', fontSize: '13.5px', fontWeight: 600, color: 'var(--text-primary)', background: 'rgba(255,255,255,0.04)' };

// 샘플 목차 깊이 설명 (예: 2단계 · 절 › 항)
function depthInfo(slide) {
    if (!slide || !slide.levels) return null;
    const names = slide.topIsChapter ? ['장', '절', '항'] : ['절', '항', '세부 항목'];
    const chain = names.slice(0, slide.levels).join(' › ');
    const overviewChain = ['장', '절', '항'].slice(0, slide.levels).join(' › ');
    return { label: `${slide.levels}단계 · ${chain}`, overview: overviewChain, perChapter: chain, usesSubsections: slide.topIsChapter ? slide.levels >= 3 : slide.levels >= 2 };
}

export default function PptTocGeneratorTab({ setErrorMsg, setSuccessMsg }) {
    const [deliverables, setDeliverables] = useState([]);
    const [chapters, setChapters] = useState([]);
    const [isExtracting, setIsExtracting] = useState(false);
    const [sampleFile, setSampleFile] = useState(null);
    const [sampleInfo, setSampleInfo] = useState(null);
    const [slideIndex, setSlideIndex] = useState(null);
    const [isAnalyzingSample, setIsAnalyzingSample] = useState(false);
    const [options, setOptions] = useState({ overview: true, perChapter: true, showPage: true, overviewTitle: '목차' });
    const [isBuilding, setIsBuilding] = useState(false);
    const [warnings, setWarnings] = useState([]);

    const resetMsg = () => { setErrorMsg?.(null); setSuccessMsg?.(null); };

    const onPickDeliverables = (e) => {
        const files = Array.from(e.target.files || []).filter(f => /\.pptx$/i.test(f.name) && !f.name.startsWith('~$'));
        e.target.value = '';
        if (!files.length) return;
        setDeliverables(prev => {
            const names = new Set(prev.map(f => f.name));
            return [...prev, ...files.filter(f => !names.has(f.name))].sort((a, b) => a.name.localeCompare(b.name, 'ko'));
        });
        setChapters([]);
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

    const onPickSample = async (e) => {
        const file = e.target.files?.[0];
        e.target.value = '';
        if (!file) return;
        resetMsg();
        setSampleFile(file);
        setSampleInfo(null);
        setIsAnalyzingSample(true);
        try {
            const info = await analyzeTocSample(file);
            setSampleInfo(info);
            setSlideIndex(info.bestIndex);
            if (info.bestIndex == null) {
                setErrorMsg?.('샘플에서 목차 항목 칸(같은 모양으로 반복되는 제목 칸)을 찾지 못했습니다. 목차 슬라이드에 {{제목}} 표시를 넣어 다시 올려 주세요.');
            }
        } catch (err) {
            console.error(err);
            setErrorMsg?.(`목차 샘플 분석 중 오류가 발생했습니다: ${err.message}`);
        } finally {
            setIsAnalyzingSample(false);
        }
    };

    const updateChapter = (ci, patch) => setChapters(prev => prev.map((c, i) => (i === ci ? { ...c, ...patch } : c)));
    const updateSection = (ci, si, patch) => setChapters(prev => prev.map((c, i) => (i === ci
        ? { ...c, sections: c.sections.map((s, j) => (j === si ? { ...s, ...patch } : s)) }
        : c)));

    const build = async () => {
        resetMsg();
        setWarnings([]);
        setIsBuilding(true);
        try {
            const { blob, warnings: w, slideCount } = await buildTocPptx({ sampleFile, slideIndex, chapters, options });
            const d = new Date();
            const stamp = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
            saveAs(blob, `목차_${stamp}.pptx`);
            setWarnings(w);
            setSuccessMsg?.(`목차 슬라이드 ${slideCount}장을 샘플 디자인으로 생성했습니다.`);
        } catch (err) {
            console.error(err);
            setErrorMsg?.(`목차 PPT 생성 중 오류가 발생했습니다: ${err.message}`);
        } finally {
            setIsBuilding(false);
        }
    };

    const selectedSlide = sampleInfo?.slides.find(s => s.index === slideIndex) || null;
    const depth = depthInfo(selectedSlide);
    const updateSub = (ci, si, xi, patch) => setChapters(prev => prev.map((c, i) => (i === ci
        ? { ...c, sections: c.sections.map((s, j) => (j === si ? { ...s, subsections: s.subsections.map((x, k) => (k === xi ? { ...x, ...patch } : x)) } : s)) }
        : c)));
    const includedSections = chapters.filter(c => c.include).reduce((n, c) => n + c.sections.filter(s => s.include).length, 0);
    const includedSubs = chapters.filter(c => c.include).reduce((n, c) => n + c.sections.filter(s => s.include).reduce((m, s) => m + (s.subsections || []).filter(x => x.include).length, 0), 0);
    const canBuild = chapters.length > 0 && sampleFile && slideIndex != null && (options.overview || options.perChapter) && !isBuilding;

    return (
        <div className="animate-fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
            <div style={{ display: 'flex', alignItems: 'flex-start', gap: '10px' }}>
                <ListTree size={20} color="#f97316" style={{ marginTop: 2 }} />
                <p style={muted}>
                    최종 산출물(PPTX)의 장·절 구조와 시작 페이지(슬라이드 하단 표기 기준)를 읽어 목차를 만들고, 올린 <strong>목차 샘플 파일의 디자인</strong>(마스터·서식·도형)을 그대로 복제해 채웁니다.
                    샘플의 목차 칸이 모자라면 같은 모양으로 늘리고, 남으면 지웁니다.
                </p>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(320px, 1fr))', gap: '16px' }}>
                {/* ① 최종 산출물 */}
                <section style={panel}>
                    <h3 style={h3}>① 최종 산출물 (PPTX, 여러 개)</h3>
                    <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap', alignItems: 'center' }}>
                        <label style={fileBtn}>
                            <Upload size={15} /> 파일 선택
                            <input data-testid="toc-deliverables" type="file" accept=".pptx" multiple onChange={onPickDeliverables} style={{ display: 'none' }} />
                        </label>
                        {deliverables.length > 0 && (
                            <button type="button" className="btn-outline" style={{ padding: '8px 12px', fontSize: '12.5px' }} onClick={() => { setDeliverables([]); setChapters([]); }}>
                                <Trash2 size={13} /> 비우기
                            </button>
                        )}
                    </div>
                    {deliverables.length > 0 ? (
                        <ul style={{ margin: 0, paddingLeft: '18px', fontSize: '13px', color: 'var(--text-secondary)', display: 'flex', flexDirection: 'column', gap: '3px' }}>
                            {deliverables.map(f => <li key={f.name}>{f.name}</li>)}
                        </ul>
                    ) : <p style={muted}>장별 산출물 PPT(예: 1장. 환경분석.pptx … 4장. 이행계획.pptx)를 올려 주세요.</p>}
                    <button type="button" className="btn-primary" disabled={!deliverables.length || isExtracting} onClick={extract}
                        style={{ padding: '10px 16px', fontSize: '13.5px', fontWeight: 700, display: 'inline-flex', alignItems: 'center', gap: '8px', justifyContent: 'center' }}>
                        {isExtracting ? <><Loader2 size={15} className="animate-spin" /> 장·절 구조 분석 중…</> : '목차 추출'}
                    </button>
                </section>

                {/* ② 목차 샘플 */}
                <section style={panel}>
                    <h3 style={h3}>② 목차 샘플 (PPTX 1개)</h3>
                    <label style={fileBtn}>
                        <Upload size={15} /> {sampleFile ? '샘플 바꾸기' : '샘플 파일 선택'}
                        <input data-testid="toc-sample" type="file" accept=".pptx" onChange={onPickSample} style={{ display: 'none' }} />
                    </label>
                    {isAnalyzingSample && <p style={muted}><Loader2 size={13} className="animate-spin" /> 샘플 디자인 분석 중…</p>}
                    {sampleFile && sampleInfo && (
                        <>
                            <span style={{ fontSize: '13px', color: 'var(--text-primary)' }}>{sampleFile.name} · {sampleInfo.slideCount}장</span>
                            {sampleInfo.slideCount > 1 && (
                                <label style={{ display: 'flex', flexDirection: 'column', gap: '4px', fontSize: '12.5px', color: 'var(--text-secondary)' }}>
                                    목차로 쓸 슬라이드
                                    <select value={slideIndex ?? ''} onChange={(e) => setSlideIndex(e.target.value === '' ? null : Number(e.target.value))} style={{ ...input, height: '38px' }}>
                                        {sampleInfo.slides.map(s => (
                                            <option key={s.index} value={s.index} disabled={s.rowCount < 1}>
                                                {`${s.index + 1}번 슬라이드${s.index === sampleInfo.bestIndex ? ' (추천)' : ''} · 항목 칸 ${s.rowCount}개 · ${s.preview || '(텍스트 없음)'}`}
                                            </option>
                                        ))}
                                    </select>
                                </label>
                            )}
                            {selectedSlide && (
                                <div data-testid="toc-sample-detect" style={{ display: 'flex', flexWrap: 'wrap', gap: '6px' }}>
                                    <span style={chip(true)}>{selectedSlide.mode === 'token' ? '{{토큰}} 지정 방식' : '자동 인식'}</span>
                                    {depth && <span data-testid="toc-sample-depth" style={{ ...chip(true), background: 'rgba(249,115,22,0.16)', color: '#fb923c' }}>목차 깊이 {depth.label}</span>}
                                    <span style={chip(true)}>항목 칸 {selectedSlide.rowCount}개{selectedSlide.columns > 1 ? ` · ${selectedSlide.columns}단` : ''}</span>
                                    <span style={chip(selectedSlide.hasNum)}>번호 칸 {selectedSlide.hasNum ? '✓' : '없음'}</span>
                                    <span style={chip(selectedSlide.hasPage)}>페이지 칸 {selectedSlide.hasPage ? '✓' : '없음'}</span>
                                    <span style={chip(selectedSlide.hasChapterRoman)}>장 번호 칸 {selectedSlide.hasChapterRoman ? '✓' : '없음'}</span>
                                    <span style={chip(selectedSlide.hasChapterTitle)}>장 제목 칸 {selectedSlide.hasChapterTitle ? '✓' : '없음'}</span>
                                </div>
                            )}
                        </>
                    )}
                    {depth && (
                        <p style={{ ...muted, color: 'var(--text-primary)' }}>
                            샘플과 같은 깊이로 만듭니다 → 전체 목차: <strong>{depth.overview}</strong> · 장별 목차: <strong>{depth.perChapter}</strong>
                        </p>
                    )}
                    <p style={muted}>
                        인식이 어긋나면 샘플 목차 슬라이드의 글자를 <code>{'{{장번호}}'}</code> <code>{'{{장제목}}'}</code> <code>{'{{번호}}'}</code> <code>{'{{제목}}'}</code> <code>{'{{페이지}}'}</code>로 바꿔 올려 주세요. 목차 행은 한 줄만 만들어도 됩니다.
                    </p>
                </section>

                {/* ③ 생성 옵션 */}
                <section style={panel}>
                    <h3 style={h3}>③ 생성</h3>
                    <label style={{ display: 'flex', gap: '8px', alignItems: 'center', fontSize: '13.5px', color: 'var(--text-primary)' }}>
                        <input type="checkbox" checked={options.overview} onChange={e => setOptions(o => ({ ...o, overview: e.target.checked }))} style={{ width: 16, height: 16 }} />
                        전체 목차 1장 (장 목록)
                    </label>
                    {options.overview && (
                        <label style={{ display: 'flex', gap: '8px', alignItems: 'center', fontSize: '12.5px', color: 'var(--text-secondary)', paddingLeft: '24px' }}>
                            장 제목 칸에 넣을 문구
                            <input value={options.overviewTitle} onChange={e => setOptions(o => ({ ...o, overviewTitle: e.target.value }))} style={{ ...input, width: '110px' }} />
                        </label>
                    )}
                    <label style={{ display: 'flex', gap: '8px', alignItems: 'center', fontSize: '13.5px', color: 'var(--text-primary)' }}>
                        <input type="checkbox" checked={options.perChapter} onChange={e => setOptions(o => ({ ...o, perChapter: e.target.checked }))} style={{ width: 16, height: 16 }} />
                        장별 목차 (장마다 1장, 절 목록)
                    </label>
                    <label style={{ display: 'flex', gap: '8px', alignItems: 'center', fontSize: '13.5px', color: 'var(--text-primary)' }}>
                        <input type="checkbox" checked={options.showPage} onChange={e => setOptions(o => ({ ...o, showPage: e.target.checked }))} style={{ width: 16, height: 16 }} />
                        시작 페이지 표기 (샘플에 페이지 칸이 있을 때)
                    </label>
                    <button type="button" className="btn-primary" disabled={!canBuild} onClick={build}
                        style={{ padding: '11px 16px', fontSize: '14px', fontWeight: 800, display: 'inline-flex', alignItems: 'center', gap: '8px', justifyContent: 'center', background: canBuild ? 'linear-gradient(135deg, #f97316, #ea580c)' : undefined }}>
                        {isBuilding ? <><Loader2 size={15} className="animate-spin" /> 생성 중…</> : <><FileDown size={16} /> 샘플 디자인으로 목차 PPT 생성</>}
                    </button>
                    <p style={muted}>
                        {chapters.length ? `포함: ${chapters.filter(c => c.include).length}개 장 · ${includedSections}개 절${depth?.usesSubsections ? ` · ${includedSubs}개 항` : ''}` : '① 목차 추출과 ② 샘플 선택을 마치면 생성할 수 있습니다.'}
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

            {/* 목차 미리보기·편집 */}
            {chapters.length > 0 && (
                <section style={{ ...panel, padding: 0, overflow: 'hidden' }} data-testid="toc-preview">
                    <div style={{ padding: '14px 18px', borderBottom: '1px solid var(--panel-border)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '8px' }}>
                        <h3 style={h3}>목차 미리보기 · 편집</h3>
                        <span style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>제목을 고치거나 체크를 해제하면 생성 결과에 반영됩니다.{depth ? (depth.usesSubsections ? ' 샘플 깊이에 맞춰 항(1.1 …)까지 표시합니다.' : ' 샘플 깊이에 맞춰 항은 제외합니다.') : ''}</span>
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
                                                {depth?.usesSubsections && (s.subsections || []).map((x, xi) => (
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
