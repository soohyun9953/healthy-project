import React, { useState } from 'react';
import { ShieldAlert, CheckCircle2, XCircle, FileWarning, AlertTriangle, ClipboardList, ArrowRightLeft, Download, PenTool, RotateCcw, Copy, Check, FileText, Eye } from 'lucide-react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
/** 매핑 결과를 엑셀 파일로 내보내기 */
async function exportToExcel(data, isTypoMode = false) {
    const wb = new ExcelJS.Workbook();
    
    // helper 
    const addSheet = (sheetName, headers, rowsData) => {
        const ws = wb.addWorksheet(sheetName, {
            views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }] // 1행(헤더) 고정
        });
        
        // set headers
        ws.columns = headers;
        
        // add rows
        ws.addRows(rowsData);
        
        // auto filter (A1: lastColumn1)
        ws.autoFilter = {
            from: { row: 1, column: 1 },
            to: { row: 1, column: headers.length }
        };
        
        // style header
        ws.getRow(1).font = { bold: true };
        ws.getRow(1).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFEFEFEF' }
        };
        // apply border to header
        ws.getRow(1).eachCell(cell => {
            cell.border = {
                top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}
            };
        });
    };

    // ── 시트 준비: 오탈자/교정교열 결과 시트 생성 함수 ──
    const buildTypoSheet = () => {
        const typoHeaders = [
            { header: '순번', key: '순번', width: 6 },
            { header: '페이지/위치', key: '위치', width: 25 },
            { header: '원문 문장 전체', key: '원문', width: 40 },
            { header: '수정 제안 문장', key: '수정', width: 40 },
            { header: '오류 유형/사유', key: '사유', width: 60 }
        ];

        const validTypos = (data.typos || []).map(item => {
            const orig = String(item.originalText || item.original || item.errorText || item.before || item.wrong || item.source || '').trim();
            const corr = String(item.correction || item.correct || item.after || item.suggestion || item.target || item.right || '').trim();
            const page = String(item.page || item.location || item.type || item.section || item.path || '1페이지');
            const reason = String(item.errorType || item.reason || item.context || item.category || '[표현 품질] 교정');
            return { orig, corr, page, reason };
        }).filter(t => t.orig && t.corr && t.orig !== t.corr && t.corr !== '문맥 검토 및 구체적 명세 보완 권고');

        const typoRows = validTypos.map((item, idx) => ({
            '순번': idx + 1,
            '위치': item.page,
            '원문': item.orig,
            '수정': item.corr,
            '사유': item.reason,
        }));

        if (typoRows.length === 0) {
            typoRows.push({
                '순번': 1,
                '위치': '전체 산출물',
                '원문': '특이 결함 없음',
                '수정': '원문 유지',
                '사유': '[품질 완료] 지적할 오탈자, 띄어쓰기 및 표현 결함이 발견되지 않았습니다. (정상 문서)'
            });
        }

        addSheet('교정교열_결과', typoHeaders, typoRows);
    };

    if (isTypoMode) {
        // 교정교열 전용 모드에서는 '교정교열_결과' 시트를 맨 첫번째(시트 1번)로 배치!
        buildTypoSheet();
        const summaryHeaders = [
            { header: '항목', key: '항목', width: 20 },
            { header: '내용', key: '내용', width: 80 }
        ];
        const summaryRows = [
            { '항목': '검수 모드', '내용': 'ISMP 산출물 전문 교정교열 모드' },
            { '항목': '총점', '내용': `${data.score || 85}점 / 100점` },
            { '항목': '종합 분석 의견', '내용': data.summary || '분석 완료' }
        ];
        addSheet('종합_요약', summaryHeaders, summaryRows);
    } else {
        // 일반 문서 검증 모드
        const summaryHeaders = [
            { header: '항목', key: '항목', width: 20 },
            { header: '내용', key: '내용', width: 80 }
        ];
        const summaryRows = [
            { '항목': '총점', '내용': `${data.score || 85}점 / 100점` },
            { '항목': '점검 범위', '내용': data.inspectionScope || '전체 문장' },
            { '항목': '종합 분석 의견', '내용': data.summary || '분석 완료' }
        ];
        addSheet('종합_요약', summaryHeaders, summaryRows);

        // ── 시트1: 요구사항 매핑 현황 (RTM) ──
        const rtmHeaders = [
            { header: '번호', key: '번호', width: 6 },
            { header: '분류', key: '분류', width: 8 },
            { header: '기준 문서 요건', key: '기준문서요건', width: 50 },
            { header: '카테고리', key: '카테고리', width: 15 },
            { header: '수준', key: '수준', width: 10 },
            { header: '상태', key: '상태', width: 12 },
            { header: '산출물 증빙 위치', key: '산출물증빙위치', width: 40 }
        ];
        const rtmRows = (data.rtm || []).map((item, idx) => ({
            '번호': idx + 1,
            '분류': String(item.type || '필수'),
            '기준문서요건': String(item.requirement || ''),
            '카테고리': String(item.category || ''),
            '수준': String(item.levelLabel || ''),
            '상태': String(item.status || ''),
            '산출물증빙위치': String(item.location || ''),
        }));
        addSheet('매핑현황(RTM)', rtmHeaders, rtmRows);

        // ── 시트2: 매핑 상세 ──
        const detailHeaders = [
            { header: '번호', key: '번호', width: 6 },
            { header: 'ID', key: 'ID', width: 12 },
            { header: '분류', key: '분류', width: 8 },
            { header: '카테고리', key: '카테고리', width: 15 },
            { header: '수준', key: '수준', width: 10 },
            { header: '계층 경로', key: '계층경로', width: 40 },
            { header: '요구사항', key: '요구사항', width: 50 },
            { header: '산출물 대응 섹션', key: '산출물대응섹션', width: 30 },
            { header: '산출물 기술 내용', key: '산출물기술내용', width: 40 },
            { header: '상태', key: '상태', width: 12 },
            { header: '차이점', key: '차이점', width: 50 }
        ];
        const detailRows = (data.requirementMapping || []).map((item, idx) => ({
            '번호': idx + 1,
            'ID': String(item.id || ''),
            '분류': String(item.type || '필수'),
            '카테고리': String(item.category || ''),
            '수준': String(item.levelLabel || ''),
            '계층경로': String(item.path || ''),
            '요구사항': String(item.requirement || ''),
            '산출물대응섹션': String(item.artifactSection || ''),
            '산출물기술내용': String(item.artifactContent || '').replace(/^"|"$/g, ''),
            '상태': String(item.status || ''),
            '차이점': String(item.gap || ''),
        }));
        addSheet('매핑상세', detailHeaders, detailRows);

        // ── 시트3: 누락 사항 ──
        if (data.omissions && data.omissions.length > 0) {
            const omiHeaders = [
                { header: '번호', key: '번호', width: 6 },
                { header: '항목', key: '항목', width: 30 },
                { header: '근거', key: '근거', width: 50 },
                { header: '사유', key: '사유', width: 50 },
                { header: '권고사항', key: '권고사항', width: 50 }
            ];
            const omiRows = data.omissions.map((item, idx) => ({
                '번호': idx + 1,
                '항목': String(item.title || ''),
                '근거': String(item.evidence || ''),
                '사유': String(item.reason || ''),
                '권고사항': String(item.recommendation || ''),
            }));
            addSheet('누락사항', omiHeaders, omiRows);
        }

        buildTypoSheet();
    }

    const now = new Date();
    const dateStr = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}`;
    const rawFileName = data.artifactFileName ? String(data.artifactFileName).replace(/[\\/:*?"<>|]/g, '_') : '';
    const filePrefix = rawFileName ? `[${rawFileName}]_` : '';
    const fileName = isTypoMode ? `${filePrefix}교정교열_결과_${dateStr}.xlsx` : `${filePrefix}기준문서_검증결과_${dateStr}.xlsx`;

    try {
        const buffer = await wb.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        
        if (window.showSaveFilePicker) {
            try {
                const handle = await window.showSaveFilePicker({
                    suggestedName: fileName,
                    types: [{
                        description: 'Excel 파일',
                        accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] },
                    }],
                });
                const writable = await handle.createWritable();
                await writable.write(buffer);
                await writable.close();
                return;
            } catch (err) {
                if (err.name === 'AbortError') return;
                console.warn('showSaveFilePicker fallback:', err);
            }
        }
        
        saveAs(blob, fileName);
    } catch (e) {
        console.error('Excel export failed:', e);
    }
}

function renderInlineDiff(originalText, correctionText) {
    const orig = String(originalText || '').trim();
    const corr = String(correctionText || '').trim();
    if (!orig || !corr || orig === corr) {
        return <span style={{ color: 'var(--text-primary)' }}>{corr || orig}</span>;
    }

    const origWords = orig.split(/\s+/);
    const corrWords = corr.split(/\s+/);

    return (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: '4px', alignItems: 'center', lineHeight: '1.5' }}>
                <span style={{ fontSize: '11px', color: '#f87171', fontWeight: 800, background: 'rgba(239,68,68,0.1)', padding: '1px 5px', borderRadius: '4px' }}>원문</span>
                {origWords.map((w, i) => {
                    const isDiff = !corrWords.includes(w);
                    return isDiff ? (
                        <del key={i} style={{ color: '#f87171', background: 'rgba(239, 68, 68, 0.2)', padding: '1px 5px', borderRadius: '4px', textDecoration: 'line-through', fontWeight: 700 }}>
                            {w}
                        </del>
                    ) : (
                        <span key={i} style={{ color: 'var(--text-secondary)' }}>{w}</span>
                    );
                })}
            </div>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: '4px', alignItems: 'center', lineHeight: '1.5' }}>
                <span style={{ fontSize: '11px', color: '#34d399', fontWeight: 800, background: 'rgba(16,185,129,0.1)', padding: '1px 5px', borderRadius: '4px' }}>교정</span>
                {corrWords.map((w, i) => {
                    const isDiff = !origWords.includes(w);
                    return isDiff ? (
                        <ins key={i} style={{ color: '#34d399', fontWeight: 800, background: 'rgba(16, 185, 129, 0.2)', padding: '1px 5px', borderRadius: '4px', textDecoration: 'none' }}>
                            {w}
                        </ins>
                    ) : (
                        <span key={i} style={{ color: 'var(--text-primary)' }}>{w}</span>
                    );
                })}
            </div>
        </div>
    );
}

// orig/corr 문자열의 공통 접두사·접미사를 제거해 실제로 달라진 핵심 구간만 좁혀낸다.
// (예: "새호 구성된"/"새로 구성된" -> 실제 차이는 "호"/"로" 한 글자뿐)
function extractCoreDiff(orig, corr) {
    const minLen = Math.min(orig.length, corr.length);
    let start = 0;
    while (start < minLen && orig[start] === corr[start]) start++;

    let endOrig = orig.length;
    let endCorr = corr.length;
    while (endOrig > start && endCorr > start && orig[endOrig - 1] === corr[endCorr - 1]) {
        endOrig--;
        endCorr--;
    }
    if (start >= endOrig && start >= endCorr) return null; // 완전히 동일

    return { start, endOrig, origCore: orig.slice(start, endOrig), corrCore: corr.slice(start, endCorr) };
}

// 원문 전체에서 각 교정 항목이 실제로 등장하는 위치를 찾아 취소선(원문)/삽입(교정) 구간으로 분해한다.
// 원문에 정확히 일치하지 않는 항목(문맥 요약 등)은 억지로 추정하지 않고 건너뛴다 - 잘못된 위치 강조가
// 오히려 신뢰를 해칠 수 있기 때문이다. 건너뛴 항목도 기존 표(위 섹션)에는 그대로 남아 확인할 수 있다.
//
// 사전 기반 검출은 같은 줄에 여러 오탈자가 있어도 각 항목의 originalText가 "줄 전체"로 동일하게 잡히므로
// (예: 새호->새로, 컨텐츠->콘텐츠가 한 줄에 있으면 두 항목 모두 원문이 그 줄 전체), 같은 원문을 공유하는
// 항목들을 그룹으로 묶어 그 줄 안에서 서로 다른 치환 구간을 함께 찾아낸다.
function buildFullDocumentTrackChanges(originalFullText, typosList) {
    if (!originalFullText) return null;

    const candidates = (typosList || [])
        .map(t => ({
            orig: String(t.originalText || t.original || t.errorText || t.before || t.wrong || '').trim(),
            corr: String(t.correction || t.correct || t.after || t.suggestion || '').trim()
        }))
        .filter(t => t.orig && t.corr && t.orig !== t.corr);

    const groups = new Map(); // orig(원문 문장/구절) -> [corr, ...]
    candidates.forEach(({ orig, corr }) => {
        if (!groups.has(orig)) groups.set(orig, []);
        groups.get(orig).push(corr);
    });

    // 긴 문장부터 처리해 짧은 문장이 다른 문장 내부에 우연히 걸리는 것을 방지
    const origList = [...groups.keys()].sort((a, b) => b.length - a.length);

    const usedRanges = [];
    const matches = [];

    origList.forEach((orig) => {
        const spans = groups.get(orig)
            .map((corr) => extractCoreDiff(orig, corr))
            .filter(Boolean);
        if (spans.length === 0) return;

        let searchFrom = 0;
        let lineIdx = -1;
        while (searchFrom <= originalFullText.length) {
            const idx = originalFullText.indexOf(orig, searchFrom);
            if (idx === -1) break;
            const end = idx + orig.length;
            const overlaps = usedRanges.some(([s, e]) => idx < e && end > s);
            if (!overlaps) { lineIdx = idx; break; }
            searchFrom = idx + 1;
        }
        if (lineIdx === -1) return; // 문서 내에서 해당 문장을 찾지 못하면 건너뜀

        spans.forEach((sp) => {
            matches.push({ start: lineIdx + sp.start, end: lineIdx + sp.endOrig, orig: sp.origCore, corr: sp.corrCore });
        });
        usedRanges.push([lineIdx, lineIdx + orig.length]);
    });

    matches.sort((a, b) => a.start - b.start);

    const segments = [];
    let cursor = 0;
    matches.forEach((m) => {
        if (m.start < cursor) return; // 안전장치: 겹치는 구간은 건너뜀
        if (m.start > cursor) segments.push({ type: 'text', content: originalFullText.slice(cursor, m.start) });
        segments.push({ type: 'change', orig: m.orig, corr: m.corr });
        cursor = m.end;
    });
    if (cursor < originalFullText.length) segments.push({ type: 'text', content: originalFullText.slice(cursor) });

    return { segments, matchedCount: matches.length, totalCandidates: candidates.length };
}

export default function ResultDashboard({ data, isTypoMode = false, onRetry }) {
    const [copiedToast, setCopiedToast] = useState(false);
    const [showCompareModal, setShowCompareModal] = useState(false);

    if (!data) return null;

    const displayScore = isNaN(data.score) || data.score === undefined || data.score === null ? 0 : Math.round(Number(data.score));
    const typosList = data.typos || [];
    const styleIssues = data.styleIssues || [];
    const trackChanges = isTypoMode && data.originalFullText ? buildFullDocumentTrackChanges(data.originalFullText, typosList) : null;

    // 교정 완료본 클립보드 복사 핸들러
    const handleCopyFullCorrectedText = () => {
        const textToCopy = data.correctedFullText || (typosList.length > 0 
            ? typosList.map((t, idx) => `${idx + 1}. [${t.page || '위치'}] ${t.correction || t.originalText}`).join('\n')
            : '교정 사항 없음');
        
        navigator.clipboard.writeText(textToCopy).then(() => {
            setCopiedToast(true);
            setTimeout(() => setCopiedToast(false), 2500);
        }).catch(() => {
            alert('클립보드 복사에 실패했습니다.');
        });
    };

    // 교정 완료본 텍스트 파일 다운로드 핸들러
    const handleDownloadCorrectedFile = () => {
        const textToDownload = data.correctedFullText || '';
        const blob = new Blob([textToDownload], { type: 'text/markdown;charset=utf-8' });
        const fileName = `${(data.artifactFileName || '문서').replace(/\.[^.]+$/, '')}_교정완료본.md`;
        saveAs(blob, fileName);
    };

    return (
        <div className="glass-panel animate-fade-in" style={{ flex: 1, display: 'flex', flexDirection: 'column', padding: '24px', gap: '24px', overflowY: 'auto' }}>
            {/* 점검범위 표시 */}
            {data.inspectionScope && (
                <section className="animate-slide-up" style={{ padding: '14px 20px', background: 'rgba(59, 130, 246, 0.08)', borderRadius: '10px', border: '1px solid rgba(59, 130, 246, 0.2)', display: 'flex', gap: '10px', alignItems: 'flex-start' }}>
                    <ClipboardList size={18} color="var(--accent-color)" style={{ flexShrink: 0, marginTop: '2px' }} />
                    <div>
                        <span style={{ fontSize: '13px', fontWeight: 600, color: 'var(--accent-color)' }}>적용된 점검범위</span>
                        <p style={{ margin: '4px 0 0', fontSize: '14px', lineHeight: '1.5', color: 'var(--text-primary)' }}>{data.inspectionScope}</p>
                    </div>
                </section>
            )}

            {/* 🔥 교정교열 모드일 경우: [오탈자 및 표현 결함 결과 표]를 화면 최상단 1순위로 즉시 노출! */}
            {isTypoMode && (
                <section className="glass-panel animate-slide-up stagger-1" style={{ padding: '24px', borderLeft: '5px solid var(--warning-color)' }}>
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '16px', flexWrap: 'wrap', gap: '12px' }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                            <PenTool size={22} color="var(--warning-color)" />
                            <h3 style={{ margin: 0, fontSize: '20px', fontWeight: 700, color: 'var(--text-primary)' }}>
                                ISMP 산출물 전문 교정교열 검수 결과
                            </h3>
                            <span style={{
                                padding: '4px 10px', borderRadius: '20px', fontSize: '13px', fontWeight: 700,
                                background: typosList.length > 0 ? 'rgba(239, 68, 68, 0.15)' : 'rgba(34, 197, 94, 0.15)',
                                color: typosList.length > 0 ? 'var(--danger-color)' : 'var(--success-color)',
                                border: `1px solid ${typosList.length > 0 ? 'rgba(239, 68, 68, 0.3)' : 'rgba(34, 197, 94, 0.3)'}`
                            }}>
                                {typosList.length > 0 ? `발견된 결함: ${typosList.length}건 (100% 전수 스캔)` : '특이 결함 없음 (완벽)'}
                            </span>
                        </div>

                        <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap', alignItems: 'center' }}>
                            {data.summary && (data.summary.includes('Gemini 검증 실패') || data.summary.includes('소진되었습니다')) && onRetry && (
                                <button
                                    onClick={onRetry}
                                    className="interactive pulse-text"
                                    style={{
                                        display: 'flex', alignItems: 'center', gap: '8px',
                                        padding: '10px 16px', fontSize: '13.5px', fontWeight: 700,
                                        background: 'rgba(59, 130, 246, 0.15)', color: 'var(--accent-blue)',
                                        border: '2px solid var(--accent-blue)', borderRadius: '10px', cursor: 'pointer',
                                    }}
                                >
                                    <RotateCcw size={16} />
                                    <span>재시도</span>
                                </button>
                            )}

                            {/* 원클릭 교정 완성본 복사 버튼 */}
                            <button
                                onClick={handleCopyFullCorrectedText}
                                className="interactive"
                                style={{
                                    display: 'flex', alignItems: 'center', gap: '6px',
                                    padding: '10px 16px', fontSize: '13.5px', fontWeight: 700,
                                    background: copiedToast ? 'rgba(34, 197, 94, 0.2)' : 'rgba(168, 85, 247, 0.15)',
                                    color: copiedToast ? 'var(--success-color)' : 'var(--accent-purple)',
                                    border: `1px solid ${copiedToast ? 'var(--success-color)' : 'rgba(168, 85, 247, 0.4)'}`,
                                    borderRadius: '10px', cursor: 'pointer', transition: 'all 0.2s ease'
                                }}
                            >
                                {copiedToast ? <Check size={16} /> : <Copy size={16} />}
                                <span>{copiedToast ? '✅ 복사 완료!' : '📋 교정 완성본 복사'}</span>
                            </button>

                            {/* 교정본 파일 다운로드 */}
                            {data.correctedFullText && (
                                <button
                                    onClick={handleDownloadCorrectedFile}
                                    className="interactive"
                                    style={{
                                        display: 'flex', alignItems: 'center', gap: '6px',
                                        padding: '10px 16px', fontSize: '13.5px', fontWeight: 700,
                                        background: 'rgba(59, 130, 246, 0.12)',
                                        color: 'var(--accent-blue)',
                                        border: '1px solid rgba(59, 130, 246, 0.3)',
                                        borderRadius: '10px', cursor: 'pointer'
                                    }}
                                >
                                    <FileText size={16} />
                                    <span>💾 교정본(.md) 다운로드</span>
                                </button>
                            )}

                            <button
                                onClick={() => exportToExcel(data, true)}
                                className="interactive"
                                style={{
                                    display: 'flex', alignItems: 'center', gap: '8px',
                                    padding: '10px 18px', fontSize: '13.5px', fontWeight: 700,
                                    background: 'linear-gradient(135deg, #10b981 0%, #059669 100%)',
                                    color: '#ffffff', border: 'none', borderRadius: '10px',
                                    cursor: 'pointer', boxShadow: '0 4px 12px rgba(16, 185, 129, 0.3)'
                                }}
                            >
                                <Download size={18} />
                                <span>📥 엑셀 내보내기</span>
                            </button>
                        </div>
                    </div>

                    {/* 전체 요약 문구 */}
                    <div style={{ padding: '14px 18px', background: 'rgba(0,0,0,0.15)', borderRadius: '8px', marginBottom: '20px', borderLeft: '4px solid var(--accent-blue)' }}>
                        <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--text-secondary)', display: 'block', marginBottom: '4px' }}>📋 종합 교열 의견 및 전수 점검 요약</span>
                        <p style={{ margin: 0, fontSize: '14px', lineHeight: '1.6', color: 'var(--text-primary)', whiteSpace: 'pre-wrap' }}>
                            {data.summary || '분석이 완료되었습니다.'}
                        </p>
                    </div>

                    {/* 오탈자 결과 명확 표시 테이블 (인라인 Diff 하이라이트 적용) */}
                    <div style={{ overflowX: 'auto' }}>
                        <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '14px' }}>
                            <thead>
                                <tr style={{ borderBottom: '2px solid var(--panel-border)', color: 'var(--text-secondary)', background: 'rgba(255,255,255,0.03)' }}>
                                    <th style={{ padding: '14px 16px', fontWeight: 600, width: '60px' }}>순번</th>
                                    <th style={{ padding: '14px 16px', fontWeight: 600, width: '12%' }}>위치/페이지</th>
                                    <th style={{ padding: '14px 16px', fontWeight: 600, width: '60%' }}>원문 vs 교정 제안 (인라인 단어 Diff)</th>
                                    <th style={{ padding: '14px 16px', fontWeight: 600, width: '22%' }}>오류 유형/사유</th>
                                </tr>
                            </thead>
                            <tbody>
                                {typosList.length > 0 ? (
                                    typosList.map((typo, idx) => (
                                        <tr key={idx} style={{ borderBottom: '1px solid rgba(255,255,255,0.08)', backgroundColor: idx % 2 === 0 ? 'rgba(0,0,0,0.15)' : 'transparent' }}>
                                            <td style={{ padding: '14px 16px', fontWeight: 700, color: 'var(--text-secondary)' }}>{idx + 1}</td>
                                            <td style={{ padding: '14px 16px', color: 'var(--warning-color)', fontWeight: 600 }}>
                                                {typo.page || typo.location || typo.type || typo.section || '1페이지'}
                                            </td>
                                            <td style={{ padding: '14px 16px' }}>
                                                {renderInlineDiff(
                                                    typo.originalText || typo.original || typo.errorText || typo.before || typo.wrong || '',
                                                    typo.correction || typo.correct || typo.after || typo.suggestion || ''
                                                )}
                                            </td>
                                            <td style={{ padding: '14px 16px', color: 'var(--text-secondary)', fontSize: '13px' }}>
                                                {typo.errorType || typo.reason || typo.context || '[표현 품질] 교정'}
                                            </td>
                                        </tr>
                                    ))
                                ) : (
                                    <tr>
                                        <td colSpan={4} style={{ padding: '30px', textAlign: 'center', color: 'var(--success-color)', fontSize: '15px', fontWeight: 600 }}>
                                            ✅ 지적할 오탈자, 띄어쓰기 및 표현 결함이 발견되지 않은 깨끗한 산출물입니다.
                                        </td>
                                    </tr>
                                )}
                            </tbody>
                        </table>
                    </div>
                </section>
            )}

            {/* 문체 일관성 검토 (합니다체/해요체 혼용 감지 - 자동교정이 아닌 검토 권고) */}
            {isTypoMode && styleIssues.length > 0 && (
                <section className="glass-panel animate-slide-up" style={{ padding: '24px', borderLeft: '5px solid #a855f7' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '12px' }}>
                        <AlertTriangle size={20} color="#a855f7" />
                        <h3 style={{ margin: 0, fontSize: '18px', fontWeight: 700, color: 'var(--text-primary)' }}>문체 일관성 검토</h3>
                        <span style={{ padding: '4px 10px', borderRadius: '20px', fontSize: '12px', fontWeight: 700, background: 'rgba(168, 85, 247, 0.15)', color: '#a855f7', border: '1px solid rgba(168, 85, 247, 0.3)' }}>
                            {styleIssues.length}건
                        </span>
                    </div>
                    <p style={{ margin: '0 0 16px', fontSize: '13px', color: 'var(--text-secondary)', lineHeight: '1.6' }}>
                        문서는 주로 <strong style={{ color: 'var(--text-primary)' }}>{styleIssues[0]?.dominantStyle}</strong>로 작성되었으나, 아래 문장은 다른 문체로 되어 있어 통일이 필요합니다.
                        문체 변경은 문장 구조 재작성이 필요해 자동으로 고치지 않으니 직접 검토해 주세요.
                    </p>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                        {styleIssues.map((issue, idx) => (
                            <div key={idx} style={{ padding: '12px 14px', background: 'rgba(168, 85, 247, 0.06)', border: '1px solid rgba(168, 85, 247, 0.15)', borderRadius: '8px' }}>
                                <div style={{ display: 'flex', gap: '8px', alignItems: 'center', marginBottom: '4px', fontSize: '12px' }}>
                                    <span style={{ color: 'var(--text-muted)' }}>{issue.page}</span>
                                    <span style={{ color: '#a855f7', fontWeight: 700 }}>{issue.currentStyle}</span>
                                    <span style={{ color: 'var(--text-muted)' }}>→ {issue.dominantStyle}로 통일 권장</span>
                                </div>
                                <div style={{ fontSize: '14px', color: 'var(--text-primary)' }}>{issue.sentence}</div>
                            </div>
                        ))}
                    </div>
                </section>
            )}

            {/* 원문 전체 대조 보기 (인라인 트랙체인지) */}
            {isTypoMode && trackChanges && trackChanges.matchedCount > 0 && (
                <section className="glass-panel animate-slide-up" style={{ padding: '24px' }}>
                    <div
                        onClick={() => setShowCompareModal(v => !v)}
                        style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', cursor: 'pointer' }}
                    >
                        <h3 style={{ margin: 0, fontSize: '18px', fontWeight: 700, color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                            <Eye size={20} color="var(--accent-color)" />
                            원문 전체 대조 보기
                            <span style={{ fontSize: '12px', fontWeight: 500, color: 'var(--text-muted)' }}>
                                (교정 {trackChanges.matchedCount}/{trackChanges.totalCandidates}건 표시)
                            </span>
                        </h3>
                        <span style={{ fontSize: '13px', color: 'var(--accent-color)', fontWeight: 600 }}>
                            {showCompareModal ? '접기 ▲' : '펼치기 ▼'}
                        </span>
                    </div>
                    {showCompareModal && (
                        <div className="animate-fade-in" style={{
                            marginTop: '16px', padding: '20px', background: 'rgba(0,0,0,0.2)', borderRadius: '10px',
                            maxHeight: '480px', overflowY: 'auto', fontSize: '14.5px', lineHeight: '2', whiteSpace: 'pre-wrap', wordBreak: 'break-word'
                        }}>
                            {trackChanges.segments.map((seg, idx) => seg.type === 'text' ? (
                                <span key={idx} style={{ color: 'var(--text-secondary)' }}>{seg.content}</span>
                            ) : (
                                <span key={idx}>
                                    <del style={{ color: '#f87171', background: 'rgba(239, 68, 68, 0.15)', padding: '1px 4px', borderRadius: '4px', textDecoration: 'line-through' }}>{seg.orig}</del>
                                    <ins style={{ color: '#34d399', background: 'rgba(16, 185, 129, 0.15)', padding: '1px 4px', borderRadius: '4px', textDecoration: 'none', fontWeight: 700, marginLeft: '2px' }}>{seg.corr}</ins>
                                </span>
                            ))}
                        </div>
                    )}
                </section>
            )}

            {/* 1. 종합 준수 현황 (일반 검증 모드일 때만 위쪽 노출) */}
            {!isTypoMode && (
                <section className="animate-slide-up stagger-1" style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
                    <div className="glass-panel" style={{ padding: '32px', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', position: 'relative', overflow: 'hidden' }}>
                        <div style={{ position: 'absolute', top: '-10%', left: '-10%', width: '40%', height: '40%', background: 'var(--success-color)', opacity: 0.1, filter: 'blur(40px)', borderRadius: '50%' }}></div>
                        <h3 style={{ margin: '0 0 24px', fontSize: '15px', color: 'var(--text-secondary)', fontWeight: 600, letterSpacing: '1px', textTransform: 'uppercase' }}>전체 요구사항 준수율</h3>
                        <div style={{ position: 'relative', width: '160px', height: '160px', display: 'flex', alignItems: 'center', justifyContent: 'center', borderRadius: '50%', background: `conic-gradient(var(--success-color) ${displayScore}%, rgba(255,255,255,0.05) 0)`, boxShadow: '0 0 30px rgba(16, 185, 129, 0.15)' }}>
                            <div style={{ position: 'absolute', width: '130px', height: '130px', background: 'var(--bg-secondary)', borderRadius: '50%', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', boxShadow: 'inset 0 0 20px rgba(0,0,0,0.4)' }}>
                                <span style={{ fontSize: '38px', fontWeight: 800, color: 'var(--text-primary)', letterSpacing: '-1px' }}>{displayScore}%</span>
                                <span style={{ fontSize: '11px', color: 'var(--success-color)', fontWeight: 600, marginTop: '-4px' }}>COMPLIANCE</span>
                            </div>
                        </div>
                    </div>

                    <div className="glass-panel animate-slide-up stagger-1" style={{ padding: '24px', display: 'flex', flexDirection: 'column', justifyContent: 'center', minHeight: '100px' }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '16px', flexWrap: 'wrap' }}>
                            <ShieldAlert size={20} color={data.summary && (data.summary.includes('Gemini 검증 실패') || data.summary.includes('소진되었습니다')) ? 'var(--danger-color)' : 'var(--warning-color)'} style={{ flexShrink: 0 }} />
                            <h3 style={{ margin: 0, fontSize: '18px', flex: '1 1 200px', whiteSpace: 'normal' }}>종합 평가 보고서</h3>
                            <div style={{ display: 'flex', gap: '8px', flexShrink: 0, flexWrap: 'wrap' }}>
                                <button
                                    onClick={() => exportToExcel(data, isTypoMode)}
                                    className="interactive"
                                    style={{
                                        display: 'flex', alignItems: 'center', gap: '8px',
                                        padding: '8px 14px', fontSize: '13px', fontWeight: 600,
                                        background: 'rgba(34, 197, 94, 0.1)', color: 'var(--success-color)',
                                        border: '1px solid rgba(34, 197, 94, 0.2)', borderRadius: '10px', cursor: 'pointer',
                                    }}
                                >
                                    <Download size={16} />
                                    <span>보고서 내보내기</span>
                                </button>
                            </div>
                        </div>
                        <p style={{ margin: 0, lineHeight: '1.6', color: 'var(--text-primary)', fontSize: '15px', fontWeight: 400, whiteSpace: 'pre-wrap' }}>
                            {data.summary}
                        </p>
                    </div>
                </section>
            )}

            {/* 2. 요구사항 추적 매트릭스 (RTM) */}
            {data.rtm && data.rtm.length > 0 && (
                <section className="glass-panel animate-slide-up stagger-2" style={{ padding: '24px' }}>
                <h3 style={{ margin: '0 0 16px', fontSize: '18px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <CheckCircle2 size={20} color="var(--accent-color)" />
                    요구사항 매핑 현황 (Semantic Map)
                </h3>
                <div style={{ overflowX: 'auto' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left' }}>
                        <thead>
                            <tr style={{ borderBottom: '1px solid var(--panel-border)', color: 'var(--text-secondary)' }}>
                                <th style={{ padding: '12px 16px', fontWeight: 500 }}>분류</th>
                                <th style={{ padding: '12px 16px', fontWeight: 500 }}>기준 문서 요건</th>
                                <th style={{ padding: '12px 16px', fontWeight: 500 }}>상태</th>
                                <th style={{ padding: '12px 16px', fontWeight: 500 }}>산출물 증빙 위치</th>
                            </tr>
                        </thead>
                        <tbody>
                            {data.rtm.map((item, idx) => (
                                <tr key={idx} style={{ borderBottom: '1px solid rgba(255,255,255,0.05)', backgroundColor: idx % 2 === 0 ? 'rgba(0,0,0,0.1)' : 'transparent' }}>
                                    <td style={{ padding: '12px 16px' }}>
                                        <span style={{
                                            padding: '4px 8px', borderRadius: '4px', fontSize: '12px', fontWeight: 500,
                                            background: item.type === '필수' ? 'rgba(239, 68, 68, 0.1)' : 'rgba(59, 130, 246, 0.1)',
                                            color: item.type === '필수' ? 'var(--danger-color)' : 'var(--accent-color)'
                                        }}>
                                            {item.type}
                                        </span>
                                    </td>
                                    <td style={{ padding: '12px 16px', fontSize: '14px' }}>{item.requirement}</td>
                                    <td style={{ padding: '12px 16px' }}>
                                        <span style={{
                                            display: 'flex', alignItems: 'center', gap: '4px', fontSize: '13px', fontWeight: 600,
                                            color: item.status === '이행(O)' ? 'var(--success-color)' : item.status === '미이행(X)' ? 'var(--danger-color)' : 'var(--warning-color)'
                                        }}>
                                            {item.status === '이행(O)' ? <CheckCircle2 size={16} /> : item.status === '미이행(X)' ? <XCircle size={16} /> : <AlertTriangle size={16} />}
                                            {item.status}
                                        </span>
                                    </td>
                                    <td style={{ padding: '12px 16px', fontSize: '14px', color: 'var(--text-secondary)' }}>{item.location}</td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
                </section>
            )}

            {/* 2.5 요구사항별 산출물 매핑 상세 */}
            {data.requirementMapping && data.requirementMapping.length > 0 && (
                <section className="glass-panel animate-slide-up stagger-3" style={{ padding: '24px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '16px' }}>
                        <h3 style={{ margin: 0, fontSize: '18px', display: 'flex', alignItems: 'center', gap: '8px', flex: 1 }}>
                            <ArrowRightLeft size={20} color="var(--accent-color)" />
                            요구사항별 산출물 매핑 상세
                        </h3>
                    </div>

                    {/* 요약 통계 */}
                    {(() => {
                        const total = data.requirementMapping.length;
                        const met = data.requirementMapping.filter(i => i.status === '이행(O)').length;
                        const partial = data.requirementMapping.filter(i => i.status === '부분 이행(△)').length;
                        const unmet = data.requirementMapping.filter(i => i.status === '미이행(X)').length;
                        return (
                            <div style={{ display: 'flex', gap: '12px', marginBottom: '20px', flexWrap: 'wrap' }}>
                                {[
                                    { label: '전체', value: total, color: 'var(--accent-color)', bg: 'rgba(59,130,246,0.1)' },
                                    { label: '이행(O)', value: met, color: 'var(--success-color)', bg: 'rgba(16,185,129,0.1)' },
                                    { label: '부분 이행(△)', value: partial, color: 'var(--warning-color)', bg: 'rgba(245,158,11,0.1)' },
                                    { label: '미이행(X)', value: unmet, color: 'var(--danger-color)', bg: 'rgba(239,68,68,0.1)' },
                                ].map((stat, i) => (
                                    <div key={i} style={{
                                        display: 'flex', alignItems: 'center', gap: '8px',
                                        padding: '8px 14px', borderRadius: '8px',
                                        background: stat.bg, border: `1px solid ${stat.color}22`,
                                    }}>
                                        <span style={{ fontSize: '22px', fontWeight: 700, color: stat.color }}>{stat.value}</span>
                                        <span style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>{stat.label}</span>
                                    </div>
                                ))}
                            </div>
                        );
                    })()}

                    {/* 매핑 카드 목록 */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '14px' }}>
                        {data.requirementMapping.map((item, idx) => {
                            const statusColor = item.status === '이행(O)'
                                ? 'var(--success-color)'
                                : item.status === '부분 이행(△)'
                                    ? 'var(--warning-color)'
                                    : 'var(--danger-color)';
                            const statusBg = item.status === '이행(O)'
                                ? 'rgba(16,185,129,0.1)'
                                : item.status === '부분 이행(△)'
                                    ? 'rgba(245,158,11,0.1)'
                                    : 'rgba(239,68,68,0.1)';
                            const statusIcon = item.status === '이행(O)'
                                ? <CheckCircle2 size={14} />
                                : item.status === '부분 이행(△)'
                                    ? <AlertTriangle size={14} />
                                    : <XCircle size={14} />;

                            return (
                                <div key={idx} style={{
                                    background: 'rgba(255,255,255,0.02)', borderRadius: '16px', padding: '24px',
                                    border: '1px solid var(--glass-border)',
                                    borderLeft: `5px solid ${statusColor}`,
                                    transition: 'transform 0.3s cubic-bezier(0.16, 1, 0.3, 1)',
                                }}
                                className="interactive"
                                >
                                    {/* 헤더: ID + 필수/선택 + 분류 + 상태 */}
                                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '8px', flexWrap: 'wrap' }}>
                                        <span style={{
                                            fontSize: '12px', fontWeight: 700, padding: '3px 8px',
                                            background: 'rgba(59,130,246,0.15)', color: 'var(--accent-color)',
                                            borderRadius: '4px', fontFamily: 'monospace',
                                        }}>{item.id}</span>
                                        <span style={{
                                            fontSize: '11px', fontWeight: 600, padding: '2px 6px',
                                            background: item.type === '필수' ? 'rgba(239,68,68,0.12)' : 'rgba(59,130,246,0.12)',
                                            color: item.type === '필수' ? 'var(--danger-color)' : 'var(--accent-color)',
                                            borderRadius: '3px',
                                            border: `1px solid ${item.type === '필수' ? 'rgba(239,68,68,0.25)' : 'rgba(59,130,246,0.25)'}`,
                                        }}>{item.type || '필수'}</span>
                                        <span style={{
                                            fontSize: '12px', fontWeight: 500, padding: '3px 8px',
                                            background: 'rgba(255,255,255,0.05)', color: 'var(--text-secondary)',
                                            borderRadius: '4px',
                                        }}>{item.category}</span>
                                        {item.levelLabel && (
                                            <span style={{
                                                fontSize: '10px', fontWeight: 500, padding: '2px 6px',
                                                background: 'rgba(168,85,247,0.1)', color: '#a855f7',
                                                borderRadius: '3px',
                                            }}>{item.levelLabel}</span>
                                        )}
                                        <span style={{
                                            marginLeft: 'auto', display: 'flex', alignItems: 'center', gap: '4px',
                                            fontSize: '13px', fontWeight: 600, padding: '3px 10px',
                                            background: statusBg, color: statusColor, borderRadius: '4px',
                                        }}>
                                            {statusIcon}
                                            {item.status}
                                        </span>
                                    </div>

                                    {/* 계층 경로 */}
                                    {item.path && (
                                        <div style={{
                                            fontSize: '11px', color: 'var(--text-secondary)',
                                            marginBottom: '8px', padding: '4px 8px',
                                            background: 'rgba(255,255,255,0.03)', borderRadius: '4px',
                                            fontFamily: 'monospace', opacity: 0.8,
                                        }}>
                                            📂 {item.path}
                                        </div>
                                    )}

                                    {/* 요구사항 원문 */}
                                    <p style={{ margin: '0 0 14px', fontSize: '15px', fontWeight: 500, lineHeight: '1.5' }}>
                                        {item.requirement}
                                    </p>

                                    {/* 매핑 정보 2열 그리드 */}
                                    <div style={{
                                        display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '10px',
                                        fontSize: '13px', marginBottom: '14px',
                                    }}>
                                        <div style={{ background: 'rgba(255,255,255,0.03)', padding: '10px', borderRadius: '6px' }}>
                                            <span style={{ color: 'var(--text-secondary)', display: 'block', marginBottom: '4px', fontSize: '12px' }}>📄 산출물 대응 섹션</span>
                                            <span style={{ color: 'var(--text-primary)' }}>{item.artifactSection}</span>
                                        </div>
                                        <div style={{ background: 'rgba(255,255,255,0.03)', padding: '10px', borderRadius: '6px' }}>
                                            <span style={{ color: 'var(--text-secondary)', display: 'block', marginBottom: '4px', fontSize: '12px' }}>📝 기술 내용</span>
                                            <span style={{ color: 'var(--text-primary)', lineHeight: '1.5' }}>{item.artifactContent}</span>
                                        </div>
                                    </div>

                                    {/* 차이점 (gap) */}
                                    {item.gap && (
                                        <div style={{
                                            padding: '10px 12px', borderRadius: '6px',
                                            background: `${statusBg}`, border: `1px solid ${statusColor}33`,
                                            fontSize: '13px', lineHeight: '1.6',
                                        }}>
                                            <span style={{ fontWeight: 600, color: statusColor, marginRight: '6px' }}>⚠ 차이점:</span>
                                            <span style={{ color: 'var(--text-primary)' }}>{item.gap}</span>
                                        </div>
                                    )}
                                </div>
                            );
                        })}
                    </div>
                </section>
            )}

            {/* 3. 주요 누락/비준수 상세 */}
            {data.omissions && data.omissions.length > 0 && (
                <section className="glass-panel animate-slide-up stagger-4" style={{ padding: '24px' }}>
                <h3 style={{ margin: '0 0 16px', fontSize: '18px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <FileWarning size={20} color="var(--danger-color)" />
                    주요 누락(Omission) 및 비준수 사항
                </h3>

                <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
                    {data.omissions.map((omission, idx) => (
                        <div key={idx} style={{ background: 'rgba(0, 0, 0, 0.2)', padding: '16px', borderRadius: '8px', borderLeft: '4px solid var(--danger-color)' }}>
                            <h4 style={{ margin: '0 0 12px', fontSize: '16px', display: 'flex', alignItems: 'center', gap: '6px' }}>
                                <span style={{ background: 'var(--danger-color)', color: '#fff', width: '24px', height: '24px', borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '12px' }}>
                                    {idx + 1}
                                </span>
                                {omission.title}
                            </h4>

                            <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', fontSize: '14px' }}>
                                <div style={{ display: 'flex', gap: '8px' }}>
                                    <span style={{ color: 'var(--text-secondary)', minWidth: '80px' }}>기준 근거:</span>
                                    <span style={{ color: 'var(--text-primary)' }}>{omission.evidence}</span>
                                </div>
                                <div style={{ display: 'flex', gap: '8px' }}>
                                    <span style={{ color: 'var(--text-secondary)', minWidth: '80px' }}>판단 이유:</span>
                                    <span style={{ color: 'var(--text-primary)' }}>{omission.reason}</span>
                                </div>
                                <div style={{ display: 'flex', gap: '8px', marginTop: '4px', padding: '8px', background: 'rgba(59, 130, 246, 0.1)', borderRadius: '4px', border: '1px solid rgba(59, 130, 246, 0.2)' }}>
                                    <span style={{ color: 'var(--accent-color)', fontWeight: 600, minWidth: '80px' }}>개선 권고:</span>
                                    <span style={{ color: 'var(--text-primary)' }}>{omission.recommendation}</span>
                                </div>
                            </div>
                        </div>
                    ))}
                </div>
                </section>
            )}

            {/* 4. 산출물 오탈자 및 용어 점검 결과 */}
            {data.typos && data.typos.length > 0 && (
                <section className="glass-panel animate-slide-up stagger-5" style={{ padding: '24px' }}>
                    <h3 style={{ margin: '0 0 16px', fontSize: '18px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <PenTool size={20} color="var(--warning-color)" />
                        ISMP 산출물 전문 교정/교열 결과
                    </h3>
                    <div style={{ overflowX: 'auto' }}>
                        <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '14px' }}>
                            <thead>
                                <tr style={{ borderBottom: '1px solid var(--panel-border)', color: 'var(--text-secondary)' }}>
                                    <th style={{ padding: '12px 16px', fontWeight: 500, width: '60px' }}>순번</th>
                                    <th style={{ padding: '12px 16px', fontWeight: 500, width: '15%' }}>위치/페이지</th>
                                    <th style={{ padding: '12px 16px', fontWeight: 500, width: '30%' }}>원문 문장 전체</th>
                                    <th style={{ padding: '12px 16px', fontWeight: 500, width: '30%' }}>수정 제안 문장</th>
                                    <th style={{ padding: '12px 16px', fontWeight: 500, width: '25%' }}>오류 유형/사유</th>
                                </tr>
                            </thead>
                            <tbody>
                                {data.typos.map((typo, idx) => (
                                    <tr key={idx} style={{ borderBottom: '1px solid rgba(255,255,255,0.05)', backgroundColor: idx % 2 === 0 ? 'rgba(0,0,0,0.1)' : 'transparent' }}>
                                        <td style={{ padding: '12px 16px' }}>{idx + 1}</td>
                                        <td style={{ padding: '12px 16px', color: 'var(--warning-color)' }}>{typo.page || typo.location || typo.type}</td>
                                        <td style={{ padding: '12px 16px', color: 'var(--danger-color)', textDecoration: 'none' }}>{typo.originalText || typo.errorText}</td>
                                        <td style={{ padding: '12px 16px', color: 'var(--success-color)', fontWeight: 600 }}>{typo.correction}</td>
                                        <td style={{ padding: '12px 16px', color: 'var(--text-secondary)' }}>{typo.errorType || typo.reason || typo.context}</td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                </section>
            )}
        </div>
    );
}
