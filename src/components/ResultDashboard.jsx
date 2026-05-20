import React from 'react';
import { ShieldAlert, CheckCircle2, XCircle, FileWarning, AlertTriangle, ClipboardList, ArrowRightLeft, Download, PenTool, RotateCcw } from 'lucide-react';
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

    if (!isTypoMode) {
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
            '분류': item.type || '필수',
            '기준문서요건': item.requirement || '',
            '카테고리': item.category || '',
            '수준': item.levelLabel || '',
            '상태': item.status || '',
            '산출물증빙위치': item.location || '',
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
            'ID': item.id || '',
            '분류': item.type || '필수',
            '카테고리': item.category || '',
            '수준': item.levelLabel || '',
            '계층경로': item.path || '',
            '요구사항': item.requirement || '',
            '산출물대응섹션': item.artifactSection || '',
            '산출물기술내용': (item.artifactContent || '').replace(/^"|"$/g, ''),
            '상태': item.status || '',
            '차이점': item.gap || '',
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
                '항목': item.title || '',
                '근거': item.evidence || '',
                '사유': item.reason || '',
                '권고사항': item.recommendation || '',
            }));
            addSheet('누락사항', omiHeaders, omiRows);
        }
    }

    // ── 시트4: 문서 품질(오탈자 등) 점검 (오직 오탈자 점검 모드에서만 출력) ──
    if (isTypoMode && data.typos && data.typos.length > 0) {
        const typoHeaders = [
            { header: '순번', key: '순번', width: 6 },
            { header: '페이지/위치', key: '위치', width: 25 },
            { header: '원문 문장 전체', key: '원문', width: 40 },
            { header: '수정 제안 문장', key: '수정', width: 40 },
            { header: '오류 유형/사유', key: '사유', width: 60 }
        ];
        const typoRows = data.typos.map((item, idx) => ({
            '순번': idx + 1,
            '위치': item.page || item.location || item.type || '',
            '원문': item.originalText || item.errorText || '',
            '수정': item.correction || '',
            '사유': item.errorType || item.reason || item.context || '',
        }));
        addSheet('교정교열_결과', typoHeaders, typoRows);
    }

    const now = new Date();
    const dateStr = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}`;
    const filePrefix = data.artifactFileName ? `[${data.artifactFileName}]_` : '';
    const fileName = isTypoMode ? `${filePrefix}교정교열_결과_${dateStr}.xlsx` : `${filePrefix}기준문서_검증결과_${dateStr}.xlsx`;

    try {
        const buffer = await wb.xlsx.writeBuffer();
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
        
        saveAs(new Blob([buffer]), fileName);
    } catch (e) {
        console.error('Excel export failed:', e);
    }
}

export default function ResultDashboard({ data, isTypoMode = false, onRetry }) {
    if (!data) return null;

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

            {/* 1. 종합 준수 현황 */}
            <section className="animate-slide-up stagger-1" style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
                <div className="glass-panel" style={{ padding: '32px', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', position: 'relative', overflow: 'hidden' }}>
                    <div style={{ position: 'absolute', top: '-10%', left: '-10%', width: '40%', height: '40%', background: 'var(--success-color)', opacity: 0.1, filter: 'blur(40px)', borderRadius: '50%' }}></div>
                    <h3 style={{ margin: '0 0 24px', fontSize: '15px', color: 'var(--text-secondary)', fontWeight: 600, letterSpacing: '1px', textTransform: 'uppercase' }}>전체 요구사항 준수율</h3>
                    <div style={{ position: 'relative', width: '160px', height: '160px', display: 'flex', alignItems: 'center', justifyContent: 'center', borderRadius: '50%', background: `conic-gradient(var(--success-color) ${data.score}%, rgba(255,255,255,0.05) 0)`, boxShadow: '0 0 30px rgba(16, 185, 129, 0.15)' }}>
                        <div style={{ position: 'absolute', width: '130px', height: '130px', background: 'var(--bg-secondary)', borderRadius: '50%', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', boxShadow: 'inset 0 0 20px rgba(0,0,0,0.4)' }}>
                            <span style={{ fontSize: '38px', fontWeight: 800, color: 'var(--text-primary)', letterSpacing: '-1px' }}>{data.score}%</span>
                            <span style={{ fontSize: '11px', color: 'var(--success-color)', fontWeight: 600, marginTop: '-4px' }}>COMPLIANCE</span>
                        </div>
                    </div>
                </div>

                <div className="glass-panel animate-slide-up stagger-1" style={{ padding: '24px', display: 'flex', flexDirection: 'column', justifyContent: 'center', minHeight: '100px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '16px', flexWrap: 'wrap' }}>
                        <ShieldAlert size={20} color={data.summary && (data.summary.includes('실패') || data.summary.includes('소진')) ? 'var(--danger-color)' : 'var(--warning-color)'} style={{ flexShrink: 0 }} />
                        <h3 style={{ margin: 0, fontSize: '18px', flex: '1 1 200px', whiteSpace: 'normal' }}>종합 평가 보고서</h3>
                        <div style={{ display: 'flex', gap: '8px', flexShrink: 0, flexWrap: 'wrap' }}>
                            {data.summary && (data.summary.includes('실패') || data.summary.includes('소진')) && onRetry && (
                                <button
                                    onClick={onRetry}
                                    className="interactive pulse-text"
                                    style={{
                                        display: 'flex', alignItems: 'center', gap: '8px',
                                        padding: '8px 14px', fontSize: '13px', fontWeight: 700,
                                        background: 'rgba(59, 130, 246, 0.15)',
                                        color: 'var(--accent-blue)',
                                        border: '2px solid var(--accent-blue)',
                                        borderRadius: '10px',
                                        cursor: 'pointer',
                                    }}
                                >
                                    <RotateCcw size={16} />
                                    <span className="mobile-hide-text">지금 다시 시도</span><span className="mobile-only-show" style={{ display: 'none' }}>재시도</span>
                                </button>
                            )}
                            <button
                                onClick={() => exportToExcel(data, isTypoMode)}
                                className="interactive"
                                style={{
                                    display: 'flex', alignItems: 'center', gap: '8px',
                                    padding: '8px 14px', fontSize: '13px', fontWeight: 600,
                                    background: 'rgba(34, 197, 94, 0.1)',
                                    color: 'var(--success-color)',
                                    border: '1px solid rgba(34, 197, 94, 0.2)',
                                    borderRadius: '10px',
                                    cursor: 'pointer',
                                }}
                            >
                                <Download size={16} />
                                <span className="mobile-hide-text">보고서 내보내기</span><span className="mobile-only-show" style={{ display: 'none' }}>다운로드</span>
                            </button>
                        </div>
                    </div>
                    <p style={{ 
                        margin: 0, lineHeight: '1.6', 
                        color: data.summary && (data.summary.includes('실패') || data.summary.includes('소진')) ? 'var(--danger-color)' : 'var(--text-primary)', 
                        fontSize: '15px',
                        fontWeight: data.summary && (data.summary.includes('실패') || data.summary.includes('소진')) ? 500 : 400
                    }}>
                        {data.summary}
                    </p>
                </div>
            </section>

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
