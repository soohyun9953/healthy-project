import React, { useState, useEffect, useRef, useMemo } from 'react';
import {
    SlidersHorizontal,
    Upload,
    Download,
    CheckCircle2,
    AlertTriangle,
    XCircle,
    FileSpreadsheet,
    FileText,
    Presentation,
    Sparkles,
    RefreshCw,
    X,
    Filter,
    Plus,
    Trash2,
    Palette,
    Layers,
    AlignLeft,
    Check,
    Copy,
    Info,
    HelpCircle
} from 'lucide-react';
import {
    DEFAULT_MASTER_SPEC,
    loadMasterSpec,
    saveMasterSpec,
    extractMasterSpec,
    inspectPptxCompliance,
    autoFixPptxFormatting,
    exportMasterComplianceReportToExcel,
    EMU_PER_INCH
} from '../utils/pptxMasterEngine';

export default function PptxMasterInspectorView() {
    // ── 1. 마스터 규격 상태 ──────────────────────────────────────────
    const [masterSpec, setMasterSpec] = useState(() => loadMasterSpec());
    const [isExtractingMaster, setIsExtractingMaster] = useState(false);
    const [newFontInput, setNewFontInput] = useState('');

    // ── 2. 검수 대상 파일 및 검수 결과 상태 ─────────────────────────
    const [targetFiles, setTargetFiles] = useState([]);
    const [activeTargetIndex, setActiveTargetIndex] = useState(0);
    const [isInspecting, setIsInspecting] = useState(false);
    const [inspectionResults, setInspectionResults] = useState([]);

    // ── 3. 필터 및 UI 상태 ─────────────────────────────────────────
    const [activeViolationTab, setActiveViolationTab] = useState('ALL');
    const [severityFilter, setSeverityFilter] = useState('ALL');
    const [isFixing, setIsFixing] = useState(false);
    const [isExportingExcel, setIsExportingExcel] = useState(false);
    const [toastMessage, setToastMessage] = useState(null);

    // 자동 교정 옵션
    const [fixOptions, setFixOptions] = useState({
        fixFonts: true,
        fixCitations: true,
        fixPositions: true,
        fixBreadcrumbDots: true
    });

    const masterFileInputRef = useRef(null);
    const targetFileInputRef = useRef(null);

    const showToast = (msg, type = 'success') => {
        setToastMessage({ msg, type });
        setTimeout(() => setToastMessage(null), 4000);
    };

    // 마스터 규격 변경 시 자동 저장
    useEffect(() => {
        saveMasterSpec(masterSpec);
    }, [masterSpec]);

    // ── 4. 기준 마스터 PPTX 업로드 및 규격 자동 추출 ────────────────
    const handleMasterFileSelected = async (file) => {
        if (!file) return;
        setIsExtractingMaster(true);
        try {
            const spec = await extractMasterSpec(file);
            setMasterSpec(spec);
            showToast(`[${file.name}]에서 슬라이드 마스터 표준 규격(폰트, 컬러, 5대 영역 좌표)을 성공적으로 추출했습니다!`);
        } catch (err) {
            console.error('마스터 규격 추출 실패:', err);
            showToast(`마스터 규격 추출 오류: ${err.message}`, 'error');
        } finally {
            setIsExtractingMaster(false);
        }
    };

    const handleResetDefaultSpec = () => {
        setMasterSpec(DEFAULT_MASTER_SPEC);
        saveMasterSpec(DEFAULT_MASTER_SPEC);
        showToast('공공의료 AI ISP 표준 골든마스터 규격 프리셋으로 복원되었습니다.');
    };

    // 허용 폰트 추가/삭제
    const handleAddFont = () => {
        const trimmed = newFontInput.trim();
        if (!trimmed) return;
        if (!masterSpec.allowedFonts.includes(trimmed)) {
            setMasterSpec(prev => ({
                ...prev,
                allowedFonts: [...prev.allowedFonts, trimmed]
            }));
            setNewFontInput('');
            showToast(`'${trimmed}' 폰트가 허용 폰트 규격에 추가되었습니다.`);
        }
    };

    const handleRemoveFont = (fontToRemove) => {
        setMasterSpec(prev => ({
            ...prev,
            allowedFonts: prev.allowedFonts.filter(f => f !== fontToRemove)
        }));
    };

    // ── 5. 검수 대상 파일 등록 및 정합성 검사 ───────────────────────
    const handleTargetFilesSelected = async (files) => {
        if (!files || files.length === 0) return;
        const newFiles = Array.from(files);
        setTargetFiles(prev => [...prev, ...newFiles]);

        setIsInspecting(true);
        try {
            const results = [];
            for (let i = 0; i < newFiles.length; i++) {
                const res = await inspectPptxCompliance(newFiles[i], masterSpec);
                results.push(res);
            }
            setInspectionResults(prev => [...prev, ...results]);
            showToast(`${newFiles.length}개 산출물에 대한 슬라이드 마스터 서식 검수가 완료되었습니다.`);
        } catch (err) {
            console.error('검수 오류:', err);
            showToast(`서식 검수 중 오류 발생: ${err.message}`, 'error');
        } finally {
            setIsInspecting(false);
        }
    };

    // 현재 선택된 파일의 검수 결과
    const currentResult = inspectionResults[activeTargetIndex] || null;
    const currentTargetFile = targetFiles[activeTargetIndex] || null;

    // 재검수
    const handleReinspect = async () => {
        if (!currentTargetFile) return;
        setIsInspecting(true);
        try {
            const res = await inspectPptxCompliance(currentTargetFile, masterSpec);
            setInspectionResults(prev => {
                const copy = [...prev];
                copy[activeTargetIndex] = res;
                return copy;
            });
            showToast(`[${currentTargetFile.name}] 서식 재검수가 완료되었습니다.`);
        } catch (err) {
            showToast(`재검수 실패: ${err.message}`, 'error');
        } finally {
            setIsInspecting(false);
        }
    };

    // ── 6. 원클릭 표준 서식 자동 교정 (.pptx 다운로드) ─────────────
    const handleAutoFix = async () => {
        if (!currentTargetFile) {
            showToast('교정할 대상 파일이 없습니다.', 'error');
            return;
        }

        setIsFixing(true);
        try {
            const fixRes = await autoFixPptxFormatting(currentTargetFile, masterSpec, fixOptions);
            showToast(`총 ${fixRes.fixedCount}개 서식 결함이 원본 보존 상태로 교정되어 다운로드되었습니다!`);
            // 교정 후 재검수 실행
            setTimeout(() => {
                handleReinspect();
            }, 800);
        } catch (err) {
            console.error('자동 교정 오류:', err);
            showToast(`자동 교정 실패: ${err.message}`, 'error');
        } finally {
            setIsFixing(false);
        }
    };

    // ── 7. 엑셀 검수 리포트 다운로드 ─────────────────────────────────
    const handleExportExcel = async () => {
        if (!currentResult) {
            showToast('다운로드할 검수 결과가 없습니다.', 'error');
            return;
        }
        setIsExportingExcel(true);
        try {
            await exportMasterComplianceReportToExcel(currentResult, masterSpec);
            showToast('마스터 서식 검수 상세 리포트(.xlsx)가 다운로드되었습니다.');
        } catch (err) {
            showToast(`엑셀 리포트 생성 오류: ${err.message}`, 'error');
        } finally {
            setIsExportingExcel(false);
        }
    };

    // ── 8. 필터링된 위반 목록 ────────────────────────────────────────
    const filteredViolations = useMemo(() => {
        if (!currentResult || !currentResult.violations) return [];
        return currentResult.violations.filter(v => {
            if (activeViolationTab !== 'ALL' && v.type !== activeViolationTab) return false;
            if (severityFilter !== 'ALL' && v.severity !== severityFilter) return false;
            return true;
        });
    }, [currentResult, activeViolationTab, severityFilter]);

    // EMU -> Inch 환산 헬퍼
    const emuToInch = (val) => (val / EMU_PER_INCH).toFixed(2);
    const szToPt = (val) => (val / 100).toFixed(1);

    return (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '20px', width: '100%' }}>
            {/* ── 알림 토스트 ────────────────────────────────────────── */}
            {toastMessage && (
                <div style={{
                    position: 'fixed',
                    top: '24px',
                    right: '24px',
                    zIndex: 9999,
                    background: toastMessage.type === 'error' ? 'rgba(239, 68, 68, 0.95)' : 'rgba(16, 185, 129, 0.95)',
                    color: '#fff',
                    padding: '12px 20px',
                    borderRadius: '10px',
                    boxShadow: '0 8px 24px rgba(0,0,0,0.3)',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '10px',
                    fontSize: '13px',
                    fontWeight: 600,
                    backdropFilter: 'blur(8px)',
                    animation: 'fadeIn 0.2s ease-in-out'
                }}>
                    {toastMessage.type === 'error' ? <AlertTriangle size={18} /> : <CheckCircle2 size={18} />}
                    <span>{toastMessage.msg}</span>
                    <button onClick={() => setToastMessage(null)} style={{ background: 'none', border: 'none', color: '#fff', cursor: 'pointer', padding: 0, marginLeft: '8px' }}>
                        <X size={14} />
                    </button>
                </div>
            )}

            {/* ── 상단 헤더 패널 ────────────────────────────────────────── */}
            <div className="glass-panel" style={{
                padding: '24px',
                borderRadius: '16px',
                border: '1px solid var(--panel-border)',
                background: 'linear-gradient(135deg, rgba(249, 115, 22, 0.08) 0%, rgba(99, 102, 241, 0.04) 100%)',
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'flex-start',
                flexWrap: 'wrap',
                gap: '16px'
            }}>
                <div>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '8px' }}>
                        <div style={{
                            width: '38px',
                            height: '38px',
                            borderRadius: '10px',
                            background: 'linear-gradient(135deg, #f97316, #ea580c)',
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                            color: '#fff',
                            boxShadow: '0 4px 12px rgba(249, 115, 22, 0.3)'
                        }}>
                            <SlidersHorizontal size={22} />
                        </div>
                        <h2 style={{ margin: 0, fontSize: '20px', fontWeight: 800, color: 'var(--text-primary)' }}>
                            PPTX 슬라이드 마스터 표준 규격 추출 및 서식 검수·자동교정 (Master Template Inspector)
                        </h2>
                        <span style={{
                            padding: '3px 10px',
                            borderRadius: '20px',
                            background: 'rgba(249, 115, 22, 0.2)',
                            color: '#fb923c',
                            fontSize: '11px',
                            fontWeight: 700,
                            border: '1px solid rgba(249, 115, 22, 0.4)'
                        }}>
                            OpenXML 서식 엔진
                        </span>
                    </div>
                    <p style={{ margin: 0, fontSize: '13px', color: 'var(--text-secondary)', lineHeight: 1.6, maxWidth: '850px' }}>
                        표준 골든 마스터 PPTX(예: <code>v0.95D.pptx</code>)로부터 <strong>허용 폰트, 테마 컬러, 5대 표준 영역(제목/헤드메시지/브레드크럼/출처) 좌표</strong>를 자동 추출하고,
                        작성된 산출물 장표의 <strong>비표준 폰트·좌표 이탈·출처 표기 오류</strong>를 정밀 적발하여 원본 이미지와 차트 훼손 없이 <strong>원클릭으로 자동 교정</strong>합니다.
                    </p>
                </div>

                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
                    <button
                        onClick={handleResetDefaultSpec}
                        style={{
                            padding: '8px 14px',
                            borderRadius: '9px',
                            border: '1px solid rgba(255, 255, 255, 0.15)',
                            background: 'rgba(255, 255, 255, 0.05)',
                            color: 'var(--text-primary)',
                            fontSize: '12px',
                            fontWeight: 600,
                            cursor: 'pointer',
                            display: 'flex',
                            alignItems: 'center',
                            gap: '6px'
                        }}
                    >
                        <RefreshCw size={13} /> 공공의료 ISP 마스터 프리셋 로드
                    </button>
                </div>
            </div>

            {/* ── 좌/우 분할 레이아웃 (좌: 기준 규격 관리, 우: 검수 및 자동교정) ── */}
            <div style={{ display: 'grid', gridTemplateColumns: 'minmax(320px, 380px) 1fr', gap: '20px', alignItems: 'start' }}>

                {/* ── [좌측 패널]: 기준 마스터 양식(Golden Master) 규격 ── */}
                <div className="glass-panel" style={{
                    padding: '20px',
                    borderRadius: '14px',
                    border: '1px solid var(--panel-border)',
                    display: 'flex',
                    flexDirection: 'column',
                    gap: '16px'
                }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                            <Layers size={17} color="#f97316" />
                            <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 700, color: 'var(--text-primary)' }}>
                                기준 마스터 규격 (Golden Spec)
                            </h3>
                        </div>
                    </div>

                    {/* 마스터 PPTX 업로드 드롭존 */}
                    <div
                        onClick={() => masterFileInputRef.current?.click()}
                        style={{
                            border: '2px dashed rgba(249, 115, 22, 0.4)',
                            borderRadius: '10px',
                            padding: '16px',
                            textAlign: 'center',
                            background: 'rgba(249, 115, 22, 0.04)',
                            cursor: 'pointer',
                            transition: 'all 0.2s'
                        }}
                    >
                        <input
                            ref={masterFileInputRef}
                            type="file"
                            accept=".pptx"
                            style={{ display: 'none' }}
                            onChange={(e) => handleMasterFileSelected(e.target.files?.[0])}
                        />
                        <Upload size={22} color="#fb923c" style={{ margin: '0 auto 6px' }} />
                        <div style={{ fontSize: '12.5px', fontWeight: 700, color: 'var(--text-primary)' }}>
                            {isExtractingMaster ? '마스터 규격 추출 중...' : '새 기준 마스터 PPTX 업로드'}
                        </div>
                        <div style={{ fontSize: '11px', color: 'var(--text-muted)', marginTop: '2px' }}>
                            현재: {masterSpec.fileName || '공공의료 AI ISP 표준'}
                        </div>
                    </div>

                    {/* 1) 허용 폰트 규격 관리 */}
                    <div style={{ borderTop: '1px solid var(--panel-border)', paddingTop: '12px' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
                            <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--text-secondary)' }}>
                                🔤 표준 허용 폰트 규격 ({masterSpec.allowedFonts.length}종)
                            </span>
                        </div>
                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px', marginBottom: '8px' }}>
                            {masterSpec.allowedFonts.map((font, idx) => (
                                <span key={idx} style={{
                                    display: 'inline-flex',
                                    alignItems: 'center',
                                    gap: '5px',
                                    padding: '4px 8px',
                                    borderRadius: '6px',
                                    background: idx === 0 ? 'rgba(249, 115, 22, 0.2)' : 'rgba(255, 255, 255, 0.08)',
                                    color: idx === 0 ? '#fb923c' : 'var(--text-primary)',
                                    fontSize: '11.5px',
                                    fontWeight: idx === 0 ? 700 : 500,
                                    border: idx === 0 ? '1px solid rgba(249, 115, 22, 0.4)' : '1px solid rgba(255,255,255,0.1)'
                                }}>
                                    {idx === 0 && '★ 대표 '}
                                    {font}
                                    {masterSpec.allowedFonts.length > 1 && (
                                        <button
                                            onClick={() => handleRemoveFont(font)}
                                            style={{ background: 'none', border: 'none', color: 'var(--text-muted)', cursor: 'pointer', padding: 0 }}
                                        >
                                            <X size={11} />
                                        </button>
                                    )}
                                </span>
                            ))}
                        </div>
                        <div style={{ display: 'flex', gap: '6px' }}>
                            <input
                                type="text"
                                placeholder="추가할 폰트명..."
                                value={newFontInput}
                                onChange={(e) => setNewFontInput(e.target.value)}
                                onKeyDown={(e) => e.key === 'Enter' && handleAddFont()}
                                style={{
                                    flex: 1,
                                    padding: '5px 8px',
                                    borderRadius: '6px',
                                    border: '1px solid var(--panel-border)',
                                    background: 'rgba(0,0,0,0.2)',
                                    color: 'var(--text-primary)',
                                    fontSize: '11.5px'
                                }}
                            />
                            <button
                                onClick={handleAddFont}
                                style={{
                                    padding: '5px 10px',
                                    borderRadius: '6px',
                                    border: '1px solid rgba(255,255,255,0.15)',
                                    background: 'rgba(255,255,255,0.06)',
                                    color: 'var(--text-primary)',
                                    fontSize: '11.5px',
                                    cursor: 'pointer'
                                }}
                            >
                                <Plus size={12} /> 추가
                            </button>
                        </div>
                    </div>

                    {/* 2) 테마 컬러 칩 */}
                    <div style={{ borderTop: '1px solid var(--panel-border)', paddingTop: '12px' }}>
                        <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--text-secondary)', display: 'block', marginBottom: '8px' }}>
                            🎨 대표 테마 컬러 팔레트
                        </span>
                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px' }}>
                            {(masterSpec.themeColors || []).map((clr, i) => (
                                <div
                                    key={i}
                                    title={`#${clr} (클릭 시 복사)`}
                                    onClick={() => {
                                        navigator.clipboard.writeText(`#${clr}`);
                                        showToast(`색상 #${clr} 복사됨`);
                                    }}
                                    style={{
                                        display: 'flex',
                                        alignItems: 'center',
                                        gap: '4px',
                                        padding: '3px 6px',
                                        borderRadius: '5px',
                                        background: 'rgba(0,0,0,0.3)',
                                        border: '1px solid rgba(255,255,255,0.1)',
                                        cursor: 'pointer',
                                        fontSize: '10.5px'
                                    }}
                                >
                                    <div style={{ width: '12px', height: '12px', borderRadius: '3px', background: `#${clr}` }} />
                                    <span style={{ color: 'var(--text-secondary)', fontFamily: 'monospace' }}>#{clr}</span>
                                </div>
                            ))}
                        </div>
                    </div>

                    {/* 3) 5대 표준 영역 좌표 카드 */}
                    <div style={{ borderTop: '1px solid var(--panel-border)', paddingTop: '12px' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
                            <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--text-secondary)' }}>
                                📐 5대 표준 영역 좌표 규격
                            </span>
                            <span style={{ fontSize: '10.5px', color: '#fb923c' }}>
                                허용오차 ±{Math.round((masterSpec.toleranceRatio || 0.08) * 100)}%
                            </span>
                        </div>

                        <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                            {Object.entries(masterSpec.standardZones || {}).map(([key, zone]) => (
                                <div key={key} style={{
                                    padding: '8px 10px',
                                    borderRadius: '7px',
                                    background: 'rgba(255,255,255,0.025)',
                                    border: '1px solid rgba(255,255,255,0.06)',
                                    fontSize: '11px'
                                }}>
                                    <div style={{ fontWeight: 700, color: '#f8fafc', marginBottom: '2px' }}>
                                        {zone.name || key}
                                    </div>
                                    <div style={{ color: 'var(--text-muted)', display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
                                        <span>X: {emuToInch(zone.x)}"</span>
                                        <span>Y: {emuToInch(zone.y)}"</span>
                                        {zone.cx && <span>W: {emuToInch(zone.cx)}"</span>}
                                        {zone.sz && <span style={{ color: '#38bdf8' }}>{szToPt(zone.sz)}pt</span>}
                                    </div>
                                </div>
                            ))}
                        </div>
                    </div>

                    {/* 허용 오차 조절 슬라이더 */}
                    <div style={{ borderTop: '1px solid var(--panel-border)', paddingTop: '10px' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '11.5px', marginBottom: '4px' }}>
                            <span style={{ color: 'var(--text-secondary)' }}>좌표 위치 허용오차율</span>
                            <span style={{ fontWeight: 700, color: '#fb923c' }}>
                                ±{Math.round((masterSpec.toleranceRatio || 0.08) * 100)}%
                            </span>
                        </div>
                        <input
                            type="range"
                            min="0.03"
                            max="0.15"
                            step="0.01"
                            value={masterSpec.toleranceRatio || 0.08}
                            onChange={(e) => setMasterSpec(prev => ({ ...prev, toleranceRatio: parseFloat(e.target.value) }))}
                            style={{ width: '100%', accentColor: '#f97316' }}
                        />
                    </div>
                </div>

                {/* ── [우측 패널]: 산출물 검수 요약 스코어보드 & 세부 결함 목록 ── */}
                <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>

                    {/* 1. 산출물 업로드 & 파일 선택기 */}
                    <div className="glass-panel" style={{
                        padding: '16px 20px',
                        borderRadius: '14px',
                        border: '1px solid var(--panel-border)',
                        display: 'flex',
                        justifyContent: 'space-between',
                        alignItems: 'center',
                        flexWrap: 'wrap',
                        gap: '12px'
                    }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                            <input
                                ref={targetFileInputRef}
                                type="file"
                                multiple
                                accept=".pptx"
                                style={{ display: 'none' }}
                                onChange={(e) => handleTargetFilesSelected(e.target.files)}
                            />
                            <button
                                onClick={() => targetFileInputRef.current?.click()}
                                style={{
                                    padding: '8px 16px',
                                    borderRadius: '8px',
                                    border: 'none',
                                    background: 'linear-gradient(135deg, #f97316, #ea580c)',
                                    color: '#fff',
                                    fontSize: '12.5px',
                                    fontWeight: 700,
                                    cursor: 'pointer',
                                    display: 'flex',
                                    alignItems: 'center',
                                    gap: '6px',
                                    boxShadow: '0 4px 12px rgba(249, 115, 22, 0.3)'
                                }}
                            >
                                <Upload size={14} /> 검수 대상 PPTX 업로드 (다중 지원)
                            </button>

                            {targetFiles.length > 0 && (
                                <select
                                    value={activeTargetIndex}
                                    onChange={(e) => setActiveTargetIndex(Number(e.target.value))}
                                    style={{
                                        padding: '7px 12px',
                                        borderRadius: '8px',
                                        border: '1px solid var(--panel-border)',
                                        background: '#1e1b4b',
                                        color: '#fff',
                                        fontSize: '12px',
                                        maxWidth: '320px'
                                    }}
                                >
                                    {targetFiles.map((f, i) => (
                                        <option key={i} value={i}>
                                            #{i + 1} {f.name} ({(f.size / 1024 / 1024).toFixed(1)}MB)
                                        </option>
                                    ))}
                                </select>
                            )}
                        </div>

                        {/* 작업 액션 버튼군 */}
                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                            {currentResult && (
                                <>
                                    <button
                                        onClick={handleReinspect}
                                        disabled={isInspecting}
                                        style={{
                                            padding: '8px 12px',
                                            borderRadius: '8px',
                                            border: '1px solid rgba(255,255,255,0.15)',
                                            background: 'rgba(255,255,255,0.05)',
                                            color: 'var(--text-primary)',
                                            fontSize: '12px',
                                            cursor: 'pointer',
                                            display: 'flex',
                                            alignItems: 'center',
                                            gap: '5px'
                                        }}
                                    >
                                        <RefreshCw size={13} /> 재검수
                                    </button>

                                    <button
                                        onClick={handleExportExcel}
                                        disabled={isExportingExcel}
                                        style={{
                                            padding: '8px 14px',
                                            borderRadius: '8px',
                                            border: '1px solid rgba(16, 185, 129, 0.4)',
                                            background: 'rgba(16, 185, 129, 0.12)',
                                            color: '#34d399',
                                            fontSize: '12px',
                                            fontWeight: 600,
                                            cursor: 'pointer',
                                            display: 'flex',
                                            alignItems: 'center',
                                            gap: '6px'
                                        }}
                                    >
                                        <FileSpreadsheet size={13} /> 리포트 (.xlsx)
                                    </button>

                                    <button
                                        onClick={handleAutoFix}
                                        disabled={isFixing}
                                        style={{
                                            padding: '8px 16px',
                                            borderRadius: '8px',
                                            border: 'none',
                                            background: 'linear-gradient(135deg, #10b981, #059669)',
                                            color: '#fff',
                                            fontSize: '12.5px',
                                            fontWeight: 700,
                                            cursor: isFixing ? 'not-allowed' : 'pointer',
                                            display: 'flex',
                                            alignItems: 'center',
                                            gap: '6px',
                                            boxShadow: '0 4px 14px rgba(16, 185, 129, 0.35)'
                                        }}
                                    >
                                        <Sparkles size={14} /> {isFixing ? '서식 교정 중...' : '원클릭 표준 서식 자동 교정 (.pptx)'}
                                    </button>
                                </>
                            )}
                        </div>
                    </div>

                    {/* 2. 요약 스코어보드 카드 */}
                    {currentResult && (
                        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))', gap: '12px' }}>
                            <div className="glass-panel" style={{
                                padding: '14px 16px',
                                borderRadius: '12px',
                                border: '1px solid var(--panel-border)',
                                background: currentResult.complianceScore >= 80 ? 'rgba(16, 185, 129, 0.08)' : 'rgba(239, 68, 68, 0.08)'
                            }}>
                                <span style={{ fontSize: '11px', color: 'var(--text-muted)', fontWeight: 600 }}>서식 준수율</span>
                                <div style={{
                                    fontSize: '24px',
                                    fontWeight: 800,
                                    color: currentResult.complianceScore >= 80 ? '#34d399' : '#f87171',
                                    marginTop: '4px'
                                }}>
                                    {currentResult.complianceScore} <span style={{ fontSize: '14px' }}>점</span>
                                </div>
                            </div>

                            <div className="glass-panel" style={{ padding: '14px 16px', borderRadius: '12px', border: '1px solid var(--panel-border)' }}>
                                <span style={{ fontSize: '11px', color: 'var(--text-muted)', fontWeight: 600 }}>총 위반 건수</span>
                                <div style={{ fontSize: '22px', fontWeight: 800, color: '#f87171', marginTop: '4px' }}>
                                    {currentResult.violationsSummary.totalViolations} <span style={{ fontSize: '12px', fontWeight: 500, color: 'var(--text-muted)' }}>건</span>
                                </div>
                            </div>

                            <div className="glass-panel" style={{ padding: '14px 16px', borderRadius: '12px', border: '1px solid var(--panel-border)' }}>
                                <span style={{ fontSize: '11px', color: 'var(--text-muted)', fontWeight: 600 }}>비표준 폰트 위반</span>
                                <div style={{ fontSize: '22px', fontWeight: 800, color: '#fb923c', marginTop: '4px' }}>
                                    {currentResult.violationsSummary.fontViolations} <span style={{ fontSize: '12px', fontWeight: 500, color: 'var(--text-muted)' }}>건</span>
                                </div>
                            </div>

                            <div className="glass-panel" style={{ padding: '14px 16px', borderRadius: '12px', border: '1px solid var(--panel-border)' }}>
                                <span style={{ fontSize: '11px', color: 'var(--text-muted)', fontWeight: 600 }}>좌표 / 크기 이탈</span>
                                <div style={{ fontSize: '22px', fontWeight: 800, color: '#60a5fa', marginTop: '4px' }}>
                                    {currentResult.violationsSummary.layoutOffsetErrors} <span style={{ fontSize: '12px', fontWeight: 500, color: 'var(--text-muted)' }}>건</span>
                                </div>
                            </div>

                            <div className="glass-panel" style={{ padding: '14px 16px', borderRadius: '12px', border: '1px solid var(--panel-border)' }}>
                                <span style={{ fontSize: '11px', color: 'var(--text-muted)', fontWeight: 600 }}>출처 / 브레드크럼</span>
                                <div style={{ fontSize: '22px', fontWeight: 800, color: '#a78bfa', marginTop: '4px' }}>
                                    {currentResult.violationsSummary.citationErrors + currentResult.violationsSummary.headerBreadcrumbErrors} <span style={{ fontSize: '12px', fontWeight: 500, color: 'var(--text-muted)' }}>건</span>
                                </div>
                            </div>
                        </div>
                    )}

                    {/* 3. 자동 교정 적용 옵션 체크박스 바 */}
                    {currentResult && (
                        <div className="glass-panel" style={{
                            padding: '12px 16px',
                            borderRadius: '10px',
                            border: '1px solid var(--panel-border)',
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'space-between',
                            flexWrap: 'wrap',
                            gap: '12px',
                            background: 'rgba(255, 255, 255, 0.02)'
                        }}>
                            <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--text-secondary)' }}>
                                🛠 자동 교정 적용 항목:
                            </span>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '14px', flexWrap: 'wrap' }}>
                                <label style={{ display: 'flex', alignItems: 'center', gap: '5px', fontSize: '12px', cursor: 'pointer', color: 'var(--text-primary)' }}>
                                    <input
                                        type="checkbox"
                                        checked={fixOptions.fixFonts}
                                        onChange={(e) => setFixOptions(prev => ({ ...prev, fixFonts: e.target.checked }))}
                                    />
                                    <span>비표준 폰트 일괄 교정</span>
                                </label>

                                <label style={{ display: 'flex', alignItems: 'center', gap: '5px', fontSize: '12px', cursor: 'pointer', color: 'var(--text-primary)' }}>
                                    <input
                                        type="checkbox"
                                        checked={fixOptions.fixCitations}
                                        onChange={(e) => setFixOptions(prev => ({ ...prev, fixCitations: e.target.checked }))}
                                    />
                                    <span>‘※ 출처 : ’ 표기 일괄 교정</span>
                                </label>

                                <label style={{ display: 'flex', alignItems: 'center', gap: '5px', fontSize: '12px', cursor: 'pointer', color: 'var(--text-primary)' }}>
                                    <input
                                        type="checkbox"
                                        checked={fixOptions.fixPositions}
                                        onChange={(e) => setFixOptions(prev => ({ ...prev, fixPositions: e.target.checked }))}
                                    />
                                    <span>제목/헤드메시지 좌표 스냅</span>
                                </label>

                                <label style={{ display: 'flex', alignItems: 'center', gap: '5px', fontSize: '12px', cursor: 'pointer', color: 'var(--text-primary)' }}>
                                    <input
                                        type="checkbox"
                                        checked={fixOptions.fixBreadcrumbDots}
                                        onChange={(e) => setFixOptions(prev => ({ ...prev, fixBreadcrumbDots: e.target.checked }))}
                                    />
                                    <span>섹션 번호 마침표 통일</span>
                                </label>
                            </div>
                        </div>
                    )}

                    {/* 4. 위반 분류 탭 & 상세 테이블 */}
                    {currentResult ? (
                        <div className="glass-panel" style={{
                            padding: '16px 20px',
                            borderRadius: '14px',
                            border: '1px solid var(--panel-border)'
                        }}>
                            {/* 필터 탭 네비게이션 */}
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '8px', marginBottom: '14px' }}>
                                <div style={{ display: 'flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap' }}>
                                    {[
                                        { key: 'ALL', label: '전체 위반', count: currentResult.violations.length },
                                        { key: 'FONT_VIOLATION', label: '비표준 폰트', count: currentResult.violationsSummary.fontViolations },
                                        { key: 'LAYOUT_OFFSET_ERROR', label: '좌표/크기 이탈', count: currentResult.violationsSummary.layoutOffsetErrors },
                                        { key: 'HEADER_BREADCRUMB_ERROR', label: '섹션/마침표 오류', count: currentResult.violationsSummary.headerBreadcrumbErrors },
                                        { key: 'CITATION_FORMAT_ERROR', label: '출처 표기 불일치', count: currentResult.violationsSummary.citationErrors },
                                        { key: 'KEY_FRAME_MISSING', label: '필수 프레임 누락', count: currentResult.violationsSummary.keyFrameMissing },
                                        { key: 'TABLE_STYLE_VIOLATION', label: '표 서식', count: currentResult.violationsSummary.tableStyleViolations }
                                    ].map(tab => (
                                        <button
                                            key={tab.key}
                                            onClick={() => setActiveViolationTab(tab.key)}
                                            style={{
                                                padding: '6px 10px',
                                                borderRadius: '7px',
                                                border: activeViolationTab === tab.key ? '1px solid #f97316' : '1px solid transparent',
                                                background: activeViolationTab === tab.key ? 'rgba(249, 115, 22, 0.15)' : 'rgba(255,255,255,0.03)',
                                                color: activeViolationTab === tab.key ? '#fb923c' : 'var(--text-secondary)',
                                                fontSize: '11.5px',
                                                fontWeight: activeViolationTab === tab.key ? 700 : 500,
                                                cursor: 'pointer',
                                                display: 'flex',
                                                alignItems: 'center',
                                                gap: '5px'
                                            }}
                                        >
                                            {tab.label}
                                            <span style={{
                                                fontSize: '10.5px',
                                                padding: '1px 5px',
                                                borderRadius: '10px',
                                                background: activeViolationTab === tab.key ? '#f97316' : 'rgba(255,255,255,0.08)',
                                                color: '#fff'
                                            }}>
                                                {tab.count}
                                            </span>
                                        </button>
                                    ))}
                                </div>

                                <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                                    <Filter size={12} color="var(--text-muted)" />
                                    <select
                                        value={severityFilter}
                                        onChange={(e) => setSeverityFilter(e.target.value)}
                                        style={{
                                            padding: '4px 8px',
                                            borderRadius: '6px',
                                            border: '1px solid var(--panel-border)',
                                            background: '#1e1b4b',
                                            color: '#fff',
                                            fontSize: '11.5px'
                                        }}
                                    >
                                        <option value="ALL">심각도 전체</option>
                                        <option value="HIGH">HIGH (중대)</option>
                                        <option value="MEDIUM">MEDIUM (보통)</option>
                                        <option value="LOW">LOW (경미)</option>
                                    </select>
                                </div>
                            </div>

                            {/* 슬라이드별 상세 위반 테이블 */}
                            {filteredViolations.length === 0 ? (
                                <div style={{ padding: '36px', textAlign: 'center', color: '#10b981', fontSize: '13px', fontWeight: 600 }}>
                                    ✓ 선택한 조건에 해당하는 서식 위반 사항이 없습니다.
                                </div>
                            ) : (
                                <div style={{ overflowX: 'auto', maxHeight: '480px' }}>
                                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px', textAlign: 'left' }}>
                                        <thead>
                                            <tr style={{ background: 'rgba(255, 255, 255, 0.04)', borderBottom: '1px solid var(--panel-border)', color: 'var(--text-secondary)' }}>
                                                <th style={{ padding: '10px 10px', width: '50px', textAlign: 'center' }}>슬라이드</th>
                                                <th style={{ padding: '10px 10px', width: '120px' }}>위반 유형</th>
                                                <th style={{ padding: '10px 10px', width: '60px', textAlign: 'center' }}>심각도</th>
                                                <th style={{ padding: '10px 10px', width: '25%' }}>현재 기재값 (발췌)</th>
                                                <th style={{ padding: '10px 10px', width: '22%' }}>표준 기준 규격</th>
                                                <th style={{ padding: '10px 10px' }}>세부 판정 사유</th>
                                                <th style={{ padding: '10px 10px', width: '70px', textAlign: 'center' }}>교정여부</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {filteredViolations.map((v) => (
                                                <tr key={v.id} style={{ borderBottom: '1px solid rgba(255, 255, 255, 0.05)' }}>
                                                    <td style={{ padding: '8px 10px', textAlign: 'center' }}>
                                                        <span style={{
                                                            padding: '2px 6px',
                                                            borderRadius: '4px',
                                                            background: 'rgba(255, 255, 255, 0.08)',
                                                            fontSize: '11px',
                                                            fontWeight: 700,
                                                            color: '#38bdf8'
                                                        }}>
                                                            p.{v.slideNum}
                                                        </span>
                                                    </td>
                                                    <td style={{ padding: '8px 10px' }}>
                                                        <span style={{
                                                            fontSize: '11px',
                                                            fontWeight: 600,
                                                            color: v.type === 'FONT_VIOLATION' ? '#fb923c' :
                                                                   v.type === 'LAYOUT_OFFSET_ERROR' ? '#60a5fa' :
                                                                   v.type === 'CITATION_FORMAT_ERROR' ? '#a78bfa' : '#f87171'
                                                        }}>
                                                            {v.typeLabel}
                                                        </span>
                                                    </td>
                                                    <td style={{ padding: '8px 10px', textAlign: 'center' }}>
                                                        <span style={{
                                                            padding: '2px 6px',
                                                            borderRadius: '4px',
                                                            fontSize: '10.5px',
                                                            fontWeight: 800,
                                                            background: v.severity === 'HIGH' ? 'rgba(239, 68, 68, 0.2)' :
                                                                       v.severity === 'MEDIUM' ? 'rgba(251, 146, 60, 0.2)' : 'rgba(59, 130, 246, 0.2)',
                                                            color: v.severity === 'HIGH' ? '#f87171' :
                                                                   v.severity === 'MEDIUM' ? '#fb923c' : '#60a5fa'
                                                        }}>
                                                            {v.severity}
                                                        </span>
                                                    </td>
                                                    <td style={{ padding: '8px 10px' }}>
                                                        <span style={{ color: '#fca5a5', textDecoration: v.canAutoFix ? 'line-through' : 'none' }}>
                                                            {v.currentValue}
                                                        </span>
                                                        {v.snippet && (
                                                            <div style={{ fontSize: '10.5px', color: 'var(--text-muted)', marginTop: '2px' }}>
                                                                "{v.snippet}"
                                                            </div>
                                                        )}
                                                    </td>
                                                    <td style={{ padding: '8px 10px' }}>
                                                        <span style={{ color: '#86efac', fontWeight: 600 }}>
                                                            {v.standardValue}
                                                        </span>
                                                    </td>
                                                    <td style={{ padding: '8px 10px', color: 'var(--text-secondary)', fontSize: '11px', lineHeight: 1.4 }}>
                                                        {v.reason}
                                                    </td>
                                                    <td style={{ padding: '8px 10px', textAlign: 'center' }}>
                                                        {v.canAutoFix ? (
                                                            <span style={{
                                                                display: 'inline-flex',
                                                                alignItems: 'center',
                                                                gap: '3px',
                                                                fontSize: '10.5px',
                                                                fontWeight: 700,
                                                                color: '#34d399'
                                                            }}>
                                                                <Check size={11} /> 자동
                                                            </span>
                                                        ) : (
                                                            <span style={{ fontSize: '10.5px', color: 'var(--text-muted)' }}>수동</span>
                                                        )}
                                                    </td>
                                                </tr>
                                            ))}
                                        </tbody>
                                    </table>
                                </div>
                            )}
                        </div>
                    ) : (
                        /* 검수 전 안내 플레이스홀더 */
                        <div className="glass-panel" style={{
                            padding: '48px 20px',
                            borderRadius: '14px',
                            border: '1px dashed var(--panel-border)',
                            textAlign: 'center',
                            color: 'var(--text-muted)'
                        }}>
                            <Presentation size={36} color="var(--text-muted)" style={{ margin: '0 auto 10px' }} />
                            <div style={{ fontSize: '14px', fontWeight: 600, color: 'var(--text-secondary)' }}>
                                검수할 PPTX 산출물 파일을 상단에서 업로드해 주세요
                            </div>
                            <div style={{ fontSize: '12px', marginTop: '4px' }}>
                                기준 마스터 규격과 1:1로 대조하여 비표준 폰트, 좌표 이탈, 출처 포맷 오류를 실시간 적발합니다.
                            </div>
                        </div>
                    )}
                </div>
            </div>
        </div>
    );
}
