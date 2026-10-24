import React, { useState, useEffect, useMemo, useRef } from 'react';
import {
    MessagesSquare,
    Upload,
    Download,
    Plus,
    Trash2,
    RefreshCw,
    Sparkles,
    AlertTriangle,
    CheckCircle2,
    FileSpreadsheet,
    FileText,
    Presentation,
    Layers,
    Search,
    Eye,
    X,
    Filter,
    ShieldAlert,
    Copy,
    ArrowRight,
    ExternalLink,
    FileCheck2,
    HelpCircle
} from 'lucide-react';
import { saveAs } from 'file-saver';
import {
    parseInterviewFileMeta,
    extractMultipleInterviewFiles,
    runInterviewCascadePipeline,
    syncToSSOTRequirements,
    generateInterviewExcel,
    generateInterviewPptx,
    generateInterviewDocx,
    generateInterviewCascadeZip,
    DEFAULT_SAMPLE_INTERVIEWS,
    DEFAULT_SAMPLE_CASCADE_DATA
} from '../utils/interviewPipeline';
import { loadDesignTokens } from '../utils/designTokens';

export default function InterviewCascadeView({ apiKey, llmProvider, omniRouteModel }) {
    // ── 상태 관리 ───────────────────────────────────────────────────
    // 업로드된 인터뷰 파일 목록
    const [interviewFiles, setInterviewFiles] = useState([]);
    const [isParsingFiles, setIsParsingFiles] = useState(false);
    const [isDragging, setIsDragging] = useState(false);

    // AI 연쇄 추출 진행 상태
    const [isExtracting, setIsExtracting] = useState(false);
    const [progress, setProgress] = useState({ percent: 0, message: '' });

    // 연쇄 추출된 4개 영역 데이터
    const [cascadeData, setCascadeData] = useState(() => {
        const saved = localStorage.getItem('healthy_project_cascade_data');
        if (saved) {
            try { return JSON.parse(saved); } catch (_) {}
        }
        return DEFAULT_SAMPLE_CASCADE_DATA;
    });

    // 활성 탭: 'summary' | 'issues' | 'directions' | 'requirements'
    const [activeTab, setActiveTab] = useState('issues');

    // 검색/필터 텍스트
    const [searchKeyword, setSearchKeyword] = useState('');

    // 원문 미리보기 모달
    const [previewDoc, setPreviewDoc] = useState(null);

    // 파일 내보내기 진행 상태
    const [isExporting, setIsExporting] = useState(false);

    // 알림 토스트 상태
    const [toastMessage, setToastMessage] = useState(null);

    const fileInputRef = useRef(null);

    // cascadeData 변경 시 LocalStorage 영속화
    useEffect(() => {
        try {
            localStorage.setItem('healthy_project_cascade_data', JSON.stringify(cascadeData));
        } catch (_) {}
    }, [cascadeData]);

    const showToast = (msg, type = 'success') => {
        setToastMessage({ msg, type });
        setTimeout(() => setToastMessage(null), 4000);
    };

    // ── 1. 파일 업로드 및 파싱 핸들러 ──────────────────────────────
    const handleFilesSelected = async (files) => {
        if (!files || files.length === 0) return;
        setIsParsingFiles(true);
        try {
            const extracted = await extractMultipleInterviewFiles(Array.from(files));
            setInterviewFiles(prev => [...prev, ...extracted]);
            showToast(`${extracted.length}개의 회의록 파일을 성공적으로 분석 및 파싱했습니다.`);
        } catch (err) {
            console.error('파일 파싱 오류:', err);
            showToast(`파일 파싱 중 오류가 발생했습니다: ${err.message}`, 'error');
        } finally {
            setIsParsingFiles(false);
        }
    };

    const handleDrop = (e) => {
        e.preventDefault();
        setIsDragging(false);
        if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
            handleFilesSelected(e.dataTransfer.files);
        }
    };

    const handleRemoveFile = (id) => {
        setInterviewFiles(prev => prev.filter(f => f.id !== id));
    };

    const handleUpdateFileMeta = (id, field, value) => {
        setInterviewFiles(prev => prev.map(f => {
            if (f.id === id) {
                return {
                    ...f,
                    meta: { ...f.meta, [field]: value }
                };
            }
            return f;
        }));
    };

    // 샘플 데이터 로드 (체험용)
    const handleLoadSamples = () => {
        setInterviewFiles(DEFAULT_SAMPLE_INTERVIEWS);
        setCascadeData(DEFAULT_SAMPLE_CASCADE_DATA);
        showToast('공공의료/ISP 인터뷰 회의록 3종 및 연쇄 추출 결과 샘플을 로드했습니다.');
    };

    // ── 2. AI 1-Click 연쇄 데이터 추출 실행 ────────────────────────
    const handleRunCascade = async () => {
        if (interviewFiles.length === 0) {
            showToast('분석할 인터뷰 회의록 파일을 먼저 업로드하거나 샘플을 불러와주세요.', 'error');
            return;
        }

        const validKey = (apiKey || '').split(',').map(k => k.trim()).find(k => k.startsWith('AIza') || k.startsWith('AQ.'));
        if (!validKey) {
            showToast('Gemini API 키가 필요합니다. 상단 헤더의 설정창에서 API 키를 입력해 주세요.', 'error');
            return;
        }

        setIsExtracting(true);
        setProgress({ percent: 10, message: '인터뷰 회의록 텍스트 취합 중...' });

        try {
            const result = await runInterviewCascadePipeline({
                interviewFiles,
                apiKey: validKey,
                modelId: omniRouteModel && omniRouteModel !== 'auto' ? omniRouteModel : 'gemini-2.5-flash',
                onProgress: (p) => setProgress(p)
            });

            setCascadeData(result);
            showToast('인터뷰 4대 연쇄 데이터(요약/As-Is/To-Be/요구사항) 추출이 성공적으로 완료되었습니다!');
            setActiveTab('issues');
        } catch (err) {
            console.error('연쇄 추출 오류:', err);
            showToast(`연쇄 추출 실패: ${err.message}`, 'error');
        } finally {
            setIsExtracting(false);
            setProgress({ percent: 0, message: '' });
        }
    };

    // ── 3. 1순위 SSOT Validator 요구사항 마스터로 1클릭 동기화 ──────
    const handleSyncToSSOT = () => {
        if (!cascadeData.derivedRequirements || cascadeData.derivedRequirements.length === 0) {
            showToast('동기화할 도출 요구사항이 없습니다.', 'error');
            return;
        }

        const addedCount = syncToSSOTRequirements(cascadeData.derivedRequirements);
        if (addedCount > 0) {
            showToast(`성공: ${addedCount}개의 신규 요구사항이 [1순위 SSOT 검증기] 마스터 리스트에 등록되었습니다!`);
        } else {
            showToast('모든 요구사항이 이미 SSOT 마스터에 등록되어 있습니다 (중복 제외).', 'info');
        }
    };

    // ── 4. 3종 산출물 개별 및 ZIP 일괄 다운로드 ──────────────────────
    const handleDownloadZip = async () => {
        setIsExporting(true);
        try {
            const tokens = loadDesignTokens();
            await generateInterviewCascadeZip(cascadeData, tokens);
            showToast('3종 산출물(Excel, PPTX, Word) 일괄 패키징 ZIP 다운로드가 완료되었습니다!');
        } catch (err) {
            console.error('ZIP 내보내기 오류:', err);
            showToast(`ZIP 내보내기 오류: ${err.message}`, 'error');
        } finally {
            setIsExporting(false);
        }
    };

    const handleDownloadExcel = async () => {
        setIsExporting(true);
        try {
            const tokens = loadDesignTokens();
            const blob = await generateInterviewExcel(cascadeData, tokens);
            saveAs(blob, `01_인터뷰결과_및_이슈도출_종합표_${new Date().toISOString().slice(0, 10)}.xlsx`);
            showToast('엑셀 종합표(.xlsx) 다운로드가 완료되었습니다!');
        } catch (err) {
            console.error('엑셀 오류:', err);
            showToast(`엑셀 생성 오류: ${err.message}`, 'error');
        } finally {
            setIsExporting(false);
        }
    };

    const handleDownloadPptx = async () => {
        setIsExporting(true);
        try {
            const tokens = loadDesignTokens();
            const blob = await generateInterviewPptx(cascadeData, tokens);
            saveAs(blob, `02_현황분석_주요문제점_및_시사점_보고서장표_${new Date().toISOString().slice(0, 10)}.pptx`);
            showToast('PPTX 보고서 장표(.pptx) 다운로드가 완료되었습니다!');
        } catch (err) {
            console.error('PPTX 오류:', err);
            showToast(`PPTX 생성 오류: ${err.message}`, 'error');
        } finally {
            setIsExporting(false);
        }
    };

    const handleDownloadDocx = async () => {
        setIsExporting(true);
        try {
            const tokens = loadDesignTokens();
            const blob = await generateInterviewDocx(cascadeData, tokens);
            saveAs(blob, `03_요구사항정의서_초안_${new Date().toISOString().slice(0, 10)}.docx`);
            showToast('Word 요구사항정의서(.docx) 다운로드가 완료되었습니다!');
        } catch (err) {
            console.error('Word 오류:', err);
            showToast(`Word 생성 오류: ${err.message}`, 'error');
        } finally {
            setIsExporting(false);
        }
    };

    // ── 5. 테이블 데이터 인라인 수정 / 행 추가 / 삭제 ─────────────
    const handleAddRow = () => {
        if (activeTab === 'summary') {
            const newRow = {
                sourceFile: '수동 추가',
                date: new Date().toISOString().slice(0, 10).replace(/-/g, '.'),
                department: '신규 부서',
                interviewee: '면담자명',
                summary: '핵심 논의 사항 요약',
                actionItems: '후속 조치 내용'
            };
            setCascadeData(prev => ({ ...prev, interviewSummary: [newRow, ...prev.interviewSummary] }));
        } else if (activeTab === 'issues') {
            const nextIdx = (cascadeData.asIsIssues.length + 1).toString().padStart(2, '0');
            const newRow = {
                id: `ISSUE-${nextIdx}`,
                category: '응용시스템',
                targetSystem: '대상 업무/시스템명',
                asIsStatus: '현행 운영 현황',
                painPoint: '핵심 문제점 및 Pain Point',
                quote: '원문 발언 인용',
                sourceFile: '수동 입력'
            };
            setCascadeData(prev => ({ ...prev, asIsIssues: [newRow, ...prev.asIsIssues] }));
        } else if (activeTab === 'directions') {
            const newRow = {
                issueId: 'ISSUE-01',
                implication: '도출된 핵심 시사점',
                toBeDirection: 'To-Be 목표 개선 방향',
                relatedProject: '연관 정보화 과제명'
            };
            setCascadeData(prev => ({ ...prev, toBeDirections: [newRow, ...prev.toBeDirections] }));
        } else if (activeTab === 'requirements') {
            const nextIdx = (cascadeData.derivedRequirements.length + 1).toString().padStart(3, '0');
            const newRow = {
                id: `SFR-${nextIdx}`,
                name: '신규 기능 요구사항',
                type: '기능',
                description: '구체적인 요구사항 명세 설명',
                priority: '상',
                sourceFile: '수동 추가'
            };
            setCascadeData(prev => ({ ...prev, derivedRequirements: [newRow, ...prev.derivedRequirements] }));
        }
    };

    const handleDeleteRow = (tabKey, index) => {
        setCascadeData(prev => ({
            ...prev,
            [tabKey]: prev[tabKey].filter((_, idx) => idx !== index)
        }));
    };

    const handleUpdateCell = (tabKey, index, field, value) => {
        setCascadeData(prev => {
            const list = [...prev[tabKey]];
            list[index] = { ...list[index], [field]: value };
            return { ...prev, [tabKey]: list };
        });
    };

    // ── 6. 필터링된 테이블 데이터 ───────────────────────────────────
    const filteredSummary = useMemo(() => {
        if (!searchKeyword.trim()) return cascadeData.interviewSummary;
        const kw = searchKeyword.toLowerCase();
        return cascadeData.interviewSummary.filter(r =>
            (r.department || '').toLowerCase().includes(kw) ||
            (r.interviewee || '').toLowerCase().includes(kw) ||
            (r.summary || '').toLowerCase().includes(kw) ||
            (r.actionItems || '').toLowerCase().includes(kw)
        );
    }, [cascadeData.interviewSummary, searchKeyword]);

    const filteredIssues = useMemo(() => {
        if (!searchKeyword.trim()) return cascadeData.asIsIssues;
        const kw = searchKeyword.toLowerCase();
        return cascadeData.asIsIssues.filter(r =>
            (r.id || '').toLowerCase().includes(kw) ||
            (r.category || '').toLowerCase().includes(kw) ||
            (r.targetSystem || '').toLowerCase().includes(kw) ||
            (r.painPoint || '').toLowerCase().includes(kw) ||
            (r.quote || '').toLowerCase().includes(kw)
        );
    }, [cascadeData.asIsIssues, searchKeyword]);

    const filteredDirections = useMemo(() => {
        if (!searchKeyword.trim()) return cascadeData.toBeDirections;
        const kw = searchKeyword.toLowerCase();
        return cascadeData.toBeDirections.filter(r =>
            (r.issueId || '').toLowerCase().includes(kw) ||
            (r.implication || '').toLowerCase().includes(kw) ||
            (r.toBeDirection || '').toLowerCase().includes(kw) ||
            (r.relatedProject || '').toLowerCase().includes(kw)
        );
    }, [cascadeData.toBeDirections, searchKeyword]);

    const filteredRequirements = useMemo(() => {
        if (!searchKeyword.trim()) return cascadeData.derivedRequirements;
        const kw = searchKeyword.toLowerCase();
        return cascadeData.derivedRequirements.filter(r =>
            (r.id || '').toLowerCase().includes(kw) ||
            (r.name || '').toLowerCase().includes(kw) ||
            (r.type || '').toLowerCase().includes(kw) ||
            (r.description || '').toLowerCase().includes(kw)
        );
    }, [cascadeData.derivedRequirements, searchKeyword]);

    const tokens = loadDesignTokens();

    return (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '24px', width: '100%' }}>
            {/* ── 토스트 알림 ────────────────────────────────────────── */}
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

            {/* ── 헤더 패널 ────────────────────────────────────────── */}
            <div className="glass-panel" style={{
                padding: '24px',
                borderRadius: '16px',
                border: '1px solid var(--panel-border)',
                background: 'linear-gradient(135deg, rgba(139, 92, 246, 0.08) 0%, rgba(59, 130, 246, 0.04) 100%)',
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
                            background: 'linear-gradient(135deg, #8b5cf6, #3b82f6)',
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                            color: '#fff',
                            boxShadow: '0 4px 12px rgba(139, 92, 246, 0.3)'
                        }}>
                            <MessagesSquare size={22} />
                        </div>
                        <h2 style={{ margin: 0, fontSize: '20px', fontWeight: 800, color: 'var(--text-primary)' }}>
                            인터뷰·회의록 기반 연쇄 산출물 자동 추출 파이프라인 (OSMU Engine)
                        </h2>
                        <span style={{
                            padding: '3px 10px',
                            borderRadius: '20px',
                            background: 'rgba(139, 92, 246, 0.2)',
                            color: '#a78bfa',
                            fontSize: '11px',
                            fontWeight: 700,
                            border: '1px solid rgba(139, 92, 246, 0.4)'
                        }}>
                            3순위 OSMU
                        </span>
                    </div>
                    <p style={{ margin: 0, fontSize: '13px', color: 'var(--text-secondary)', lineHeight: 1.6, maxWidth: '850px' }}>
                        다중 인터뷰/회의록 원문(<strong>.hwpx, .docx, .txt, .md</strong>)을 일괄 파싱하여 메타데이터를 자동 식별하고,
                        AI가 <strong>①인터뷰요약 ➔ ②As-Is문제점 ➔ ③To-Be시사점 ➔ ④요구사항명세</strong>를 연쇄 도출합니다.
                        도출된 데이터는 <strong>엑셀 종합표·PPTX 비교장표·Word 요구사항서</strong> 3종 세트로 일괄 패키징(ZIP) 다운로드됩니다.
                    </p>
                </div>

                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
                    <button
                        onClick={handleLoadSamples}
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
                        <RefreshCw size={13} /> 샘플 회의록 3종 로드
                    </button>

                    <button
                        onClick={handleDownloadZip}
                        disabled={isExporting}
                        style={{
                            padding: '8px 16px',
                            borderRadius: '9px',
                            border: '1px solid rgba(139, 92, 246, 0.5)',
                            background: 'linear-gradient(135deg, #8b5cf6, #6366f1)',
                            color: '#fff',
                            fontSize: '12.5px',
                            fontWeight: 700,
                            cursor: isExporting ? 'not-allowed' : 'pointer',
                            display: 'flex',
                            alignItems: 'center',
                            gap: '7px',
                            boxShadow: '0 4px 14px rgba(139, 92, 246, 0.35)'
                        }}
                    >
                        <Download size={14} />
                        {isExporting ? '패키징 생성 중...' : '연쇄 산출물 일괄 다운로드 (ZIP)'}
                    </button>
                </div>
            </div>

            {/* ── STEP 1: 다중 회의록 파일 업로드 및 메타 관리 ─────────── */}
            <div className="glass-panel" style={{
                padding: '20px 24px',
                borderRadius: '14px',
                border: '1px solid var(--panel-border)',
                background: 'rgba(255, 255, 255, 0.02)'
            }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '14px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <span style={{
                            display: 'inline-flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                            width: '22px',
                            height: '22px',
                            borderRadius: '50%',
                            background: '#8b5cf6',
                            color: '#fff',
                            fontSize: '11px',
                            fontWeight: 800
                        }}>
                            1
                        </span>
                        <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 700, color: 'var(--text-primary)' }}>
                            다중 인터뷰/회의록 파일 업로드 & 메타데이터 자동 파싱
                        </h3>
                        <span style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>
                            (등록된 회의록: <strong>{interviewFiles.length}</strong>건)
                        </span>
                    </div>

                    <button
                        onClick={handleRunCascade}
                        disabled={isExtracting || interviewFiles.length === 0}
                        style={{
                            padding: '9px 18px',
                            borderRadius: '10px',
                            border: 'none',
                            background: interviewFiles.length === 0 ? 'rgba(255,255,255,0.08)' : 'linear-gradient(135deg, #10b981, #059669)',
                            color: interviewFiles.length === 0 ? 'var(--text-muted)' : '#fff',
                            fontSize: '13px',
                            fontWeight: 800,
                            cursor: interviewFiles.length === 0 || isExtracting ? 'not-allowed' : 'pointer',
                            display: 'flex',
                            alignItems: 'center',
                            gap: '8px',
                            boxShadow: interviewFiles.length > 0 ? '0 4px 16px rgba(16, 185, 129, 0.35)' : 'none'
                        }}
                    >
                        <Sparkles size={15} />
                        {isExtracting ? 'AI 연쇄 분석 및 도출 중...' : '1-Click 연쇄 데이터 추출 시작'}
                    </button>
                </div>

                {/* 진행률 바 (AI 작동 시) */}
                {isExtracting && (
                    <div style={{
                        marginBottom: '16px',
                        padding: '14px 16px',
                        borderRadius: '10px',
                        background: 'rgba(139, 92, 246, 0.08)',
                        border: '1px solid rgba(139, 92, 246, 0.3)'
                    }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '12px', marginBottom: '6px', fontWeight: 600 }}>
                            <span style={{ color: '#a78bfa', display: 'flex', alignItems: 'center', gap: '6px' }}>
                                <Sparkles size={13} /> {progress.message || '인터뷰 원문 분석 중...'}
                            </span>
                            <span style={{ color: 'var(--text-primary)' }}>{progress.percent}%</span>
                        </div>
                        <div style={{ width: '100%', height: '8px', background: 'rgba(255,255,255,0.06)', borderRadius: '4px', overflow: 'hidden' }}>
                            <div style={{
                                width: `${progress.percent}%`,
                                height: '100%',
                                background: 'linear-gradient(90deg, #8b5cf6, #3b82f6)',
                                transition: 'width 0.3s ease'
                            }} />
                        </div>
                    </div>
                )}

                {/* 드래그 앤 드롭 영역 */}
                <div
                    onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
                    onDragLeave={() => setIsDragging(false)}
                    onDrop={handleDrop}
                    onClick={() => fileInputRef.current?.click()}
                    style={{
                        border: `2px dashed ${isDragging ? '#8b5cf6' : 'rgba(255, 255, 255, 0.15)'}`,
                        borderRadius: '12px',
                        padding: '24px 20px',
                        textAlign: 'center',
                        background: isDragging ? 'rgba(139, 92, 246, 0.08)' : 'rgba(255, 255, 255, 0.015)',
                        cursor: 'pointer',
                        transition: 'all 0.2s ease',
                        marginBottom: interviewFiles.length > 0 ? '16px' : '0'
                    }}
                >
                    <input
                        ref={fileInputRef}
                        type="file"
                        multiple
                        accept=".hwpx,.docx,.txt,.md"
                        style={{ display: 'none' }}
                        onChange={(e) => handleFilesSelected(e.target.files)}
                    />
                    <Upload size={32} color={isDragging ? '#8b5cf6' : 'var(--text-secondary)'} style={{ margin: '0 auto 8px' }} />
                    <p style={{ margin: 0, fontSize: '13.5px', fontWeight: 600, color: 'var(--text-primary)' }}>
                        인터뷰 회의록 파일들을 이곳에 드래그하거나 클릭하여 일괄 등록하세요
                    </p>
                    <p style={{ margin: '4px 0 0', fontSize: '12px', color: 'var(--text-secondary)' }}>
                        지원 형식: <strong>.hwpx, .docx, .txt, .md</strong> (다중 선택 가능 · 파일명에서 일자/부서/면담자 자동 추출)
                    </p>
                </div>

                {/* 등록된 파일 목록 테이블 */}
                {interviewFiles.length > 0 && (
                    <div style={{ overflowX: 'auto', borderRadius: '10px', border: '1px solid var(--panel-border)' }}>
                        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12.5px', textAlign: 'left' }}>
                            <thead>
                                <tr style={{ background: 'rgba(255, 255, 255, 0.04)', borderBottom: '1px solid var(--panel-border)', color: 'var(--text-secondary)' }}>
                                    <th style={{ padding: '10px 14px', width: '30%' }}>파일명 / 포맷</th>
                                    <th style={{ padding: '10px 14px', width: '16%' }}>면담일자</th>
                                    <th style={{ padding: '10px 14px', width: '22%' }}>대상기관 / 부서</th>
                                    <th style={{ padding: '10px 14px', width: '16%' }}>면담대상(인터뷰이)</th>
                                    <th style={{ padding: '10px 14px', width: '16%', textAlign: 'center' }}>작업</th>
                                </tr>
                            </thead>
                            <tbody>
                                {interviewFiles.map((item) => (
                                    <tr key={item.id} style={{ borderBottom: '1px solid rgba(255, 255, 255, 0.05)' }}>
                                        <td style={{ padding: '10px 14px' }}>
                                            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                                <FileText size={15} color="#8b5cf6" />
                                                <span style={{ fontWeight: 600, color: 'var(--text-primary)' }}>{item.meta.filename}</span>
                                                <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>({(item.text || '').length.toLocaleString()}자)</span>
                                            </div>
                                        </td>
                                        <td style={{ padding: '8px 14px' }}>
                                            <input
                                                type="text"
                                                value={item.meta.date}
                                                onChange={(e) => handleUpdateFileMeta(item.id, 'date', e.target.value)}
                                                style={{
                                                    width: '100%',
                                                    padding: '5px 8px',
                                                    borderRadius: '6px',
                                                    border: '1px solid var(--panel-border)',
                                                    background: 'rgba(0,0,0,0.2)',
                                                    color: 'var(--text-primary)',
                                                    fontSize: '12px'
                                                }}
                                            />
                                        </td>
                                        <td style={{ padding: '8px 14px' }}>
                                            <input
                                                type="text"
                                                value={item.meta.department}
                                                onChange={(e) => handleUpdateFileMeta(item.id, 'department', e.target.value)}
                                                style={{
                                                    width: '100%',
                                                    padding: '5px 8px',
                                                    borderRadius: '6px',
                                                    border: '1px solid var(--panel-border)',
                                                    background: 'rgba(0,0,0,0.2)',
                                                    color: 'var(--text-primary)',
                                                    fontSize: '12px'
                                                }}
                                            />
                                        </td>
                                        <td style={{ padding: '8px 14px' }}>
                                            <input
                                                type="text"
                                                value={item.meta.interviewee}
                                                onChange={(e) => handleUpdateFileMeta(item.id, 'interviewee', e.target.value)}
                                                style={{
                                                    width: '100%',
                                                    padding: '5px 8px',
                                                    borderRadius: '6px',
                                                    border: '1px solid var(--panel-border)',
                                                    background: 'rgba(0,0,0,0.2)',
                                                    color: 'var(--text-primary)',
                                                    fontSize: '12px'
                                                }}
                                            />
                                        </td>
                                        <td style={{ padding: '8px 14px', textAlign: 'center' }}>
                                            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '6px' }}>
                                                <button
                                                    onClick={() => setPreviewDoc(item)}
                                                    title="본문 텍스트 미리보기"
                                                    style={{
                                                        padding: '5px 8px',
                                                        borderRadius: '6px',
                                                        border: '1px solid rgba(255, 255, 255, 0.1)',
                                                        background: 'rgba(255, 255, 255, 0.05)',
                                                        color: 'var(--text-secondary)',
                                                        cursor: 'pointer',
                                                        display: 'flex',
                                                        alignItems: 'center',
                                                        gap: '4px',
                                                        fontSize: '11px'
                                                    }}
                                                >
                                                    <Eye size={12} /> 미리보기
                                                </button>
                                                <button
                                                    onClick={() => handleRemoveFile(item.id)}
                                                    title="삭제"
                                                    style={{
                                                        padding: '5px 8px',
                                                        borderRadius: '6px',
                                                        border: '1px solid rgba(239, 68, 68, 0.2)',
                                                        background: 'rgba(239, 68, 68, 0.08)',
                                                        color: '#f87171',
                                                        cursor: 'pointer'
                                                    }}
                                                >
                                                    <Trash2 size={12} />
                                                </button>
                                            </div>
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                )}
            </div>

            {/* ── STEP 2: 4단계 연쇄 추출 그리드 에디터 ───────────────── */}
            <div className="glass-panel" style={{
                padding: '24px',
                borderRadius: '16px',
                border: '1px solid var(--panel-border)',
                background: 'rgba(255, 255, 255, 0.02)'
            }}>
                {/* 탭 네비게이션 & 액션 툴바 */}
                <div style={{
                    display: 'flex',
                    justifyContent: 'space-between',
                    alignItems: 'center',
                    flexWrap: 'wrap',
                    gap: '12px',
                    marginBottom: '18px',
                    borderBottom: '1px solid var(--panel-border)',
                    paddingBottom: '14px'
                }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
                        <button
                            onClick={() => setActiveTab('summary')}
                            style={{
                                padding: '9px 16px',
                                borderRadius: '10px',
                                border: activeTab === 'summary' ? '1px solid #8b5cf6' : '1px solid transparent',
                                background: activeTab === 'summary' ? 'rgba(139, 92, 246, 0.15)' : 'rgba(255, 255, 255, 0.04)',
                                color: activeTab === 'summary' ? '#a78bfa' : 'var(--text-secondary)',
                                fontSize: '13px',
                                fontWeight: activeTab === 'summary' ? 700 : 500,
                                cursor: 'pointer',
                                display: 'flex',
                                alignItems: 'center',
                                gap: '8px'
                            }}
                        >
                            📋 인터뷰 결과 요약
                            <span style={{
                                padding: '2px 7px',
                                borderRadius: '12px',
                                background: activeTab === 'summary' ? '#8b5cf6' : 'rgba(255,255,255,0.1)',
                                color: '#fff',
                                fontSize: '11px',
                                fontWeight: 700
                            }}>
                                {cascadeData.interviewSummary.length}
                            </span>
                        </button>

                        <button
                            onClick={() => setActiveTab('issues')}
                            style={{
                                padding: '9px 16px',
                                borderRadius: '10px',
                                border: activeTab === 'issues' ? '1px solid #ef4444' : '1px solid transparent',
                                background: activeTab === 'issues' ? 'rgba(239, 68, 68, 0.15)' : 'rgba(255, 255, 255, 0.04)',
                                color: activeTab === 'issues' ? '#f87171' : 'var(--text-secondary)',
                                fontSize: '13px',
                                fontWeight: activeTab === 'issues' ? 700 : 500,
                                cursor: 'pointer',
                                display: 'flex',
                                alignItems: 'center',
                                gap: '8px'
                            }}
                        >
                            🚨 현황 및 문제점 (As-Is)
                            <span style={{
                                padding: '2px 7px',
                                borderRadius: '12px',
                                background: activeTab === 'issues' ? '#ef4444' : 'rgba(255,255,255,0.1)',
                                color: '#fff',
                                fontSize: '11px',
                                fontWeight: 700
                            }}>
                                {cascadeData.asIsIssues.length}
                            </span>
                        </button>

                        <button
                            onClick={() => setActiveTab('directions')}
                            style={{
                                padding: '9px 16px',
                                borderRadius: '10px',
                                border: activeTab === 'directions' ? '1px solid #10b981' : '1px solid transparent',
                                background: activeTab === 'directions' ? 'rgba(16, 185, 129, 0.15)' : 'rgba(255, 255, 255, 0.04)',
                                color: activeTab === 'directions' ? '#34d399' : 'var(--text-secondary)',
                                fontSize: '13px',
                                fontWeight: activeTab === 'directions' ? 700 : 500,
                                cursor: 'pointer',
                                display: 'flex',
                                alignItems: 'center',
                                gap: '8px'
                            }}
                        >
                            🚀 개선방향 및 시사점 (To-Be)
                            <span style={{
                                padding: '2px 7px',
                                borderRadius: '12px',
                                background: activeTab === 'directions' ? '#10b981' : 'rgba(255,255,255,0.1)',
                                color: '#fff',
                                fontSize: '11px',
                                fontWeight: 700
                            }}>
                                {cascadeData.toBeDirections.length}
                            </span>
                        </button>

                        <button
                            onClick={() => setActiveTab('requirements')}
                            style={{
                                padding: '9px 16px',
                                borderRadius: '10px',
                                border: activeTab === 'requirements' ? '1px solid #3b82f6' : '1px solid transparent',
                                background: activeTab === 'requirements' ? 'rgba(59, 130, 246, 0.15)' : 'rgba(255, 255, 255, 0.04)',
                                color: activeTab === 'requirements' ? '#60a5fa' : 'var(--text-secondary)',
                                fontSize: '13px',
                                fontWeight: activeTab === 'requirements' ? 700 : 500,
                                cursor: 'pointer',
                                display: 'flex',
                                alignItems: 'center',
                                gap: '8px'
                            }}
                        >
                            📌 도출 요구사항 (SSOT 연계)
                            <span style={{
                                padding: '2px 7px',
                                borderRadius: '12px',
                                background: activeTab === 'requirements' ? '#3b82f6' : 'rgba(255,255,255,0.1)',
                                color: '#fff',
                                fontSize: '11px',
                                fontWeight: 700
                            }}>
                                {cascadeData.derivedRequirements.length}
                            </span>
                        </button>
                    </div>

                    {/* 개별 산출물 다운로드 & SSOT 동기화 버튼군 */}
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
                        {activeTab === 'requirements' && (
                            <button
                                onClick={handleSyncToSSOT}
                                title="1순위 SSOT Validator 요구사항 마스터로 즉시 동기화"
                                style={{
                                    padding: '7px 14px',
                                    borderRadius: '8px',
                                    border: '1px solid rgba(16, 185, 129, 0.5)',
                                    background: 'rgba(16, 185, 129, 0.15)',
                                    color: '#34d399',
                                    fontSize: '12px',
                                    fontWeight: 700,
                                    cursor: 'pointer',
                                    display: 'flex',
                                    alignItems: 'center',
                                    gap: '6px'
                                }}
                            >
                                <FileCheck2 size={13} /> SSOT 마스터로 등록
                            </button>
                        )}

                        <button
                            onClick={handleDownloadExcel}
                            disabled={isExporting}
                            style={{
                                padding: '7px 12px',
                                borderRadius: '8px',
                                border: '1px solid rgba(16, 185, 129, 0.3)',
                                background: 'rgba(16, 185, 129, 0.08)',
                                color: '#34d399',
                                fontSize: '12px',
                                fontWeight: 600,
                                cursor: 'pointer',
                                display: 'flex',
                                alignItems: 'center',
                                gap: '5px'
                            }}
                        >
                            <FileSpreadsheet size={13} /> 엑셀(.xlsx)
                        </button>

                        <button
                            onClick={handleDownloadPptx}
                            disabled={isExporting}
                            style={{
                                padding: '7px 12px',
                                borderRadius: '8px',
                                border: '1px solid rgba(249, 115, 22, 0.3)',
                                background: 'rgba(249, 115, 22, 0.08)',
                                color: '#fb923c',
                                fontSize: '12px',
                                fontWeight: 600,
                                cursor: 'pointer',
                                display: 'flex',
                                alignItems: 'center',
                                gap: '5px'
                            }}
                        >
                            <Presentation size={13} /> PPTX 비교장표
                        </button>

                        <button
                            onClick={handleDownloadDocx}
                            disabled={isExporting}
                            style={{
                                padding: '7px 12px',
                                borderRadius: '8px',
                                border: '1px solid rgba(59, 130, 246, 0.3)',
                                background: 'rgba(59, 130, 246, 0.08)',
                                color: '#60a5fa',
                                fontSize: '12px',
                                fontWeight: 600,
                                cursor: 'pointer',
                                display: 'flex',
                                alignItems: 'center',
                                gap: '5px'
                            }}
                        >
                            <FileText size={13} /> Word 요구사항서
                        </button>
                    </div>
                </div>

                {/* 검색 바 & 행 추가 컨트롤 */}
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '14px', gap: '12px' }}>
                    <div style={{ position: 'relative', width: '320px' }}>
                        <Search size={14} color="var(--text-muted)" style={{ position: 'absolute', left: '10px', top: '50%', transform: 'translateY(-50%)' }} />
                        <input
                            type="text"
                            placeholder="현재 탭 데이터 검색 (키워드, 부서, ID 등)..."
                            value={searchKeyword}
                            onChange={(e) => setSearchKeyword(e.target.value)}
                            style={{
                                width: '100%',
                                padding: '7px 10px 7px 32px',
                                borderRadius: '8px',
                                border: '1px solid var(--panel-border)',
                                background: 'rgba(0, 0, 0, 0.2)',
                                color: 'var(--text-primary)',
                                fontSize: '12px'
                            }}
                        />
                    </div>

                    <button
                        onClick={handleAddRow}
                        style={{
                            padding: '7px 14px',
                            borderRadius: '8px',
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
                        <Plus size={13} /> 행 추가
                    </button>
                </div>

                {/* ── TAB 1: 인터뷰 요약 그리드 ────────────────────────── */}
                {activeTab === 'summary' && (
                    <div style={{ overflowX: 'auto', borderRadius: '10px', border: '1px solid var(--panel-border)' }}>
                        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px', textAlign: 'left' }}>
                            <thead>
                                <tr style={{ background: 'rgba(255, 255, 255, 0.04)', borderBottom: '1px solid var(--panel-border)', color: 'var(--text-secondary)' }}>
                                    <th style={{ padding: '10px', width: '50px', textAlign: 'center' }}>No</th>
                                    <th style={{ padding: '10px', width: '110px' }}>면담일자</th>
                                    <th style={{ padding: '10px', width: '150px' }}>대상기관/부서</th>
                                    <th style={{ padding: '10px', width: '120px' }}>면담대상자</th>
                                    <th style={{ padding: '10px' }}>핵심 논의 안건 요약</th>
                                    <th style={{ padding: '10px', width: '220px' }}>후속 조치 (Action Items)</th>
                                    <th style={{ padding: '10px', width: '150px' }}>출처 파일</th>
                                    <th style={{ padding: '10px', width: '50px', textAlign: 'center' }}>삭제</th>
                                </tr>
                            </thead>
                            <tbody>
                                {filteredSummary.map((row, idx) => (
                                    <tr key={idx} style={{ borderBottom: '1px solid rgba(255, 255, 255, 0.05)' }}>
                                        <td style={{ padding: '8px', textAlign: 'center', color: 'var(--text-muted)' }}>{idx + 1}</td>
                                        <td style={{ padding: '8px' }}>
                                            <input
                                                type="text"
                                                value={row.date || ''}
                                                onChange={(e) => handleUpdateCell('interviewSummary', idx, 'date', e.target.value)}
                                                style={{ width: '100%', padding: '5px', borderRadius: '5px', border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)', color: 'var(--text-primary)', fontSize: '12px' }}
                                            />
                                        </td>
                                        <td style={{ padding: '8px' }}>
                                            <input
                                                type="text"
                                                value={row.department || ''}
                                                onChange={(e) => handleUpdateCell('interviewSummary', idx, 'department', e.target.value)}
                                                style={{ width: '100%', padding: '5px', borderRadius: '5px', border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)', color: 'var(--text-primary)', fontSize: '12px', fontWeight: 600 }}
                                            />
                                        </td>
                                        <td style={{ padding: '8px' }}>
                                            <input
                                                type="text"
                                                value={row.interviewee || ''}
                                                onChange={(e) => handleUpdateCell('interviewSummary', idx, 'interviewee', e.target.value)}
                                                style={{ width: '100%', padding: '5px', borderRadius: '5px', border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)', color: 'var(--text-primary)', fontSize: '12px' }}
                                            />
                                        </td>
                                        <td style={{ padding: '8px' }}>
                                            <textarea
                                                rows={2}
                                                value={row.summary || ''}
                                                onChange={(e) => handleUpdateCell('interviewSummary', idx, 'summary', e.target.value)}
                                                style={{ width: '100%', padding: '5px', borderRadius: '5px', border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)', color: 'var(--text-primary)', fontSize: '12px', resize: 'vertical' }}
                                            />
                                        </td>
                                        <td style={{ padding: '8px' }}>
                                            <textarea
                                                rows={2}
                                                value={row.actionItems || ''}
                                                onChange={(e) => handleUpdateCell('interviewSummary', idx, 'actionItems', e.target.value)}
                                                style={{ width: '100%', padding: '5px', borderRadius: '5px', border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)', color: '#38bdf8', fontSize: '11.5px', resize: 'vertical' }}
                                            />
                                        </td>
                                        <td style={{ padding: '8px', color: 'var(--text-muted)', fontSize: '11px' }}>
                                            {row.sourceFile || '-'}
                                        </td>
                                        <td style={{ padding: '8px', textAlign: 'center' }}>
                                            <button
                                                onClick={() => handleDeleteRow('interviewSummary', idx)}
                                                style={{ background: 'none', border: 'none', color: '#f87171', cursor: 'pointer' }}
                                            >
                                                <Trash2 size={13} />
                                            </button>
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                )}

                {/* ── TAB 2: 현황 및 문제점 (As-Is) 그리드 ─────────────── */}
                {activeTab === 'issues' && (
                    <div style={{ overflowX: 'auto', borderRadius: '10px', border: '1px solid var(--panel-border)' }}>
                        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px', textAlign: 'left' }}>
                            <thead>
                                <tr style={{ background: 'rgba(255, 255, 255, 0.04)', borderBottom: '1px solid var(--panel-border)', color: 'var(--text-secondary)' }}>
                                    <th style={{ padding: '10px', width: '90px' }}>이슈 ID</th>
                                    <th style={{ padding: '10px', width: '110px' }}>분류</th>
                                    <th style={{ padding: '10px', width: '160px' }}>대상 업무/시스템</th>
                                    <th style={{ padding: '10px', width: '220px' }}>현황 (As-Is)</th>
                                    <th style={{ padding: '10px' }}>핵심 문제점 & Pain Point</th>
                                    <th style={{ padding: '10px', width: '240px' }}>원문 발언 근거 (Quote)</th>
                                    <th style={{ padding: '10px', width: '50px', textAlign: 'center' }}>삭제</th>
                                </tr>
                            </thead>
                            <tbody>
                                {filteredIssues.map((row, idx) => (
                                    <tr key={idx} style={{ borderBottom: '1px solid rgba(255, 255, 255, 0.05)' }}>
                                        <td style={{ padding: '8px' }}>
                                            <input
                                                type="text"
                                                value={row.id || ''}
                                                onChange={(e) => handleUpdateCell('asIsIssues', idx, 'id', e.target.value)}
                                                style={{ width: '100%', padding: '5px', borderRadius: '5px', border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)', color: '#f87171', fontSize: '12px', fontWeight: 700 }}
                                            />
                                        </td>
                                        <td style={{ padding: '8px' }}>
                                            <select
                                                value={row.category || '응용시스템'}
                                                onChange={(e) => handleUpdateCell('asIsIssues', idx, 'category', e.target.value)}
                                                style={{ width: '100%', padding: '5px', borderRadius: '5px', border: '1px solid var(--panel-border)', background: '#1e1b4b', color: 'var(--text-primary)', fontSize: '12px' }}
                                            >
                                                <option value="업무">업무</option>
                                                <option value="응용시스템">응용시스템</option>
                                                <option value="데이터">데이터</option>
                                                <option value="인프라">인프라</option>
                                                <option value="거버넌스">거버넌스</option>
                                            </select>
                                        </td>
                                        <td style={{ padding: '8px' }}>
                                            <input
                                                type="text"
                                                value={row.targetSystem || ''}
                                                onChange={(e) => handleUpdateCell('asIsIssues', idx, 'targetSystem', e.target.value)}
                                                style={{ width: '100%', padding: '5px', borderRadius: '5px', border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)', color: 'var(--text-primary)', fontSize: '12px' }}
                                            />
                                        </td>
                                        <td style={{ padding: '8px' }}>
                                            <textarea
                                                rows={2}
                                                value={row.asIsStatus || ''}
                                                onChange={(e) => handleUpdateCell('asIsIssues', idx, 'asIsStatus', e.target.value)}
                                                style={{ width: '100%', padding: '5px', borderRadius: '5px', border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)', color: 'var(--text-secondary)', fontSize: '11.5px', resize: 'vertical' }}
                                            />
                                        </td>
                                        <td style={{ padding: '8px' }}>
                                            <textarea
                                                rows={2}
                                                value={row.painPoint || ''}
                                                onChange={(e) => handleUpdateCell('asIsIssues', idx, 'painPoint', e.target.value)}
                                                style={{ width: '100%', padding: '5px', borderRadius: '5px', border: '1px solid rgba(239, 68, 68, 0.3)', background: 'rgba(239, 68, 68, 0.05)', color: 'var(--text-primary)', fontSize: '12px', fontWeight: 600, resize: 'vertical' }}
                                            />
                                        </td>
                                        <td style={{ padding: '8px' }}>
                                            <textarea
                                                rows={2}
                                                value={row.quote || ''}
                                                onChange={(e) => handleUpdateCell('asIsIssues', idx, 'quote', e.target.value)}
                                                placeholder="💬 발언 인용..."
                                                style={{ width: '100%', padding: '5px', borderRadius: '5px', border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)', color: '#a78bfa', fontSize: '11px', fontStyle: 'italic', resize: 'vertical' }}
                                            />
                                        </td>
                                        <td style={{ padding: '8px', textAlign: 'center' }}>
                                            <button
                                                onClick={() => handleDeleteRow('asIsIssues', idx)}
                                                style={{ background: 'none', border: 'none', color: '#f87171', cursor: 'pointer' }}
                                            >
                                                <Trash2 size={13} />
                                            </button>
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                )}

                {/* ── TAB 3: 개선방향 및 시사점 (To-Be) 그리드 ────────── */}
                {activeTab === 'directions' && (
                    <div style={{ overflowX: 'auto', borderRadius: '10px', border: '1px solid var(--panel-border)' }}>
                        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px', textAlign: 'left' }}>
                            <thead>
                                <tr style={{ background: 'rgba(255, 255, 255, 0.04)', borderBottom: '1px solid var(--panel-border)', color: 'var(--text-secondary)' }}>
                                    <th style={{ padding: '10px', width: '110px' }}>연계 이슈ID</th>
                                    <th style={{ padding: '10px', width: '32%' }}>도출 시사점 (Implications)</th>
                                    <th style={{ padding: '10px', width: '35%' }}>목표모델 및 개선방향 (To-Be)</th>
                                    <th style={{ padding: '10px', width: '23%' }}>연관 정보화 추진과제명</th>
                                    <th style={{ padding: '10px', width: '50px', textAlign: 'center' }}>삭제</th>
                                </tr>
                            </thead>
                            <tbody>
                                {filteredDirections.map((row, idx) => (
                                    <tr key={idx} style={{ borderBottom: '1px solid rgba(255, 255, 255, 0.05)' }}>
                                        <td style={{ padding: '8px' }}>
                                            <input
                                                type="text"
                                                value={row.issueId || ''}
                                                onChange={(e) => handleUpdateCell('toBeDirections', idx, 'issueId', e.target.value)}
                                                style={{ width: '100%', padding: '5px', borderRadius: '5px', border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)', color: '#34d399', fontSize: '12px', fontWeight: 700 }}
                                            />
                                        </td>
                                        <td style={{ padding: '8px' }}>
                                            <textarea
                                                rows={2}
                                                value={row.implication || ''}
                                                onChange={(e) => handleUpdateCell('toBeDirections', idx, 'implication', e.target.value)}
                                                style={{ width: '100%', padding: '5px', borderRadius: '5px', border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)', color: 'var(--text-secondary)', fontSize: '12px', resize: 'vertical' }}
                                            />
                                        </td>
                                        <td style={{ padding: '8px' }}>
                                            <textarea
                                                rows={2}
                                                value={row.toBeDirection || ''}
                                                onChange={(e) => handleUpdateCell('toBeDirections', idx, 'toBeDirection', e.target.value)}
                                                style={{ width: '100%', padding: '5px', borderRadius: '5px', border: '1px solid rgba(16, 185, 129, 0.3)', background: 'rgba(16, 185, 129, 0.05)', color: 'var(--text-primary)', fontSize: '12px', fontWeight: 600, resize: 'vertical' }}
                                            />
                                        </td>
                                        <td style={{ padding: '8px' }}>
                                            <input
                                                type="text"
                                                value={row.relatedProject || ''}
                                                onChange={(e) => handleUpdateCell('toBeDirections', idx, 'relatedProject', e.target.value)}
                                                style={{ width: '100%', padding: '5px', borderRadius: '5px', border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)', color: '#38bdf8', fontSize: '12px', fontWeight: 600 }}
                                            />
                                        </td>
                                        <td style={{ padding: '8px', textAlign: 'center' }}>
                                            <button
                                                onClick={() => handleDeleteRow('toBeDirections', idx)}
                                                style={{ background: 'none', border: 'none', color: '#f87171', cursor: 'pointer' }}
                                            >
                                                <Trash2 size={13} />
                                            </button>
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                )}

                {/* ── TAB 4: 도출 요구사항 (SSOT) 그리드 ───────────────── */}
                {activeTab === 'requirements' && (
                    <div style={{ overflowX: 'auto', borderRadius: '10px', border: '1px solid var(--panel-border)' }}>
                        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px', textAlign: 'left' }}>
                            <thead>
                                <tr style={{ background: 'rgba(255, 255, 255, 0.04)', borderBottom: '1px solid var(--panel-border)', color: 'var(--text-secondary)' }}>
                                    <th style={{ padding: '10px', width: '110px' }}>요구사항 ID</th>
                                    <th style={{ padding: '10px', width: '22%' }}>요구사항 명칭</th>
                                    <th style={{ padding: '10px', width: '90px' }}>구분</th>
                                    <th style={{ padding: '10px' }}>상세 요건 정의</th>
                                    <th style={{ padding: '10px', width: '80px', textAlign: 'center' }}>우선순위</th>
                                    <th style={{ padding: '10px', width: '140px' }}>출처 회의록</th>
                                    <th style={{ padding: '10px', width: '50px', textAlign: 'center' }}>삭제</th>
                                </tr>
                            </thead>
                            <tbody>
                                {filteredRequirements.map((row, idx) => (
                                    <tr key={idx} style={{ borderBottom: '1px solid rgba(255, 255, 255, 0.05)' }}>
                                        <td style={{ padding: '8px' }}>
                                            <input
                                                type="text"
                                                value={row.id || ''}
                                                onChange={(e) => handleUpdateCell('derivedRequirements', idx, 'id', e.target.value)}
                                                style={{ width: '100%', padding: '5px', borderRadius: '5px', border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)', color: '#60a5fa', fontSize: '12px', fontWeight: 800 }}
                                            />
                                        </td>
                                        <td style={{ padding: '8px' }}>
                                            <input
                                                type="text"
                                                value={row.name || ''}
                                                onChange={(e) => handleUpdateCell('derivedRequirements', idx, 'name', e.target.value)}
                                                style={{ width: '100%', padding: '5px', borderRadius: '5px', border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)', color: 'var(--text-primary)', fontSize: '12px', fontWeight: 600 }}
                                            />
                                        </td>
                                        <td style={{ padding: '8px' }}>
                                            <select
                                                value={row.type || '기능'}
                                                onChange={(e) => handleUpdateCell('derivedRequirements', idx, 'type', e.target.value)}
                                                style={{ width: '100%', padding: '5px', borderRadius: '5px', border: '1px solid var(--panel-border)', background: '#1e1b4b', color: 'var(--text-primary)', fontSize: '12px' }}
                                            >
                                                <option value="기능">기능</option>
                                                <option value="비기능">비기능</option>
                                                <option value="데이터">데이터</option>
                                                <option value="연계">연계</option>
                                                <option value="보안">보안</option>
                                                <option value="인프라">인프라</option>
                                            </select>
                                        </td>
                                        <td style={{ padding: '8px' }}>
                                            <textarea
                                                rows={2}
                                                value={row.description || ''}
                                                onChange={(e) => handleUpdateCell('derivedRequirements', idx, 'description', e.target.value)}
                                                style={{ width: '100%', padding: '5px', borderRadius: '5px', border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)', color: 'var(--text-primary)', fontSize: '12px', resize: 'vertical' }}
                                            />
                                        </td>
                                        <td style={{ padding: '8px', textAlign: 'center' }}>
                                            <select
                                                value={row.priority || '상'}
                                                onChange={(e) => handleUpdateCell('derivedRequirements', idx, 'priority', e.target.value)}
                                                style={{
                                                    padding: '4px 8px',
                                                    borderRadius: '6px',
                                                    border: '1px solid var(--panel-border)',
                                                    background: row.priority === '상' ? 'rgba(239, 68, 68, 0.2)' : 'rgba(255,255,255,0.05)',
                                                    color: row.priority === '상' ? '#f87171' : 'var(--text-primary)',
                                                    fontSize: '11.5px',
                                                    fontWeight: 700
                                                }}
                                            >
                                                <option value="상">상</option>
                                                <option value="중">중</option>
                                                <option value="하">하</option>
                                            </select>
                                        </td>
                                        <td style={{ padding: '8px', color: 'var(--text-muted)', fontSize: '11px' }}>
                                            {row.sourceFile || '-'}
                                        </td>
                                        <td style={{ padding: '8px', textAlign: 'center' }}>
                                            <button
                                                onClick={() => handleDeleteRow('derivedRequirements', idx)}
                                                style={{ background: 'none', border: 'none', color: '#f87171', cursor: 'pointer' }}
                                            >
                                                <Trash2 size={13} />
                                            </button>
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                )}
            </div>

            {/* ── 원문 텍스트 미리보기 모달 ─────────────────────────── */}
            {previewDoc && (
                <div style={{
                    position: 'fixed',
                    top: 0,
                    left: 0,
                    width: '100%',
                    height: '100%',
                    background: 'rgba(0, 0, 0, 0.7)',
                    backdropFilter: 'blur(8px)',
                    zIndex: 10000,
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    padding: '24px'
                }}>
                    <div className="glass-panel" style={{
                        width: '100%',
                        maxWidth: '800px',
                        maxHeight: '85vh',
                        display: 'flex',
                        flexDirection: 'column',
                        borderRadius: '16px',
                        border: '1px solid var(--panel-border)',
                        background: '#0f172a',
                        overflow: 'hidden',
                        boxShadow: '0 24px 64px rgba(0,0,0,0.6)'
                    }}>
                        <div style={{
                            padding: '16px 20px',
                            borderBottom: '1px solid var(--panel-border)',
                            display: 'flex',
                            justifyContent: 'space-between',
                            alignItems: 'center'
                        }}>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                <FileText size={18} color="#8b5cf6" />
                                <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 700, color: '#fff' }}>
                                    {previewDoc.meta.filename} (원문 미리보기)
                                </h3>
                            </div>
                            <button
                                onClick={() => setPreviewDoc(null)}
                                style={{ background: 'none', border: 'none', color: 'var(--text-muted)', cursor: 'pointer', padding: '4px' }}
                            >
                                <X size={18} />
                            </button>
                        </div>

                        <div style={{ padding: '20px', overflowY: 'auto', flex: 1, whiteSpace: 'pre-wrap', fontSize: '13px', lineHeight: 1.7, color: '#cbd5e1', fontFamily: 'monospace' }}>
                            {previewDoc.text}
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
}
