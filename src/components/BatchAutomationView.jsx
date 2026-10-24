import React, { useState, useEffect, useRef, useMemo } from 'react';
import {
    Workflow,
    Upload,
    Download,
    Play,
    Pause,
    RotateCcw,
    Trash2,
    CheckCircle2,
    AlertTriangle,
    XCircle,
    Clock,
    FileSpreadsheet,
    FileText,
    BookOpen,
    Layers,
    Sliders,
    Search,
    Eye,
    X,
    Sparkles,
    ShieldAlert,
    ExternalLink,
    Database,
    Plus,
    RefreshCw
} from 'lucide-react';
import { saveAs } from 'file-saver';
import {
    BatchQueueRunner,
    BATCH_ITEM_STATUS,
    STATUS_LABELS,
    STATUS_COLORS,
    generateBatchReportZip,
    buildSingleFileDocxReport
} from '../utils/batchQueueRunner';
import {
    indexDocumentFile,
    getAllIndexedDocs,
    deleteIndexedDoc,
    clearAllIndexedDocs,
    searchRagHybrid
} from '../utils/clientRagIndexer';
import { loadDesignTokens } from '../utils/designTokens';
import { loadSSOTData } from '../utils/ssotValidator';

export default function BatchAutomationView({ apiKey, llmProvider, omniRouteModel }) {
    // ── 큐 상태 ──────────────────────────────────────────────────────
    const [queueItems, setQueueItems] = useState([]);
    const [concurrency, setConcurrency] = useState(2);
    const [useAiCheck, setUseAiCheck] = useState(true);
    const [useRagContext, setUseRagContext] = useState(true);
    const [queueState, setQueueState] = useState('IDLE'); // 'IDLE' | 'RUNNING' | 'PAUSED' | 'COMPLETED'
    const [progressStats, setProgressStats] = useState({ total: 0, completed: 0, error: 0, running: 0, pending: 0, percent: 0 });

    // RAG 관리 모달 상태
    const [showRagModal, setShowRagModal] = useState(false);
    const [indexedDocs, setIndexedDocs] = useState([]);
    const [isIndexingFile, setIsIndexingFile] = useState(false);
    const [indexProgress, setIndexProgress] = useState({ percent: 0, message: '' });
    const [ragTestQuery, setRagTestQuery] = useState('ISP 수립 대상 및 총구축비 산정 기준');
    const [ragTestResults, setRagTestResults] = useState([]);
    const [isSearchingRag, setIsSearchingRag] = useState(false);

    // 결함 상세 조회 모달
    const [inspectingItem, setInspectingItem] = useState(null);
    const [issueFilterSeverity, setIssueFilterSeverity] = useState('ALL');

    // 토스트 및 내보내기 상태
    const [toastMessage, setToastMessage] = useState(null);
    const [isExporting, setIsExporting] = useState(false);
    const [isDragging, setIsDragging] = useState(false);

    const fileInputRef = useRef(null);
    const ragFileInputRef = useRef(null);
    const runnerRef = useRef(null);

    const showToast = (msg, type = 'success') => {
        setToastMessage({ msg, type });
        setTimeout(() => setToastMessage(null), 4000);
    };

    // ── 1. RAG IndexedDB 문서 목록 로드 ─────────────────────────────
    const refreshIndexedDocs = async () => {
        const docs = await getAllIndexedDocs();
        setIndexedDocs(docs);
    };

    useEffect(() => {
        refreshIndexedDocs();
    }, []);

    // ── 2. 배치 러너 초기화 ─────────────────────────────────────────
    useEffect(() => {
        const masterData = loadSSOTData();
        const validKey = (apiKey || '').split(',').map(k => k.trim()).find(k => k.startsWith('AIza') || k.startsWith('AQ.'));

        runnerRef.current = new BatchQueueRunner({
            concurrency,
            apiKey: validKey || '',
            modelId: omniRouteModel && omniRouteModel !== 'auto' ? omniRouteModel : 'gemini-2.5-flash',
            masterData,
            useAiCheck,
            useRagContext,
            onItemChange: (updatedItem) => {
                setQueueItems(prev => prev.map(item => item.id === updatedItem.id ? { ...updatedItem } : item));
            },
            onProgress: (stats) => {
                setProgressStats(stats);
            },
            onComplete: (items) => {
                setQueueState('COMPLETED');
                showToast(`모든 배치 검수 작업이 완료되었습니다! (총 ${items.length}개 파일)`);
            }
        });
    }, [concurrency, apiKey, omniRouteModel, useAiCheck, useRagContext]);

    // ── 3. 파일 드롭 및 추가 ────────────────────────────────────────
    const handleFilesAdded = (files) => {
        if (!files || files.length === 0) return;
        if (!runnerRef.current) return;

        const added = runnerRef.current.addFiles(files);
        setQueueItems(prev => [...prev, ...added]);
        setQueueState('IDLE');
        showToast(`${added.length}개의 산출물 파일이 배치 검수 큐에 등록되었습니다.`);
    };

    const handleDrop = (e) => {
        e.preventDefault();
        setIsDragging(false);
        if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
            handleFilesAdded(e.dataTransfer.files);
        }
    };

    // 체험용 가상 샘플 파일 4종 등록
    const handleAddSampleFiles = () => {
        const sample1 = new File([
            `Ⅰ. 사업 개요 및 추진 배경
본 사업은 지역의료원과 권역책임의료기관 간 환자 전원 및 진료정보 고속도로 연계를 목표로 한다.
현재 지방의료원에서는 환자 전원 시 진료의뢰서를 수기 복사하여 전달하고 있으며, 백만원 단위 표기 오류가 발생함.
요구사항 고유번호 CNR-001, SFR-001 등이 반영되어야 하나 일부 문서에서 누락됨.`
        ], '01_사업추진계획서_최종본.docx', { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });

        const sample2 = new File([
            `Ⅱ. 현황 분석 및 문제점
원내 EMR 시스템은 14년째 노후화 운영 중이며, 건강정보 고속도로 연계 규격이 미비함.
지역의료원 통합 인프라 전환 시 소요 예산은 총 45억원이나 세부 항목 합계는 42억원으로 3억원 오차가 발생함.
요구사항 ID USR-001, GHOST-999 표기 오류 확인 필요.`
        ], '02_현황분석_보고서_장표.pptx', { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });

        const sample3 = new File([
            `Ⅲ. 목표모델 및 시스템 요건
클라우드 전환 및 마이데이터 연계 체계 구축.
지역의료원 명칭을 지방의료원으로 통일해야 하며, 총 구축비 산정 기준을 준수해야 함.`
        ], '03_목표시스템_기능정의서.hwpx', { type: 'application/hwp+zip' });

        const sample4 = new File([
            `Ⅳ. 이행계획 및 소요예산
단위시스템별 하드웨어 도입비 15억원, 소프트웨어 구입비 12억원, 개발비 20억원으로 계 47억원.`
        ], '04_소요예산산출내역서.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

        handleFilesAdded([sample1, sample2, sample3, sample4]);
    };

    // ── 4. 큐 제어 핸들러 ──────────────────────────────────────────
    const handleStartQueue = () => {
        if (!runnerRef.current) return;
        if (queueItems.length === 0) {
            showToast('검수할 파일이 없습니다. 파일을 먼저 등록해 주세요.', 'error');
            return;
        }
        setQueueState('RUNNING');
        runnerRef.current.start();
        showToast('배치 검수 큐 작업이 시작되었습니다.');
    };

    const handlePauseQueue = () => {
        if (!runnerRef.current) return;
        runnerRef.current.pause();
        setQueueState('PAUSED');
        showToast('배치 검수가 일시정지되었습니다.');
    };

    const handleResumeQueue = () => {
        if (!runnerRef.current) return;
        runnerRef.current.resume();
        setQueueState('RUNNING');
        showToast('배치 검수가 재개되었습니다.');
    };

    const handleRetryFailed = () => {
        if (!runnerRef.current) return;
        runnerRef.current.retryFailed();
        setQueueState('RUNNING');
        showToast('실패한 파일들에 대한 재검수를 시작합니다.');
    };

    const handleClearQueue = () => {
        if (runnerRef.current) {
            runnerRef.current.cancel();
        }
        setQueueItems([]);
        setQueueState('IDLE');
        setProgressStats({ total: 0, completed: 0, error: 0, running: 0, pending: 0, percent: 0 });
        showToast('큐 목록이 초기화되었습니다.');
    };

    const handleRemoveItem = (id) => {
        setQueueItems(prev => prev.filter(item => item.id !== id));
        if (runnerRef.current) {
            runnerRef.current.items = runnerRef.current.items.filter(item => item.id !== id);
            runnerRef.current._notifyProgress();
        }
    };

    // ── 5. ZIP 일괄 다운로드 & 단일 교정본 다운로드 ────────────────
    const handleDownloadZip = async () => {
        const completedItems = queueItems.filter(i => i.status === BATCH_ITEM_STATUS.COMPLETED);
        if (completedItems.length === 0) {
            showToast('완료된 검수 결과가 없습니다. 배치 검수를 먼저 완료해 주세요.', 'error');
            return;
        }

        setIsExporting(true);
        try {
            const tokens = loadDesignTokens();
            await generateBatchReportZip(queueItems, tokens);
            showToast('전체 산출물 통합 대시보드(Excel) 및 개별 교정리포트(Word) ZIP 패키지 다운로드가 완료되었습니다!');
        } catch (err) {
            console.error('ZIP 생성 에러:', err);
            showToast(`ZIP 생성 오류: ${err.message}`, 'error');
        } finally {
            setIsExporting(false);
        }
    };

    const handleDownloadSingleDocx = async (item) => {
        try {
            const tokens = loadDesignTokens();
            const blob = await buildSingleFileDocxReport(item, tokens);
            const cleanName = (item.filename || 'doc').replace(/\.[^/.]+$/, '');
            saveAs(blob, `교정리포트_${cleanName}_${new Date().toISOString().slice(0, 10)}.docx`);
            showToast(`[${cleanName}] Word 교정 리포트가 다운로드되었습니다.`);
        } catch (err) {
            showToast(`리포트 생성 실패: ${err.message}`, 'error');
        }
    };

    // ── 6. RAG 인덱싱 및 테스트 검색 ───────────────────────────────
    const handleUploadRagGuideline = async (e) => {
        const file = e.target.files?.[0];
        if (!file) return;

        setIsIndexingFile(true);
        setIndexProgress({ percent: 10, message: `[${file.name}] 가이드라인 브라우저 직접 파싱 시작...` });

        try {
            const res = await indexDocumentFile(file, (p) => setIndexProgress(p));
            showToast(`[${res.filename}] IndexedDB 인덱싱 완료 (${res.chunkCount}개 청크 저장)`);
            await refreshIndexedDocs();
        } catch (err) {
            console.error('RAG 인덱싱 실패:', err);
            showToast(`인덱싱 오류: ${err.message}`, 'error');
        } finally {
            setIsIndexingFile(false);
            setIndexProgress({ percent: 0, message: '' });
            if (ragFileInputRef.current) ragFileInputRef.current.value = '';
        }
    };

    const handleDeleteRagDoc = async (docId) => {
        if (!window.confirm('해당 가이드라인 문서와 인덱싱된 청크를 IndexedDB에서 삭제하시겠습니까?')) return;
        await deleteIndexedDoc(docId);
        await refreshIndexedDocs();
        showToast('문서가 삭제되었습니다.');
    };

    const handleClearAllRag = async () => {
        if (!window.confirm('IndexedDB의 모든 RAG 가이드라인 데이터를 초기화하시겠습니까?')) return;
        await clearAllIndexedDocs();
        await refreshIndexedDocs();
        showToast('모든 인덱싱 데이터가 초기화되었습니다.');
    };

    const handleTestSearchRag = async () => {
        if (!ragTestQuery.trim()) return;
        setIsSearchingRag(true);
        try {
            const results = await searchRagHybrid(ragTestQuery, { topK: 5 });
            setRagTestResults(results);
            if (results.length === 0) {
                showToast('검색어와 일치하는 청크가 없습니다.', 'info');
            }
        } catch (err) {
            showToast(`검색 오류: ${err.message}`, 'error');
        } finally {
            setIsSearchingRag(false);
        }
    };

    // 통계 집계
    const totalDefectsAll = useMemo(() => {
        return queueItems.reduce((acc, item) => acc + (item.totalDefectsCount || 0), 0);
    }, [queueItems]);

    const totalRuleDefects = useMemo(() => {
        return queueItems.reduce((acc, item) => acc + (item.ruleDefectsCount || 0), 0);
    }, [queueItems]);

    const totalAiDefects = useMemo(() => {
        return queueItems.reduce((acc, item) => acc + (item.aiDefectsCount || 0), 0);
    }, [queueItems]);

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

            {/* ── 헤더 패널 ────────────────────────────────────────── */}
            <div className="glass-panel" style={{
                padding: '24px',
                borderRadius: '16px',
                border: '1px solid var(--panel-border)',
                background: 'linear-gradient(135deg, rgba(236, 72, 153, 0.08) 0%, rgba(99, 102, 241, 0.05) 100%)',
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
                            background: 'linear-gradient(135deg, #ec4899, #8b5cf6)',
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                            color: '#fff',
                            boxShadow: '0 4px 12px rgba(236, 72, 153, 0.3)'
                        }}>
                            <Workflow size={22} />
                        </div>
                        <h2 style={{ margin: 0, fontSize: '20px', fontWeight: 800, color: 'var(--text-primary)' }}>
                            다중 산출물 배치(Batch) 자동 검수 큐 & 브라우저 내장형 RAG 인덱서
                        </h2>
                        <span style={{
                            padding: '3px 10px',
                            borderRadius: '20px',
                            background: 'rgba(236, 72, 153, 0.2)',
                            color: '#f472b6',
                            fontSize: '11px',
                            fontWeight: 700,
                            border: '1px solid rgba(236, 72, 153, 0.4)'
                        }}>
                            4순위 Batch Queue
                        </span>
                    </div>
                    <p style={{ margin: 0, fontSize: '13px', color: 'var(--text-secondary)', lineHeight: 1.6, maxWidth: '850px' }}>
                        산출물 파일 10~30개(<strong>.pptx, .docx, .hwpx, .xlsx, .pdf</strong>)를 일괄 등록하여 동시성 제어 큐로 안전하게 검수합니다.
                        Gemini 429 초과 시 <strong>지수 백오프(Exponential Backoff)</strong>로 자동 재시도하며,
                        <strong>브라우저 IndexedDB 기반 RAG</strong>를 통해 Vercel 환경에서도 가이드라인을 참조한 정밀 교정본을 <strong>통합 대시보드(Excel)와 개별 Redline 보고서(Word) ZIP 패키지</strong>로 다운로드합니다.
                    </p>
                </div>

                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
                    <button
                        onClick={handleAddSampleFiles}
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
                        <RefreshCw size={13} /> 체험용 샘플 4종 추가
                    </button>

                    <button
                        onClick={() => setShowRagModal(true)}
                        style={{
                            padding: '8px 14px',
                            borderRadius: '9px',
                            border: '1px solid rgba(99, 102, 241, 0.4)',
                            background: 'rgba(99, 102, 241, 0.12)',
                            color: '#818cf8',
                            fontSize: '12px',
                            fontWeight: 700,
                            cursor: 'pointer',
                            display: 'flex',
                            alignItems: 'center',
                            gap: '6px'
                        }}
                    >
                        <BookOpen size={13} /> 브라우저 RAG 가이드라인 ({indexedDocs.length}건)
                    </button>

                    <button
                        onClick={handleDownloadZip}
                        disabled={isExporting || queueItems.every(i => i.status !== BATCH_ITEM_STATUS.COMPLETED)}
                        style={{
                            padding: '8px 16px',
                            borderRadius: '9px',
                            border: '1px solid rgba(236, 72, 153, 0.5)',
                            background: 'linear-gradient(135deg, #ec4899, #8b5cf6)',
                            color: '#fff',
                            fontSize: '12.5px',
                            fontWeight: 700,
                            cursor: isExporting ? 'not-allowed' : 'pointer',
                            display: 'flex',
                            alignItems: 'center',
                            gap: '7px',
                            boxShadow: '0 4px 14px rgba(236, 72, 153, 0.35)',
                            opacity: queueItems.some(i => i.status === BATCH_ITEM_STATUS.COMPLETED) ? 1 : 0.6
                        }}
                    >
                        <Download size={14} />
                        {isExporting ? '패키징 빌드 중...' : '산출물 일괄검수 패키지 (ZIP)'}
                    </button>
                </div>
            </div>

            {/* ── 주요 통계 대시보드 카드 ────────────────────────────── */}
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '14px' }}>
                <div className="glass-panel" style={{ padding: '16px 20px', borderRadius: '12px', border: '1px solid var(--panel-border)' }}>
                    <span style={{ fontSize: '11px', color: 'var(--text-muted)', fontWeight: 600 }}>큐 등록 산출물</span>
                    <div style={{ fontSize: '24px', fontWeight: 800, color: 'var(--text-primary)', marginTop: '4px' }}>
                        {queueItems.length} <span style={{ fontSize: '13px', fontWeight: 500, color: 'var(--text-secondary)' }}>개 파일</span>
                    </div>
                </div>

                <div className="glass-panel" style={{ padding: '16px 20px', borderRadius: '12px', border: '1px solid var(--panel-border)' }}>
                    <span style={{ fontSize: '11px', color: 'var(--text-muted)', fontWeight: 600 }}>전체 검수 진행률</span>
                    <div style={{ fontSize: '24px', fontWeight: 800, color: '#38bdf8', marginTop: '4px' }}>
                        {progressStats.percent}%
                        <span style={{ fontSize: '12px', fontWeight: 500, color: 'var(--text-secondary)', marginLeft: '8px' }}>
                            ({progressStats.completed}완료 / {progressStats.error}오류)
                        </span>
                    </div>
                </div>

                <div className="glass-panel" style={{ padding: '16px 20px', borderRadius: '12px', border: '1px solid var(--panel-border)' }}>
                    <span style={{ fontSize: '11px', color: 'var(--text-muted)', fontWeight: 600 }}>발견된 총 결함 건수</span>
                    <div style={{ fontSize: '24px', fontWeight: 800, color: '#f87171', marginTop: '4px' }}>
                        {totalDefectsAll} <span style={{ fontSize: '12px', fontWeight: 500, color: 'var(--text-secondary)' }}>(Rule: {totalRuleDefects} / AI: {totalAiDefects})</span>
                    </div>
                </div>

                <div className="glass-panel" style={{ padding: '16px 20px', borderRadius: '12px', border: '1px solid var(--panel-border)' }}>
                    <span style={{ fontSize: '11px', color: 'var(--text-muted)', fontWeight: 600 }}>동시 처리 및 큐 상태</span>
                    <div style={{ fontSize: '16px', fontWeight: 800, color: queueState === 'RUNNING' ? '#10b981' : queueState === 'PAUSED' ? '#fb923c' : 'var(--text-primary)', marginTop: '8px', display: 'flex', alignItems: 'center', gap: '6px' }}>
                        <span style={{ width: '8px', height: '8px', borderRadius: '50%', background: queueState === 'RUNNING' ? '#10b981' : queueState === 'PAUSED' ? '#fb923c' : '#94a3b8' }} />
                        {queueState === 'RUNNING' ? `실행중 (${concurrency}열 동시)` : queueState === 'PAUSED' ? '일시정지됨' : queueState === 'COMPLETED' ? '검수 완료' : '대기중'}
                    </div>
                </div>
            </div>

            {/* ── 큐 컨트롤 바 & 설정 ─────────────────────────────────── */}
            <div className="glass-panel" style={{
                padding: '16px 20px',
                borderRadius: '12px',
                border: '1px solid var(--panel-border)',
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center',
                flexWrap: 'wrap',
                gap: '12px'
            }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '14px', flexWrap: 'wrap' }}>
                    {/* 동시성 선택 */}
                    <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                        <Sliders size={13} color="var(--text-muted)" />
                        <span style={{ fontSize: '12px', color: 'var(--text-secondary)', fontWeight: 600 }}>동시 검수 수:</span>
                        <select
                            value={concurrency}
                            onChange={(e) => setConcurrency(Number(e.target.value))}
                            style={{
                                padding: '4px 8px',
                                borderRadius: '6px',
                                border: '1px solid var(--panel-border)',
                                background: '#1e1b4b',
                                color: 'var(--text-primary)',
                                fontSize: '12px'
                            }}
                        >
                            <option value={1}>1개 (순차 안전)</option>
                            <option value={2}>2개 (권장 기본)</option>
                            <option value={3}>3개 (고속)</option>
                            <option value={4}>4개 (초고속)</option>
                        </select>
                    </div>

                    {/* AI 문맥 검수 토글 */}
                    <label style={{ display: 'flex', alignItems: 'center', gap: '6px', cursor: 'pointer', fontSize: '12px', color: 'var(--text-primary)' }}>
                        <input
                            type="checkbox"
                            checked={useAiCheck}
                            onChange={(e) => setUseAiCheck(e.target.checked)}
                        />
                        <span>AI 문맥·지침 검수 포함</span>
                    </label>

                    {/* RAG 가이드라인 주입 토글 */}
                    <label style={{ display: 'flex', alignItems: 'center', gap: '6px', cursor: 'pointer', fontSize: '12px', color: '#818cf8', fontWeight: 600 }}>
                        <input
                            type="checkbox"
                            checked={useRagContext}
                            onChange={(e) => setUseRagContext(e.target.checked)}
                        />
                        <span>RAG 가이드라인 Top-5 주입</span>
                    </label>
                </div>

                {/* 큐 실행 액션 버튼군 */}
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    {queueState !== 'RUNNING' ? (
                        <button
                            onClick={handleStartQueue}
                            disabled={queueItems.length === 0}
                            style={{
                                padding: '8px 16px',
                                borderRadius: '8px',
                                border: 'none',
                                background: queueItems.length === 0 ? 'rgba(255,255,255,0.08)' : 'linear-gradient(135deg, #10b981, #059669)',
                                color: '#fff',
                                fontSize: '12.5px',
                                fontWeight: 700,
                                cursor: queueItems.length === 0 ? 'not-allowed' : 'pointer',
                                display: 'flex',
                                alignItems: 'center',
                                gap: '6px'
                            }}
                        >
                            <Play size={13} /> {queueState === 'PAUSED' ? '큐 재개' : '배치 검수 시작'}
                        </button>
                    ) : (
                        <button
                            onClick={handlePauseQueue}
                            style={{
                                padding: '8px 16px',
                                borderRadius: '8px',
                                border: '1px solid rgba(251, 146, 60, 0.4)',
                                background: 'rgba(251, 146, 60, 0.15)',
                                color: '#fb923c',
                                fontSize: '12.5px',
                                fontWeight: 700,
                                cursor: 'pointer',
                                display: 'flex',
                                alignItems: 'center',
                                gap: '6px'
                            }}
                        >
                            <Pause size={13} /> 일시정지
                        </button>
                    )}

                    {progressStats.error > 0 && (
                        <button
                            onClick={handleRetryFailed}
                            style={{
                                padding: '8px 14px',
                                borderRadius: '8px',
                                border: '1px solid rgba(239, 68, 68, 0.3)',
                                background: 'rgba(239, 68, 68, 0.12)',
                                color: '#f87171',
                                fontSize: '12px',
                                fontWeight: 600,
                                cursor: 'pointer',
                                display: 'flex',
                                alignItems: 'center',
                                gap: '5px'
                            }}
                        >
                            <RotateCcw size={12} /> 실패건 재시도 ({progressStats.error})
                        </button>
                    )}

                    <button
                        onClick={handleClearQueue}
                        disabled={queueItems.length === 0}
                        style={{
                            padding: '8px 12px',
                            borderRadius: '8px',
                            border: '1px solid rgba(255, 255, 255, 0.1)',
                            background: 'rgba(255, 255, 255, 0.05)',
                            color: 'var(--text-muted)',
                            fontSize: '12px',
                            cursor: queueItems.length === 0 ? 'not-allowed' : 'pointer',
                            display: 'flex',
                            alignItems: 'center',
                            gap: '4px'
                        }}
                    >
                        <Trash2 size={12} /> 비우기
                    </button>
                </div>
            </div>

            {/* ── 드래그 앤 드롭 영역 ─────────────────────────────────── */}
            <div
                onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
                onDragLeave={() => setIsDragging(false)}
                onDrop={handleDrop}
                onClick={() => fileInputRef.current?.click()}
                style={{
                    border: `2px dashed ${isDragging ? '#ec4899' : 'rgba(255, 255, 255, 0.15)'}`,
                    borderRadius: '12px',
                    padding: '24px 20px',
                    textAlign: 'center',
                    background: isDragging ? 'rgba(236, 72, 153, 0.08)' : 'rgba(255, 255, 255, 0.015)',
                    cursor: 'pointer',
                    transition: 'all 0.2s ease'
                }}
            >
                <input
                    ref={fileInputRef}
                    type="file"
                    multiple
                    accept=".pptx,.docx,.hwpx,.xlsx,.pdf,.txt,.md"
                    style={{ display: 'none' }}
                    onChange={(e) => handleFilesAdded(e.target.files)}
                />
                <Upload size={32} color={isDragging ? '#ec4899' : 'var(--text-secondary)'} style={{ margin: '0 auto 8px' }} />
                <p style={{ margin: 0, fontSize: '13.5px', fontWeight: 600, color: 'var(--text-primary)' }}>
                    검수할 산출물 파일 10~30개를 이곳에 드래그하거나 클릭하여 일괄 등록하세요
                </p>
                <p style={{ margin: '4px 0 0', fontSize: '12px', color: 'var(--text-secondary)' }}>
                    지원 형식: <strong>.pptx, .docx, .hwpx, .xlsx, .pdf, .md, .txt</strong> (동시 다중 선택 가능)
                </p>
            </div>

            {/* ── 큐 파일 테이블 ──────────────────────────────────────── */}
            {queueItems.length > 0 && (
                <div className="glass-panel" style={{
                    padding: '20px',
                    borderRadius: '14px',
                    border: '1px solid var(--panel-border)',
                    overflowX: 'auto'
                }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '14px' }}>
                        <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 700, color: 'var(--text-primary)' }}>
                            배치 검수 작업 큐 목록 ({queueItems.length}건)
                        </h3>
                        <span style={{ fontSize: '12px', color: 'var(--text-muted)' }}>
                            완료: {progressStats.completed} / 진행중: {progressStats.running} / 대기: {progressStats.pending}
                        </span>
                    </div>

                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12.5px', textAlign: 'left' }}>
                        <thead>
                            <tr style={{ background: 'rgba(255, 255, 255, 0.04)', borderBottom: '1px solid var(--panel-border)', color: 'var(--text-secondary)' }}>
                                <th style={{ padding: '10px 12px', width: '50px', textAlign: 'center' }}>No</th>
                                <th style={{ padding: '10px 12px', width: '28%' }}>산출물 파일명</th>
                                <th style={{ padding: '10px 12px', width: '18%' }}>검수 단계 & 상태</th>
                                <th style={{ padding: '10px 12px', width: '24%' }}>실시간 진행 메시지</th>
                                <th style={{ padding: '10px 12px', width: '14%', textAlign: 'center' }}>식별 결함</th>
                                <th style={{ padding: '10px 12px', width: '12%', textAlign: 'center' }}>작업</th>
                            </tr>
                        </thead>
                        <tbody>
                            {queueItems.map((item, idx) => (
                                <tr key={item.id} style={{ borderBottom: '1px solid rgba(255, 255, 255, 0.05)' }}>
                                    <td style={{ padding: '10px 12px', textAlign: 'center', color: 'var(--text-muted)' }}>{idx + 1}</td>
                                    <td style={{ padding: '10px 12px' }}>
                                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                            <span style={{
                                                padding: '2px 6px',
                                                borderRadius: '4px',
                                                background: 'rgba(255, 255, 255, 0.1)',
                                                fontSize: '10.5px',
                                                fontWeight: 700,
                                                color: '#38bdf8'
                                            }}>
                                                {(item.ext || '').toUpperCase()}
                                            </span>
                                            <span style={{ fontWeight: 600, color: 'var(--text-primary)' }}>{item.filename}</span>
                                            <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>({(item.fileSize / 1024).toFixed(1)}KB)</span>
                                        </div>
                                    </td>
                                    <td style={{ padding: '10px 12px' }}>
                                        <span style={{
                                            display: 'inline-flex',
                                            alignItems: 'center',
                                            gap: '5px',
                                            padding: '3px 9px',
                                            borderRadius: '12px',
                                            fontSize: '11.5px',
                                            fontWeight: 700,
                                            background: `${STATUS_COLORS[item.status]}20`,
                                            color: STATUS_COLORS[item.status],
                                            border: `1px solid ${STATUS_COLORS[item.status]}40`
                                        }}>
                                            <span style={{ width: '6px', height: '6px', borderRadius: '50%', background: STATUS_COLORS[item.status] }} />
                                            {STATUS_LABELS[item.status]}
                                        </span>
                                    </td>
                                    <td style={{ padding: '10px 12px' }}>
                                        <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                                            <span style={{ fontSize: '11.5px', color: 'var(--text-secondary)' }}>{item.statusMessage}</span>
                                            {item.status !== BATCH_ITEM_STATUS.PENDING && item.status !== BATCH_ITEM_STATUS.COMPLETED && (
                                                <div style={{ width: '100%', height: '4px', background: 'rgba(255,255,255,0.06)', borderRadius: '2px', overflow: 'hidden' }}>
                                                    <div style={{
                                                        width: `${item.progress}%`,
                                                        height: '100%',
                                                        background: 'linear-gradient(90deg, #ec4899, #8b5cf6)',
                                                        transition: 'width 0.3s ease'
                                                    }} />
                                                </div>
                                            )}
                                        </div>
                                    </td>
                                    <td style={{ padding: '10px 12px', textAlign: 'center' }}>
                                        {item.status === BATCH_ITEM_STATUS.COMPLETED ? (
                                            <button
                                                onClick={() => setInspectingItem(item)}
                                                style={{
                                                    padding: '3px 8px',
                                                    borderRadius: '6px',
                                                    border: '1px solid rgba(248, 113, 113, 0.4)',
                                                    background: item.totalDefectsCount > 0 ? 'rgba(239, 68, 68, 0.15)' : 'rgba(16, 185, 129, 0.15)',
                                                    color: item.totalDefectsCount > 0 ? '#f87171' : '#34d399',
                                                    fontSize: '11.5px',
                                                    fontWeight: 700,
                                                    cursor: 'pointer'
                                                }}
                                            >
                                                {item.totalDefectsCount > 0 ? `결함 ${item.totalDefectsCount}건` : '결함 0건 (완벽)'}
                                            </button>
                                        ) : item.status === BATCH_ITEM_STATUS.ERROR ? (
                                            <span style={{ fontSize: '11px', color: '#ef4444' }}>오류 발생</span>
                                        ) : (
                                            <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>-</span>
                                        )}
                                    </td>
                                    <td style={{ padding: '10px 12px', textAlign: 'center' }}>
                                        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '6px' }}>
                                            {item.status === BATCH_ITEM_STATUS.COMPLETED && (
                                                <button
                                                    onClick={() => handleDownloadSingleDocx(item)}
                                                    title="Word 교정 리포트 다운로드"
                                                    style={{
                                                        padding: '4px 7px',
                                                        borderRadius: '6px',
                                                        border: '1px solid rgba(59, 130, 246, 0.3)',
                                                        background: 'rgba(59, 130, 246, 0.1)',
                                                        color: '#60a5fa',
                                                        cursor: 'pointer'
                                                    }}
                                                >
                                                    <FileText size={12} />
                                                </button>
                                            )}
                                            <button
                                                onClick={() => handleRemoveItem(item.id)}
                                                title="항목 삭제"
                                                style={{
                                                    padding: '4px 7px',
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

            {/* ── 브라우저 내장 RAG 가이드라인 관리 모달 ──────────────── */}
            {showRagModal && (
                <div style={{
                    position: 'fixed',
                    top: 0,
                    left: 0,
                    width: '100%',
                    height: '100%',
                    background: 'rgba(0, 0, 0, 0.75)',
                    backdropFilter: 'blur(8px)',
                    zIndex: 10000,
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    padding: '20px'
                }}>
                    <div className="glass-panel" style={{
                        width: '100%',
                        maxWidth: '850px',
                        maxHeight: '90vh',
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
                            alignItems: 'center',
                            background: 'rgba(255, 255, 255, 0.02)'
                        }}>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                <Database size={18} color="#818cf8" />
                                <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 700, color: '#fff' }}>
                                    브라우저 내장형 RAG 가이드라인 인덱서 (IndexedDB)
                                </h3>
                            </div>
                            <button
                                onClick={() => setShowRagModal(false)}
                                style={{ background: 'none', border: 'none', color: 'var(--text-muted)', cursor: 'pointer', padding: '4px' }}
                            >
                                <X size={18} />
                            </button>
                        </div>

                        <div style={{ padding: '20px', overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '16px' }}>
                            <div style={{
                                padding: '12px 16px',
                                borderRadius: '10px',
                                background: 'rgba(99, 102, 241, 0.08)',
                                border: '1px solid rgba(99, 102, 241, 0.25)',
                                fontSize: '12.5px',
                                color: '#cbd5e1',
                                lineHeight: 1.6
                            }}>
                                💡 <strong>로컬 스크립트/Python 의존성 제로</strong>: Vercel 배포 환경에서도 브라우저가 직접 PDF, HWPX, DOCX 문서를 800자 단위로 스마트 청킹하여 브라우저 로컬 저장소(IndexedDB)에 저장합니다. 검수 시 질문 및 장표 키워드와 가장 관련도가 높은 <strong>상위 5개 청크(BM25)</strong>를 AI 프롬프트에 자동 주입합니다.
                            </div>

                            {/* 가이드라인 파일 등록 버튼 */}
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                    <input
                                        ref={ragFileInputRef}
                                        type="file"
                                        accept=".pdf,.hwpx,.docx,.xlsx,.txt,.md"
                                        style={{ display: 'none' }}
                                        onChange={handleUploadRagGuideline}
                                    />
                                    <button
                                        onClick={() => ragFileInputRef.current?.click()}
                                        disabled={isIndexingFile}
                                        style={{
                                            padding: '8px 14px',
                                            borderRadius: '8px',
                                            border: 'none',
                                            background: 'linear-gradient(135deg, #6366f1, #8b5cf6)',
                                            color: '#fff',
                                            fontSize: '12.5px',
                                            fontWeight: 700,
                                            cursor: isIndexingFile ? 'not-allowed' : 'pointer',
                                            display: 'flex',
                                            alignItems: 'center',
                                            gap: '6px'
                                        }}
                                    >
                                        <Plus size={14} /> 기준 가이드라인 문서 등록
                                    </button>
                                </div>

                                {indexedDocs.length > 0 && (
                                    <button
                                        onClick={handleClearAllRag}
                                        style={{ background: 'none', border: 'none', color: '#f87171', fontSize: '12px', cursor: 'pointer' }}
                                    >
                                        전체 초기화
                                    </button>
                                )}
                            </div>

                            {/* 인덱싱 진행 바 */}
                            {isIndexingFile && (
                                <div style={{ padding: '12px', borderRadius: '8px', background: 'rgba(255,255,255,0.05)', border: '1px solid rgba(255,255,255,0.1)' }}>
                                    <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '12px', marginBottom: '4px' }}>
                                        <span style={{ color: '#818cf8' }}>{indexProgress.message}</span>
                                        <span>{indexProgress.percent}%</span>
                                    </div>
                                    <div style={{ width: '100%', height: '6px', background: 'rgba(0,0,0,0.3)', borderRadius: '3px', overflow: 'hidden' }}>
                                        <div style={{ width: `${indexProgress.percent}%`, height: '100%', background: '#6366f1', transition: 'width 0.3s ease' }} />
                                    </div>
                                </div>
                            )}

                            {/* 등록된 가이드라인 문서 목록 */}
                            <div>
                                <h4 style={{ margin: '0 0 8px', fontSize: '13px', color: 'var(--text-secondary)' }}>
                                    현재 등록된 가이드라인 문서 ({indexedDocs.length}건)
                                </h4>
                                {indexedDocs.length === 0 ? (
                                    <div style={{ padding: '20px', textAlign: 'center', color: 'var(--text-muted)', fontSize: '12.5px', border: '1px dashed var(--panel-border)', borderRadius: '8px' }}>
                                        등록된 가이드라인 문서가 없습니다. ISP 수립 공통가이드 또는 제안요청서를 업로드해 보세요.
                                    </div>
                                ) : (
                                    <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                                        {indexedDocs.map(doc => (
                                            <div key={doc.id} style={{
                                                padding: '10px 14px',
                                                borderRadius: '8px',
                                                background: 'rgba(255, 255, 255, 0.03)',
                                                border: '1px solid rgba(255, 255, 255, 0.08)',
                                                display: 'flex',
                                                justifyContent: 'space-between',
                                                alignItems: 'center'
                                            }}>
                                                <div>
                                                    <span style={{ fontWeight: 600, color: '#fff', fontSize: '13px' }}>{doc.filename}</span>
                                                    <div style={{ fontSize: '11px', color: 'var(--text-muted)', marginTop: '2px' }}>
                                                        청크: {doc.chunkCount}개 ({doc.totalLength?.toLocaleString()}자) | {new Date(doc.uploadedAt).toLocaleDateString()}
                                                    </div>
                                                </div>
                                                <button
                                                    onClick={() => handleDeleteRagDoc(doc.id)}
                                                    style={{ background: 'none', border: 'none', color: '#f87171', cursor: 'pointer', padding: '4px' }}
                                                >
                                                    <Trash2 size={13} />
                                                </button>
                                            </div>
                                        ))}
                                    </div>
                                )}
                            </div>

                            {/* BM25 검색 테스트 */}
                            <div style={{ borderTop: '1px solid var(--panel-border)', paddingTop: '16px' }}>
                                <h4 style={{ margin: '0 0 8px', fontSize: '13px', color: 'var(--text-secondary)' }}>
                                    실시간 BM25 하이브리드 검색 시뮬레이션
                                </h4>
                                <div style={{ display: 'flex', gap: '8px', marginBottom: '10px' }}>
                                    <input
                                        type="text"
                                        value={ragTestQuery}
                                        onChange={(e) => setRagTestQuery(e.target.value)}
                                        onKeyDown={(e) => e.key === 'Enter' && handleTestSearchRag()}
                                        placeholder="검색할 질문 또는 키워드 입력..."
                                        style={{
                                            flex: 1,
                                            padding: '8px 12px',
                                            borderRadius: '8px',
                                            border: '1px solid var(--panel-border)',
                                            background: 'rgba(0,0,0,0.3)',
                                            color: '#fff',
                                            fontSize: '12.5px'
                                        }}
                                    />
                                    <button
                                        onClick={handleTestSearchRag}
                                        disabled={isSearchingRag}
                                        style={{
                                            padding: '8px 16px',
                                            borderRadius: '8px',
                                            border: 'none',
                                            background: '#3b82f6',
                                            color: '#fff',
                                            fontSize: '12.5px',
                                            fontWeight: 600,
                                            cursor: 'pointer'
                                        }}
                                    >
                                        검색 (Top-5)
                                    </button>
                                </div>

                                {ragTestResults.length > 0 && (
                                    <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                                        {ragTestResults.map((chunk, i) => (
                                            <div key={chunk.id} style={{
                                                padding: '10px 12px',
                                                borderRadius: '8px',
                                                background: 'rgba(0,0,0,0.25)',
                                                border: '1px solid rgba(59, 130, 246, 0.2)',
                                                fontSize: '12px'
                                            }}>
                                                <div style={{ display: 'flex', justifyContent: 'space-between', color: '#60a5fa', fontWeight: 600, marginBottom: '4px' }}>
                                                    <span>#{i + 1} {chunk.filename} (청크 #{chunk.chunkIndex})</span>
                                                    <span>적합도: {chunk.score}점</span>
                                                </div>
                                                <div style={{ color: '#cbd5e1', lineHeight: 1.5, whiteSpace: 'pre-wrap', maxHeight: '90px', overflowY: 'auto' }}>
                                                    {chunk.text}
                                                </div>
                                            </div>
                                        ))}
                                    </div>
                                )}
                            </div>
                        </div>
                    </div>
                </div>
            )}

            {/* ── 결함 상세 조회 모달 ─────────────────────────────────── */}
            {inspectingItem && (
                <div style={{
                    position: 'fixed',
                    top: 0,
                    left: 0,
                    width: '100%',
                    height: '100%',
                    background: 'rgba(0, 0, 0, 0.75)',
                    backdropFilter: 'blur(8px)',
                    zIndex: 10000,
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    padding: '20px'
                }}>
                    <div className="glass-panel" style={{
                        width: '100%',
                        maxWidth: '900px',
                        maxHeight: '90vh',
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
                            <div>
                                <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 700, color: '#fff' }}>
                                    {inspectingItem.filename} - 결함 상세 분석 ({inspectingItem.totalDefectsCount}건)
                                </h3>
                                <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>
                                    소요시간: {(inspectingItem.durationMs / 1000).toFixed(1)}초 | Rule 검증: {inspectingItem.ruleDefectsCount}건 | AI 검증: {inspectingItem.aiDefectsCount}건
                                </span>
                            </div>
                            <button
                                onClick={() => setInspectingItem(null)}
                                style={{ background: 'none', border: 'none', color: 'var(--text-muted)', cursor: 'pointer', padding: '4px' }}
                            >
                                <X size={18} />
                            </button>
                        </div>

                        <div style={{ padding: '16px 20px', overflowY: 'auto', flex: 1 }}>
                            {inspectingItem.allIssues.length === 0 ? (
                                <div style={{ padding: '40px', textAlign: 'center', color: '#10b981', fontSize: '14px', fontWeight: 600 }}>
                                    ✓ 위반 사항이나 결함이 발견되지 않은 완벽한 산출물입니다.
                                </div>
                            ) : (
                                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px', textAlign: 'left' }}>
                                    <thead>
                                        <tr style={{ background: 'rgba(255,255,255,0.04)', borderBottom: '1px solid var(--panel-border)', color: 'var(--text-secondary)' }}>
                                            <th style={{ padding: '8px 10px', width: '40px' }}>No</th>
                                            <th style={{ padding: '8px 10px', width: '90px' }}>단계</th>
                                            <th style={{ padding: '8px 10px', width: '100px' }}>유형</th>
                                            <th style={{ padding: '8px 10px', width: '50px' }}>위치</th>
                                            <th style={{ padding: '8px 10px', width: '30%' }}>원문 (오표기/결함)</th>
                                            <th style={{ padding: '8px 10px', width: '25%' }}>수정 권장안</th>
                                            <th style={{ padding: '8px 10px' }}>사유 및 근거</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {inspectingItem.allIssues.map((iss, i) => (
                                            <tr key={i} style={{ borderBottom: '1px solid rgba(255,255,255,0.05)' }}>
                                                <td style={{ padding: '8px 10px', color: 'var(--text-muted)' }}>{i + 1}</td>
                                                <td style={{ padding: '8px 10px', color: '#a78bfa' }}>{iss.stage?.includes('Stage 1') ? 'Rule 검증' : 'AI 문맥검증'}</td>
                                                <td style={{ padding: '8px 10px', color: '#f87171', fontWeight: 600 }}>{iss.category || iss.errorType}</td>
                                                <td style={{ padding: '8px 10px' }}>p.{iss.page || 1}</td>
                                                <td style={{ padding: '8px 10px', color: '#fca5a5' }}>{iss.originalText || iss.matchedTerm}</td>
                                                <td style={{ padding: '8px 10px', color: '#86efac', fontWeight: 600 }}>{iss.suggestedText || '-'}</td>
                                                <td style={{ padding: '8px 10px', color: 'var(--text-secondary)', fontSize: '11px' }}>{iss.reason || '-'}</td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            )}
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
}
