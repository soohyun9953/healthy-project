import React, { useState, useEffect, useMemo, useRef } from 'react';
import {
    FileCheck2,
    Upload,
    Download,
    Plus,
    Trash2,
    RefreshCw,
    Sparkles,
    AlertTriangle,
    CheckCircle2,
    XCircle,
    Search,
    ChevronDown,
    ChevronUp,
    FileSpreadsheet,
    FileText,
    Settings,
    Layers,
    ListChecks,
    BookOpen,
    HelpCircle,
    Filter,
    ShieldAlert,
    ExternalLink,
    Maximize2
} from 'lucide-react';
import * as XLSX from 'xlsx';
import {
    loadSSOTData,
    saveSSOTData,
    DEFAULT_PROJECT_METADATA,
    DEFAULT_GLOSSARY,
    DEFAULT_REQUIREMENTS,
    extractTablesFromFile,
    runStage1Validation,
    runStage2AIValidation,
    exportSSOTReportToExcel
} from '../utils/ssotValidator';
import { extractStructuredFileContent } from '../utils/fileExtractor';

export default function SSOTValidatorView({ apiKey, llmProvider, omniRouteModel }) {
    // ── 마스터 데이터 상태 (LocalStorage 영속화) ──────────────────────
    const [ssotData, setSsotData] = useState(() => loadSSOTData());
    const [isConfigExpanded, setIsConfigExpanded] = useState(false); // 마스터 설정창 접기/펼치기
    const [configTab, setConfigTab] = useState('metadata'); // 'metadata' | 'glossary' | 'requirements'

    // 변경 시 LocalStorage 즉시 저장
    useEffect(() => {
        saveSSOTData(ssotData);
    }, [ssotData]);

    // ── 검수 대상 파일 및 실행 상태 ──────────────────────────────
    const [targetFile, setTargetFile] = useState(null);
    const [isExtracting, setIsExtracting] = useState(false);
    const [isValidating, setIsValidating] = useState(false);
    const [isDragging, setIsDragging] = useState(false);

    // 검증 옵션
    const [runStage1, setRunStage1] = useState(true);
    const [runStage2, setRunStage2] = useState(false);
    const [selectedModel, setSelectedModel] = useState('gemini-2.5-flash');

    // AI 진행 상태
    const [aiProgress, setAiProgress] = useState({ current: 0, total: 0, percent: 0, message: '' });

    // 검증 결과 상태
    const [validationResult, setValidationResult] = useState(null);
    // { fileName, stage1, stage2Issues, allIssues }

    // 결과 뷰 필터 및 검색
    const [activeResultTab, setActiveResultTab] = useState('all'); // 'all' | 'requirements' | 'glossary' | 'numbers' | 'ai' | 'rtm'
    const [searchQuery, setSearchQuery] = useState('');
    const [severityFilter, setSeverityFilter] = useState('ALL'); // 'ALL' | 'HIGH' | 'MEDIUM' | 'LOW'

    // 텍스트 일괄 붙여넣기 모달
    const [showBulkReqModal, setShowBulkReqModal] = useState(false);
    const [bulkReqText, setBulkReqText] = useState('');

    const fileInputRef = useRef(null);
    const glossaryExcelInputRef = useRef(null);
    const reqExcelInputRef = useRef(null);

    // ── 마스터 데이터 변경 핸들러 ────────────────────────────────
    const handleMetadataChange = (field, value) => {
        setSsotData(prev => ({
            ...prev,
            metadata: {
                ...prev.metadata,
                [field]: value
            }
        }));
    };

    const handleResetPreset = () => {
        if (window.confirm('공공의료/ISP 표준 기본 프리셋으로 초기화하시겠습니까? 현재 입력된 사용자 항목이 기본값으로 대체됩니다.')) {
            setSsotData({
                metadata: { ...DEFAULT_PROJECT_METADATA },
                glossary: [...DEFAULT_GLOSSARY],
                requirements: [...DEFAULT_REQUIREMENTS]
            });
        }
    };

    // 용어사전 행 조작
    const handleAddGlossaryItem = () => {
        const newItem = {
            id: `glo-${Date.now()}`,
            category: '행정용어',
            approved: '',
            forbidden: '',
            note: ''
        };
        setSsotData(prev => ({
            ...prev,
            glossary: [newItem, ...prev.glossary]
        }));
    };

    const handleUpdateGlossaryItem = (index, field, value) => {
        setSsotData(prev => {
            const next = [...prev.glossary];
            next[index] = { ...next[index], [field]: value };
            return { ...prev, glossary: next };
        });
    };

    const handleDeleteGlossaryItem = (index) => {
        setSsotData(prev => {
            const next = prev.glossary.filter((_, idx) => idx !== index);
            return { ...prev, glossary: next };
        });
    };

    // 용어사전 엑셀 다운로드 / 업로드
    const handleExportGlossaryExcel = () => {
        const rows = ssotData.glossary.map((g, idx) => ({
            '순번': idx + 1,
            '분류': g.category || '행정용어',
            '표준용어(Approved)': g.approved || '',
            '금지동의어/오표기(Forbidden, 쉼표구분)': g.forbidden || '',
            '비고/설명': g.note || ''
        }));
        const worksheet = XLSX.utils.json_to_sheet(rows);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, '표준용어사전');
        XLSX.writeFile(workbook, `프로젝트_표준용어사전_${new Date().toISOString().slice(0, 10)}.xlsx`);
    };

    const handleImportGlossaryExcel = (e) => {
        const file = e.target.files?.[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (evt) => {
            try {
                const data = new Uint8Array(evt.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet);

                const imported = jsonData.map((r, i) => ({
                    id: `glo-imp-${Date.now()}-${i}`,
                    category: r['분류'] || r['카테고리'] || '행정용어',
                    approved: (r['표준용어(Approved)'] || r['표준용어'] || r['표준'] || '').trim(),
                    forbidden: (r['금지동의어/오표기(Forbidden, 쉼표구분)'] || r['금지동의어'] || r['금지어'] || r['오표기'] || '').trim(),
                    note: (r['비고/설명'] || r['비고'] || '').trim()
                })).filter(g => g.approved.length > 0);

                if (imported.length > 0) {
                    setSsotData(prev => ({
                        ...prev,
                        glossary: [...imported, ...prev.glossary]
                    }));
                    alert(`성공적으로 ${imported.length}건의 표준용어를 추가 등록했습니다.`);
                } else {
                    alert('유효한 표준용어 데이터를 찾을 수 없습니다. 양식을 확인해주세요.');
                }
            } catch (err) {
                console.error(err);
                alert('엑셀 파일 파싱 중 오류가 발생했습니다.');
            }
        };
        reader.readAsArrayBuffer(file);
        e.target.value = '';
    };

    // 요구사항 행 조작
    const handleAddRequirement = () => {
        const newReq = {
            id: `SFR-${String(ssotData.requirements.length + 1).padStart(3, '0')}`,
            name: '',
            keywords: ''
        };
        setSsotData(prev => ({
            ...prev,
            requirements: [newReq, ...prev.requirements]
        }));
    };

    const handleUpdateRequirement = (index, field, value) => {
        setSsotData(prev => {
            const next = [...prev.requirements];
            next[index] = { ...next[index], [field]: value };
            return { ...prev, requirements: next };
        });
    };

    const handleDeleteRequirement = (index) => {
        setSsotData(prev => {
            const next = prev.requirements.filter((_, idx) => idx !== index);
            return { ...prev, requirements: next };
        });
    };

    // 요구사항 텍스트 일괄 파싱 등록
    const handleApplyBulkRequirements = () => {
        if (!bulkReqText.trim()) return;
        const lines = bulkReqText.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
        const parsed = [];

        lines.forEach((line, idx) => {
            // "SFR-001 | 명칭 | 키워드" 또는 탭/쉼표 구분
            const parts = line.split(/[|\t,]/).map(p => p.trim());
            if (parts.length >= 2) {
                parsed.push({
                    id: parts[0].toUpperCase(),
                    name: parts[1],
                    keywords: parts[2] || ''
                });
            } else if (parts.length === 1 && /^[A-Z]{2,4}-\d{3,4}/i.test(parts[0])) {
                parsed.push({
                    id: parts[0].toUpperCase(),
                    name: `요구사항 ${parts[0]}`,
                    keywords: ''
                });
            }
        });

        if (parsed.length > 0) {
            setSsotData(prev => ({
                ...prev,
                requirements: [...parsed, ...prev.requirements]
            }));
            setShowBulkReqModal(false);
            setBulkReqText('');
            alert(`총 ${parsed.length}건의 요구사항이 마스터에 등록되었습니다.`);
        } else {
            alert('구분자(| 또는 탭 또는 쉼표)로 [요구사항ID | 명칭 | 키워드] 형식으로 입력해주세요.');
        }
    };

    // ── 파일 드롭 및 선택 ────────────────────────────────────────
    const handleFileDrop = (e) => {
        e.preventDefault();
        e.stopPropagation();
        setIsDragging(false);
        const files = Array.from(e.dataTransfer.files);
        if (files.length > 0) {
            setTargetFile(files[0]);
            setValidationResult(null);
        }
    };

    const handleFileSelect = (e) => {
        const file = e.target.files?.[0];
        if (file) {
            setTargetFile(file);
            setValidationResult(null);
        }
        e.target.value = '';
    };

    // ── 검증 실행 ───────────────────────────────────────────────
    const handleStartValidation = async () => {
        if (!targetFile) {
            alert('검수할 문서를 먼저 업로드해주세요.');
            return;
        }

        if (!runStage1 && !runStage2) {
            alert('최소 1개 이상의 검증 단계(Stage 1 또는 Stage 2)를 선택해야 합니다.');
            return;
        }

        if (runStage2 && !apiKey) {
            alert('Stage 2 AI 정밀 검증을 실행하려면 상단에 Gemini API Key가 등록되어 있어야 합니다.');
            return;
        }

        setIsValidating(true);
        setIsExtracting(true);
        setAiProgress({ current: 0, total: 0, percent: 0, message: '문서 구조 및 텍스트/표 데이터 추출 중...' });

        try {
            // 1. 문서 텍스트 및 페이지 구조 추출
            const fileContent = await extractStructuredFileContent(targetFile);
            // 2. 표 데이터(2D 배열) 추출
            const tables = await extractTablesFromFile(targetFile);
            setIsExtracting(false);

            let stage1Result = {
                issues: [],
                rtmMatrix: [],
                rtmMappingRate: 100,
                coveredReqCount: 0,
                totalMasterReqCount: ssotData.requirements.length
            };

            // 3. Stage 1 결정론적 Rule 검증
            if (runStage1) {
                setAiProgress({ current: 0, total: 100, percent: 30, message: 'Stage 1 결정론적 룰 및 RTM/수치 정합성 검증 중...' });
                stage1Result = await runStage1Validation(fileContent, tables, ssotData);
            }

            // 4. Stage 2 Gemini AI 문맥·오탈자 검증
            let stage2Issues = [];
            if (runStage2) {
                stage2Issues = await runStage2AIValidation({
                    fileContent,
                    masterData: ssotData,
                    apiKey,
                    modelId: selectedModel,
                    onProgress: (p) => setAiProgress(p)
                });
            }

            const allIssues = [...stage1Result.issues, ...stage2Issues];

            setValidationResult({
                fileName: targetFile.name,
                fileSize: targetFile.size,
                pages: fileContent.pages,
                stage1: stage1Result,
                stage2Issues,
                allIssues
            });
        } catch (err) {
            console.error('검증 중 오류 발생:', err);
            alert(`검증 중 오류가 발생했습니다: ${err.message || '알 수 없는 오류'}`);
        } finally {
            setIsValidating(false);
            setIsExtracting(false);
            setAiProgress({ current: 0, total: 0, percent: 100, message: '검증 완료' });
        }
    };

    // ── 리포트 다운로드 ──────────────────────────────────────────
    const handleDownloadExcelReport = async () => {
        if (!validationResult) return;
        try {
            await exportSSOTReportToExcel({
                fileName: validationResult.fileName,
                masterData: ssotData,
                stage1Result: validationResult.stage1,
                allIssues: validationResult.allIssues
            });
        } catch (err) {
            console.error('엑셀 리포트 생성 실패:', err);
            alert('리포트 엑셀 생성 중 오류가 발생했습니다.');
        }
    };

    // ── 결과 필터링 ─────────────────────────────────────────────
    const filteredIssues = useMemo(() => {
        if (!validationResult) return [];
        return validationResult.allIssues.filter(issue => {
            // 탭 필터
            if (activeResultTab === 'requirements' && !issue.category.includes('요구사항')) return false;
            if (activeResultTab === 'glossary' && issue.category !== '용어 불일치') return false;
            if (activeResultTab === 'numbers' && issue.category !== '수치 오류') return false;
            if (activeResultTab === 'ai' && !issue.stage.includes('Stage 2')) return false;

            // 심각도 필터
            if (severityFilter !== 'ALL' && issue.severity !== severityFilter) return false;

            // 검색어 필터
            if (searchQuery.trim()) {
                const q = searchQuery.toLowerCase();
                const matched = (issue.matchedTerm || '').toLowerCase().includes(q) ||
                    (issue.originalText || '').toLowerCase().includes(q) ||
                    (issue.suggestedText || '').toLowerCase().includes(q) ||
                    (issue.reason || '').toLowerCase().includes(q) ||
                    (issue.errorType || '').toLowerCase().includes(q);
                if (!matched) return false;
            }

            return true;
        });
    }, [validationResult, activeResultTab, severityFilter, searchQuery]);

    // 요약 지표 통계
    const stats = useMemo(() => {
        if (!validationResult) {
            return { total: 0, rtmRate: 100, glossary: 0, aiOrNum: 0 };
        }
        const issues = validationResult.allIssues;
        const glossaryCount = issues.filter(i => i.category === '용어 불일치').length;
        const aiCount = issues.filter(i => i.stage.includes('Stage 2') || i.category === '수치 오류').length;
        return {
            total: issues.length,
            rtmRate: validationResult.stage1.rtmMappingRate,
            glossary: glossaryCount,
            aiOrNum: aiCount
        };
    }, [validationResult]);

    return (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
            {/* 상단 타이틀 배너 */}
            <div className="glass-panel" style={{ padding: '20px 24px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '16px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '14px' }}>
                    <div style={{
                        width: '46px',
                        height: '46px',
                        borderRadius: '12px',
                        background: 'linear-gradient(135deg, rgba(16, 185, 129, 0.2), rgba(59, 130, 246, 0.2))',
                        border: '1px solid rgba(16, 185, 129, 0.4)',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        color: '#10b981'
                    }}>
                        <FileCheck2 size={26} />
                    </div>
                    <div>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                            <h2 style={{ fontSize: '18px', fontWeight: 700, margin: 0, color: 'var(--text-primary)' }}>
                                프로젝트 표준 용어 & 요구사항 ID/수치 정합성 검증기 (SSOT)
                            </h2>
                            <span style={{
                                padding: '2px 8px',
                                borderRadius: '12px',
                                fontSize: '11px',
                                fontWeight: 700,
                                background: 'rgba(16, 185, 129, 0.15)',
                                color: '#10b981',
                                border: '1px solid rgba(16, 185, 129, 0.3)'
                            }}>
                                SSOT Validator
                            </span>
                        </div>
                        <p style={{ margin: '4px 0 0', fontSize: '13px', color: 'var(--text-secondary)' }}>
                            단일 진실 원천(Single Source of Truth) 기반으로 제안서·보고서 내 금지용어, 요구사항 ID 누락(RTM), 표 산술 덧셈 및 AI 문맥 모순을 전수 검증합니다.
                        </p>
                    </div>
                </div>

                <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                    <button
                        onClick={handleResetPreset}
                        className="btn-outline"
                        style={{ padding: '8px 14px', fontSize: '12.5px', display: 'flex', alignItems: 'center', gap: '6px' }}
                        title="공공의료/ISP 표준 기본 프리셋으로 초기화"
                    >
                        <RefreshCw size={14} /> 프리셋 복원
                    </button>
                    <button
                        onClick={() => setIsConfigExpanded(!isConfigExpanded)}
                        className="btn-primary"
                        style={{
                            padding: '8px 16px',
                            fontSize: '13px',
                            display: 'flex',
                            alignItems: 'center',
                            gap: '6px',
                            background: isConfigExpanded ? 'var(--accent-purple)' : undefined
                        }}
                    >
                        <Settings size={15} />
                        {isConfigExpanded ? '마스터 설정 닫기' : '마스터 & 용어사전 관리'}
                        {isConfigExpanded ? <ChevronUp size={15} /> : <ChevronDown size={15} />}
                    </button>
                </div>
            </div>

            {/* 마스터 메타데이터 & 표준용어 & 요구사항 관리 패널 (접기/펼치기) */}
            {isConfigExpanded && (
                <div className="glass-panel" style={{ padding: '20px', display: 'flex', flexDirection: 'column', gap: '16px', border: '1px solid rgba(168, 85, 247, 0.3)' }}>
                    {/* 탭 네비게이션 */}
                    <div style={{ display: 'flex', gap: '8px', borderBottom: '1px solid var(--panel-border)', paddingBottom: '12px' }}>
                        <button
                            onClick={() => setConfigTab('metadata')}
                            className={`btn-tab ${configTab === 'metadata' ? 'active' : ''}`}
                            style={{
                                padding: '6px 14px',
                                borderRadius: '8px',
                                fontSize: '13px',
                                fontWeight: 600,
                                background: configTab === 'metadata' ? 'var(--accent-blue)' : 'rgba(255,255,255,0.05)',
                                color: configTab === 'metadata' ? '#fff' : 'var(--text-secondary)',
                                border: 'none',
                                cursor: 'pointer'
                            }}
                        >
                            ① 프로젝트 기본 메타정보
                        </button>
                        <button
                            onClick={() => setConfigTab('glossary')}
                            className={`btn-tab ${configTab === 'glossary' ? 'active' : ''}`}
                            style={{
                                padding: '6px 14px',
                                borderRadius: '8px',
                                fontSize: '13px',
                                fontWeight: 600,
                                background: configTab === 'glossary' ? 'var(--accent-blue)' : 'rgba(255,255,255,0.05)',
                                color: configTab === 'glossary' ? '#fff' : 'var(--text-secondary)',
                                border: 'none',
                                cursor: 'pointer'
                            }}
                        >
                            ② 표준 용어사전 ({ssotData.glossary.length}건)
                        </button>
                        <button
                            onClick={() => setConfigTab('requirements')}
                            className={`btn-tab ${configTab === 'requirements' ? 'active' : ''}`}
                            style={{
                                padding: '6px 14px',
                                borderRadius: '8px',
                                fontSize: '13px',
                                fontWeight: 600,
                                background: configTab === 'requirements' ? 'var(--accent-blue)' : 'rgba(255,255,255,0.05)',
                                color: configTab === 'requirements' ? '#fff' : 'var(--text-secondary)',
                                border: 'none',
                                cursor: 'pointer'
                            }}
                        >
                            ③ 요구사항 마스터 리스트 ({ssotData.requirements.length}건)
                        </button>
                    </div>

                    {/* 탭 1: 프로젝트 메타정보 */}
                    {configTab === 'metadata' && (
                        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: '14px' }}>
                            <div>
                                <label style={{ display: 'block', fontSize: '12px', fontWeight: 600, color: 'var(--text-muted)', marginBottom: '4px' }}>
                                    사업명 (공식 프로젝트명)
                                </label>
                                <input
                                    type="text"
                                    value={ssotData.metadata.projectName || ''}
                                    onChange={(e) => handleMetadataChange('projectName', e.target.value)}
                                    placeholder="예: 공공보건의료 지능형 플랫폼 구축 ISP"
                                    style={{ width: '100%', padding: '8px 12px', borderRadius: '6px', border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)', color: 'var(--text-primary)', fontSize: '13px' }}
                                />
                            </div>
                            <div>
                                <label style={{ display: 'block', fontSize: '12px', fontWeight: 600, color: 'var(--text-muted)', marginBottom: '4px' }}>
                                    발주기관 (공식 명칭)
                                </label>
                                <input
                                    type="text"
                                    value={ssotData.metadata.clientAgency || ''}
                                    onChange={(e) => handleMetadataChange('clientAgency', e.target.value)}
                                    placeholder="예: 국립중앙의료원"
                                    style={{ width: '100%', padding: '8px 12px', borderRadius: '6px', border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)', color: 'var(--text-primary)', fontSize: '13px' }}
                                />
                            </div>
                            <div>
                                <label style={{ display: 'block', fontSize: '12px', fontWeight: 600, color: 'var(--text-muted)', marginBottom: '4px' }}>
                                    수행기간
                                </label>
                                <input
                                    type="text"
                                    value={ssotData.metadata.period || ''}
                                    onChange={(e) => handleMetadataChange('period', e.target.value)}
                                    placeholder="예: 2026.04 ~ 2026.11"
                                    style={{ width: '100%', padding: '8px 12px', borderRadius: '6px', border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)', color: 'var(--text-primary)', fontSize: '13px' }}
                                />
                            </div>
                            <div>
                                <label style={{ display: 'block', fontSize: '12px', fontWeight: 600, color: 'var(--text-muted)', marginBottom: '4px' }}>
                                    총사업비 (숫자 및 한글 표기)
                                </label>
                                <input
                                    type="text"
                                    value={ssotData.metadata.budget || ''}
                                    onChange={(e) => handleMetadataChange('budget', e.target.value)}
                                    placeholder="예: 500,000,000원 (금오억원)"
                                    style={{ width: '100%', padding: '8px 12px', borderRadius: '6px', border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)', color: 'var(--text-primary)', fontSize: '13px' }}
                                />
                            </div>
                        </div>
                    )}

                    {/* 탭 2: 표준 용어사전 테이블 */}
                    {configTab === 'glossary' && (
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '8px' }}>
                                <span style={{ fontSize: '12.5px', color: 'var(--text-secondary)' }}>
                                    💡 문서 내에서 금지동의어가 발견되면 즉시 표준용어로 치환 권고가 발생합니다.
                                </span>
                                <div style={{ display: 'flex', gap: '8px' }}>
                                    <input
                                        type="file"
                                        accept=".xlsx,.xls"
                                        ref={glossaryExcelInputRef}
                                        style={{ display: 'none' }}
                                        onChange={handleImportGlossaryExcel}
                                    />
                                    <button
                                        onClick={() => glossaryExcelInputRef.current?.click()}
                                        className="btn-outline"
                                        style={{ padding: '6px 10px', fontSize: '12px', display: 'flex', alignItems: 'center', gap: '5px' }}
                                    >
                                        <Upload size={13} /> 엑셀 일괄 업로드
                                    </button>
                                    <button
                                        onClick={handleExportGlossaryExcel}
                                        className="btn-outline"
                                        style={{ padding: '6px 10px', fontSize: '12px', display: 'flex', alignItems: 'center', gap: '5px' }}
                                    >
                                        <Download size={13} /> 용어사전 다운로드
                                    </button>
                                    <button
                                        onClick={handleAddGlossaryItem}
                                        className="btn-primary"
                                        style={{ padding: '6px 12px', fontSize: '12px', display: 'flex', alignItems: 'center', gap: '5px' }}
                                    >
                                        <Plus size={14} /> 단어 추가
                                    </button>
                                </div>
                            </div>

                            <div style={{ maxHeight: '280px', overflowY: 'auto', border: '1px solid var(--panel-border)', borderRadius: '6px' }}>
                                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12.5px' }}>
                                    <thead>
                                        <tr style={{ background: 'rgba(255,255,255,0.06)', borderBottom: '1px solid var(--panel-border)', textAlign: 'left' }}>
                                            <th style={{ padding: '8px 10px', width: '110px' }}>분류</th>
                                            <th style={{ padding: '8px 10px', width: '180px' }}>표준용어(Approved)</th>
                                            <th style={{ padding: '8px 10px' }}>금지동의어/오표기(Forbidden, 쉼표구분)</th>
                                            <th style={{ padding: '8px 10px', width: '220px' }}>비고</th>
                                            <th style={{ padding: '8px 10px', width: '50px', textAlign: 'center' }}>삭제</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {ssotData.glossary.map((item, idx) => (
                                            <tr key={item.id || idx} style={{ borderBottom: '1px solid rgba(255,255,255,0.05)' }}>
                                                <td style={{ padding: '6px 8px' }}>
                                                    <select
                                                        value={item.category || '행정용어'}
                                                        onChange={(e) => handleUpdateGlossaryItem(idx, 'category', e.target.value)}
                                                        style={{ width: '100%', padding: '4px', borderRadius: '4px', background: 'rgba(0,0,0,0.3)', border: '1px solid var(--panel-border)', color: 'var(--text-primary)', fontSize: '12px' }}
                                                    >
                                                        <option value="기관명">기관명</option>
                                                        <option value="시스템명">시스템명</option>
                                                        <option value="행정용어">행정용어</option>
                                                        <option value="약어">약어</option>
                                                        <option value="기술용어">기술용어</option>
                                                    </select>
                                                </td>
                                                <td style={{ padding: '6px 8px' }}>
                                                    <input
                                                        type="text"
                                                        value={item.approved}
                                                        onChange={(e) => handleUpdateGlossaryItem(idx, 'approved', e.target.value)}
                                                        placeholder="예: 지방의료원"
                                                        style={{ width: '100%', padding: '4px 8px', borderRadius: '4px', background: 'rgba(0,0,0,0.3)', border: '1px solid var(--panel-border)', color: '#10b981', fontWeight: 600, fontSize: '12px' }}
                                                    />
                                                </td>
                                                <td style={{ padding: '6px 8px' }}>
                                                    <input
                                                        type="text"
                                                        value={item.forbidden}
                                                        onChange={(e) => handleUpdateGlossaryItem(idx, 'forbidden', e.target.value)}
                                                        placeholder="예: 지역의료원, 도립병원"
                                                        style={{ width: '100%', padding: '4px 8px', borderRadius: '4px', background: 'rgba(0,0,0,0.3)', border: '1px solid var(--panel-border)', color: 'var(--danger-color)', fontSize: '12px' }}
                                                    />
                                                </td>
                                                <td style={{ padding: '6px 8px' }}>
                                                    <input
                                                        type="text"
                                                        value={item.note || ''}
                                                        onChange={(e) => handleUpdateGlossaryItem(idx, 'note', e.target.value)}
                                                        placeholder="사유 또는 법적 근거"
                                                        style={{ width: '100%', padding: '4px 8px', borderRadius: '4px', background: 'rgba(0,0,0,0.3)', border: '1px solid var(--panel-border)', color: 'var(--text-secondary)', fontSize: '12px' }}
                                                    />
                                                </td>
                                                <td style={{ padding: '6px 8px', textAlign: 'center' }}>
                                                    <button
                                                        onClick={() => handleDeleteGlossaryItem(idx)}
                                                        style={{ background: 'none', border: 'none', color: 'var(--text-muted)', cursor: 'pointer', padding: '4px' }}
                                                        title="삭제"
                                                    >
                                                        <Trash2 size={14} />
                                                    </button>
                                                </td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    )}

                    {/* 탭 3: 요구사항 마스터 리스트 */}
                    {configTab === 'requirements' && (
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '8px' }}>
                                <span style={{ fontSize: '12.5px', color: 'var(--text-secondary)' }}>
                                    💡 과업지시서/RFP 요구사항 ID(예: CNR-001)를 등록하면 산출물 내 누락(Missing ID) 및 유령(Ghost ID)을 검출합니다.
                                </span>
                                <div style={{ display: 'flex', gap: '8px' }}>
                                    <button
                                        onClick={() => setShowBulkReqModal(true)}
                                        className="btn-outline"
                                        style={{ padding: '6px 10px', fontSize: '12px', display: 'flex', alignItems: 'center', gap: '5px' }}
                                    >
                                        <FileText size={13} /> 텍스트 일괄 붙여넣기
                                    </button>
                                    <button
                                        onClick={handleAddRequirement}
                                        className="btn-primary"
                                        style={{ padding: '6px 12px', fontSize: '12px', display: 'flex', alignItems: 'center', gap: '5px' }}
                                    >
                                        <Plus size={14} /> 요구사항 추가
                                    </button>
                                </div>
                            </div>

                            <div style={{ maxHeight: '280px', overflowY: 'auto', border: '1px solid var(--panel-border)', borderRadius: '6px' }}>
                                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12.5px' }}>
                                    <thead>
                                        <tr style={{ background: 'rgba(255,255,255,0.06)', borderBottom: '1px solid var(--panel-border)', textAlign: 'left' }}>
                                            <th style={{ padding: '8px 10px', width: '130px' }}>요구사항 ID</th>
                                            <th style={{ padding: '8px 10px' }}>요구사항명</th>
                                            <th style={{ padding: '8px 10px', width: '280px' }}>필수 키워드 (쉼표 구분)</th>
                                            <th style={{ padding: '8px 10px', width: '50px', textAlign: 'center' }}>삭제</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {ssotData.requirements.map((req, idx) => (
                                            <tr key={req.id || idx} style={{ borderBottom: '1px solid rgba(255,255,255,0.05)' }}>
                                                <td style={{ padding: '6px 8px' }}>
                                                    <input
                                                        type="text"
                                                        value={req.id}
                                                        onChange={(e) => handleUpdateRequirement(idx, 'id', e.target.value)}
                                                        placeholder="예: SFR-001"
                                                        style={{ width: '100%', padding: '4px 8px', borderRadius: '4px', background: 'rgba(0,0,0,0.3)', border: '1px solid var(--panel-border)', color: 'var(--accent-blue)', fontWeight: 700, fontSize: '12px' }}
                                                    />
                                                </td>
                                                <td style={{ padding: '6px 8px' }}>
                                                    <input
                                                        type="text"
                                                        value={req.name}
                                                        onChange={(e) => handleUpdateRequirement(idx, 'name', e.target.value)}
                                                        placeholder="요구사항 상세 명칭"
                                                        style={{ width: '100%', padding: '4px 8px', borderRadius: '4px', background: 'rgba(0,0,0,0.3)', border: '1px solid var(--panel-border)', color: 'var(--text-primary)', fontSize: '12px' }}
                                                    />
                                                </td>
                                                <td style={{ padding: '6px 8px' }}>
                                                    <input
                                                        type="text"
                                                        value={req.keywords}
                                                        onChange={(e) => handleUpdateRequirement(idx, 'keywords', e.target.value)}
                                                        placeholder="예: 데이터, 연계, 표준"
                                                        style={{ width: '100%', padding: '4px 8px', borderRadius: '4px', background: 'rgba(0,0,0,0.3)', border: '1px solid var(--panel-border)', color: 'var(--text-secondary)', fontSize: '12px' }}
                                                    />
                                                </td>
                                                <td style={{ padding: '6px 8px', textAlign: 'center' }}>
                                                    <button
                                                        onClick={() => handleDeleteRequirement(idx)}
                                                        style={{ background: 'none', border: 'none', color: 'var(--text-muted)', cursor: 'pointer', padding: '4px' }}
                                                        title="삭제"
                                                    >
                                                        <Trash2 size={14} />
                                                    </button>
                                                </td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    )}
                </div>
            )}

            {/* 검수 대상 파일 업로드 및 검증 제어 패널 */}
            <div className="glass-panel" style={{ padding: '24px', display: 'flex', flexDirection: 'column', gap: '20px' }}>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: '12px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <Upload size={18} style={{ color: 'var(--accent-blue)' }} />
                        <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 700, color: 'var(--text-primary)' }}>
                            검수 대상 산출물 문서 업로드
                        </h3>
                        <span style={{ fontSize: '12px', color: 'var(--text-muted)' }}>
                            (.hwpx, .docx, .pptx, .xlsx, .pdf 지원)
                        </span>
                    </div>

                    {/* 검증 단계 선택 체크박스 */}
                    <div style={{ display: 'flex', alignItems: 'center', gap: '16px', background: 'rgba(0,0,0,0.2)', padding: '6px 14px', borderRadius: '8px', border: '1px solid var(--panel-border)' }}>
                        <label style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '12.5px', fontWeight: 600, cursor: 'pointer', color: 'var(--text-primary)' }}>
                            <input
                                type="checkbox"
                                checked={runStage1}
                                onChange={(e) => setRunStage1(e.target.checked)}
                                style={{ accentColor: '#10b981' }}
                            />
                            Stage 1: 결정론적 Rule 검증 (무료·초고속)
                        </label>
                        <label style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '12.5px', fontWeight: 600, cursor: 'pointer', color: 'var(--text-primary)' }}>
                            <input
                                type="checkbox"
                                checked={runStage2}
                                onChange={(e) => setRunStage2(e.target.checked)}
                                style={{ accentColor: '#a855f7' }}
                            />
                            Stage 2: Gemini AI 문맥·오탈자 검증
                        </label>
                        {runStage2 && (
                            <select
                                value={selectedModel}
                                onChange={(e) => setSelectedModel(e.target.value)}
                                style={{ padding: '3px 8px', borderRadius: '4px', background: 'rgba(0,0,0,0.4)', border: '1px solid var(--panel-border)', color: '#a855f7', fontSize: '11.5px', fontWeight: 600 }}
                            >
                                <option value="gemini-2.5-flash">Gemini 2.5 Flash (권장)</option>
                                <option value="gemini-2.0-flash">Gemini 2.0 Flash</option>
                                <option value="gemini-1.5-flash">Gemini 1.5 Flash</option>
                                <option value="gemini-1.5-pro">Gemini 1.5 Pro</option>
                            </select>
                        )}
                    </div>
                </div>

                {/* 드래그 앤 드롭 영역 */}
                <div
                    onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
                    onDragLeave={() => setIsDragging(false)}
                    onDrop={handleFileDrop}
                    onClick={() => fileInputRef.current?.click()}
                    style={{
                        border: `2px dashed ${isDragging ? '#10b981' : targetFile ? 'var(--accent-blue)' : 'var(--panel-border)'}`,
                        borderRadius: '12px',
                        padding: '30px 20px',
                        textAlign: 'center',
                        cursor: 'pointer',
                        background: isDragging ? 'rgba(16, 185, 129, 0.05)' : targetFile ? 'rgba(59, 130, 246, 0.04)' : 'rgba(255,255,255,0.02)',
                        transition: 'all 0.2s ease',
                        display: 'flex',
                        flexDirection: 'column',
                        alignItems: 'center',
                        gap: '10px'
                    }}
                >
                    <input
                        type="file"
                        ref={fileInputRef}
                        accept=".hwpx,.docx,.pptx,.xlsx,.xls,.pdf,.txt,.md"
                        style={{ display: 'none' }}
                        onChange={handleFileSelect}
                    />

                    {targetFile ? (
                        <>
                            <div style={{ width: '48px', height: '48px', borderRadius: '50%', background: 'rgba(59, 130, 246, 0.15)', display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'var(--accent-blue)' }}>
                                <FileSpreadsheet size={26} />
                            </div>
                            <div>
                                <span style={{ fontSize: '15px', fontWeight: 700, color: 'var(--text-primary)' }}>
                                    {targetFile.name}
                                </span>
                                <span style={{ marginLeft: '8px', fontSize: '12px', color: 'var(--text-muted)' }}>
                                    ({(targetFile.size / 1024).toFixed(1)} KB)
                                </span>
                            </div>
                            <span style={{ fontSize: '12px', color: '#10b981', fontWeight: 600 }}>
                                ✓ 검수 준비 완료. 아래 '정합성 검증 시작' 버튼을 눌러주세요. (클릭하여 다른 파일 선택)
                            </span>
                        </>
                    ) : (
                        <>
                            <div style={{ width: '48px', height: '48px', borderRadius: '50%', background: 'rgba(255,255,255,0.05)', display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'var(--text-muted)' }}>
                                <Upload size={24} />
                            </div>
                            <div>
                                <span style={{ fontSize: '14px', fontWeight: 600, color: 'var(--text-primary)' }}>
                                    클릭하여 검수 대상 파일을 선택하거나 이곳으로 드래그하세요
                                </span>
                                <p style={{ margin: '4px 0 0', fontSize: '12px', color: 'var(--text-muted)' }}>
                                    HWPX, DOCX, PPTX, XLSX, PDF 등 공공기관 표준 산출물 형식 지원
                                </p>
                            </div>
                        </>
                    )}
                </div>

                {/* 검증 진행 상태바 (실행 중일 때) */}
                {isValidating && (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', background: 'rgba(0,0,0,0.25)', padding: '14px 18px', borderRadius: '8px', border: '1px solid var(--panel-border)' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: '13px' }}>
                            <span style={{ display: 'flex', alignItems: 'center', gap: '8px', color: '#10b981', fontWeight: 600 }}>
                                <Sparkles size={16} className="animate-spin" />
                                {aiProgress.message || '문서 정합성 전수 검증 진행 중...'}
                            </span>
                            <span style={{ fontWeight: 700, color: 'var(--text-primary)' }}>
                                {aiProgress.percent}%
                            </span>
                        </div>
                        <div style={{ width: '100%', height: '6px', background: 'rgba(255,255,255,0.1)', borderRadius: '3px', overflow: 'hidden' }}>
                            <div
                                style={{
                                    width: `${aiProgress.percent}%`,
                                    height: '100%',
                                    background: 'linear-gradient(90deg, #10b981, #3b82f6, #a855f7)',
                                    transition: 'width 0.3s ease'
                                }}
                            />
                        </div>
                    </div>
                )}

                {/* 실행 버튼 영역 */}
                <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '10px' }}>
                    {validationResult && (
                        <button
                            onClick={() => { setTargetFile(null); setValidationResult(null); }}
                            className="btn-outline"
                            style={{ padding: '10px 18px', fontSize: '13px' }}
                            disabled={isValidating}
                        >
                            초기화
                        </button>
                    )}
                    <button
                        onClick={handleStartValidation}
                        disabled={isValidating || !targetFile}
                        className="btn-primary"
                        style={{
                            padding: '10px 24px',
                            fontSize: '14px',
                            fontWeight: 700,
                            display: 'flex',
                            alignItems: 'center',
                            gap: '8px',
                            background: 'linear-gradient(135deg, #10b981, #059669)',
                            boxShadow: '0 4px 12px rgba(16, 185, 129, 0.3)',
                            opacity: (isValidating || !targetFile) ? 0.6 : 1,
                            cursor: (isValidating || !targetFile) ? 'not-allowed' : 'pointer'
                        }}
                    >
                        <FileCheck2 size={18} />
                        {isValidating ? '정합성 정밀 검증 중...' : '정합성 검증 시작 (Run SSOT)'}
                    </button>
                </div>
            </div>

            {/* 검증 결과 대시보드 */}
            {validationResult && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
                    {/* 상단 4대 핵심 지표 요약 카드 */}
                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '14px' }}>
                        {/* 카드 1: 총 결함 건수 */}
                        <div className="glass-panel" style={{ padding: '18px 20px', display: 'flex', flexDirection: 'column', gap: '6px', borderLeft: '4px solid #ef4444' }}>
                            <span style={{ fontSize: '12px', fontWeight: 600, color: 'var(--text-muted)' }}>
                                총 결함 검출 건수
                            </span>
                            <div style={{ display: 'flex', alignItems: 'baseline', gap: '8px' }}>
                                <span style={{ fontSize: '28px', fontWeight: 800, color: stats.total > 0 ? '#ef4444' : '#10b981' }}>
                                    {stats.total}
                                </span>
                                <span style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>건</span>
                            </div>
                            <span style={{ fontSize: '11.5px', color: stats.total === 0 ? '#10b981' : 'var(--danger-color)' }}>
                                {stats.total === 0 ? '✓ 결함 없음 (CLEAN)' : '즉시 수정 조치 필요'}
                            </span>
                        </div>

                        {/* 카드 2: 요구사항 ID 매핑률 (RTM) */}
                        <div className="glass-panel" style={{ padding: '18px 20px', display: 'flex', flexDirection: 'column', gap: '6px', borderLeft: '4px solid #3b82f6' }}>
                            <span style={{ fontSize: '12px', fontWeight: 600, color: 'var(--text-muted)' }}>
                                요구사항 ID 매핑률 (RTM)
                            </span>
                            <div style={{ display: 'flex', alignItems: 'baseline', gap: '8px' }}>
                                <span style={{ fontSize: '28px', fontWeight: 800, color: stats.rtmRate === 100 ? '#10b981' : stats.rtmRate >= 80 ? '#f59e0b' : '#ef4444' }}>
                                    {stats.rtmRate}%
                                </span>
                                <span style={{ fontSize: '12.5px', color: 'var(--text-secondary)' }}>
                                    ({validationResult.stage1.coveredReqCount} / {validationResult.stage1.totalMasterReqCount} 식별)
                                </span>
                            </div>
                            <div style={{ width: '100%', height: '5px', background: 'rgba(255,255,255,0.1)', borderRadius: '2px', overflow: 'hidden', marginTop: '2px' }}>
                                <div style={{ width: `${stats.rtmRate}%`, height: '100%', background: stats.rtmRate === 100 ? '#10b981' : stats.rtmRate >= 80 ? '#f59e0b' : '#ef4444' }} />
                            </div>
                        </div>

                        {/* 카드 3: 표준용어 사전 위반 */}
                        <div className="glass-panel" style={{ padding: '18px 20px', display: 'flex', flexDirection: 'column', gap: '6px', borderLeft: '4px solid #f59e0b' }}>
                            <span style={{ fontSize: '12px', fontWeight: 600, color: 'var(--text-muted)' }}>
                                용어사전 위반 (금지어/오표기)
                            </span>
                            <div style={{ display: 'flex', alignItems: 'baseline', gap: '8px' }}>
                                <span style={{ fontSize: '28px', fontWeight: 800, color: stats.glossary > 0 ? '#f59e0b' : '#10b981' }}>
                                    {stats.glossary}
                                </span>
                                <span style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>건</span>
                            </div>
                            <span style={{ fontSize: '11.5px', color: stats.glossary === 0 ? '#10b981' : '#f59e0b' }}>
                                {stats.glossary === 0 ? '✓ 표준용어 100% 준수' : '금지동의어 치환 권고'}
                            </span>
                        </div>

                        {/* 카드 4: AI 문맥 / 수치 오류 */}
                        <div className="glass-panel" style={{ padding: '18px 20px', display: 'flex', flexDirection: 'column', gap: '6px', borderLeft: '4px solid #a855f7' }}>
                            <span style={{ fontSize: '12px', fontWeight: 600, color: 'var(--text-muted)' }}>
                                수치 오류 & AI 문맥 결함
                            </span>
                            <div style={{ display: 'flex', alignItems: 'baseline', gap: '8px' }}>
                                <span style={{ fontSize: '28px', fontWeight: 800, color: stats.aiOrNum > 0 ? '#a855f7' : '#10b981' }}>
                                    {stats.aiOrNum}
                                </span>
                                <span style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>건</span>
                            </div>
                            <span style={{ fontSize: '11.5px', color: stats.aiOrNum === 0 ? '#10b981' : '#a855f7' }}>
                                {stats.aiOrNum === 0 ? '✓ 표 산술 및 문맥 일치' : '표 합계 및 문맥 모순 검출'}
                            </span>
                        </div>
                    </div>

                    {/* 결과 상세 리스트 & RTM 매트릭스 패널 */}
                    <div className="glass-panel" style={{ padding: '20px', display: 'flex', flexDirection: 'column', gap: '16px' }}>
                        {/* 탭 및 액션 바 */}
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px' }}>
                            {/* 필터 탭 */}
                            <div style={{ display: 'flex', gap: '6px', flexWrap: 'wrap' }}>
                                {[
                                    { id: 'all', label: `전체 (${validationResult.allIssues.length})` },
                                    { id: 'requirements', label: `요구사항 ID (${validationResult.allIssues.filter(i => i.category.includes('요구사항')).length})` },
                                    { id: 'glossary', label: `용어 불일치 (${stats.glossary})` },
                                    { id: 'numbers', label: `수치 오류 (${validationResult.allIssues.filter(i => i.category === '수치 오류').length})` },
                                    { id: 'ai', label: `AI 문맥/오탈자 (${validationResult.stage2Issues.length})` },
                                    { id: 'rtm', label: `RTM 매트릭스 표 (${validationResult.stage1.rtmMatrix.length})` }
                                ].map(tab => (
                                    <button
                                        key={tab.id}
                                        onClick={() => setActiveResultTab(tab.id)}
                                        style={{
                                            padding: '6px 12px',
                                            borderRadius: '6px',
                                            fontSize: '12.5px',
                                            fontWeight: 600,
                                            background: activeResultTab === tab.id ? 'var(--accent-blue)' : 'rgba(255,255,255,0.05)',
                                            color: activeResultTab === tab.id ? '#fff' : 'var(--text-secondary)',
                                            border: 'none',
                                            cursor: 'pointer'
                                        }}
                                    >
                                        {tab.label}
                                    </button>
                                ))}
                            </div>

                            {/* 우측 엑셀 다운로드 버튼 */}
                            <button
                                onClick={handleDownloadExcelReport}
                                className="btn-primary"
                                style={{
                                    padding: '8px 16px',
                                    fontSize: '13px',
                                    fontWeight: 700,
                                    display: 'flex',
                                    alignItems: 'center',
                                    gap: '6px',
                                    background: 'linear-gradient(135deg, #10b981, #059669)',
                                    boxShadow: '0 2px 8px rgba(16, 185, 129, 0.3)'
                                }}
                            >
                                <Download size={15} />
                                통합 정합성 검수 리포트 (.xlsx) 다운로드
                            </button>
                        </div>

                        {/* 검색창 및 심각도 필터 (RTM 표가 아닐 때 노출) */}
                        {activeResultTab !== 'rtm' && (
                            <div style={{ display: 'flex', gap: '10px', alignItems: 'center', flexWrap: 'wrap' }}>
                                <div style={{ position: 'relative', flex: 1, minWidth: '220px' }}>
                                    <Search size={15} style={{ position: 'absolute', left: '10px', top: '50%', transform: 'translateY(-50%)', color: 'var(--text-muted)' }} />
                                    <input
                                        type="text"
                                        value={searchQuery}
                                        onChange={(e) => setSearchQuery(e.target.value)}
                                        placeholder="검출된 단어, 원문, 수정안, 사유 검색..."
                                        style={{
                                            width: '100%',
                                            padding: '8px 12px 8px 34px',
                                            borderRadius: '6px',
                                            border: '1px solid var(--panel-border)',
                                            background: 'rgba(0,0,0,0.2)',
                                            color: 'var(--text-primary)',
                                            fontSize: '12.5px'
                                        }}
                                    />
                                </div>
                                <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                                    <span style={{ fontSize: '12px', color: 'var(--text-muted)' }}>심각도:</span>
                                    <select
                                        value={severityFilter}
                                        onChange={(e) => setSeverityFilter(e.target.value)}
                                        style={{ padding: '6px 10px', borderRadius: '6px', background: 'rgba(0,0,0,0.3)', border: '1px solid var(--panel-border)', color: 'var(--text-primary)', fontSize: '12px' }}
                                    >
                                        <option value="ALL">전체 심각도</option>
                                        <option value="HIGH">HIGH (치명적)</option>
                                        <option value="MEDIUM">MEDIUM (주의)</option>
                                        <option value="LOW">LOW (권고)</option>
                                    </select>
                                </div>
                            </div>
                        )}

                        {/* RTM 매트릭스 전용 표 보기 */}
                        {activeResultTab === 'rtm' ? (
                            <div style={{ overflowX: 'auto', border: '1px solid var(--panel-border)', borderRadius: '8px' }}>
                                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12.5px' }}>
                                    <thead>
                                        <tr style={{ background: 'rgba(255,255,255,0.06)', borderBottom: '1px solid var(--panel-border)', textAlign: 'left' }}>
                                            <th style={{ padding: '10px 12px', width: '60px', textAlign: 'center' }}>No.</th>
                                            <th style={{ padding: '10px 12px', width: '140px' }}>요구사항 ID</th>
                                            <th style={{ padding: '10px 12px' }}>요구사항명</th>
                                            <th style={{ padding: '10px 12px', width: '220px' }}>필수 키워드</th>
                                            <th style={{ padding: '10px 12px', width: '120px', textAlign: 'center' }}>출현 여부</th>
                                            <th style={{ padding: '10px 12px', width: '140px', textAlign: 'center' }}>위치(페이지)</th>
                                            <th style={{ padding: '10px 12px', width: '150px', textAlign: 'center' }}>정합성 상태</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        {validationResult.stage1.rtmMatrix.map((item, idx) => (
                                            <tr key={item.id} style={{ borderBottom: '1px solid rgba(255,255,255,0.04)', background: !item.isPresent ? 'rgba(239, 68, 68, 0.04)' : undefined }}>
                                                <td style={{ padding: '10px', textAlign: 'center', color: 'var(--text-muted)' }}>{idx + 1}</td>
                                                <td style={{ padding: '10px', fontWeight: 700, color: 'var(--accent-blue)' }}>{item.id}</td>
                                                <td style={{ padding: '10px', color: 'var(--text-primary)' }}>{item.name}</td>
                                                <td style={{ padding: '10px', color: 'var(--text-secondary)', fontSize: '12px' }}>{item.keywords || '-'}</td>
                                                <td style={{ padding: '10px', textAlign: 'center' }}>
                                                    {item.isPresent ? (
                                                        <span style={{ color: '#10b981', fontWeight: 700 }}>출현 (O)</span>
                                                    ) : (
                                                        <span style={{ color: '#ef4444', fontWeight: 700 }}>누락 (X)</span>
                                                    )}
                                                </td>
                                                <td style={{ padding: '10px', textAlign: 'center', color: 'var(--text-muted)' }}>
                                                    {item.pages.length > 0 ? `p.${item.pages.join(', ')}` : '-'}
                                                </td>
                                                <td style={{ padding: '10px', textAlign: 'center' }}>
                                                    <span style={{
                                                        padding: '3px 8px',
                                                        borderRadius: '10px',
                                                        fontSize: '11px',
                                                        fontWeight: 700,
                                                        background: !item.isPresent ? 'rgba(239, 68, 68, 0.15)' : item.status.includes('누락') ? 'rgba(245, 158, 11, 0.15)' : 'rgba(16, 185, 129, 0.15)',
                                                        color: !item.isPresent ? '#ef4444' : item.status.includes('누락') ? '#f59e0b' : '#10b981'
                                                    }}>
                                                        {item.status}
                                                    </span>
                                                </td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            </div>
                        ) : (
                            /* 결함 목록 카드/테이블 뷰 */
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                                {filteredIssues.length === 0 ? (
                                    <div style={{ padding: '40px', textAlign: 'center', color: 'var(--text-muted)' }}>
                                        <CheckCircle2 size={36} style={{ color: '#10b981', margin: '0 auto 10px', display: 'block' }} />
                                        선택된 조건에 해당하는 결함이 없습니다.
                                    </div>
                                ) : (
                                    filteredIssues.map((issue) => (
                                        <div
                                            key={issue.id}
                                            style={{
                                                padding: '16px',
                                                borderRadius: '8px',
                                                background: 'rgba(255,255,255,0.02)',
                                                border: '1px solid var(--panel-border)',
                                                borderLeft: `4px solid ${issue.severity === 'HIGH' ? '#ef4444' : issue.severity === 'MEDIUM' ? '#f59e0b' : '#3b82f6'}`,
                                                display: 'flex',
                                                flexDirection: 'column',
                                                gap: '8px'
                                            }}
                                        >
                                            {/* 상단 뱃지 및 메타 */}
                                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '8px' }}>
                                                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                                    <span style={{
                                                        padding: '2px 6px',
                                                        borderRadius: '4px',
                                                        fontSize: '11px',
                                                        fontWeight: 700,
                                                        background: issue.severity === 'HIGH' ? 'rgba(239, 68, 68, 0.15)' : issue.severity === 'MEDIUM' ? 'rgba(245, 158, 11, 0.15)' : 'rgba(59, 130, 246, 0.15)',
                                                        color: issue.severity === 'HIGH' ? '#ef4444' : issue.severity === 'MEDIUM' ? '#f59e0b' : '#3b82f6'
                                                    }}>
                                                        {issue.severity}
                                                    </span>
                                                    <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--text-primary)' }}>
                                                        {issue.errorType}
                                                    </span>
                                                    <span style={{ fontSize: '11.5px', color: 'var(--text-muted)' }}>
                                                        ({issue.stage})
                                                    </span>
                                                </div>
                                                <span style={{ fontSize: '12px', color: 'var(--text-muted)' }}>
                                                    발생 위치: p.{issue.page}
                                                </span>
                                            </div>

                                            {/* 원문 및 수정안 대조 */}
                                            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: '10px', background: 'rgba(0,0,0,0.2)', padding: '10px 12px', borderRadius: '6px' }}>
                                                <div>
                                                    <span style={{ display: 'block', fontSize: '11px', color: 'var(--text-muted)', marginBottom: '3px' }}>
                                                        발견된 원문 / 오표기:
                                                    </span>
                                                    <span style={{ fontSize: '13px', color: '#ef4444', fontWeight: 600, wordBreak: 'break-all' }}>
                                                        {issue.originalText || issue.matchedTerm}
                                                    </span>
                                                </div>
                                                <div>
                                                    <span style={{ display: 'block', fontSize: '11px', color: 'var(--text-muted)', marginBottom: '3px' }}>
                                                        수정 권고안 (표준):
                                                    </span>
                                                    <span style={{ fontSize: '13px', color: '#10b981', fontWeight: 700, wordBreak: 'break-all' }}>
                                                        {issue.suggestedText || '-'}
                                                    </span>
                                                </div>
                                            </div>

                                            {/* 사유 설명 */}
                                            <p style={{ margin: 0, fontSize: '12.5px', color: 'var(--text-secondary)', lineHeight: '1.4' }}>
                                                💬 {issue.reason}
                                            </p>
                                        </div>
                                    ))
                                )}
                            </div>
                        )}
                    </div>
                </div>
            )}

            {/* 텍스트 일괄 붙여넣기 모달 */}
            {showBulkReqModal && (
                <div style={{
                    position: 'fixed', top: 0, left: 0, width: '100%', height: '100%',
                    background: 'rgba(0,0,0,0.6)', backdropFilter: 'blur(4px)',
                    zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '20px'
                }}>
                    <div className="glass-panel" style={{ width: '100%', maxWidth: '640px', padding: '24px', display: 'flex', flexDirection: 'column', gap: '16px' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                            <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 700, color: 'var(--text-primary)' }}>
                                요구사항 마스터 일괄 붙여넣기
                            </h3>
                            <button onClick={() => setShowBulkReqModal(false)} style={{ background: 'none', border: 'none', color: 'var(--text-muted)', cursor: 'pointer' }}>
                                ✕
                            </button>
                        </div>
                        <p style={{ margin: 0, fontSize: '12.5px', color: 'var(--text-secondary)', lineHeight: '1.4' }}>
                            한 줄에 하나씩 <b>[요구사항ID | 명칭 | 필수키워드]</b> 형식(구분자: | 또는 탭 또는 쉼표)으로 입력해주세요.
                        </p>
                        <textarea
                            rows={8}
                            value={bulkReqText}
                            onChange={(e) => setBulkReqText(e.target.value)}
                            placeholder={`CNR-001 | 공공보건의료 데이터 통합 연계 체계 구축 | 데이터, 연계, EMR\nSFR-001 | AI 기반 진료지원(CDSS) 알고리즘 실증 | AI, 알고리즘, CDSS\nPMR-001 | 개인정보보호 및 의료법 규제 준수 거버넌스 | 개인정보, 의료법`}
                            style={{
                                width: '100%',
                                padding: '12px',
                                borderRadius: '6px',
                                border: '1px solid var(--panel-border)',
                                background: 'rgba(0,0,0,0.3)',
                                color: 'var(--text-primary)',
                                fontSize: '12.5px',
                                fontFamily: 'monospace'
                            }}
                        />
                        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px' }}>
                            <button onClick={() => setShowBulkReqModal(false)} className="btn-outline" style={{ padding: '8px 14px', fontSize: '12.5px' }}>
                                취소
                            </button>
                            <button onClick={handleApplyBulkRequirements} className="btn-primary" style={{ padding: '8px 18px', fontSize: '12.5px' }}>
                                일괄 등록하기
                            </button>
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
}
