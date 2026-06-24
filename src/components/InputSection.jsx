import React, { useState, useRef, useCallback } from 'react';
import { Target, FileText, Play, Loader2, Upload, X, ClipboardList, AlertCircle, FileSpreadsheet, ChevronLeft, ChevronRight, Trash2, Sparkles } from 'lucide-react';
import { processFile, ALL_ACCEPT, getFileExtension, classifyFile } from '../utils/fileExtractor';

// ── 파일 타입 아이콘/라벨 ──────────────────────────────────────
function getFileTypeLabel(ext) {
    const type = classifyFile(ext);
    switch (type) {
        case 'pdf': return { label: 'PDF', color: '#ef4444' };
        case 'excel': return { label: 'Excel', color: '#22c55e' };
        case 'pptx': return { label: 'PPTX', color: '#f97316' };
        case 'hwpx': return { label: 'HWPX', color: '#0ea5e9' };
        case 'text': return { label: ext.toUpperCase(), color: '#3b82f6' };
        default: return { label: ext.toUpperCase(), color: '#9ca3af' };
    }
}

// ── FileUploadArea 컴포넌트 ──────────────────────────────────
function FileUploadArea({ label, icon: Icon, fileName, onFileSelect, onFileClear, textValue, onTextChange, placeholder, isLoading, loadError }) {
    const fileRef = useRef(null);
    const [isDragging, setIsDragging] = useState(false);

    // 파일 처리 공통 핸들러
    const handleFile = useCallback(async (file) => {
        if (!file) return;
        onFileSelect(file.name, '', true); // 로딩 시작
        try {
            const result = await processFile(file);
            // 객체 전체가 아닌 추출된 텍스트 내용(result.text)만 전달
            onFileSelect(file.name, result.text, false);
        } catch (err) {
            console.error('파일 처리 오류:', err);
            onFileSelect(file.name, '', false, err.message);
        }
    }, [onFileSelect]);

    // 파일 선택 핸들러
    const handleFileChange = (e) => {
        const file = e.target.files[0];
        handleFile(file);
        e.target.value = ''; // input 초기화
    };

    // 드래그앤드롭 핸들러
    const handleDragEnter = (e) => { e.preventDefault(); e.stopPropagation(); setIsDragging(true); };
    const handleDragLeave = (e) => { e.preventDefault(); e.stopPropagation(); setIsDragging(false); };
    const handleDragOver = (e) => { e.preventDefault(); e.stopPropagation(); };
    const handleDrop = (e) => {
        e.preventDefault();
        e.stopPropagation();
        setIsDragging(false);
        const file = e.dataTransfer.files[0];
        if (file) handleFile(file);
    };

    const fileExt = fileName ? getFileExtension(fileName) : null;
    const fileTypeInfo = fileExt ? getFileTypeLabel(fileExt) : null;
    const loadingLabel = fileExt
        ? classifyFile(fileExt) === 'pdf' ? 'PDF 추출 중...'
            : classifyFile(fileExt) === 'excel' ? 'Excel 변환 중...'
                : classifyFile(fileExt) === 'pptx' ? 'PPTX 추출 중...'
                    : classifyFile(fileExt) === 'hwpx' ? 'HWPX 파싱 중...'
                        : '파일 읽는 중...'
        : '파일 처리 중...';

    return (
        <div
            style={{
                display: 'flex', flexDirection: 'column', flex: 1, gap: '12px',
                position: 'relative', minHeight: 0, width: '100%',
            }}
            onDragEnter={handleDragEnter}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
            className={isDragging ? 'cursor-copy' : ''}
        >
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                <label style={{
                    fontSize: '14px', fontWeight: 500, color: 'var(--text-secondary)',
                    display: 'flex', alignItems: 'center', gap: '6px',
                }}>
                    <Icon size={16} /> {label}
                </label>
                {/* 개별 지우기 버튼 추가 */}
                {(fileName || textValue) && !isLoading && (
                    <button
                        onClick={onFileClear}
                        className="interactive"
                        style={{
                            background: 'transparent', border: 'none', color: 'var(--text-secondary)',
                            fontSize: '12px', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '4px',
                            opacity: 0.7
                        }}
                    >
                        <Trash2 size={13} /> 내용 지우기
                    </button>
                )}
            </div>

            {/* 파일 선택 버튼 + 파일명 */}
            <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                <button
                    onClick={() => fileRef.current?.click()}
                    disabled={isLoading}
                    className="interactive"
                    style={{
                        display: 'flex', alignItems: 'center', gap: '8px',
                        padding: '10px 16px', fontSize: '13px', fontWeight: 600,
                        background: 'rgba(59, 130, 246, 0.1)',
                        color: 'var(--accent-blue)',
                        border: '1px solid rgba(59, 130, 246, 0.2)',
                        borderRadius: '10px',
                        cursor: isLoading ? 'wait' : 'pointer',
                        flexShrink: 0,
                        opacity: isLoading ? 0.6 : 1,
                    }}
                >
                    {isLoading
                        ? <Loader2 size={16} className="animate-spin" />
                        : <Upload size={16} />}
                    {isLoading ? loadingLabel : '파일 업로드'}
                </button>
                {fileName && (
                    <div style={{
                        display: 'flex', alignItems: 'center', gap: '6px',
                        padding: '6px 10px', fontSize: '12px',
                        background: loadError ? 'rgba(239, 68, 68, 0.1)' : 'rgba(16, 185, 129, 0.1)',
                        color: loadError ? 'var(--danger-color)' : 'var(--success-color)',
                        borderRadius: '6px',
                        border: `1px solid ${loadError ? 'rgba(239, 68, 68, 0.2)' : 'rgba(16, 185, 129, 0.2)'}`,
                        overflow: 'hidden', flex: 1,
                    }}>
                        {fileTypeInfo && (
                            <span style={{
                                fontSize: '10px', padding: '1px 4px', borderRadius: '3px',
                                background: `${fileTypeInfo.color}22`, color: fileTypeInfo.color,
                                fontWeight: 600, flexShrink: 0,
                            }}>{fileTypeInfo.label}</span>
                        )}
                        <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                            {fileName}
                        </span>
                        <button
                            onClick={onFileClear}
                            style={{
                                background: 'transparent', border: 'none', color: 'var(--text-secondary)',
                                cursor: 'pointer', padding: '2px', display: 'flex', flexShrink: 0,
                            }}
                        >
                            <X size={12} />
                        </button>
                    </div>
                )}
                <input
                    ref={fileRef}
                    type="file"
                    accept={ALL_ACCEPT}
                    onChange={handleFileChange}
                    style={{ display: 'none' }}
                />
            </div>

            {/* 에러 메시지 */}
            {loadError && (
                <div style={{
                    padding: '8px 12px', fontSize: '12px', lineHeight: '1.5',
                    background: 'rgba(239, 68, 68, 0.08)',
                    color: 'var(--danger-color)',
                    borderRadius: '6px',
                    border: '1px solid rgba(239, 68, 68, 0.15)',
                    display: 'flex', gap: '6px', alignItems: 'flex-start',
                }}>
                    <AlertCircle size={14} style={{ flexShrink: 0, marginTop: '1px' }} />
                    {loadError}
                </div>
            )}

            {/* 텍스트 영역 + 드래그앤드롭 오버레이 */}
            <div style={{ flex: 1, position: 'relative', display: 'flex' }}>
                <textarea
                    placeholder={placeholder}
                    value={textValue}
                    onChange={e => onTextChange(e.target.value)}
                    style={{ 
                        flex: 1, resize: 'none', lineHeight: '1.6', width: '100%',
                        background: 'rgba(0, 0, 0, 0.2)', color: 'var(--text-primary)',
                        border: '1px solid var(--glass-border)', borderRadius: '12px',
                        padding: '16px', outline: 'none'
                    }}
                />
                {/* 드래그 오버레이 */}
                {isDragging && (
                    <div style={{
                        position: 'absolute', inset: 0,
                        background: 'rgba(59, 130, 246, 0.12)',
                        border: '2px dashed var(--accent-color)',
                        borderRadius: '8px',
                        display: 'flex', flexDirection: 'column',
                        alignItems: 'center', justifyContent: 'center',
                        gap: '8px',
                        zIndex: 10,
                        backdropFilter: 'blur(4px)',
                        pointerEvents: 'none',
                    }}>
                        <Upload size={32} color="var(--accent-color)" />
                        <span style={{
                            fontSize: '14px', fontWeight: 600,
                            color: 'var(--accent-color)',
                        }}>
                            파일을 여기에 놓으세요
                        </span>
                        <span style={{ fontSize: '11px', color: 'var(--text-secondary)' }}>
                            PDF, Excel, TXT 파일 지원
                        </span>
                    </div>
                )}
            </div>
        </div>
    );
}

// ── 메인 InputSection ──────────────────────────────────────
export default function InputSection({ onAnalyze, isAnalyzing, isTypoMode = false, onReset }) {
    const [activeTab, setActiveTab] = useState(isTypoMode ? 1 : 0);
    const [guideline, setGuideline] = useState('');
    const [artifact, setArtifact] = useState('');
    const [inspectionScope, setInspectionScope] = useState('');
    const [glossary, setGlossary] = useState(() => localStorage.getItem('rfp_glossary') || '');
    const [useRag, setUseRag] = useState(false);
    
    // 파일 정보 및 로딩 상태
    const [guidelineFile, setGuidelineFile] = useState('');
    const [artifactFile, setArtifactFile] = useState('');
    const [glossaryFile, setGlossaryFile] = useState(() => localStorage.getItem('rfp_glossary_file') || '');
    const [guidelineLoading, setGuidelineLoading] = useState(false);
    const [artifactLoading, setArtifactLoading] = useState(false);
    const [glossaryLoading, setGlossaryLoading] = useState(false);
    const [guidelineError, setGuidelineError] = useState(null);
    const [artifactError, setArtifactError] = useState(null);
    const [glossaryError, setGlossaryError] = useState(null);
    
    // 용어사전 영속성 동기화
    React.useEffect(() => {
        localStorage.setItem('rfp_glossary', glossary);
    }, [glossary]);

    React.useEffect(() => {
        localStorage.setItem('rfp_glossary_file', glossaryFile);
    }, [glossaryFile]);

    const handleGuidelineFile = (name, content, loading = false, error = null) => {
        setGuidelineFile(name);
        setGuidelineLoading(loading);
        setGuidelineError(error);
        if (content) setGuideline(content);
        if (error) setGuideline('');
    };

    const handleArtifactFile = (name, content, loading = false, error = null) => {
        setArtifactFile(name);
        setArtifactLoading(loading);
        setArtifactError(error);
        if (content) setArtifact(content);
        if (error) setArtifact('');
    };

    const handleGlossaryFile = (name, content, loading = false, error = null) => {
        setGlossaryFile(name);
        setGlossaryLoading(loading);
        setGlossaryError(error);
        if (content) setGlossary(content);
        if (error) setGlossary('');
    };

    const handleReset = () => {
        if (window.confirm('입력된 문서 데이터(기준문서, 산출물, 점검범위)를 초기화하시겠습니까?\n(용어사전은 보존됩니다)')) {
            setGuideline('');
            setArtifact('');
            setInspectionScope('');
            setGuidelineFile('');
            setArtifactFile('');
            setGuidelineError(null);
            setArtifactError(null);
            if (onReset) onReset();
        }
    };

    // ── 최소화 모드 렌더링 ──────────────────────────────────

    return (
        <div className="glass-panel animate-fade-in" style={{
            flex: 1, display: 'flex', flexDirection: 'column',
            padding: '28px', gap: '24px',
            minWidth: '450px', height: '100%',
            cursor: isAnalyzing ? 'wait' : 'default',
            transition: 'all 0.3s ease',
            overflow: 'hidden', position: 'relative'
        }}>
            <div style={{
                display: 'flex', alignItems: 'center', gap: '8px',
                borderBottom: '1px solid var(--panel-border)', paddingBottom: '16px',
            }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexShrink: 0 }}>
                    <FileText size={20} color="var(--accent-color)" />
                    <h2 style={{ margin: 0, fontSize: '18px', fontWeight: 600, whiteSpace: 'nowrap' }}>데이터 입력</h2>
                </div>
                
                <button
                    onClick={handleReset}
                    title="전체 입력 초기화"
                    className="interactive"
                    style={{
                        background: 'rgba(239, 68, 68, 0.1)',
                        border: '1px solid rgba(239, 68, 68, 0.2)',
                        padding: '6px 12px',
                        borderRadius: '8px',
                        color: 'var(--danger-color)',
                        cursor: 'pointer',
                        fontSize: '12px',
                        fontWeight: 600,
                        marginLeft: '12px',
                        display: 'flex',
                        alignItems: 'center',
                        gap: '6px',
                        whiteSpace: 'nowrap'
                    }}
                >
                    <X size={14} /> 초기화
                </button>


                <div style={{ marginLeft: 'auto', display: 'flex', gap: '4px', flexWrap: 'wrap', justifyContent: 'flex-end' }}>
                    {['PDF', 'Excel', 'PPTX', 'HWPX', 'TXT'].map((ext, idx) => {
                        const colors = ['#ef4444', '#22c55e', '#f97316', '#0ea5e9', '#3b82f6'];
                        return (
                            <span key={ext} style={{
                                fontSize: '10px', padding: '3px 6px',
                                background: `${colors[idx]}11`, color: colors[idx],
                                borderRadius: '4px', fontWeight: 600, whiteSpace: 'nowrap'
                            }}>{ext}</span>
                        );
                    })}
                </div>
            </div>

            {/* 탭 헤더 */}
            <div style={{ 
                display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '8px', 
                borderBottom: '1px solid var(--glass-border)', paddingBottom: '12px', marginBottom: '8px' 
            }}>
                {(isTypoMode ? [
                    { id: 1, label: '1. 검증 대상 문서', icon: FileText, hasData: !!artifact || !!artifactFile },
                    { id: 2, label: '2. 용어 사전', icon: FileSpreadsheet, hasData: !!glossary || !!glossaryFile },
                    { id: 3, label: '3. 점검 범위', icon: ClipboardList, hasData: !!inspectionScope }
                ] : [
                    { id: 0, label: '1. 기준 문서', icon: Target, hasData: !!guideline || !!guidelineFile },
                    { id: 1, label: '2. 산출물', icon: FileText, hasData: !!artifact || !!artifactFile },
                    { id: 3, label: '3. 점검 범위', icon: ClipboardList, hasData: !!inspectionScope }
                ]).map(tab => (
                    <button
                        key={tab.id}
                        onClick={() => setActiveTab(tab.id)}
                        style={{
                            padding: '10px 4px',
                            background: activeTab === tab.id ? 'rgba(59, 130, 246, 0.12)' : 'rgba(255, 255, 255, 0.02)',
                            border: '1px solid', borderColor: activeTab === tab.id ? 'var(--accent-blue)' : 'var(--glass-border)',
                            borderRadius: '10px',
                            color: activeTab === tab.id ? 'var(--text-primary)' : 'var(--text-secondary)',
                            fontWeight: activeTab === tab.id ? 700 : 500, cursor: 'pointer',
                            display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px',
                            transition: 'all 0.2s cubic-bezier(0.16, 1, 0.3, 1)', fontSize: '12px', position: 'relative',
                        }}
                    >
                        <tab.icon size={14} />
                        <span style={{ whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{tab.label}</span>
                        {tab.hasData && (
                            <div style={{ 
                                width: '6px', height: '6px', borderRadius: '50%', 
                                background: 'var(--success-color)', marginLeft: '4px',
                                boxShadow: '0 0 6px var(--success-color)'
                            }} />
                        )}
                    </button>
                ))}
            </div>

            {/* 탭 컨텐츠 */}
            <div style={{ flex: 1, display: 'flex', flexDirection: 'column' }}>
                {activeTab === 0 && !isTypoMode && (
                    <FileUploadArea
                        label="비교/검증의 잣대가 될 기준 문서 (RFP, 분석서 등)"
                        icon={Target}
                        fileName={guidelineFile}
                        onFileSelect={handleGuidelineFile}
                        onFileClear={() => { setGuidelineFile(''); setGuideline(''); setGuidelineError(null); }}
                        textValue={guideline}
                        onTextChange={setGuideline}
                        placeholder="기준 문서를 업로드하거나 직접 텍스트를 입력하세요."
                        isLoading={guidelineLoading}
                        loadError={guidelineError}
                    />
                )}
                {activeTab === 1 && (
                    <FileUploadArea
                        label="진단 및 오류 검증이 필요한 결과 산출물 단위"
                        icon={FileText}
                        fileName={artifactFile}
                        onFileSelect={handleArtifactFile}
                        onFileClear={() => { setArtifactFile(''); setArtifact(''); setArtifactError(null); }}
                        textValue={artifact}
                        onTextChange={setArtifact}
                        placeholder="검증할 산출물을 여기에 끌어다 놓거나 작성하세요."
                        isLoading={artifactLoading}
                        loadError={artifactError}
                    />
                )}
                {activeTab === 2 && (
                    <FileUploadArea
                        label="최우선 기준이 되는 맞춤형 도메인 용어집"
                        icon={FileSpreadsheet}
                        fileName={glossaryFile}
                        onFileSelect={handleGlossaryFile}
                        onFileClear={() => { setGlossaryFile(''); setGlossary(''); setGlossaryError(null); }}
                        textValue={glossary}
                        onTextChange={setGlossary}
                        placeholder="커스텀 용어 사전을 엑셀/PDF로 업로드하거나 직접 붙여넣으세요."
                        isLoading={glossaryLoading}
                        loadError={glossaryError}
                    />
                )}
                {activeTab === 3 && (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', flex: 1, background: 'var(--bg-secondary)', padding: '16px', borderRadius: '8px', border: '1px solid var(--panel-border)' }}>
                        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '8px' }}>
                            <label style={{
                                fontSize: '15px', fontWeight: 600, color: 'var(--text-primary)',
                                display: 'flex', alignItems: 'center', gap: '8px'
                            }}>
                                <ClipboardList size={18} color="var(--accent-color)" /> 검증 포커스 영역 및 개별 테스트 규칙
                            </label>
                            {inspectionScope && (
                                <button
                                    onClick={() => setInspectionScope('')}
                                    style={{
                                        background: 'transparent', border: 'none', color: 'var(--text-secondary)',
                                        fontSize: '12px', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '4px'
                                    }}
                                >
                                    <Trash2 size={12} /> 필드 지우기
                                </button>
                            )}
                        </div>
                        <textarea
                            placeholder="예시: CSR-011 요구사항 항목만을 중심으로 집중 점검하세요. 목차 일관성보다 문장 간 논리성에 더 큰 가중치를 두어 검사해주세요."
                            value={inspectionScope}
                            onChange={e => setInspectionScope(e.target.value)}
                            style={{ 
                                flex: 1, resize: 'none', lineHeight: '1.6', fontSize: '15px',
                                background: 'rgba(0, 0, 0, 0.2)', color: 'var(--text-primary)',
                                border: '1px solid var(--glass-border)', borderRadius: '12px',
                                padding: '16px', outline: 'none'
                            }}
                        />
                    </div>
                )}
            </div>

            {/* RAG Option Checkbox */}
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px', padding: '12px 16px', background: useRag ? 'rgba(59, 130, 246, 0.08)' : 'rgba(255,255,255,0.02)', borderRadius: '12px', border: `1px solid ${useRag ? 'rgba(59, 130, 246, 0.3)' : 'var(--glass-border)'}`, cursor: 'pointer', transition: 'all 0.2s' }} onClick={() => setUseRag(!useRag)}>
                <input 
                    type="checkbox" 
                    checked={useRag} 
                    onChange={(e) => setUseRag(e.target.checked)} 
                    style={{ cursor: 'pointer', width: '18px', height: '18px' }} 
                    onClick={(e) => e.stopPropagation()}
                />
                <div style={{ display: 'flex', flexDirection: 'column' }}>
                    <span style={{ fontSize: '14px', fontWeight: 700, color: useRag ? 'var(--accent-blue)' : 'var(--text-primary)' }}>ISMP RAG 지식베이스 연동</span>
                    <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>데스크탑 산출물 데이터를 분석의 배경 지식으로 활용합니다.</span>
                </div>
            </div>

            <button
                className="primary interactive"
                onClick={() => onAnalyze(guideline, artifact, inspectionScope, glossary, artifactFile, useRag)}
                disabled={isAnalyzing || guidelineLoading || artifactLoading || glossaryLoading || ((!isTypoMode && !guideline && !artifact && !glossary) || (isTypoMode && !artifact))}
                style={{
                    marginTop: '12px',
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                    gap: '10px', padding: '18px', fontSize: '17px', fontWeight: 700,
                    background: 'linear-gradient(135deg, #2563eb, #3b82f6)',
                    color: 'white',
                    border: 'none',
                    borderRadius: '16px',
                    boxShadow: '0 8px 24px rgba(37, 99, 235, 0.3)',
                    opacity: (isAnalyzing || guidelineLoading || artifactLoading || glossaryLoading || ((!isTypoMode && !guideline && !artifact && !glossary) || (isTypoMode && !artifact))) ? 0.6 : 1,
                    cursor: (isAnalyzing || guidelineLoading || artifactLoading || glossaryLoading || ((!isTypoMode && !guideline && !artifact && !glossary) || (isTypoMode && !artifact))) ? 'not-allowed' : 'pointer',
                }}
            >
                {isAnalyzing ? (
                    <>
                        <Loader2 size={24} className="animate-spin" />
                        {isTypoMode ? 'AI 분석 중...' : 'AI 검증 중...'}
                    </>
                ) : (
                    <>
                        <Sparkles size={22} style={{ color: '#38bdf8', filter: 'drop-shadow(0 0 2px rgba(56, 189, 248, 0.5))' }} />
                        {isTypoMode ? '문서 품질 정밀 점검 시작' : '엄격한 AI 검증 시작'}
                    </>
                )}
            </button>

            <style>{`
        @keyframes spin {
          from { transform: rotate(0deg); }
          to { transform: rotate(360deg); }
        }
      `}</style>
        </div>
    );
}
