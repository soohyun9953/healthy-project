import React, { useState, useRef, useEffect } from 'react';
import {
    FileText, Upload, Download, Trash2, FolderDown, Archive,
    CheckCircle2, AlertCircle, Loader2, Sparkles, RefreshCw,
    Settings, Eye, Edit3, ArrowRight, Layers, FileCode, Check,
    AlignLeft, BookOpen, ExternalLink, ChevronDown, ChevronUp
} from 'lucide-react';
import {
    DOCX_THEMES,
    DOCX_FONTS,
    DEFAULT_DOCX_OPTIONS,
    convertMdToDocxBlob,
    convertSingleMdFile,
    convertBatchMdFiles,
    downloadAllAsZip,
    saveFilesToDirectory,
    downloadSingleFile
} from '../utils/mdToDocxConverter.js';
import { marked } from 'marked';

// 샘플 마크다운 템플릿
const SAMPLE_MARKDOWN = `# 차세대 클라우드 기반 통합의료정보시스템 구축 제안서

## 1. 사업 개요 및 목표
본 사업은 **공공의료 혁신** 및 *데이터 기반 스마트 진료* 체계 구현을 위하여 최신 클라우드 인프라와 AI 기술을 융합한 표준 정보시스템을 구축하는 데 목적이 있습니다.

- **사업명**: 2026년도 스마트 공공병원 차세대 의료정보시스템 통합 구축
- **사업 기간**: 계약체결일로부터 12개월
- **추진 방식**: 클라우드 네이티브 기반 표준화 개발

---

## 2. 주요 추진 전략 및 핵심 과제

### 2.1 핵심 추진 영역
1. **클라우드 인프라 고도화**: 무중단 이중화 아키텍처 적용
2. **AI 스마트 어시스턴트 도입**: 진료 기록 요약 및 처방 보조
3. **표준 데이터 레이크 구축**: FHIR/HL7 표준 기반 데이터 상호운용성 확보

> 💡 **중요 추진 원칙**  
> 모든 의료 데이터는 개인정보보호법 및 의료법 가이드라인을 100% 준수하여 암호화 전송 및 저장되어야 합니다.

---

## 3. 시스템 구성 및 비교 분석

| 구분 | 현행(AS-IS) | 목표(TO-BE) | 기대효과 |
| :--- | :--- | :--- | :---: |
| **인프라** | 로컬 레거시 서버 | 멀티 클라우드 분산 환경 | 가용성 99.99% 달성 |
| **데이터 연계** | 개별 기관 단절 구조 | 표준 FHIR 기반 실시간 연계 | 연계 속도 3배 향상 |
| **보안 체계** | 방화벽 중심 단일 통제 | 제로 트러스트 통합 보안 | 보안 위협 90% 이상 차단 |

---

## 4. API 인터페이스 예시 코드
\`\`\`json
{
  "service": "SmartHospitalCore",
  "version": "v2.11.2",
  "status": "active",
  "features": ["AI_EMR", "HL7_FHIR", "CloudNative_DB"]
}
\`\`\`

자세한 표준 규격은 [보건복지부 마이헬스웨이 공식 사이트](https://www.mohw.go.kr)를 참조하십시오.
`;

export default function MdToDocxConverter() {
    const [mode, setMode] = useState('file'); // 'file' | 'text'
    const [files, setFiles] = useState([]);
    const [isDragging, setIsDragging] = useState(false);
    const [isConverting, setIsConverting] = useState(false);
    const [progress, setProgress] = useState({ current: 0, total: 0, currentFile: '', status: '' });
    const [results, setResults] = useState([]);
    
    // 텍스트 모드 State
    const [editorText, setEditorText] = useState(SAMPLE_MARKDOWN);
    const [textConverting, setTextConverting] = useState(false);
    
    // 서식 설정 State
    const [options, setOptions] = useState(DEFAULT_DOCX_OPTIONS);
    const [showOptions, setShowOptions] = useState(true);
    const [successMsg, setSuccessMsg] = useState(null);
    const [errorMsg, setErrorMsg] = useState(null);
    
    const fileInputRef = useRef(null);

    // 알림 메시지 자동 타이머
    useEffect(() => {
        if (successMsg) {
            const timer = setTimeout(() => setSuccessMsg(null), 4000);
            return () => clearTimeout(timer);
        }
    }, [successMsg]);

    // 파일 필터링: .md, .markdown, .txt
    const filterMdFiles = (fileList) => {
        const validExtensions = ['.md', '.markdown', '.txt'];
        return Array.from(fileList).filter(file => {
            const name = file.name.toLowerCase();
            return validExtensions.some(ext => name.endsWith(ext));
        });
    };

    // 드래그 앤 드롭 핸들러
    const handleDragOver = (e) => {
        e.preventDefault();
        e.stopPropagation();
        setIsDragging(true);
    };

    const handleDragLeave = (e) => {
        e.preventDefault();
        e.stopPropagation();
        setIsDragging(false);
    };

    const handleDrop = async (e) => {
        e.preventDefault();
        e.stopPropagation();
        setIsDragging(false);

        if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
            const validFiles = filterMdFiles(e.dataTransfer.files);
            if (validFiles.length === 0) {
                setErrorMsg('변환 가능한 마크다운 파일(.md, .markdown, .txt)을 선택해 주세요.');
                return;
            }
            await processBatchFiles(validFiles);
        }
    };

    const handleFileChange = async (e) => {
        if (e.target.files && e.target.files.length > 0) {
            const validFiles = filterMdFiles(e.target.files);
            if (validFiles.length === 0) {
                setErrorMsg('변환 가능한 마크다운 파일(.md, .markdown, .txt)을 선택해 주세요.');
                return;
            }
            await processBatchFiles(validFiles);
            e.target.value = '';
        }
    };

    // 파일 일괄 변환 처리
    const processBatchFiles = async (fileList) => {
        setErrorMsg(null);
        setSuccessMsg(null);
        setIsConverting(true);
        setFiles(fileList);
        setResults([]);

        try {
            const convertedResults = await convertBatchMdFiles(fileList, options, (p) => {
                setProgress(p);
            });

            setResults(convertedResults);
            const successCount = convertedResults.filter(r => r.status === 'success').length;
            setSuccessMsg(`총 ${successCount}개의 마크다운 파일이 MS Word(.docx)로 완벽하게 변환되었습니다.`);
        } catch (err) {
            console.error('Batch convert error:', err);
            setErrorMsg(`변환 중 오류가 발생했습니다: ${err.message}`);
        } finally {
            setIsConverting(false);
        }
    };

    // 직접 입력 텍스트 단일 변환 다운로드
    const handleConvertEditorText = async () => {
        if (!editorText.trim()) {
            setErrorMsg('변환할 마크다운 텍스트를 입력해 주세요.');
            return;
        }

        setTextConverting(true);
        setErrorMsg(null);
        try {
            const blob = await convertMdToDocxBlob(editorText, options);
            const docTitle = options.documentTitle || '마크다운_변환_문서';
            const fileName = `${options.prefix || ''}${docTitle}.docx`;
            downloadSingleFile(blob, fileName);
            setSuccessMsg(`'${fileName}' 파일이 성공적으로 다운로드되었습니다.`);
        } catch (err) {
            console.error('Editor convert error:', err);
            setErrorMsg(`Word 변환 실패: ${err.message}`);
        } finally {
            setTextConverting(false);
        }
    };

    // 전체 ZIP 다운로드
    const handleDownloadAllZip = async () => {
        try {
            const count = await downloadAllAsZip(results, '변환완료_Word문서목록.zip');
            setSuccessMsg(`총 ${count}개 Word 파일이 포함된 ZIP 압축 파일이 다운로드되었습니다.`);
        } catch (err) {
            setErrorMsg(err.message);
        }
    };

    // 원하는 폴더 직접 일괄 저장
    const handleSaveToDirectory = async () => {
        try {
            const savedCount = await saveFilesToDirectory(results);
            setSuccessMsg(`선택하신 로컬 폴더에 총 ${savedCount}개의 Word 파일이 성공적으로 저장되었습니다.`);
        } catch (err) {
            if (err.name !== 'AbortError') {
                setErrorMsg(err.message);
            }
        }
    };

    // 옵션 변경 핸들러
    const updateOption = (key, value) => {
        setOptions(prev => ({ ...prev, [key]: value }));
    };

    // 재변환 실행 (파일 목록이 있는 상태에서 옵션 변경 후)
    const handleReconvert = () => {
        if (files.length > 0) {
            processBatchFiles(files);
        }
    };

    return (
        <div className="animate-fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
            {/* 상단 서브 모드 탭 & 안내 배너 */}
            <div style={{
                background: 'linear-gradient(135deg, rgba(37, 99, 235, 0.08), rgba(59, 130, 246, 0.03))',
                border: '1px solid rgba(59, 130, 246, 0.25)',
                borderRadius: '16px',
                padding: '20px 24px',
                display: 'flex',
                flexDirection: 'column',
                gap: '16px'
            }}>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: '12px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                        <div style={{
                            width: '42px',
                            height: '42px',
                            borderRadius: '10px',
                            background: 'linear-gradient(135deg, #2563eb, #1d4ed8)',
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                            boxShadow: '0 4px 12px rgba(37, 99, 235, 0.3)'
                        }}>
                            <FileText size={24} color="#ffffff" />
                        </div>
                        <div>
                            <h2 style={{ margin: 0, fontSize: '18px', fontWeight: 800, color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                                마크다운(MD) ➔ MS Word(DOCX) 변환기
                                <span style={{ fontSize: '11px', background: 'rgba(37, 99, 235, 0.15)', color: '#60a5fa', padding: '2px 8px', borderRadius: '12px', fontWeight: 700 }}>
                                    무손실 표준 DOCX
                                </span>
                            </h2>
                            <p style={{ margin: '4px 0 0', fontSize: '13px', color: 'var(--text-secondary)' }}>
                                마크다운 문서(.md, .markdown, .txt)의 제목 계층, 표(Table), 목록, 인용구, 코드 블록을 깔끔한 MS Word 서식으로 즉시 자동 변환합니다.
                            </p>
                        </div>
                    </div>

                    {/* 모드 전환 버튼 */}
                    <div style={{
                        display: 'flex',
                        background: 'rgba(0,0,0,0.25)',
                        padding: '4px',
                        borderRadius: '10px',
                        border: '1px solid var(--panel-border)'
                    }}>
                        <button
                            onClick={() => setMode('file')}
                            className="interactive"
                            style={{
                                display: 'flex',
                                alignItems: 'center',
                                gap: '6px',
                                padding: '8px 14px',
                                borderRadius: '7px',
                                border: 'none',
                                background: mode === 'file' ? '#2563eb' : 'transparent',
                                color: mode === 'file' ? '#ffffff' : 'var(--text-secondary)',
                                fontWeight: 700,
                                fontSize: '13px',
                                cursor: 'pointer',
                                transition: 'all 0.2s'
                            }}
                        >
                            <Upload size={14} /> MD 파일 일괄 업로드
                        </button>
                        <button
                            onClick={() => setMode('text')}
                            className="interactive"
                            style={{
                                display: 'flex',
                                alignItems: 'center',
                                gap: '6px',
                                padding: '8px 14px',
                                borderRadius: '7px',
                                border: 'none',
                                background: mode === 'text' ? '#2563eb' : 'transparent',
                                color: mode === 'text' ? '#ffffff' : 'var(--text-secondary)',
                                fontWeight: 700,
                                fontSize: '13px',
                                cursor: 'pointer',
                                transition: 'all 0.2s'
                            }}
                        >
                            <Edit3 size={14} /> 텍스트 직접 입력 / 미리보기
                        </button>
                    </div>
                </div>
            </div>

            {/* 메시지 알림 배너 */}
            {successMsg && (
                <div className="animate-fade-in" style={{
                    padding: '12px 16px',
                    background: 'rgba(16, 185, 129, 0.12)',
                    border: '1px solid rgba(16, 185, 129, 0.3)',
                    borderRadius: '10px',
                    color: 'var(--success-color)',
                    fontSize: '13.5px',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '10px',
                    fontWeight: 600
                }}>
                    <CheckCircle2 size={18} /> {successMsg}
                </div>
            )}
            {errorMsg && (
                <div className="animate-fade-in" style={{
                    padding: '12px 16px',
                    background: 'rgba(239, 68, 68, 0.12)',
                    border: '1px solid rgba(239, 68, 68, 0.3)',
                    borderRadius: '10px',
                    color: 'var(--danger-color)',
                    fontSize: '13.5px',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '10px',
                    fontWeight: 600
                }}>
                    <AlertCircle size={18} /> {errorMsg}
                </div>
            )}

            {/* Word 서식 및 디자인 옵션 패널 (접기/펼치기) */}
            <div style={{
                background: 'rgba(255,255,255,0.02)',
                border: '1px solid var(--panel-border)',
                borderRadius: '14px',
                overflow: 'hidden'
            }}>
                <button
                    onClick={() => setShowOptions(!showOptions)}
                    className="interactive"
                    style={{
                        width: '100%',
                        padding: '14px 20px',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'space-between',
                        background: 'transparent',
                        border: 'none',
                        color: 'var(--text-primary)',
                        fontWeight: 700,
                        fontSize: '14px',
                        cursor: 'pointer'
                    }}
                >
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <Settings size={16} color="#3b82f6" />
                        <span>Word 문서 서식 및 스타일 테마 설정</span>
                        <span style={{ fontSize: '11.5px', color: 'var(--text-muted)', fontWeight: 400 }}>
                            (폰트: {options.fontFamily}, 크기: {options.fontSizePt}pt, 테마: {DOCX_THEMES[options.themeId]?.name})
                        </span>
                    </div>
                    {showOptions ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
                </button>

                {showOptions && (
                    <div style={{
                        padding: '18px 20px 22px',
                        borderTop: '1px solid var(--panel-border)',
                        display: 'grid',
                        gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))',
                        gap: '16px',
                        background: 'rgba(0,0,0,0.1)'
                    }}>
                        {/* 1. 디자인 테마 선택 */}
                        <div>
                            <label style={{ display: 'block', fontSize: '12.5px', fontWeight: 700, color: 'var(--text-secondary)', marginBottom: '8px' }}>
                                🎨 문서 색상 테마
                            </label>
                            <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px' }}>
                                {Object.values(DOCX_THEMES).map((t) => (
                                    <button
                                        key={t.id}
                                        type="button"
                                        onClick={() => updateOption('themeId', t.id)}
                                        style={{
                                            padding: '6px 10px',
                                            borderRadius: '6px',
                                            fontSize: '12px',
                                            cursor: 'pointer',
                                            display: 'flex',
                                            alignItems: 'center',
                                            gap: '6px',
                                            background: options.themeId === t.id ? `#${t.primary}22` : 'var(--bg-secondary)',
                                            border: `1px solid ${options.themeId === t.id ? '#' + t.primary : 'var(--panel-border)'}`,
                                            color: options.themeId === t.id ? '#ffffff' : 'var(--text-secondary)',
                                            fontWeight: options.themeId === t.id ? 700 : 500
                                        }}
                                    >
                                        <span style={{ width: '10px', height: '10px', borderRadius: '50%', background: `#${t.primary}`, display: 'inline-block' }} />
                                        {t.name}
                                    </button>
                                ))}
                            </div>
                        </div>

                        {/* 2. 본문 기본 글꼴 */}
                        <div>
                            <label style={{ display: 'block', fontSize: '12.5px', fontWeight: 700, color: 'var(--text-secondary)', marginBottom: '8px' }}>
                                🔤 본문 기본 폰트
                            </label>
                            <select
                                value={options.fontFamily}
                                onChange={(e) => updateOption('fontFamily', e.target.value)}
                                style={{
                                    width: '100%',
                                    padding: '8px 12px',
                                    borderRadius: '8px',
                                    background: 'var(--bg-secondary)',
                                    border: '1px solid var(--panel-border)',
                                    color: 'var(--text-primary)',
                                    fontSize: '13px',
                                    outline: 'none',
                                    cursor: 'pointer'
                                }}
                            >
                                {DOCX_FONTS.map(f => (
                                    <option key={f.id} value={f.id} style={{ background: '#0f172a', color: '#ffffff' }}>
                                        {f.name}
                                    </option>
                                ))}
                            </select>
                        </div>

                        {/* 3. 글자 크기 & 줄 간격 */}
                        <div>
                            <label style={{ display: 'block', fontSize: '12.5px', fontWeight: 700, color: 'var(--text-secondary)', marginBottom: '8px' }}>
                                📏 본문 크기 / 줄 간격
                            </label>
                            <div style={{ display: 'flex', gap: '8px' }}>
                                <select
                                    value={options.fontSizePt}
                                    onChange={(e) => updateOption('fontSizePt', parseFloat(e.target.value))}
                                    style={{
                                        flex: 1,
                                        padding: '8px',
                                        borderRadius: '8px',
                                        background: 'var(--bg-secondary)',
                                        border: '1px solid var(--panel-border)',
                                        color: 'var(--text-primary)',
                                        fontSize: '12.5px'
                                    }}
                                >
                                    <option value={10}>10.0 pt</option>
                                    <option value={10.5}>10.5 pt (기본)</option>
                                    <option value={11}>11.0 pt</option>
                                    <option value={12}>12.0 pt</option>
                                </select>
                                <select
                                    value={options.lineSpacing}
                                    onChange={(e) => updateOption('lineSpacing', parseInt(e.target.value))}
                                    style={{
                                        flex: 1,
                                        padding: '8px',
                                        borderRadius: '8px',
                                        background: 'var(--bg-secondary)',
                                        border: '1px solid var(--panel-border)',
                                        color: 'var(--text-primary)',
                                        fontSize: '12.5px'
                                    }}
                                >
                                    <option value={240}>1.0배 (기본)</option>
                                    <option value={276}>1.15배 (표준)</option>
                                    <option value={312}>1.3배 (여유)</option>
                                    <option value={360}>1.5배 (넓게)</option>
                                </select>
                            </div>
                        </div>

                        {/* 4. 머리글/바닥글 & 접두사 */}
                        <div>
                            <label style={{ display: 'block', fontSize: '12.5px', fontWeight: 700, color: 'var(--text-secondary)', marginBottom: '8px' }}>
                                📑 페이지 부가 설정
                            </label>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                                <label style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '12.5px', color: 'var(--text-primary)', cursor: 'pointer' }}>
                                    <input
                                        type="checkbox"
                                        checked={options.includeHeaderFooter}
                                        onChange={(e) => updateOption('includeHeaderFooter', e.target.checked)}
                                        style={{ accentColor: '#2563eb' }}
                                    />
                                    머리글(제목) 및 바닥글(페이지 번호) 삽입
                                </label>
                                <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                                    <span style={{ fontSize: '12px', color: 'var(--text-muted)' }}>파일명 접두사:</span>
                                    <input
                                        type="text"
                                        value={options.prefix}
                                        onChange={(e) => updateOption('prefix', e.target.value)}
                                        placeholder="예: 변환_"
                                        style={{
                                            flex: 1,
                                            padding: '4px 8px',
                                            borderRadius: '6px',
                                            background: 'var(--bg-secondary)',
                                            border: '1px solid var(--panel-border)',
                                            color: 'var(--text-primary)',
                                            fontSize: '12px'
                                        }}
                                    />
                                </div>
                            </div>
                        </div>
                    </div>
                )}
            </div>

            {/* ── 모드 A: 파일 일괄 업로드 UI ── */}
            {mode === 'file' && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
                    {/* 드롭존 영역 */}
                    <div
                        onDragOver={handleDragOver}
                        onDragLeave={handleDragLeave}
                        onDrop={handleDrop}
                        style={{
                            border: `2px dashed ${isDragging ? '#2563eb' : 'rgba(59, 130, 246, 0.3)'}`,
                            borderRadius: '16px',
                            padding: '40px 24px',
                            background: isDragging ? 'rgba(37, 99, 235, 0.12)' : 'rgba(37, 99, 235, 0.03)',
                            display: 'flex',
                            flexDirection: 'column',
                            alignItems: 'center',
                            justifyContent: 'center',
                            gap: '14px',
                            cursor: 'pointer',
                            transition: 'all 0.2s ease',
                            textAlign: 'center'
                        }}
                        onClick={() => fileInputRef.current?.click()}
                    >
                        <input
                            type="file"
                            multiple
                            accept=".md,.markdown,.txt"
                            ref={fileInputRef}
                            onChange={handleFileChange}
                            style={{ display: 'none' }}
                        />

                        <div style={{
                            width: '56px',
                            height: '56px',
                            borderRadius: '50%',
                            background: 'rgba(37, 99, 235, 0.15)',
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                            border: '1px solid rgba(59, 130, 246, 0.3)'
                        }}>
                            <Upload size={28} color="#3b82f6" />
                        </div>

                        <div>
                            <div style={{ fontSize: '16px', fontWeight: 700, color: 'var(--text-primary)', marginBottom: '4px' }}>
                                변환할 마크다운 파일(.md, .markdown, .txt)들을 여기에 끌어다 놓으세요
                            </div>
                            <div style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>
                                복수 파일 동시 선택 지원 • 드롭하는 즉시 자동으로 초고속 MS Word 변환 진행
                            </div>
                        </div>

                        <button
                            type="button"
                            className="interactive"
                            style={{
                                padding: '10px 22px',
                                borderRadius: '10px',
                                background: 'linear-gradient(135deg, #2563eb, #1d4ed8)',
                                border: 'none',
                                color: '#ffffff',
                                fontWeight: 700,
                                fontSize: '13.5px',
                                cursor: 'pointer',
                                display: 'flex',
                                alignItems: 'center',
                                gap: '8px',
                                boxShadow: '0 4px 14px rgba(37, 99, 235, 0.35)'
                            }}
                        >
                            <FileText size={16} /> 내 PC에서 마크다운 파일 선택
                        </button>
                    </div>

                    {/* 변환 진행 중 프로그레스 바 */}
                    {isConverting && (
                        <div className="animate-fade-in" style={{
                            background: 'rgba(37, 99, 235, 0.08)',
                            border: '1px solid rgba(59, 130, 246, 0.2)',
                            borderRadius: '12px',
                            padding: '16px 20px',
                            display: 'flex',
                            flexDirection: 'column',
                            gap: '10px'
                        }}>
                            <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '13px', fontWeight: 600 }}>
                                <span style={{ color: '#60a5fa', display: 'flex', alignItems: 'center', gap: '6px' }}>
                                    <Loader2 size={16} className="animate-spin" /> 변환 처리 중: {progress.currentFile}
                                </span>
                                <span style={{ color: 'var(--text-primary)' }}>
                                    {progress.current} / {progress.total} ({Math.round((progress.current / (progress.total || 1)) * 100)}%)
                                </span>
                            </div>
                            <div style={{ width: '100%', height: '8px', background: 'rgba(255,255,255,0.1)', borderRadius: '4px', overflow: 'hidden' }}>
                                <div style={{
                                    width: `${(progress.current / (progress.total || 1)) * 100}%`,
                                    height: '100%',
                                    background: 'linear-gradient(90deg, #3b82f6, #60a5fa)',
                                    transition: 'width 0.3s ease'
                                }} />
                            </div>
                        </div>
                    )}

                    {/* 변환 완료 목록 및 일괄 저장 액션 바 */}
                    {results.length > 0 && (
                        <div className="animate-fade-in" style={{
                            background: 'rgba(255,255,255,0.02)',
                            border: '1px solid var(--panel-border)',
                            borderRadius: '14px',
                            padding: '20px',
                            display: 'flex',
                            flexDirection: 'column',
                            gap: '16px'
                        }}>
                            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: '12px' }}>
                                <div>
                                    <h3 style={{ margin: 0, fontSize: '15.5px', fontWeight: 700, color: 'var(--text-primary)' }}>
                                        변환 완료 문서 목록 ({results.filter(r => r.status === 'success').length} / {results.length})
                                    </h3>
                                    <p style={{ margin: '2px 0 0', fontSize: '12.5px', color: 'var(--text-secondary)' }}>
                                        각 파일을 개별 다운로드하거나, 전체를 한 번에 원하는 로컬 폴더에 저장할 수 있습니다.
                                    </p>
                                    <div style={{ fontSize: '11.5px', color: '#fbbf24', marginTop: '4px' }}>
                                        💡 <strong>폴더 선택 팁:</strong> 브라우저 보안 정책상 '바탕화면' 루트 대신 <strong>'새 폴더'</strong>를 만들어 선택해 주세요.
                                    </div>
                                </div>

                                <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
                                    <button
                                        onClick={handleSaveToDirectory}
                                        className="interactive"
                                        style={{
                                            padding: '9px 16px',
                                            borderRadius: '8px',
                                            background: '#10b981',
                                            border: 'none',
                                            color: '#ffffff',
                                            fontWeight: 700,
                                            fontSize: '13px',
                                            cursor: 'pointer',
                                            display: 'flex',
                                            alignItems: 'center',
                                            gap: '6px',
                                            boxShadow: '0 2px 8px rgba(16, 185, 129, 0.3)'
                                        }}
                                    >
                                        <FolderDown size={16} /> 원하는 폴더 선택하여 전체 저장
                                    </button>
                                    <button
                                        onClick={handleDownloadAllZip}
                                        className="interactive"
                                        style={{
                                            padding: '9px 16px',
                                            borderRadius: '8px',
                                            background: 'var(--bg-secondary)',
                                            border: '1px solid var(--panel-border)',
                                            color: 'var(--text-primary)',
                                            fontWeight: 600,
                                            fontSize: '13px',
                                            cursor: 'pointer',
                                            display: 'flex',
                                            alignItems: 'center',
                                            gap: '6px'
                                        }}
                                    >
                                        <Archive size={16} /> 전체 ZIP 다운로드
                                    </button>
                                </div>
                            </div>

                            {/* 변환된 파일 리스트 테이블 */}
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', maxHeight: '360px', overflowY: 'auto' }}>
                                {results.map((item, idx) => (
                                    <div
                                        key={idx}
                                        style={{
                                            display: 'flex',
                                            alignItems: 'center',
                                            justifyContent: 'space-between',
                                            padding: '12px 16px',
                                            borderRadius: '10px',
                                            background: 'rgba(0,0,0,0.2)',
                                            border: '1px solid var(--panel-border)'
                                        }}
                                    >
                                        <div style={{ display: 'flex', alignItems: 'center', gap: '12px', minWidth: 0 }}>
                                            <div style={{
                                                width: '32px',
                                                height: '32px',
                                                borderRadius: '6px',
                                                background: item.status === 'success' ? 'rgba(37, 99, 235, 0.15)' : 'rgba(239, 68, 68, 0.15)',
                                                display: 'flex',
                                                alignItems: 'center',
                                                justifyContent: 'center',
                                                flexShrink: 0
                                            }}>
                                                <FileText size={18} color={item.status === 'success' ? '#3b82f6' : '#ef4444'} />
                                            </div>
                                            <div style={{ minWidth: 0 }}>
                                                <div style={{ fontSize: '13.5px', fontWeight: 700, color: 'var(--text-primary)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                                                    {item.fileName}
                                                </div>
                                                <div style={{ fontSize: '11.5px', color: 'var(--text-muted)' }}>
                                                    원본: {item.file.name} ({(item.file.size / 1024).toFixed(1)} KB)
                                                </div>
                                            </div>
                                        </div>

                                        <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexShrink: 0 }}>
                                            {item.status === 'success' ? (
                                                <button
                                                    onClick={() => downloadSingleFile(item.blob, item.fileName)}
                                                    className="interactive"
                                                    style={{
                                                        padding: '6px 12px',
                                                        borderRadius: '6px',
                                                        background: 'rgba(37, 99, 235, 0.15)',
                                                        border: '1px solid rgba(37, 99, 235, 0.35)',
                                                        color: '#60a5fa',
                                                        fontWeight: 600,
                                                        fontSize: '12px',
                                                        cursor: 'pointer',
                                                        display: 'flex',
                                                        alignItems: 'center',
                                                        gap: '4px'
                                                    }}
                                                >
                                                    <Download size={13} /> 다운로드
                                                </button>
                                            ) : (
                                                <span style={{ fontSize: '12px', color: '#ef4444' }}>
                                                    실패: {item.error}
                                                </span>
                                            )}
                                        </div>
                                    </div>
                                ))}
                            </div>
                        </div>
                    )}
                </div>
            )}

            {/* ── 모드 B: 텍스트 직접 입력 / 실시간 변환 UI ── */}
            {mode === 'text' && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '10px' }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                            <button
                                onClick={() => setEditorText(SAMPLE_MARKDOWN)}
                                className="interactive"
                                style={{
                                    padding: '6px 12px',
                                    borderRadius: '6px',
                                    background: 'var(--bg-secondary)',
                                    border: '1px solid var(--panel-border)',
                                    color: 'var(--text-secondary)',
                                    fontSize: '12px',
                                    cursor: 'pointer',
                                    display: 'flex',
                                    alignItems: 'center',
                                    gap: '4px'
                                }}
                            >
                                <Sparkles size={13} color="#f59e0b" /> 표준 제안서 샘플 마크다운 불러오기
                            </button>
                            <button
                                onClick={() => setEditorText('')}
                                className="interactive"
                                style={{
                                    padding: '6px 12px',
                                    borderRadius: '6px',
                                    background: 'transparent',
                                    border: '1px solid var(--panel-border)',
                                    color: 'var(--text-muted)',
                                    fontSize: '12px',
                                    cursor: 'pointer'
                                }}
                            >
                                내용 비우기
                            </button>
                        </div>

                        <button
                            onClick={handleConvertEditorText}
                            disabled={textConverting || !editorText.trim()}
                            className="interactive"
                            style={{
                                padding: '10px 22px',
                                borderRadius: '10px',
                                background: 'linear-gradient(135deg, #2563eb, #1d4ed8)',
                                border: 'none',
                                color: '#ffffff',
                                fontWeight: 700,
                                fontSize: '14px',
                                cursor: textConverting || !editorText.trim() ? 'not-allowed' : 'pointer',
                                display: 'flex',
                                alignItems: 'center',
                                gap: '8px',
                                opacity: textConverting || !editorText.trim() ? 0.6 : 1,
                                boxShadow: '0 4px 14px rgba(37, 99, 235, 0.35)'
                            }}
                        >
                            {textConverting ? <Loader2 size={16} className="animate-spin" /> : <Download size={16} />}
                            지금 바로 MS Word(.docx)로 다운로드
                        </button>
                    </div>

                    <div style={{
                        display: 'grid',
                        gridTemplateColumns: '1fr 1fr',
                        gap: '16px',
                        minHeight: '480px'
                    }}>
                        {/* 좌측: 마크다운 입력창 */}
                        <div style={{
                            display: 'flex',
                            flexDirection: 'column',
                            background: 'var(--bg-secondary)',
                            border: '1px solid var(--panel-border)',
                            borderRadius: '12px',
                            overflow: 'hidden'
                        }}>
                            <div style={{
                                padding: '10px 16px',
                                background: 'rgba(0,0,0,0.2)',
                                borderBottom: '1px solid var(--panel-border)',
                                fontSize: '12.5px',
                                fontWeight: 700,
                                color: 'var(--text-primary)',
                                display: 'flex',
                                justifyContent: 'space-between',
                                alignItems: 'center'
                            }}>
                                <span style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                                    <FileCode size={14} color="#3b82f6" /> 마크다운 원본 텍스트
                                </span>
                                <span style={{ fontSize: '11.5px', color: 'var(--text-muted)' }}>
                                    {editorText.length.toLocaleString()} 자
                                </span>
                            </div>
                            <textarea
                                value={editorText}
                                onChange={(e) => setEditorText(e.target.value)}
                                placeholder="이곳에 마크다운(# 제목, ## 소제목, - 목록, | 표 | 등)을 입력하거나 붙여넣으세요..."
                                style={{
                                    flex: 1,
                                    width: '100%',
                                    padding: '16px',
                                    background: 'transparent',
                                    border: 'none',
                                    color: 'var(--text-primary)',
                                    fontSize: '13.5px',
                                    fontFamily: 'Consolas, Monaco, "Courier New", monospace',
                                    lineHeight: 1.6,
                                    resize: 'none',
                                    outline: 'none',
                                    boxSizing: 'border-box'
                                }}
                            />
                        </div>

                        {/* 우측: 실시간 렌더링 미리보기 */}
                        <div style={{
                            display: 'flex',
                            flexDirection: 'column',
                            background: '#ffffff',
                            color: '#1e293b',
                            border: '1px solid var(--panel-border)',
                            borderRadius: '12px',
                            overflow: 'hidden'
                        }}>
                            <div style={{
                                padding: '10px 16px',
                                background: '#f8fafc',
                                borderBottom: '1px solid #e2e8f0',
                                fontSize: '12.5px',
                                fontWeight: 700,
                                color: '#334155',
                                display: 'flex',
                                justifyContent: 'space-between',
                                alignItems: 'center'
                            }}>
                                <span style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                                    <Eye size={14} color="#2563eb" /> Word 스타일 문서 미리보기
                                </span>
                                <span style={{ fontSize: '11px', color: '#64748b' }}>
                                    A4 용지 규격 렌더링
                                </span>
                            </div>
                            <div
                                style={{
                                    flex: 1,
                                    padding: '24px 28px',
                                    overflowY: 'auto',
                                    fontFamily: options.fontFamily || 'Malgun Gothic',
                                    fontSize: `${options.fontSizePt || 10.5}pt`,
                                    lineHeight: (options.lineSpacing || 276) / 240,
                                }}
                                dangerouslySetInnerHTML={{ __html: marked.parse(editorText || '') }}
                            />
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
}
