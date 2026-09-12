import React, { useState, useRef } from 'react';
import { FileSpreadsheet, Loader2, Upload, Presentation, Info, Sparkles } from 'lucide-react';
import { parseExcelData, generatePptFromTemplate } from '../utils/pptExporter';

// 엑셀 ➜ PPT 데이터 매핑 생성 탭. PptGenerator.jsx의 excel_mapping 서브탭에서 분리됨.
export default function PptExcelMappingTab({ setErrorMsg, setSuccessMsg }) {
    const [excelFile, setExcelFile] = useState(null);
    const [pptTemplate, setPptTemplate] = useState(null);
    const [excelDataPreview, setExcelDataPreview] = useState(null);
    const [templateLabel, setTemplateLabel] = useState('');
    const [isParsing, setIsParsing] = useState(false);
    const [isGenerating, setIsGenerating] = useState(false);
    // 향후 분할 생성 모드 확장을 위한 값 - 현재는 UI 토글이 없어 'single' 고정
    const [generationMode] = useState('single');
    const [isDraggingExcel, setIsDraggingExcel] = useState(false);
    const [isDraggingTemplate, setIsDraggingTemplate] = useState(false);
    const excelInputRef = useRef(null);
    const pptInputRef = useRef(null);

    const processExcelFile = async (file) => {
        if (!file) return;
        setErrorMsg(null);
        setSuccessMsg(null);
        setExcelFile(file);
        setIsParsing(true);
        setExcelDataPreview(null);
        try {
            const data = await parseExcelData(file);
            if (data.length === 0) {
                setErrorMsg('엑셀 파일에 데이터가 없습니다.');
            } else {
                setExcelDataPreview(data);
            }
        } catch (err) {
            console.error(err);
            setErrorMsg('엑셀 파일 분석 중 오류가 발생했습니다. 올바른 파일인지 확인해주세요.');
        } finally {
            setIsParsing(false);
        }
    };

    const handleExcelChange = (e) => processExcelFile(e.target.files[0]);

    const processTemplateFile = (file) => {
        if (!file) return;
        const ext = file.name.split('.').pop().toLowerCase();
        if (ext !== 'pptx') {
            setErrorMsg('PPT 템플릿은 .pptx 확장자만 지원합니다.');
            return;
        }
        setErrorMsg(null);
        setSuccessMsg(null);
        setPptTemplate(file);
        setTemplateLabel(file.name);
    };

    const handleTemplateChange = (e) => processTemplateFile(e.target.files[0]);

    const handleExcelDragEvents = {
        onDragOver: (e) => { e.preventDefault(); e.stopPropagation(); },
        onDragEnter: (e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingExcel(true); },
        onDragLeave: (e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingExcel(false); },
        onDrop: (e) => {
            e.preventDefault(); e.stopPropagation(); setIsDraggingExcel(false);
            const file = e.dataTransfer.files[0];
            if (file) processExcelFile(file);
        }
    };

    const handleTemplateDragEvents = {
        onDragOver: (e) => { e.preventDefault(); e.stopPropagation(); },
        onDragEnter: (e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingTemplate(true); },
        onDragLeave: (e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingTemplate(false); },
        onDrop: (e) => {
            e.preventDefault(); e.stopPropagation(); setIsDraggingTemplate(false);
            const file = e.dataTransfer.files[0];
            if (file) processTemplateFile(file);
        }
    };

    const handleGenerate = async () => {
        if (!excelFile || !pptTemplate || !excelDataPreview) {
            setErrorMsg('엑셀 파일과 PPT 템플릿을 모두 등록해주세요.');
            return;
        }
        setErrorMsg(null);
        setSuccessMsg(null);
        setIsGenerating(true);
        await new Promise(r => setTimeout(r, 800));
        try {
            await generatePptFromTemplate(pptTemplate, excelDataPreview, generationMode);
            setSuccessMsg('성공적으로 PPT 파일이 생성되어 다운로드되었습니다.');
            setExcelFile(null);
            setPptTemplate(null);
            setExcelDataPreview(null);
            setTemplateLabel('');
        } catch (err) {
            console.error(err);
            if (err.message && err.message.includes("Can't find end of central directory")) {
                setErrorMsg('유효하지 않은 PPTX 파일입니다. 손상되었거나 암호가 걸려있을 수 있습니다.');
            } else {
                setErrorMsg(`변환 중 오류 발생: ${err.message || '알 수 없는 오류'}`);
            }
        } finally {
            setIsGenerating(false);
        }
    };

    const columns = excelDataPreview && excelDataPreview.length > 0 ? Object.keys(excelDataPreview[0]) : [];

    return (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
                {/* 데이터 엑셀 업로드 영역 */}
                <div
                    {...handleExcelDragEvents}
                    style={{
                        border: `2px dashed ${isDraggingExcel ? 'var(--success-color)' : 'rgba(34, 197, 94, 0.3)'}`,
                        borderRadius: '12px', padding: '24px',
                        background: isDraggingExcel ? 'rgba(34, 197, 94, 0.1)' : 'rgba(34, 197, 94, 0.05)',
                        display: 'flex', flexDirection: 'column', gap: '16px',
                        transition: 'all 0.2s ease', cursor: isDraggingExcel ? 'copy' : 'default'
                    }}
                >
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <FileSpreadsheet size={20} color="var(--success-color)" />
                        <h3 style={{ margin: 0, fontSize: '16px', color: 'var(--text-primary)' }}>1. 엑셀 데이터 등록</h3>
                    </div>

                    <input
                        type="file"
                        accept=".xlsx, .xls"
                        onChange={handleExcelChange}
                        ref={excelInputRef}
                        style={{ display: 'none' }}
                    />

                    <div style={{ display: 'flex', gap: '12px', alignItems: 'center' }}>
                        <button
                            onClick={() => excelInputRef.current?.click()}
                            className="interactive"
                            style={{
                                display: 'flex', alignItems: 'center', gap: '8px',
                                padding: '10px 16px', borderRadius: '8px', cursor: 'pointer',
                                background: 'var(--bg-secondary)', border: '1px solid var(--panel-border)',
                                color: 'var(--text-primary)', fontWeight: 600, fontSize: '13px'
                            }}
                        >
                            {isParsing ? <Loader2 size={16} className="animate-spin" /> : <Upload size={16} />}
                            엑셀 찾아보기 (.xlsx)
                        </button>
                        {excelFile && (
                            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', background: 'rgba(0,0,0,0.2)', padding: '6px 12px', borderRadius: '6px', fontSize: '13px', border: '1px solid var(--success-color)' }}>
                                <span style={{ color: 'var(--success-color)' }}>✔</span>
                                {excelFile.name}
                            </div>
                        )}
                    </div>

                    <div style={{ fontSize: '12px', color: 'var(--text-secondary)', lineHeight: '1.5' }}>
                        💡 1행이 라벨(변수명)이 되고, 2행부터 실제 데이터로 간주합니다.
                    </div>
                </div>

                {/* PPT 양식 파일 업로드 영역 */}
                <div
                    {...handleTemplateDragEvents}
                    style={{
                        border: `2px dashed ${isDraggingTemplate ? 'var(--accent-blue)' : 'rgba(59, 130, 246, 0.3)'}`,
                        borderRadius: '12px', padding: '24px',
                        background: isDraggingTemplate ? 'rgba(59, 130, 246, 0.1)' : 'rgba(59, 130, 246, 0.05)',
                        display: 'flex', flexDirection: 'column', gap: '16px',
                        transition: 'all 0.2s ease', cursor: isDraggingTemplate ? 'copy' : 'default'
                    }}
                >
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <Presentation size={20} color="var(--accent-blue)" />
                        <h3 style={{ margin: 0, fontSize: '16px', color: 'var(--text-primary)' }}>2. PPT 양식(.pptx) 등록</h3>
                    </div>

                    <input
                        type="file"
                        accept=".pptx"
                        onChange={handleTemplateChange}
                        ref={pptInputRef}
                        style={{ display: 'none' }}
                    />

                    <div style={{ display: 'flex', gap: '12px', alignItems: 'center' }}>
                        <button
                            onClick={() => pptInputRef.current?.click()}
                            className="interactive"
                            style={{
                                display: 'flex', alignItems: 'center', gap: '8px',
                                padding: '10px 16px', borderRadius: '8px', cursor: 'pointer',
                                background: 'var(--bg-secondary)', border: '1px solid var(--panel-border)',
                                color: 'var(--text-primary)', fontWeight: 600, fontSize: '13px'
                            }}
                        >
                            <Upload size={16} /> PPT 템플릿 찾기
                        </button>
                        {pptTemplate && (
                            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', background: 'rgba(0,0,0,0.2)', padding: '6px 12px', borderRadius: '6px', fontSize: '13px', border: '1px solid var(--accent-blue)' }}>
                                <span style={{ color: 'var(--accent-blue)' }}>✔</span>
                                {templateLabel}
                            </div>
                        )}
                    </div>

                    <div style={{
                        fontSize: '13px', color: 'var(--text-secondary)', lineHeight: '1.6',
                        marginTop: '4px', background: 'rgba(255,255,255,0.03)',
                        padding: '14px 18px', borderRadius: '8px', borderLeft: '3px solid var(--accent-blue)'
                    }}>
                        <div style={{ fontWeight: 700, color: 'var(--text-primary)', marginBottom: '8px', display: 'flex', alignItems: 'center', gap: '6px' }}>
                            💡 PPT 템플릿 매핑 가이드
                        </div>
                        <p style={{ margin: 0 }}>
                            PPT 내 텍스트 상자에 <code>{`{열이름}`}</code> (예: <code>{`{성명}`}</code>, <code>{`{부서}`}</code>) 형식으로 입력하세요.<br/>
                            <strong style={{ color: 'var(--accent-blue)' }}>* 자동 슬라이드 복제:</strong> 엑셀 행 개수만큼 슬라이드가 자동으로 생성되며 각 행의 데이터가 각 슬라이드에 채워집니다.
                        </p>
                    </div>
                </div>
            </div>

            {/* 데이터 컬럼 매핑 프리뷰 */}
            {excelDataPreview && (
                <div className="animate-fade-in" style={{ padding: '20px', background: 'rgba(255,255,255,0.03)', border: '1px solid var(--panel-border)', borderRadius: '12px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '12px' }}>
                        <Info size={18} color="var(--warning-color)" />
                        <h4 style={{ margin: 0, fontSize: '15px' }}>사용 가능한 템플릿 태그 (총 {excelDataPreview.length}개 행 인식됨)</h4>
                    </div>
                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px' }}>
                        {columns.map(col => (
                            <div key={col} style={{
                                padding: '4px 8px', background: 'rgba(168,85,247,0.1)', color: '#c084fc',
                                borderRadius: '6px', fontSize: '13px', fontFamily: 'monospace', fontWeight: 600, border: '1px solid rgba(168,85,247,0.2)'
                            }}>
                                {`{${col}}`}
                            </div>
                        ))}
                    </div>
                </div>
            )}

            <div style={{ marginTop: 'auto', paddingTop: '20px' }}>
                <button
                    className="interactive"
                    onClick={handleGenerate}
                    disabled={!excelFile || !pptTemplate || isGenerating}
                    style={{
                        width: '100%',
                        padding: '16px',
                        background: (!excelFile || !pptTemplate || isGenerating) ? 'rgba(255,255,255,0.05)' : 'linear-gradient(135deg, #0284c7, #3b82f6)',
                        color: (!excelFile || !pptTemplate || isGenerating) ? 'var(--text-muted)' : 'white',
                        border: 'none',
                        borderRadius: '12px',
                        fontSize: '16px',
                        fontWeight: 700,
                        cursor: (!excelFile || !pptTemplate || isGenerating) ? 'not-allowed' : 'pointer',
                        display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '10px'
                    }}
                >
                    {isGenerating ? (
                        <><Loader2 size={20} className="animate-spin" /> 파워포인트 문서 자동 치환 및 생성 중...</>
                    ) : (
                        <><Sparkles size={20} /> 엑셀 ↔ PPT 자동 매핑 및 파일 변환 시작</>
                    )}
                </button>
            </div>
        </div>
    );
}
