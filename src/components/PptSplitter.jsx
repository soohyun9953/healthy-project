import React, { useState, useRef, useMemo } from 'react';
import { 
    Split, Upload, FileText, Download, CheckCircle2, AlertCircle, 
    RefreshCw, Layers, Plus, Trash2, Edit3, Save, FileArchive, ArrowRight,
    Sliders, Check, FolderDown, CheckSquare, Square, MinusSquare
} from 'lucide-react';
import { analyzePptxSections, createSubPptx, downloadAllSectionsAsZip } from '../utils/pptSplitter';
import { saveAs } from 'file-saver';

export default function PptSplitter() {
    const [file, setFile] = useState(null);
    const [fileBuffer, setFileBuffer] = useState(null);
    const [isDragging, setIsDragging] = useState(false);
    const [isAnalyzing, setIsAnalyzing] = useState(false);
    const [isSplitting, setIsSplitting] = useState(false);
    const [splitProgress, setSplitProgress] = useState({ text: '', percent: 0 });
    const [analysisResult, setAnalysisResult] = useState(null); // { totalSlides, sections, slidesInfo }
    const [sections, setSections] = useState([]);
    const [fileNamePrefix, setFileNamePrefix] = useState('[분할]');
    const [statusMessage, setStatusMessage] = useState(null);
    const fileInputRef = useRef(null);

    // 선택된 섹션 계산
    const selectedSections = useMemo(() => {
        return sections.filter(sec => sec.selected !== false);
    }, [sections]);

    const isAllSelected = sections.length > 0 && selectedSections.length === sections.length;
    const isSomeSelected = selectedSections.length > 0 && selectedSections.length < sections.length;

    // 파일 선택 및 분석 시작
    const handleFile = async (uploadedFile) => {
        if (!uploadedFile) return;
        if (!uploadedFile.name.toLowerCase().endsWith('.pptx')) {
            alert('PPTX 형식의 파워포인트 파일만 지원됩니다.');
            return;
        }

        setFile(uploadedFile);
        setIsAnalyzing(true);
        setStatusMessage(null);
        setAnalysisResult(null);
        setSections([]);

        try {
            const buffer = await uploadedFile.arrayBuffer();
            setFileBuffer(buffer);

            const result = await analyzePptxSections(buffer);
            setAnalysisResult(result);
            // 모든 섹션 기본 선택 상태로 초기화
            const initializedSections = result.sections.map(s => ({ ...s, selected: true }));
            setSections(initializedSections);
            setStatusMessage({
                type: 'success',
                text: `총 ${result.totalSlides}장의 슬라이드에서 ${result.sections.length}개의 주요 목차(섹션)를 자동 감지하였습니다. 분할할 목차를 선택 후 다운로드하세요.`
            });
        } catch (err) {
            console.error('PPTX 목차 분석 실패:', err);
            setStatusMessage({
                type: 'error',
                text: `PPTX 파일 분석 중 오류가 발생했습니다: ${err.message}`
            });
        } finally {
            setIsAnalyzing(false);
        }
    };

    const handleDrop = (e) => {
        e.preventDefault();
        setIsDragging(false);
        if (e.dataTransfer.files && e.dataTransfer.files[0]) {
            handleFile(e.dataTransfer.files[0]);
        }
    };

    // 개별 섹션 선택 토글
    const handleToggleSelect = (index) => {
        const updated = [...sections];
        updated[index] = {
            ...updated[index],
            selected: !updated[index].selected
        };
        setSections(updated);
    };

    // 전체 선택 / 전체 해제 토글
    const handleToggleSelectAll = () => {
        const targetState = !isAllSelected;
        const updated = sections.map(s => ({ ...s, selected: targetState }));
        setSections(updated);
    };

    // 선택 반전
    const handleInvertSelect = () => {
        const updated = sections.map(s => ({ ...s, selected: !s.selected }));
        setSections(updated);
    };

    // 섹션 추가
    const handleAddSection = () => {
        const lastSection = sections[sections.length - 1];
        const nextStart = lastSection ? Math.min(lastSection.endSlide + 1, analysisResult?.totalSlides || 1) : 1;
        const total = analysisResult?.totalSlides || 1;

        const newSec = {
            id: sections.length + 1,
            title: `새 목차 섹션 ${sections.length + 1}`,
            startSlide: nextStart,
            endSlide: total,
            slideCount: Math.max(1, total - nextStart + 1),
            previewText: '사용자 추가 섹션',
            selected: true
        };
        setSections([...sections, newSec]);
    };

    // 단일 섹션 삭제
    const handleDeleteSection = (index) => {
        if (sections.length <= 1) {
            alert('최소 1개 이상의 섹션이 존재해야 합니다.');
            return;
        }
        const updated = sections.filter((_, i) => i !== index).map((sec, idx) => ({
            ...sec,
            id: idx + 1
        }));
        setSections(updated);
    };

    // 선택된 섹션 일괄 삭제
    const handleDeleteSelected = () => {
        if (selectedSections.length === 0) {
            alert('삭제할 섹션을 선택해주세요.');
            return;
        }
        if (selectedSections.length === sections.length) {
            alert('모든 섹션을 삭제할 수는 없습니다. 최소 1개는 유지해야 합니다.');
            return;
        }
        if (!confirm(`선택한 ${selectedSections.length}개 목차를 목록에서 삭제하시겠습니까?`)) {
            return;
        }
        const remaining = sections.filter(s => !s.selected).map((sec, idx) => ({
            ...sec,
            id: idx + 1
        }));
        setSections(remaining);
    };

    // 섹션 정보 수정
    const handleSectionChange = (index, field, value) => {
        const updated = [...sections];
        const target = { ...updated[index] };

        if (field === 'startSlide' || field === 'endSlide') {
            const numVal = parseInt(value, 10) || 1;
            const clamped = Math.max(1, Math.min(numVal, analysisResult?.totalSlides || 1));
            target[field] = clamped;
            if (target.endSlide < target.startSlide) {
                target.endSlide = target.startSlide;
            }
            target.slideCount = target.endSlide - target.startSlide + 1;
        } else {
            target[field] = value;
        }

        updated[index] = target;
        setSections(updated);
    };

    // 개별 섹션 단독 다운로드
    const handleDownloadSingleSection = async (sec, index) => {
        if (!fileBuffer || !file) return;
        try {
            const subBlob = await createSubPptx(fileBuffer, sec.startSlide, sec.endSlide);
            const baseName = file.name.replace(/\.[^.]+$/, '');
            const cleanTitle = (sec.title || `섹션_${sec.id}`).replace(/[\/:*?"<>|]/g, '_').trim();
            const prefix = fileNamePrefix ? `${fileNamePrefix.trim()}_` : '';
            const downloadName = `${prefix}[${String(index + 1).padStart(2, '0')}]_${cleanTitle}_${baseName}.pptx`;
            saveAs(subBlob, downloadName);
        } catch (err) {
            console.error('단일 섹션 분할 다운로드 실패:', err);
            alert(`분할 파일 생성 실패: ${err.message}`);
        }
    };

    // 선택된 섹션들 일괄 ZIP 다운로드
    const handleDownloadSelectedZip = async () => {
        if (!fileBuffer || !file) return;
        if (selectedSections.length === 0) {
            alert('분할하여 다운로드할 목차를 1개 이상 선택해주세요.');
            return;
        }

        setIsSplitting(true);
        setSplitProgress({ text: '선택된 목차 분할 준비 중...', percent: 5 });

        try {
            await downloadAllSectionsAsZip(
                fileBuffer,
                file.name,
                selectedSections,
                (msg, pct) => setSplitProgress({ text: msg, percent: pct }),
                fileNamePrefix
            );
            setStatusMessage({
                type: 'success',
                text: `성공적으로 선택된 ${selectedSections.length}개의 목차별 PPTX 파일을 생성하여 ZIP 압축 다운로드하였습니다.`
            });
        } catch (err) {
            console.error('ZIP 분할 다운로드 실패:', err);
            setStatusMessage({
                type: 'error',
                text: `ZIP 일괄 분할 생성 중 오류가 발생했습니다: ${err.message}`
            });
        } finally {
            setIsSplitting(false);
        }
    };

    const handleReset = () => {
        setFile(null);
        setFileBuffer(null);
        setAnalysisResult(null);
        setSections([]);
        setStatusMessage(null);
    };

    return (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '24px', color: 'var(--text-primary)' }}>
            
            {/* 타이틀 및 기능 안내 배너 */}
            <div style={{
                display: 'flex', justifyContent: 'space-between', alignItems: 'center',
                background: 'linear-gradient(135deg, rgba(168, 85, 247, 0.15), rgba(59, 130, 246, 0.08))',
                border: '1px solid rgba(168, 85, 247, 0.25)', padding: '24px', borderRadius: '16px', backdropFilter: 'blur(10px)'
            }}>
                <div>
                    <h2 style={{ display: 'flex', alignItems: 'center', gap: '10px', fontSize: '22px', fontWeight: 800, margin: 0 }}>
                        <Split size={26} color="#a855f7" /> PPT 목차 분석 기반 하위 파일 선택 분리 생성
                    </h2>
                    <p style={{ margin: '8px 0 0 0', fontSize: '14px', color: 'var(--text-secondary)', lineHeight: '1.6' }}>
                        대용량 PPTX 파일을 업로드하면 내부 목차(대주제/장/간지)를 AI 및 구조 분석 엔진이 자동 감지합니다.<br />
                        <strong>분할을 원하는 목차만 체크박스로 선택하여 독립 PPTX 파일들로 무손실 분할 생성 및 일괄 ZIP 다운로드</strong>할 수 있습니다.
                    </p>
                </div>
                {file && (
                    <button
                        onClick={handleReset}
                        className="interactive"
                        style={{
                            padding: '10px 16px', background: 'rgba(255,255,255,0.05)', color: 'var(--text-primary)',
                            border: '1px solid var(--panel-border)', borderRadius: '8px', cursor: 'pointer',
                            display: 'flex', alignItems: 'center', gap: '8px', fontSize: '13px', fontWeight: 600
                        }}
                    >
                        <RefreshCw size={15} /> 새로 올리기
                    </button>
                )}
            </div>

            {/* 메인 파일 업로드 영역 (파일이 없을 때) */}
            {!file && (
                <div
                    onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
                    onDragLeave={() => setIsDragging(false)}
                    onDrop={handleDrop}
                    onClick={() => fileInputRef.current?.click()}
                    style={{
                        border: `2px dashed ${isDragging ? 'var(--accent-purple)' : 'rgba(168, 85, 247, 0.35)'}`,
                        borderRadius: '16px', padding: '60px 20px',
                        background: isDragging ? 'rgba(168, 85, 247, 0.12)' : 'rgba(168, 85, 247, 0.03)',
                        display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center',
                        gap: '16px', cursor: 'pointer', transition: 'all 0.2s ease', textAlign: 'center'
                    }}
                >
                    <div style={{
                        width: '64px', height: '64px', borderRadius: '50%',
                        background: 'rgba(168, 85, 247, 0.15)', display: 'flex', alignItems: 'center', justifyContent: 'center'
                    }}>
                        {isAnalyzing ? <RefreshCw size={32} color="#a855f7" className="animate-spin" /> : <Upload size={32} color="#a855f7" />}
                    </div>

                    <div>
                        <h3 style={{ margin: '0 0 6px 0', fontSize: '18px', fontWeight: 700 }}>
                            {isAnalyzing ? 'PPTX 목차 및 슬라이드 구조 분석 중...' : '분할할 PPTX 파일을 여기에 끌어다 놓으세요'}
                        </h3>
                        <p style={{ margin: 0, fontSize: '13.5px', color: 'var(--text-secondary)' }}>
                            또는 클릭하여 컴퓨터에서 파일 선택 (.pptx)
                        </p>
                    </div>

                    <input
                        type="file"
                        accept=".pptx"
                        ref={fileInputRef}
                        onChange={(e) => e.target.files && handleFile(e.target.files[0])}
                        style={{ display: 'none' }}
                    />
                </div>
            )}

            {/* 상태 메시지 배너 */}
            {statusMessage && (
                <div style={{
                    padding: '14px 18px', borderRadius: '10px', fontSize: '13.5px', fontWeight: 600,
                    display: 'flex', alignItems: 'center', gap: '10px',
                    background: statusMessage.type === 'success' ? 'rgba(34, 197, 94, 0.1)' : 'rgba(239, 68, 68, 0.1)',
                    border: `1px solid ${statusMessage.type === 'success' ? 'rgba(34, 197, 94, 0.25)' : 'rgba(239, 68, 68, 0.25)'}`,
                    color: statusMessage.type === 'success' ? 'var(--success-color)' : 'var(--danger-color)'
                }}>
                    {statusMessage.type === 'success' ? <CheckCircle2 size={18} /> : <AlertCircle size={18} />}
                    {statusMessage.text}
                </div>
            )}

            {/* 분석 완료 후 섹션 선택/편집 및 다운로드 대시보드 */}
            {analysisResult && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
                    
                    {/* 상단 컨트롤 바 */}
                    <div style={{
                        background: 'var(--panel-bg)', border: '1px solid var(--panel-border)',
                        borderRadius: '16px', padding: '20px', display: 'flex', justifyContent: 'space-between',
                        alignItems: 'center', flexWrap: 'wrap', gap: '16px'
                    }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '16px', flexWrap: 'wrap' }}>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                <FileText size={20} color="var(--accent-purple)" />
                                <span style={{ fontWeight: 700, fontSize: '15px' }}>{file?.name}</span>
                                <span style={{ fontSize: '12px', background: 'rgba(255,255,255,0.08)', padding: '3px 8px', borderRadius: '6px', color: 'var(--text-secondary)' }}>
                                    총 {analysisResult.totalSlides}장
                                </span>
                            </div>

                            {/* 선택 통계 뱃지 */}
                            <div style={{
                                display: 'flex', alignItems: 'center', gap: '6px',
                                background: selectedSections.length > 0 ? 'rgba(168, 85, 247, 0.15)' : 'rgba(255,255,255,0.05)',
                                border: `1px solid ${selectedSections.length > 0 ? 'rgba(168, 85, 247, 0.3)' : 'var(--panel-border)'}`,
                                padding: '4px 12px', borderRadius: '20px', fontSize: '12.5px', fontWeight: 700,
                                color: selectedSections.length > 0 ? '#c084fc' : 'var(--text-muted)'
                            }}>
                                <CheckCircle2 size={14} />
                                <span>선택: {selectedSections.length} / {sections.length}개 목차</span>
                            </div>

                            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                <span style={{ fontSize: '12.5px', color: 'var(--text-muted)' }}>파일명 접두사:</span>
                                <input
                                    type="text"
                                    value={fileNamePrefix}
                                    onChange={(e) => setFileNamePrefix(e.target.value)}
                                    placeholder="[분할]"
                                    style={{
                                        padding: '6px 10px', background: 'rgba(0,0,0,0.2)',
                                        border: '1px solid var(--panel-border)', borderRadius: '6px',
                                        color: 'var(--text-primary)', fontSize: '12.5px', width: '90px'
                                    }}
                                />
                            </div>
                        </div>

                        <div style={{ display: 'flex', gap: '10px', alignItems: 'center', flexWrap: 'wrap' }}>
                            <button
                                onClick={handleAddSection}
                                className="interactive"
                                style={{
                                    display: 'flex', alignItems: 'center', gap: '6px',
                                    padding: '10px 14px', background: 'rgba(255,255,255,0.05)',
                                    border: '1px solid var(--panel-border)', borderRadius: '10px',
                                    color: 'var(--text-primary)', fontSize: '13px', fontWeight: 600, cursor: 'pointer'
                                }}
                            >
                                <Plus size={16} /> 목차 추가
                            </button>

                            {selectedSections.length > 0 && selectedSections.length < sections.length && (
                                <button
                                    onClick={handleDeleteSelected}
                                    className="interactive"
                                    title="선택된 목차 항목들을 목록에서 삭제"
                                    style={{
                                        display: 'flex', alignItems: 'center', gap: '6px',
                                        padding: '10px 14px', background: 'rgba(239, 68, 68, 0.1)',
                                        border: '1px solid rgba(239, 68, 68, 0.25)', borderRadius: '10px',
                                        color: '#f87171', fontSize: '13px', fontWeight: 600, cursor: 'pointer'
                                    }}
                                >
                                    <Trash2 size={15} /> 선택 삭제 ({selectedSections.length})
                                </button>
                            )}

                            <button
                                onClick={handleDownloadSelectedZip}
                                disabled={isSplitting || selectedSections.length === 0}
                                className="interactive"
                                style={{
                                    display: 'flex', alignItems: 'center', gap: '8px',
                                    padding: '10px 22px',
                                    background: selectedSections.length > 0 
                                        ? 'linear-gradient(135deg, #a855f7 0%, #3b82f6 100%)' 
                                        : 'rgba(255,255,255,0.1)',
                                    border: 'none', borderRadius: '10px',
                                    color: selectedSections.length > 0 ? '#ffffff' : 'var(--text-muted)',
                                    fontSize: '14px', fontWeight: 700,
                                    cursor: (isSplitting || selectedSections.length === 0) ? 'not-allowed' : 'pointer',
                                    boxShadow: selectedSections.length > 0 ? '0 4px 14px rgba(168, 85, 247, 0.3)' : 'none',
                                    transition: 'all 0.2s ease'
                                }}
                            >
                                {isSplitting ? <RefreshCw size={17} className="animate-spin" /> : <FolderDown size={17} />}
                                <span>
                                    {isSplitting 
                                        ? '분할 파일 생성 중...' 
                                        : selectedSections.length === 0 
                                            ? '목차를 선택해주세요' 
                                            : `⚡ 선택한 ${selectedSections.length}개 목차 분리 (ZIP 다운로드)`
                                    }
                                </span>
                            </button>
                        </div>
                    </div>

                    {/* 진행 상태 게이지 */}
                    {isSplitting && (
                        <div style={{
                            background: 'var(--panel-bg)', border: '1px solid var(--panel-border)',
                            borderRadius: '12px', padding: '16px 20px', display: 'flex', flexDirection: 'column', gap: '10px'
                        }}>
                            <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '13px', fontWeight: 600 }}>
                                <span style={{ color: 'var(--accent-purple)' }}>{splitProgress.text}</span>
                                <span>{splitProgress.percent}%</span>
                            </div>
                            <div style={{ width: '100%', height: '6px', background: 'rgba(255,255,255,0.1)', borderRadius: '3px', overflow: 'hidden' }}>
                                <div style={{ width: `${splitProgress.percent}%`, height: '100%', background: 'linear-gradient(90deg, #a855f7, #3b82f6)', transition: 'width 0.3s ease' }} />
                            </div>
                        </div>
                    )}

                    {/* 목차별 분할 선택 및 구성표 */}
                    <div style={{
                        background: 'var(--panel-bg)', border: '1px solid var(--panel-border)',
                        borderRadius: '16px', padding: '24px', display: 'flex', flexDirection: 'column', gap: '16px'
                    }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px' }}>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                                <Layers size={18} color="var(--accent-purple)" />
                                <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 700 }}>
                                    분할 대상 목차 목록 ({sections.length}개 감지됨)
                                </h3>
                            </div>
                            
                            {/* 선택 보조 빠른 버튼군 */}
                            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                <button
                                    onClick={handleToggleSelectAll}
                                    className="interactive"
                                    style={{
                                        padding: '5px 10px', background: 'rgba(255,255,255,0.05)',
                                        border: '1px solid var(--panel-border)', borderRadius: '6px',
                                        color: 'var(--text-secondary)', fontSize: '12px', fontWeight: 600, cursor: 'pointer',
                                        display: 'inline-flex', alignItems: 'center', gap: '4px'
                                    }}
                                >
                                    {isAllSelected ? <Square size={13} /> : <CheckSquare size={13} />}
                                    {isAllSelected ? '선택 해제' : '전체 선택'}
                                </button>
                                <button
                                    onClick={handleInvertSelect}
                                    className="interactive"
                                    style={{
                                        padding: '5px 10px', background: 'rgba(255,255,255,0.05)',
                                        border: '1px solid var(--panel-border)', borderRadius: '6px',
                                        color: 'var(--text-secondary)', fontSize: '12px', fontWeight: 600, cursor: 'pointer'
                                    }}
                                >
                                    선택 반전
                                </button>
                                <span style={{ fontSize: '12px', color: 'var(--text-muted)', marginLeft: '6px' }}>
                                    💡 체크박스로 분할할 목차를 선택하세요.
                                </span>
                            </div>
                        </div>

                        <div style={{ overflowX: 'auto' }}>
                            <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '13.5px' }}>
                                <thead>
                                    <tr style={{ borderBottom: '2px solid var(--panel-border)', color: 'var(--text-secondary)', background: 'rgba(255,255,255,0.02)' }}>
                                        <th style={{ padding: '12px 14px', width: '50px', textAlign: 'center' }}>
                                            <input
                                                type="checkbox"
                                                checked={isAllSelected}
                                                ref={el => {
                                                    if (el) el.indeterminate = isSomeSelected;
                                                }}
                                                onChange={handleToggleSelectAll}
                                                title={isAllSelected ? '전체 해제' : '전체 선택'}
                                                style={{ cursor: 'pointer', width: '16px', height: '16px', accentColor: '#a855f7' }}
                                            />
                                        </th>
                                        <th style={{ padding: '12px 10px', width: '50px' }}>순번</th>
                                        <th style={{ padding: '12px 14px', width: '35%' }}>분할 목차(섹션)명</th>
                                        <th style={{ padding: '12px 14px', width: '15%' }}>시작 슬라이드</th>
                                        <th style={{ padding: '12px 14px', width: '15%' }}>종료 슬라이드</th>
                                        <th style={{ padding: '12px 14px', width: '10%' }}>슬라이드 수</th>
                                        <th style={{ padding: '12px 14px', width: '14%', textAlign: 'center' }}>개별 다운로드</th>
                                        <th style={{ padding: '12px 14px', width: '50px', textAlign: 'center' }}>삭제</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {sections.map((sec, idx) => {
                                        const isSelected = sec.selected !== false;
                                        return (
                                            <tr 
                                                key={idx} 
                                                style={{ 
                                                    borderBottom: '1px solid rgba(255,255,255,0.05)', 
                                                    backgroundColor: isSelected 
                                                        ? (idx % 2 === 0 ? 'rgba(168, 85, 247, 0.05)' : 'rgba(168, 85, 247, 0.02)') 
                                                        : 'rgba(0,0,0,0.2)',
                                                    opacity: isSelected ? 1 : 0.6,
                                                    transition: 'all 0.15s ease'
                                                }}
                                            >
                                                <td style={{ padding: '12px 14px', textAlign: 'center' }}>
                                                    <input
                                                        type="checkbox"
                                                        checked={isSelected}
                                                        onChange={() => handleToggleSelect(idx)}
                                                        style={{ cursor: 'pointer', width: '16px', height: '16px', accentColor: '#a855f7' }}
                                                    />
                                                </td>
                                                <td style={{ padding: '12px 10px', fontWeight: 700, color: isSelected ? 'var(--text-secondary)' : 'var(--text-muted)' }}>
                                                    {idx + 1}
                                                </td>
                                                <td style={{ padding: '12px 14px' }}>
                                                    <input
                                                        type="text"
                                                        value={sec.title}
                                                        onChange={(e) => handleSectionChange(idx, 'title', e.target.value)}
                                                        style={{
                                                            width: '95%', padding: '7px 10px',
                                                            background: 'rgba(0,0,0,0.2)', border: '1px solid var(--panel-border)',
                                                            borderRadius: '6px', color: isSelected ? 'var(--text-primary)' : 'var(--text-muted)', 
                                                            fontSize: '13.5px', fontWeight: 600
                                                        }}
                                                    />
                                                </td>
                                                <td style={{ padding: '12px 14px' }}>
                                                    <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                                                        <input
                                                            type="number"
                                                            min={1}
                                                            max={analysisResult.totalSlides}
                                                            value={sec.startSlide}
                                                            onChange={(e) => handleSectionChange(idx, 'startSlide', e.target.value)}
                                                            style={{
                                                                width: '65px', padding: '6px 8px',
                                                                background: 'rgba(0,0,0,0.2)', border: '1px solid var(--panel-border)',
                                                                borderRadius: '6px', color: 'var(--text-primary)', fontSize: '13px', textAlign: 'center'
                                                            }}
                                                        />
                                                        <span style={{ fontSize: '12px', color: 'var(--text-muted)' }}>p</span>
                                                    </div>
                                                </td>
                                                <td style={{ padding: '12px 14px' }}>
                                                    <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                                                        <input
                                                            type="number"
                                                            min={sec.startSlide}
                                                            max={analysisResult.totalSlides}
                                                            value={sec.endSlide}
                                                            onChange={(e) => handleSectionChange(idx, 'endSlide', e.target.value)}
                                                            style={{
                                                                width: '65px', padding: '6px 8px',
                                                                background: 'rgba(0,0,0,0.2)', border: '1px solid var(--panel-border)',
                                                                borderRadius: '6px', color: 'var(--text-primary)', fontSize: '13px', textAlign: 'center'
                                                            }}
                                                        />
                                                        <span style={{ fontSize: '12px', color: 'var(--text-muted)' }}>p</span>
                                                    </div>
                                                </td>
                                                <td style={{ padding: '12px 14px' }}>
                                                    <span style={{
                                                        display: 'inline-block', padding: '4px 10px', borderRadius: '12px',
                                                        fontSize: '12px', fontWeight: 700,
                                                        background: isSelected ? 'rgba(168, 85, 247, 0.12)' : 'rgba(255,255,255,0.05)', 
                                                        color: isSelected ? '#c084fc' : 'var(--text-muted)'
                                                    }}>
                                                        {sec.slideCount}장
                                                    </span>
                                                </td>
                                                <td style={{ padding: '12px 14px', textAlign: 'center' }}>
                                                    <button
                                                        onClick={() => handleDownloadSingleSection(sec, idx)}
                                                        className="interactive"
                                                        title="이 목차만 PPTX 파일로 다운로드"
                                                        style={{
                                                            padding: '6px 12px', background: 'rgba(59, 130, 246, 0.12)',
                                                            border: '1px solid rgba(59, 130, 246, 0.3)', borderRadius: '6px',
                                                            color: '#60a5fa', fontSize: '12px', fontWeight: 600, cursor: 'pointer',
                                                            display: 'inline-flex', alignItems: 'center', gap: '4px'
                                                        }}
                                                    >
                                                        <Download size={13} /> 다운로드
                                                    </button>
                                                </td>
                                                <td style={{ padding: '12px 14px', textAlign: 'center' }}>
                                                    <button
                                                        onClick={() => handleDeleteSection(idx)}
                                                        className="interactive"
                                                        title="이 섹션 삭제"
                                                        style={{
                                                            padding: '6px', background: 'transparent', border: 'none',
                                                            color: 'var(--text-muted)', cursor: 'pointer', borderRadius: '4px'
                                                        }}
                                                        onMouseEnter={(e) => e.currentTarget.style.color = '#ef4444'}
                                                        onMouseLeave={(e) => e.currentTarget.style.color = 'var(--text-muted)'}
                                                    >
                                                        <Trash2 size={15} />
                                                    </button>
                                                </td>
                                            </tr>
                                        );
                                    })}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
}
