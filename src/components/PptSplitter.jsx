import React, { useState, useRef } from 'react';
import { 
    Split, Upload, FileText, Download, CheckCircle2, AlertCircle, 
    RefreshCw, Layers, Plus, Trash2, Edit3, Save, FileArchive, ArrowRight,
    Sliders, Check, FolderDown
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
            setSections(result.sections);
            setStatusMessage({
                type: 'success',
                text: `총 ${result.totalSlides}장의 슬라이드에서 ${result.sections.length}개의 주요 목차(섹션)를 자동 감지하였습니다.`
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
            previewText: '사용자 추가 섹션'
        };
        setSections([...sections, newSec]);
    };

    // 섹션 삭제
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
            const cleanTitle = (sec.title || `섹션_${sec.id}`).replace(/[\\/:*?"<>|]/g, '_').trim();
            const downloadName = `${fileNamePrefix}_[${String(index + 1).padStart(2, '0')}]_${cleanTitle}_${baseName}.pptx`;
            saveAs(subBlob, downloadName);
        } catch (err) {
            console.error('단일 섹션 분할 다운로드 실패:', err);
            alert(`분할 파일 생성 실패: ${err.message}`);
        }
    };

    // 전체 섹션 일괄 ZIP 다운로드
    const handleDownloadAllZip = async () => {
        if (!fileBuffer || !file || sections.length === 0) return;
        setIsSplitting(true);
        setSplitProgress({ text: '분할 준비 중...', percent: 5 });

        try {
            await downloadAllSectionsAsZip(
                fileBuffer,
                file.name,
                sections,
                (msg, pct) => setSplitProgress({ text: msg, percent: pct })
            );
            setStatusMessage({
                type: 'success',
                text: `성공적으로 ${sections.length}개의 목차별 PPTX 파일을 생성하여 ZIP 압축 다운로드하였습니다.`
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
                        <Split size={26} color="#a855f7" /> PPT 목차 분석 기반 하위 파일 분리 생성
                    </h2>
                    <p style={{ margin: '8px 0 0 0', fontSize: '14px', color: 'var(--text-secondary)', lineHeight: '1.6' }}>
                        대용량 PPTX 파일을 업로드하면 내부 목차(대주제/장/간지)를 AI 및 구조 분석 엔진이 자동 감지하여 <strong>하위 목차별 독립 PPTX 파일들로 무손실 분할 생성</strong>합니다.<br />
                        (슬라이드 마스터, 테마, 특수 폰트, 이미지 및 표 서식이 100% 원본 그대로 완벽 보존됩니다.)
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

            {/* 분석 완료 후 섹션 편집 및 다운로드 대시보드 */}
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
                                        color: 'var(--text-primary)', fontSize: '12.5px', width: '100px'
                                    }}
                                />
                            </div>
                        </div>

                        <div style={{ display: 'flex', gap: '10px' }}>
                            <button
                                onClick={handleAddSection}
                                className="interactive"
                                style={{
                                    display: 'flex', alignItems: 'center', gap: '6px',
                                    padding: '10px 16px', background: 'rgba(255,255,255,0.05)',
                                    border: '1px solid var(--panel-border)', borderRadius: '10px',
                                    color: 'var(--text-primary)', fontSize: '13px', fontWeight: 600, cursor: 'pointer'
                                }}
                            >
                                <Plus size={16} /> 섹션 추가
                            </button>

                            <button
                                onClick={handleDownloadAllZip}
                                disabled={isSplitting}
                                className="interactive"
                                style={{
                                    display: 'flex', alignItems: 'center', gap: '8px',
                                    padding: '10px 22px', background: 'linear-gradient(135deg, #a855f7 0%, #3b82f6 100%)',
                                    border: 'none', borderRadius: '10px',
                                    color: '#ffffff', fontSize: '14px', fontWeight: 700,
                                    cursor: isSplitting ? 'not-allowed' : 'pointer',
                                    boxShadow: '0 4px 14px rgba(168, 85, 247, 0.3)'
                                }}
                            >
                                {isSplitting ? <RefreshCw size={17} className="animate-spin" /> : <FolderDown size={17} />}
                                <span>{isSplitting ? '분할 파일 생성 중...' : `⚡ ${sections.length}개 목차 일괄 분리 (ZIP 다운로드)`}</span>
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

                    {/* 목차별 분할 구성표 */}
                    <div style={{
                        background: 'var(--panel-bg)', border: '1px solid var(--panel-border)',
                        borderRadius: '16px', padding: '24px', display: 'flex', flexDirection: 'column', gap: '16px'
                    }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                <Layers size={18} color="var(--accent-purple)" />
                                <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 700 }}>
                                    분할 대상 하위 목차 목록 ({sections.length}개 파일로 분리)
                                </h3>
                            </div>
                            <span style={{ fontSize: '12px', color: 'var(--text-muted)' }}>
                                💡 목차명이나 슬라이드 범위를 클릭하여 직접 수정할 수 있습니다.
                            </span>
                        </div>

                        <div style={{ overflowX: 'auto' }}>
                            <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '13.5px' }}>
                                <thead>
                                    <tr style={{ borderBottom: '2px solid var(--panel-border)', color: 'var(--text-secondary)', background: 'rgba(255,255,255,0.02)' }}>
                                        <th style={{ padding: '12px 14px', width: '60px' }}>순번</th>
                                        <th style={{ padding: '12px 14px', width: '35%' }}>분할 목차(섹션)명</th>
                                        <th style={{ padding: '12px 14px', width: '15%' }}>시작 슬라이드</th>
                                        <th style={{ padding: '12px 14px', width: '15%' }}>종료 슬라이드</th>
                                        <th style={{ padding: '12px 14px', width: '12%' }}>슬라이드 수</th>
                                        <th style={{ padding: '12px 14px', width: '15%', textAlign: 'center' }}>개별 다운로드</th>
                                        <th style={{ padding: '12px 14px', width: '60px', textAlign: 'center' }}>삭제</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {sections.map((sec, idx) => (
                                        <tr key={idx} style={{ borderBottom: '1px solid rgba(255,255,255,0.05)', backgroundColor: idx % 2 === 0 ? 'rgba(0,0,0,0.1)' : 'transparent' }}>
                                            <td style={{ padding: '12px 14px', fontWeight: 700, color: 'var(--text-secondary)' }}>
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
                                                        borderRadius: '6px', color: 'var(--text-primary)', fontSize: '13.5px', fontWeight: 600
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
                                                    background: 'rgba(168, 85, 247, 0.12)', color: '#c084fc'
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
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
}
