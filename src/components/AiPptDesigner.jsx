import React, { useState, useRef } from 'react';
import { Presentation, FileText, Upload, Sparkles, Loader2, Download, AlertCircle, LayoutTemplate, X, CheckCircle2 } from 'lucide-react';
import pptxgen from 'pptxgenjs';
import JSZip from 'jszip';
import { analyzePptContent } from '../utils/aiPptAnalyzer';
import { injectSlidesIntoMaster, saveFileWithLocationPicker } from '../utils/pptExporter';

export default function AiPptDesigner({ apiKey }) {
    const [inputText, setInputText] = useState('');
    const [emphasisText, setEmphasisText] = useState('');
    const [inputFile, setInputFile] = useState(null);
    const [inputSlideCount, setInputSlideCount] = useState(0);
    const [extractedThemeColor, setExtractedThemeColor] = useState(null);
    const [status, setStatus] = useState('idle'); // idle, analyzing, generating, success, error
    const [progressMsg, setProgressMsg] = useState('');
    const [errorMsg, setErrorMsg] = useState('');
    const [isDragging, setIsDragging] = useState(false);
    const fileInputRef = useRef(null);

    const handleFileChange = async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        processFile(file);
    };

    const processFile = async (file) => {
        if (!file.name.endsWith('.pptx')) {
            setErrorMsg('지원하지 않는 파일 형식입니다. .pptx 파일만 업로드 가능합니다.');
            return;
        }
        setInputFile(file);
        setErrorMsg('');
        setStatus('idle');
        
        try {
            // Extract text from pptx using JSZip
            const arrayBuffer = await file.arrayBuffer();
            const zip = await JSZip.loadAsync(arrayBuffer);
            let extractedText = '';
            
            const slideFiles = Object.keys(zip.files)
                .filter(name => name.startsWith('ppt/slides/slide') && name.endsWith('.xml'))
                .sort((a, b) => {
                    const numA = parseInt(a.match(/slide(\d+)\.xml/)[1]);
                    const numB = parseInt(b.match(/slide(\d+)\.xml/)[1]);
                    return numA - numB;
                });
                
            setInputSlideCount(slideFiles.length);
            
            for (let i = 0; i < slideFiles.length; i++) {
                const slideFile = slideFiles[i];
                const content = await zip.file(slideFile).async('string');
                
                extractedText += `\n\n--- [원본 슬라이드 ${i + 1}] ---\n`;
                
                // Group text by <a:p> (paragraph) to preserve sentence structure
                const paragraphs = content.match(/<a:p[\s>][\s\S]*?<\/a:p>/g) || content.match(/<a:p>[\s\S]*?<\/a:p>/g);
                if (paragraphs) {
                    paragraphs.forEach(p => {
                        const texts = p.match(/<a:t[^>]*>(.*?)<\/a:t>/g);
                        if (texts) {
                            const pText = texts.map(m => m.replace(/<[^>]+>/g, '')).join('').trim();
                            if (pText) extractedText += `- ${pText}\n`;
                        }
                    });
                } else {
                    const matches = content.match(/<a:t[^>]*>(.*?)<\/a:t>/g);
                    if (matches) {
                        extractedText += matches.map(m => m.replace(/<[^>]+>/g, '')).join(' ') + '\n';
                    }
                }
            }
            setInputText(extractedText.trim());

            // Extract Theme Color
            try {
                const themeFile = Object.keys(zip.files).find(name => name.includes('ppt/theme/theme1.xml'));
                if (themeFile) {
                    const themeXml = await zip.file(themeFile).async('string');
                    const accent1Match = themeXml.match(/<a:accent1>[\s\S]*?<a:srgbClr val="([0-9A-Fa-f]{6})"/);
                    if (accent1Match && accent1Match[1]) {
                        setExtractedThemeColor(accent1Match[1]);
                    } else {
                        const accent2Match = themeXml.match(/<a:accent2>[\s\S]*?<a:srgbClr val="([0-9A-Fa-f]{6})"/);
                        if (accent2Match && accent2Match[1]) setExtractedThemeColor(accent2Match[1]);
                    }
                }
            } catch (e) {
                console.error("Theme extraction failed:", e);
            }
        } catch (err) {
            console.error(err);
            setErrorMsg('PPT 파일에서 텍스트를 추출하는 중 오류가 발생했습니다.');
            setInputFile(null);
        }
    };

    const handleDragEvents = {
        onDragOver: (e) => { e.preventDefault(); e.stopPropagation(); setIsDragging(true); },
        onDragEnter: (e) => { e.preventDefault(); e.stopPropagation(); setIsDragging(true); },
        onDragLeave: (e) => { e.preventDefault(); e.stopPropagation(); setIsDragging(false); },
        onDrop: (e) => {
            e.preventDefault(); e.stopPropagation(); setIsDragging(false);
            const file = e.dataTransfer.files[0];
            if (file) processFile(file);
        }
    };

    const clearInput = () => {
        setInputText('');
        setEmphasisText('');
        setInputFile(null);
        setInputSlideCount(0);
        setExtractedThemeColor(null);
        setErrorMsg('');
        setStatus('idle');
    };

    const generatePpt = async (jsonData) => {
        let pres = new pptxgen();
        pres.theme = { headFontFace: "Malgun Gothic", bodyFontFace: "Malgun Gothic" };
        
        const themeColor = extractedThemeColor || ((jsonData.theme && jsonData.theme.startsWith('#')) ? jsonData.theme.replace('#', '') : "0284c7");
        const subColor = "475569";
        const textColor = "1e293b";
        
        pres.defineSlideMaster({
            title: 'MASTER_SLIDE',
            background: { color: 'f8fafc' },
            objects: [
                { rect: { x: 0, y: 0, w: '100%', h: 0.08, fill: { color: themeColor } } },
                { text: { text: 'Generated by AI PPT Designer', options: { x: 0.5, y: '95%', w: 4, h: 0.3, fontSize: 9, color: '94a3b8' } } },
                { text: { text: 'Powered by Gemini 2.5', options: { x: 5.5, y: '95%', w: 4, h: 0.3, fontSize: 9, color: '94a3b8', align: 'right' } } }
            ]
        });

        const slides = jsonData.slides || [];
        
        slides.forEach(slide => {
            let pptSlide = pres.addSlide({ masterName: 'MASTER_SLIDE' });
            
            if (slide.type === 'TITLE') {
                pptSlide.addText(slide.title || '제목 없음', { x: 1, y: 2.2, w: 8, fontSize: 44, bold: true, color: themeColor, align: 'center' });
                if (slide.subtitle) {
                    pptSlide.addText(slide.subtitle, { x: 1, y: 3.2, w: 8, fontSize: 22, color: subColor, align: 'center' });
                }
                if (slide.author) {
                    pptSlide.addText(slide.author, { x: 1, y: 4.5, w: 8, fontSize: 14, color: '64748b', align: 'center' });
                }
            } else if (slide.type === 'SECTION') {
                pptSlide.background = { color: themeColor };
                pptSlide.addText(slide.title || '섹션', { x: 1, y: 2.5, w: 8, fontSize: 40, bold: true, color: 'ffffff', align: 'center' });
                pptSlide.addShape(pres.ShapeType.line, { x: 4.5, y: 3.2, w: 1.0, h: 0, line: { color: 'ffffff', width: 3 } });
            } else if (slide.type === 'TWO_COLUMN') {
                pptSlide.addText(slide.title || '제목', { x: 0.5, y: 0.5, w: 9, fontSize: 28, bold: true, color: themeColor });
                pptSlide.addShape(pres.ShapeType.line, { x: 0.5, y: 1.1, w: 9.0, h: 0, line: { color: themeColor, width: 2 } });
                
                // Left
                pptSlide.addShape(pres.ShapeType.rect, { x: 0.5, y: 1.4, w: 4.3, h: 0.4, fill: { color: 'f1f5f9' } });
                pptSlide.addText(slide.leftTitle || '항목 1', { x: 0.6, y: 1.4, w: 4.1, h: 0.4, fontSize: 16, bold: true, color: themeColor });
                const leftBullets = (slide.leftBullets || []).map(b => ({ text: b, options: { bullet: true, color: textColor, fontSize: 14, breakLine: true } }));
                if (leftBullets.length > 0) pptSlide.addText(leftBullets, { x: 0.5, y: 2.0, w: 4.3, h: 3, valign: 'top' });
                
                // Right
                pptSlide.addShape(pres.ShapeType.rect, { x: 5.2, y: 1.4, w: 4.3, h: 0.4, fill: { color: 'f1f5f9' } });
                pptSlide.addText(slide.rightTitle || '항목 2', { x: 5.3, y: 1.4, w: 4.1, h: 0.4, fontSize: 16, bold: true, color: themeColor });
                const rightBullets = (slide.rightBullets || []).map(b => ({ text: b, options: { bullet: true, color: textColor, fontSize: 14, breakLine: true } }));
                if (rightBullets.length > 0) pptSlide.addText(rightBullets, { x: 5.2, y: 2.0, w: 4.3, h: 3, valign: 'top' });
            } else if (slide.type === 'ARCHITECTURE_LAYER') {
                pptSlide.addText(slide.title || '아키텍처', { x: 0.5, y: 0.5, w: 9, fontSize: 28, bold: true, color: themeColor });
                pptSlide.addShape(pres.ShapeType.line, { x: 0.5, y: 1.1, w: 9.0, h: 0, line: { color: themeColor, width: 2 } });
                
                const layers = slide.layers || [];
                const startY = 1.5;
                const layerHeight = 3.5 / Math.max(layers.length, 1);
                
                layers.forEach((layer, idx) => {
                    const currentY = startY + (idx * layerHeight);
                    pptSlide.addShape(pres.ShapeType.rect, { 
                        x: 0.5, y: currentY, w: 2.5, h: layerHeight - 0.2, 
                        fill: { color: themeColor },
                        shadow: { type: 'outer', color: '000000', opacity: 0.3, blur: 5, offset: 3, angle: 45 }
                    });
                    pptSlide.addText(layer.name, { x: 0.5, y: currentY, w: 2.5, h: layerHeight - 0.2, align: 'center', color: 'ffffff', fontSize: 16, bold: true });
                    
                    pptSlide.addShape(pres.ShapeType.rect, { 
                        x: 3.2, y: currentY, w: 6.3, h: layerHeight - 0.2, 
                        fill: { color: 'f8fafc' }, line: { color: 'cbd5e1', width: 1 },
                        shadow: { type: 'outer', color: '000000', opacity: 0.1, blur: 4, offset: 2, angle: 45 }
                    });
                    
                    const itemsText = (layer.items || []).join('   |   ');
                    pptSlide.addText(itemsText, { x: 3.4, y: currentY, w: 5.9, h: layerHeight - 0.2, align: 'left', color: textColor, fontSize: 14 });
                });
            } else if (slide.type === 'PROCESS_FLOW') {
                pptSlide.addText(slide.title || '프로세스 흐름', { x: 0.5, y: 0.5, w: 9, fontSize: 28, bold: true, color: themeColor });
                pptSlide.addShape(pres.ShapeType.line, { x: 0.5, y: 1.1, w: 9.0, h: 0, line: { color: themeColor, width: 2 } });
                
                const steps = slide.steps || [];
                const stepWidth = 8.0 / Math.max(steps.length, 1);
                
                steps.forEach((step, idx) => {
                    const currentX = 0.5 + (idx * stepWidth);
                    
                    // 3D Circle
                    pptSlide.addShape(pres.ShapeType.ellipse, { 
                        x: currentX + (stepWidth/2) - 0.6, y: 1.5, w: 1.2, h: 1.2, 
                        fill: { color: themeColor },
                        shadow: { type: 'outer', color: themeColor, opacity: 0.4, blur: 8, offset: 4, angle: 45 }
                    });
                    pptSlide.addText(String(idx + 1), { x: currentX + (stepWidth/2) - 0.6, y: 1.5, w: 1.2, h: 1.2, align: 'center', color: 'ffffff', fontSize: 24, bold: true });
                    
                    pptSlide.addText(step.label, { x: currentX, y: 2.9, w: stepWidth, h: 0.4, align: 'center', valign: 'top', color: themeColor, fontSize: 16, bold: true });
                    pptSlide.addText(step.desc, { x: currentX + 0.1, y: 3.4, w: stepWidth - 0.2, h: 1.8, align: 'center', valign: 'top', color: textColor, fontSize: 11, breakLine: true });
                    
                    // 3D Arrow
                    if (idx < steps.length - 1) {
                        pptSlide.addShape(pres.ShapeType.rightArrow, { 
                            x: currentX + stepWidth - 0.4, y: 1.95, w: 0.8, h: 0.3, 
                            fill: { color: 'cbd5e1' },
                            shadow: { type: 'outer', color: '000000', opacity: 0.15, blur: 3, offset: 2, angle: 45 }
                        });
                    }
                });
            } else if (slide.type === 'KEYWORD_HIGHLIGHT') {
                pptSlide.addText(slide.title || '핵심 키워드', { x: 0.5, y: 0.5, w: 9, fontSize: 28, bold: true, color: themeColor });
                pptSlide.addShape(pres.ShapeType.line, { x: 0.5, y: 1.1, w: 9.0, h: 0, line: { color: themeColor, width: 2 } });
                
                const keywords = slide.keywords || [];
                const cols = Math.min(keywords.length > 0 ? keywords.length : 1, 3);
                const boxW = 8.6 / cols;
                const boxH = 1.2;
                
                keywords.forEach((kw, idx) => {
                    const row = Math.floor(idx / cols);
                    const col = idx % cols;
                    const curX = 0.7 + (col * boxW);
                    const curY = 1.5 + (row * (boxH + 0.4));
                    
                    // 3D Keyword Box (Cube)
                    pptSlide.addShape(pres.ShapeType.cube, { 
                        x: curX, y: curY, w: boxW - 0.4, h: boxH, 
                        fill: { color: 'f8fafc' }, line: { color: themeColor, width: 1.5 },
                        shadow: { type: 'outer', color: themeColor, opacity: 0.35, blur: 8, offset: 5, angle: 45 }
                    });
                    
                    pptSlide.addText(kw.word, { x: curX, y: curY + 0.1, w: boxW - 0.4, h: 0.4, align: 'center', color: themeColor, fontSize: 18, bold: true });
                    pptSlide.addText(kw.desc, { x: curX + 0.1, y: curY + 0.5, w: boxW - 0.6, h: 0.6, align: 'center', color: textColor, fontSize: 12 });
                });
            } else {
                // BULLET (Default)
                pptSlide.addText(slide.title || '제목', { x: 0.5, y: 0.5, w: 9, fontSize: 28, bold: true, color: themeColor });
                pptSlide.addShape(pres.ShapeType.line, { x: 0.5, y: 1.1, w: 9.0, h: 0, line: { color: themeColor, width: 2 } });
                
                const bullets = (slide.bullets || []).map(b => ({ text: b, options: { bullet: true, color: textColor, fontSize: 16, breakLine: true } }));
                if (bullets.length > 0) {
                    pptSlide.addText(bullets, { x: 0.5, y: 1.5, w: 9, h: 3.5, valign: 'top' });
                }
            }
        });

        const aiGenBlob = await pres.write('blob');
        
        if (inputFile) {
            try {
                const mergedBlob = await injectSlidesIntoMaster(inputFile, aiGenBlob);
                await saveFileWithLocationPicker(mergedBlob, `AI_마스터적용_${inputFile.name}`);
                return;
            } catch (err) {
                console.error("마스터 템플릿 병합 실패, 기본 다운로드로 전환:", err);
            }
        }
        
        await saveFileWithLocationPicker(aiGenBlob, "AI_Designed_Presentation.pptx");
    };

    const handleGenerate = async () => {
        if (!inputText.trim()) {
            setErrorMsg('분석할 텍스트를 입력하거나 PPT 파일을 업로드해주세요.');
            return;
        }

        setStatus('analyzing');
        setErrorMsg('');
        
        try {
            const structure = await analyzePptContent(inputText, emphasisText, inputSlideCount, apiKey, setProgressMsg);
            
            setStatus('generating');
            setProgressMsg('슬라이드 레이아웃 및 디자인을 렌더링 중입니다...');
            
            await generatePpt(structure);
            
            setStatus('success');
            setProgressMsg('성공적으로 PPT 파일이 생성되어 다운로드되었습니다.');
        } catch (err) {
            console.error(err);
            setStatus('error');
            setErrorMsg(err.message || 'PPT 생성 중 알 수 없는 오류가 발생했습니다.');
        }
    };

    return (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
            <div className="glass-panel animate-slide-up" style={{ padding: '32px', display: 'flex', flexDirection: 'column', gap: '24px', background: 'linear-gradient(to bottom right, rgba(255,255,255,0.03), rgba(0,0,0,0.1))' }}>
                
                {/* Header */}
                <div style={{ display: 'flex', alignItems: 'center', gap: '16px', borderBottom: '1px solid var(--panel-border)', paddingBottom: '20px' }}>
                    <div style={{ padding: '14px', background: 'linear-gradient(135deg, rgba(236, 72, 153, 0.2), rgba(219, 39, 119, 0.1))', borderRadius: '16px', border: '1px solid rgba(236, 72, 153, 0.3)', boxShadow: '0 8px 32px rgba(236, 72, 153, 0.1)' }}>
                        <LayoutTemplate size={28} color="#ec4899" />
                    </div>
                    <div>
                        <h2 style={{ margin: 0, fontSize: '24px', color: 'var(--text-primary)', letterSpacing: '-0.5px', fontWeight: 800, display: 'flex', alignItems: 'center', gap: '8px' }}>
                            AI PPT 디자이너 <span style={{ padding: '4px 8px', background: 'rgba(236, 72, 153, 0.15)', color: '#ec4899', fontSize: '12px', borderRadius: '12px', fontWeight: 700, letterSpacing: '0.5px' }}>BETA</span>
                        </h2>
                        <p style={{ margin: '6px 0 0', fontSize: '14.5px', color: 'var(--text-secondary)', lineHeight: 1.5 }}>
                            아무 텍스트나 기존 PPT 파일을 넣으세요. AI가 내용을 구조화하고 의미 전달이 명확한 프리미엄 디자인 PPT로 자동 재구성해 드립니다.
                        </p>
                    </div>
                </div>

                <div style={{ display: 'grid', gridTemplateColumns: '1fr', gap: '24px' }}>
                    
                    {/* Input Area */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                            <h3 style={{ margin: 0, fontSize: '16px', color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                                <FileText size={18} color="var(--accent-blue)" /> 내용 입력 또는 파일 업로드
                            </h3>
                            {inputFile && (
                                <button onClick={clearInput} style={{ background: 'none', border: 'none', color: 'var(--text-muted)', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '4px', fontSize: '13px' }}>
                                    <X size={14} /> 초기화
                                </button>
                            )}
                        </div>
                        
                        <div 
                            {...handleDragEvents}
                            style={{
                                position: 'relative',
                                background: isDragging ? 'rgba(59, 130, 246, 0.05)' : 'rgba(0,0,0,0.15)',
                                border: `2px ${isDragging ? 'dashed var(--accent-blue)' : 'solid var(--panel-border)'}`,
                                borderRadius: '16px',
                                overflow: 'hidden',
                                transition: 'all 0.3s ease',
                                boxShadow: 'inset 0 2px 10px rgba(0,0,0,0.1)'
                            }}
                        >
                            {!inputFile ? (
                                <div style={{ position: 'absolute', top: 0, left: 0, right: 0, bottom: 0, pointerEvents: 'none', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', opacity: inputText ? 0 : 1, transition: 'opacity 0.2s', gap: '12px' }}>
                                    <div style={{ padding: '16px', background: 'rgba(255,255,255,0.03)', borderRadius: '50%', border: '1px dashed var(--glass-border)' }}>
                                        <Upload size={24} color="var(--text-muted)" />
                                    </div>
                                    <p style={{ margin: 0, color: 'var(--text-muted)', fontSize: '14px', fontWeight: 500 }}>여기에 텍스트를 붙여넣거나 PPTX 파일을 드래그하세요</p>
                                    <button 
                                        onClick={(e) => { e.preventDefault(); fileInputRef.current?.click(); }}
                                        style={{ pointerEvents: 'auto', padding: '8px 16px', background: 'var(--bg-secondary)', border: '1px solid var(--glass-border)', borderRadius: '8px', color: 'var(--text-primary)', cursor: 'pointer', fontSize: '13px', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '6px' }}
                                    >
                                        <Presentation size={14} /> 찾아보기
                                    </button>
                                </div>
                            ) : null}

                            <textarea
                                value={inputText}
                                onChange={(e) => setInputText(e.target.value)}
                                placeholder=""
                                style={{
                                    width: '100%',
                                    minHeight: '280px',
                                    background: 'transparent',
                                    border: 'none',
                                    padding: '24px',
                                    color: 'var(--text-primary)',
                                    fontSize: '15px',
                                    lineHeight: 1.6,
                                    resize: 'vertical',
                                    outline: 'none',
                                    boxSizing: 'border-box'
                                }}
                            />
                            
                            <input type="file" ref={fileInputRef} onChange={handleFileChange} accept=".pptx" style={{ display: 'none' }} />
                            
                            {inputFile && (
                                <div style={{ position: 'absolute', bottom: '16px', right: '16px', background: 'rgba(16, 185, 129, 0.15)', border: '1px solid rgba(16, 185, 129, 0.3)', padding: '8px 16px', borderRadius: '12px', display: 'flex', alignItems: 'center', gap: '8px', backdropFilter: 'blur(8px)' }}>
                                    <CheckCircle2 size={16} color="var(--success-color)" />
                                    <span style={{ fontSize: '13px', color: 'var(--success-color)', fontWeight: 600 }}>{inputFile.name} 로드 완료 ({inputSlideCount}장)</span>
                                    {extractedThemeColor && (
                                        <div style={{ display: 'flex', alignItems: 'center', gap: '6px', marginLeft: '12px', paddingLeft: '12px', borderLeft: '1px solid rgba(16, 185, 129, 0.3)' }}>
                                            <div style={{ width: '14px', height: '14px', borderRadius: '50%', background: `#${extractedThemeColor}`, border: '1px solid rgba(255,255,255,0.5)' }} />
                                            <span style={{ fontSize: '12px', color: 'var(--success-color)' }}>양식 테마 적용됨</span>
                                        </div>
                                    )}
                                </div>
                            )}
                        </div>
                        
                        {/* Emphasis Area */}
                        <div style={{ marginTop: '8px' }}>
                            <h3 style={{ margin: '0 0 12px 0', fontSize: '15px', color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                                <Sparkles size={16} color="#ec4899" /> 강조하고 싶은 내용 요약 (선택사항)
                            </h3>
                            <textarea
                                value={emphasisText}
                                onChange={(e) => setEmphasisText(e.target.value)}
                                placeholder="예: '서버리스 아키텍처 도입으로 인한 비용 절감 효과를 가장 크게 강조해줘', '기술 스택에 Docker, Kubernetes를 포함해서 눈에 띄게 배치해줘'"
                                style={{
                                    width: '100%',
                                    minHeight: '80px',
                                    background: 'rgba(0,0,0,0.15)',
                                    border: '1px solid var(--panel-border)',
                                    borderRadius: '12px',
                                    padding: '16px',
                                    color: 'var(--text-primary)',
                                    fontSize: '14px',
                                    lineHeight: 1.5,
                                    resize: 'vertical',
                                    outline: 'none',
                                    boxSizing: 'border-box'
                                }}
                            />
                        </div>
                    </div>

                </div>

                {/* Status Messages */}
                {errorMsg && (
                    <div className="animate-fade-in" style={{ padding: '16px', background: 'rgba(239, 68, 68, 0.1)', color: '#fca5a5', borderRadius: '12px', border: '1px solid rgba(239, 68, 68, 0.2)', fontSize: '14px', display: 'flex', alignItems: 'center', gap: '10px' }}>
                        <AlertCircle size={18} /> {errorMsg}
                    </div>
                )}
                {status === 'success' && (
                    <div className="animate-fade-in" style={{ padding: '16px', background: 'rgba(16, 185, 129, 0.1)', color: '#6ee7b7', borderRadius: '12px', border: '1px solid rgba(16, 185, 129, 0.2)', fontSize: '14px', display: 'flex', alignItems: 'center', gap: '10px' }}>
                        <Download size={18} /> {progressMsg}
                    </div>
                )}

                {/* Generate Button */}
                <div style={{ marginTop: '10px' }}>
                    <button
                        onClick={handleGenerate}
                        disabled={status === 'analyzing' || status === 'generating' || !inputText.trim()}
                        style={{
                            width: '100%',
                            padding: '20px',
                            background: (!inputText.trim() || status === 'analyzing' || status === 'generating') ? 'rgba(255,255,255,0.05)' : 'linear-gradient(135deg, #db2777, #ec4899, #f43f5e)',
                            color: (!inputText.trim() || status === 'analyzing' || status === 'generating') ? 'var(--text-muted)' : 'white',
                            border: 'none',
                            borderRadius: '16px',
                            fontSize: '18px',
                            fontWeight: 800,
                            letterSpacing: '0.5px',
                            cursor: (!inputText.trim() || status === 'analyzing' || status === 'generating') ? 'not-allowed' : 'pointer',
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                            gap: '12px',
                            boxShadow: (!inputText.trim() || status === 'analyzing' || status === 'generating') ? 'none' : '0 10px 25px rgba(236, 72, 153, 0.3)',
                            transition: 'all 0.3s ease'
                        }}
                    >
                        {(status === 'analyzing' || status === 'generating') ? (
                            <><Loader2 size={24} className="animate-spin" /> {progressMsg}</>
                        ) : (
                            <><Sparkles size={24} /> AI 디자인 생성 시작</>
                        )}
                    </button>
                    <p style={{ textAlign: 'center', marginTop: '12px', fontSize: '13px', color: 'var(--text-muted)' }}>
                        AI가 텍스트의 맥락을 분석하여 적절한 구조와 모던한 템플릿으로 슬라이드를 자동 생성합니다.
                    </p>
                </div>

            </div>
        </div>
    );
}
