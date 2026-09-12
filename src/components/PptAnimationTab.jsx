import React, { useState, useRef } from 'react';
import {
    Presentation, Upload, Sparkles, MousePointer2, CheckCircle2,
    Box, ToggleRight, ToggleLeft, Layers, Loader2
} from 'lucide-react';
import { addSmartAnimationsToPpt, saveFileWithLocationPicker, getPptSlideCount } from '../utils/pptExporter';

// PPT 스마트 애니메이션 탭. PptGenerator.jsx의 smart_animation 서브탭에서 분리됨.
export default function PptAnimationTab({ setErrorMsg, setSuccessMsg }) {
    const [animationPptFile, setAnimationPptFile] = useState(null);
    const [isProcessingAnimation, setIsProcessingAnimation] = useState(false);
    const [isDraggingAnimation, setIsDraggingAnimation] = useState(false);
    const [animationType, setAnimationType] = useState('transition'); // Default to transition
    const [useGrouping, setUseGrouping] = useState(true);
    const [slideAnimations, setSlideAnimations] = useState([]); // [{ enabled: true, type: 'transition', useGrouping: true }, ...]
    const animInputRef = useRef(null);

    const processAnimationFile = async (file) => {
        if (!file) return;
        if (!file.name.toLowerCase().endsWith('.pptx')) {
            setErrorMsg('PPTX 파일만 지원합니다.');
            return;
        }
        setErrorMsg(null);
        setSuccessMsg(null);
        setAnimationPptFile(file);

        // 슬라이드 개수 분석
        try {
            const count = await getPptSlideCount(file);
            setSlideAnimations(Array(count).fill(null).map(() => ({ enabled: true, type: 'transition', useGrouping: true })));
        } catch (err) {
            console.error('슬라이드 분석 오류:', err);
            setErrorMsg('PPT 파일을 분석하는 중 오류가 발생했습니다.');
        }
    };

    const handleAnimationProcess = async () => {
        if (!animationPptFile) {
            setErrorMsg('애니메이션을 적용할 PPT 파일을 등록해주세요.');
            return;
        }
        setErrorMsg(null);
        setSuccessMsg(null);
        setIsProcessingAnimation(true);
        try {
            const modifiedBlob = await addSmartAnimationsToPpt(animationPptFile, {
                animationType,
                useGrouping,
                perSlideConfigs: slideAnimations
            });
            await saveFileWithLocationPicker(modifiedBlob, `애니메이션_${animationPptFile.name}`);
            setSuccessMsg('성공적으로 스마트 애니메이션이 적용되어 저장되었습니다.');
            setAnimationPptFile(null);
            setSlideAnimations([]);
        } catch (err) {
            console.error(err);
            setErrorMsg(`애니메이션 처리 중 오류 발생: ${err.message || '알 수 없는 오류'}`);
        } finally {
            setIsProcessingAnimation(false);
        }
    };

    const updateAllSlideAnimations = (enabled, type, grouping) => {
        setSlideAnimations(prev => prev.map(anim => ({
            enabled: enabled !== null ? enabled : anim.enabled,
            type: type !== null ? type : anim.type,
            useGrouping: grouping !== null ? grouping : anim.useGrouping
        })));
        if (type) setAnimationType(type);
        if (grouping !== null) setUseGrouping(grouping);
    };

    const updateSingleSlideAnimation = (index, updates) => {
        setSlideAnimations(prev => {
            const next = [...prev];
            next[index] = { ...next[index], ...updates };
            return next;
        });
    };

    return (
        <div className="animate-fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
            {/* 스마트 애니메이션 업로드 영역 */}
            <div
                onDragOver={(e) => { e.preventDefault(); e.stopPropagation(); }}
                onDragEnter={(e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingAnimation(true); }}
                onDragLeave={(e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingAnimation(false); }}
                onDrop={(e) => {
                    e.preventDefault(); e.stopPropagation(); setIsDraggingAnimation(false);
                    processAnimationFile(e.dataTransfer.files[0]);
                }}
                style={{
                    border: `2px dashed ${isDraggingAnimation ? '#a855f7' : 'rgba(168, 85, 247, 0.3)'}`,
                    borderRadius: '12px', padding: '32px',
                    background: isDraggingAnimation ? 'rgba(168, 85, 247, 0.1)' : 'rgba(168, 85, 247, 0.05)',
                    display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', gap: '16px',
                    transition: 'all 0.2s ease', cursor: isDraggingAnimation ? 'copy' : 'default',
                    minHeight: '200px'
                }}
            >
                <div style={{ padding: '16px', background: 'rgba(168, 85, 247, 0.1)', borderRadius: '50%' }}>
                    <Presentation size={32} color="#a855f7" />
                </div>
                <div style={{ textAlign: 'center' }}>
                    <h3 style={{ margin: '0 0 8px 0', fontSize: '18px', color: 'var(--text-primary)' }}>원본 PPT 파일 등록</h3>
                    <p style={{ margin: 0, fontSize: '14px', color: 'var(--text-secondary)' }}>애니메이션을 추가할 .pptx 파일을 드래그하거나 선택하세요.</p>
                </div>

                <input
                    type="file"
                    accept=".pptx"
                    onChange={(e) => processAnimationFile(e.target.files[0])}
                    ref={animInputRef}
                    style={{ display: 'none' }}
                />

                <div style={{ display: 'flex', gap: '12px', alignItems: 'center' }}>
                    <button
                        onClick={() => animInputRef.current?.click()}
                        className="interactive"
                        style={{
                            display: 'flex', alignItems: 'center', gap: '8px',
                            padding: '10px 20px', borderRadius: '8px', cursor: 'pointer',
                            background: 'var(--bg-secondary)', border: '1px solid var(--panel-border)',
                            color: 'var(--text-primary)', fontWeight: 600, fontSize: '14px'
                        }}
                    >
                        <Upload size={18} /> PPT 파일 찾기
                    </button>
                    <button
                        className={`interactive ${animationType === 'transition' ? 'active' : ''}`}
                        onClick={() => updateAllSlideAnimations(null, 'transition', null)}
                        style={{
                            flex: 1, padding: '12px', borderRadius: '8px',
                            background: animationType === 'transition' ? 'rgba(168, 85, 247, 0.2)' : 'rgba(255,255,255,0.05)',
                            color: animationType === 'transition' ? '#c084fc' : 'var(--text-secondary)',
                            cursor: 'pointer', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '8px',
                            border: `1px solid ${animationType === 'transition' ? '#a855f7' : 'transparent'}`,
                            transition: 'all 0.2s'
                        }}
                    >
                        <Presentation size={24} />
                        <div style={{ fontSize: '13px', fontWeight: 600 }}>슬라이드 전환</div>
                    </button>
                    <button
                        className={`interactive ${animationType === 'fade' ? 'active' : ''}`}
                        onClick={() => updateAllSlideAnimations(null, 'fade', null)}
                        style={{
                            flex: 1, padding: '12px', borderRadius: '8px',
                            background: animationType === 'fade' ? 'rgba(168, 85, 247, 0.2)' : 'rgba(255,255,255,0.05)',
                            color: animationType === 'fade' ? '#c084fc' : 'var(--text-secondary)',
                            cursor: 'pointer', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '8px',
                            border: `1px solid ${animationType === 'fade' ? '#a855f7' : 'transparent'}`,
                            transition: 'all 0.2s'
                        }}
                    >
                        <Sparkles size={24} />
                        <div style={{ fontSize: '13px', fontWeight: 600 }}>객체 페이드</div>
                    </button>
                    <button
                        className={`interactive ${animationType === 'appear' ? 'active' : ''}`}
                        onClick={() => updateAllSlideAnimations(null, 'appear', null)}
                        style={{
                            flex: 1, padding: '12px', borderRadius: '8px',
                            background: animationType === 'appear' ? 'rgba(168, 85, 247, 0.2)' : 'rgba(255,255,255,0.05)',
                            color: animationType === 'appear' ? '#c084fc' : 'var(--text-secondary)',
                            cursor: 'pointer', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '8px',
                            border: `1px solid ${animationType === 'appear' ? '#a855f7' : 'transparent'}`,
                            transition: 'all 0.2s'
                        }}
                    >
                        <MousePointer2 size={24} />
                        <div style={{ fontSize: '13px', fontWeight: 600 }}>객체 나타나기</div>
                    </button>
                    {animationPptFile && (
                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', background: 'rgba(168, 85, 247, 0.1)', padding: '8px 16px', borderRadius: '8px', fontSize: '14px', border: '1px solid #a855f7', color: '#c084fc' }}>
                            <CheckCircle2 size={16} />
                            {animationPptFile.name}
                        </div>
                    )}
                </div>
            </div>

            {/* 옵션 설정 */}
            <div style={{
                border: '1px solid var(--panel-border)',
                borderRadius: '12px', padding: '24px',
                background: 'rgba(255, 255, 255, 0.02)',
                display: 'flex', flexDirection: 'column', gap: '20px'
            }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <Sparkles size={20} color="#a855f7" />
                    <h3 style={{ margin: 0, fontSize: '16px', color: 'var(--text-primary)' }}>애니메이션 스타일 설정</h3>
                </div>

                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
                    <button
                        onClick={() => setAnimationType('fade')}
                        style={{
                            padding: '16px', borderRadius: '12px', cursor: 'pointer',
                            background: animationType === 'fade' ? 'rgba(168, 85, 247, 0.15)' : 'rgba(0,0,0,0.1)',
                            border: `1px solid ${animationType === 'fade' ? '#a855f7' : 'var(--panel-border)'}`,
                            color: animationType === 'fade' ? '#c084fc' : 'var(--text-secondary)',
                            display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '8px', transition: 'all 0.2s'
                        }}
                    >
                        <div style={{ fontWeight: 700, fontSize: '15px' }}>페이드 (Fade)</div>
                        <div style={{ fontSize: '12px', opacity: 0.8 }}>객체가 부드럽게 나타납니다. (권장)</div>
                    </button>
                    <button
                        onClick={() => setAnimationType('appear')}
                        style={{
                            padding: '16px', borderRadius: '12px', cursor: 'pointer',
                            background: animationType === 'appear' ? 'rgba(168, 85, 247, 0.15)' : 'rgba(0,0,0,0.1)',
                            border: `1px solid ${animationType === 'appear' ? '#a855f7' : 'var(--panel-border)'}`,
                            color: animationType === 'appear' ? '#c084fc' : 'var(--text-secondary)',
                            display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '8px', transition: 'all 0.2s'
                        }}
                    >
                        <div style={{ fontWeight: 700, fontSize: '15px' }}>나타나기 (Appear)</div>
                        <div style={{ fontSize: '12px', opacity: 0.8 }}>객체가 즉시 나타납니다.</div>
                    </button>
                </div>

                {animationType !== 'transition' && (
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '12px', background: 'rgba(255,255,255,0.05)', borderRadius: '8px' }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                            <Box size={18} color="#a855f7" />
                            <div>
                                <div style={{ fontSize: '14px', fontWeight: 600, color: 'var(--text-primary)' }}>객체 그룹화 (한번에 나타나기)</div>
                                <div style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>슬라이드의 모든 객체가 클릭 한 번에 동시에 나타납니다.</div>
                            </div>
                        </div>
                        <button
                            onClick={() => updateAllSlideAnimations(null, null, !useGrouping)}
                            style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 0, color: useGrouping ? '#a855f7' : 'var(--text-muted)' }}
                        >
                            {useGrouping ? <ToggleRight size={32} /> : <ToggleLeft size={32} />}
                        </button>
                    </div>
                )}

                <div style={{
                    padding: '16px', background: 'rgba(59, 130, 246, 0.05)',
                    borderLeft: '4px solid var(--accent-blue)', borderRadius: '4px',
                    fontSize: '13px', color: 'var(--text-secondary)', lineHeight: 1.6
                }}>
                    {animationType === 'transition' ? (
                        <>💡 <strong>전환 효과:</strong> 객체를 수정하지 않고 슬라이드가 넘어갈 때 부드러운 <strong>페이드 인</strong> 효과를 적용합니다. 가장 안전하고 깔끔한 방식입니다.</>
                    ) : (
                        <>💡 <strong>작동 원리:</strong> <strong>'그룹화'</strong>를 켜면 슬라이드 전체가 한 번에 나타나며, 끄면 객체들이 <strong>상단 → 하단</strong> 순서대로 하나씩 나타납니다.</>
                    )}
                </div>
            </div>

            {/* 페이지별 애니메이션 설정 (파일 업로드 후 노출) */}
            {animationPptFile && slideAnimations.length > 0 && (
                <div className="animate-fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                            <Layers size={18} color="#a855f7" />
                            <h3 style={{ margin: 0, fontSize: '15px', color: 'var(--text-primary)' }}>페이지별 개별 설정 ({slideAnimations.length} 슬라이드)</h3>
                        </div>
                        <div style={{ display: 'flex', gap: '8px' }}>
                            <button
                                onClick={() => updateAllSlideAnimations(true, null)}
                                style={{ padding: '4px 8px', fontSize: '12px', background: 'rgba(168, 85, 247, 0.1)', border: '1px solid #a855f7', color: '#c084fc', borderRadius: '4px', cursor: 'pointer' }}
                            >전체 켜기</button>
                            <button
                                onClick={() => updateAllSlideAnimations(false, null)}
                                style={{ padding: '4px 8px', fontSize: '12px', background: 'rgba(255,255,255,0.05)', border: '1px solid var(--panel-border)', color: 'var(--text-secondary)', borderRadius: '4px', cursor: 'pointer' }}
                            >전체 끄기</button>
                        </div>
                    </div>

                    <div style={{
                        display: 'grid',
                        gridTemplateColumns: 'repeat(auto-fill, minmax(200px, 1fr))',
                        gap: '12px',
                        maxHeight: '400px',
                        overflowY: 'auto',
                        padding: '4px'
                    }}>
                        {slideAnimations.map((anim, idx) => (
                            <div key={idx} style={{
                                background: anim.enabled ? 'rgba(168, 85, 247, 0.05)' : 'rgba(0,0,0,0.1)',
                                border: `1px solid ${anim.enabled ? 'rgba(168, 85, 247, 0.3)' : 'var(--panel-border)'}`,
                                borderRadius: '10px',
                                padding: '12px',
                                display: 'flex',
                                flexDirection: 'column',
                                gap: '10px',
                                transition: 'all 0.2s'
                            }}>
                                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                                    <span style={{ fontSize: '13px', fontWeight: 700, color: anim.enabled ? '#c084fc' : 'var(--text-muted)' }}>Slide {idx + 1}</span>
                                    <button
                                        onClick={() => updateSingleSlideAnimation(idx, { enabled: !anim.enabled })}
                                        style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 0, color: anim.enabled ? '#a855f7' : 'var(--text-muted)' }}
                                    >
                                        {anim.enabled ? <ToggleRight size={24} /> : <ToggleLeft size={24} />}
                                    </button>
                                </div>

                                {anim.enabled && (
                                    <select
                                        value={anim.type}
                                        onChange={(e) => updateSingleSlideAnimation(idx, { type: e.target.value })}
                                        style={{
                                            background: 'rgba(0,0,0,0.2)',
                                            color: 'white',
                                            border: '1px solid var(--panel-border)',
                                            borderRadius: '4px',
                                            fontSize: '12px',
                                            padding: '4px'
                                        }}
                                    >
                                        <option value="transition">슬라이드 전환 (Transition)</option>
                                        <option value="fade">객체 페이드 (Object Fade)</option>
                                        <option value="appear">객체 나타나기 (Object Appear)</option>
                                    </select>
                                )}

                                {anim.enabled && anim.type !== 'transition' && (
                                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginTop: '4px' }}>
                                        <span style={{ fontSize: '11px', color: 'var(--text-secondary)' }}>그룹화</span>
                                        <button
                                            onClick={() => updateSingleSlideAnimation(idx, { useGrouping: !anim.useGrouping })}
                                            style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 0, color: anim.useGrouping ? '#a855f7' : 'var(--text-muted)' }}
                                        >
                                            {anim.useGrouping ? <ToggleRight size={20} /> : <ToggleLeft size={20} />}
                                        </button>
                                    </div>
                                )}
                            </div>
                        ))}
                    </div>
                </div>
            )}

            <div style={{ marginTop: 'auto', paddingTop: '10px' }}>
                <button
                    className="interactive"
                    onClick={handleAnimationProcess}
                    disabled={!animationPptFile || isProcessingAnimation}
                    style={{
                        width: '100%',
                        padding: '18px',
                        background: (!animationPptFile || isProcessingAnimation) ? 'rgba(255,255,255,0.05)' : 'linear-gradient(135deg, #a855f7, #6366f1)',
                        color: (!animationPptFile || isProcessingAnimation) ? 'var(--text-muted)' : 'white',
                        border: 'none',
                        borderRadius: '12px',
                        fontSize: '17px',
                        fontWeight: 800,
                        cursor: (!animationPptFile || isProcessingAnimation) ? 'not-allowed' : 'pointer',
                        display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '12px',
                        boxShadow: (!animationPptFile || isProcessingAnimation) ? 'none' : '0 8px 20px rgba(168, 85, 247, 0.2)'
                    }}
                >
                    {isProcessingAnimation ? (
                        <><Loader2 size={22} className="animate-spin" /> 전 슬라이드 객체 분석 및 애니메이션 주입 중...</>
                    ) : (
                        <><Sparkles size={22} /> 스마트 애니메이션 적용 및 저장</>
                    )}
                </button>
            </div>
        </div>
    );
}
