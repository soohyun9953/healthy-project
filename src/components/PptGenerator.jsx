import React, { useState } from 'react';
import { Presentation, X, Info, FileDown } from 'lucide-react';
import HwpxConverter from './HwpxConverter.jsx';
import MdToDocxConverter from './MdToDocxConverter.jsx';
import PptSplitter from './PptSplitter.jsx';
import PptPdfConvertTab from './PptPdfConvertTab.jsx';
import PptAnimationTab from './PptAnimationTab.jsx';
import PptHwpxFusionTab from './PptHwpxFusionTab.jsx';
import PptExcelMappingTab from './PptExcelMappingTab.jsx';
import PptBatchEditTab from './PptBatchEditTab.jsx';

export default function PptGenerator({ apiKey, llmProvider = 'gemini', omniRouteModel = 'auto' }) {
    const [activeTab, setActiveTab] = useState('batch_edit'); // 'batch_edit' as default
    const [errorMsg, setErrorMsg] = useState(null);
    const [successMsg, setSuccessMsg] = useState(null);

    return (
        <div style={{ display: 'flex', flexDirection: 'column' }}>
            <div className="glass-panel animate-slide-up" style={{ padding: '24px', display: 'flex', flexDirection: 'column', gap: '24px' }}>
                
                {/* 헤더 */}
                <div style={{ display: 'flex', alignItems: 'center', gap: '12px', borderBottom: '1px solid var(--panel-border)', paddingBottom: '16px' }}>
                    <div style={{ padding: '10px', background: 'rgba(59, 130, 246, 0.1)', borderRadius: '10px' }}>
                        <Presentation size={24} color="var(--accent-blue)" />
                    </div>
                    <div>
                        <h2 style={{ margin: 0, fontSize: '20px', color: 'var(--text-primary)', letterSpacing: '-0.5px' }}>
                            스마트 편집기
                        </h2>
                        <p style={{ margin: '4px 0 0', fontSize: '13.5px', color: 'var(--text-secondary)' }}>
                            데이터 매핑을 통한 PPT 생성 또는 텍스트 서식 일괄 변경 기능을 제공합니다.
                        </p>
                    </div>
                </div>

                {/* 탭 메뉴 */}
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px', borderBottom: '1px solid var(--panel-border)', paddingBottom: '12px' }}>
                    <button 
                        id="tab-batch-edit"
                        onClick={() => { setActiveTab('batch_edit'); setErrorMsg(null); setSuccessMsg(null); }}
                        className="interactive"
                        style={{
                            padding: '10px 20px', borderRadius: '8px', cursor: 'pointer',
                            background: activeTab === 'batch_edit' ? 'rgba(236, 72, 153, 0.1)' : 'transparent',
                            border: '1px solid ' + (activeTab === 'batch_edit' ? '#ec4899' : 'transparent'),
                            color: activeTab === 'batch_edit' ? '#f472b6' : 'var(--text-secondary)',
                            fontWeight: 600, fontSize: '14px', transition: 'all 0.2s'
                        }}
                    >
                        PPT 일괄편집
                    </button>
                    <button 
                        id="tab-excel-mapping"
                        onClick={() => { setActiveTab('excel_mapping'); setErrorMsg(null); setSuccessMsg(null); }}
                        className="interactive"
                        style={{
                            padding: '10px 20px', borderRadius: '8px', cursor: 'pointer',
                            background: activeTab === 'excel_mapping' ? 'rgba(59, 130, 246, 0.1)' : 'transparent',
                            border: '1px solid ' + (activeTab === 'excel_mapping' ? 'var(--accent-blue)' : 'transparent'),
                            color: activeTab === 'excel_mapping' ? 'var(--accent-blue)' : 'var(--text-secondary)',
                            fontWeight: 600, fontSize: '14px', transition: 'all 0.2s'
                        }}
                    >
                        엑셀 ➜ PPT
                    </button>
                    <button 
                        id="tab-smart-animation"
                        onClick={() => { setActiveTab('smart_animation'); setErrorMsg(null); setSuccessMsg(null); }}
                        className="interactive"
                        style={{
                            padding: '10px 20px', borderRadius: '8px', cursor: 'pointer',
                            background: activeTab === 'smart_animation' ? 'rgba(168, 85, 247, 0.1)' : 'transparent',
                            border: '1px solid ' + (activeTab === 'smart_animation' ? '#a855f7' : 'transparent'),
                            color: activeTab === 'smart_animation' ? '#c084fc' : 'var(--text-secondary)',
                            fontWeight: 600, fontSize: '14px', transition: 'all 0.2s'
                        }}
                    >
                        PPT 스마트 애니메이션
                    </button>

                    <button 
                        id="tab-pptx-hwpx-fusion"
                        onClick={() => { setActiveTab('pptx_hwpx_fusion'); setErrorMsg(null); setSuccessMsg(null); }}
                        className="interactive"
                        style={{
                            padding: '10px 20px', borderRadius: '8px', cursor: 'pointer',
                            background: activeTab === 'pptx_hwpx_fusion' ? 'rgba(59, 130, 246, 0.1)' : 'transparent',
                            border: '1px solid ' + (activeTab === 'pptx_hwpx_fusion' ? 'var(--accent-color)' : 'transparent'),
                            color: activeTab === 'pptx_hwpx_fusion' ? 'var(--accent-color)' : 'var(--text-secondary)',
                            fontWeight: 600, fontSize: '14px', transition: 'all 0.2s'
                        }}
                    >
                        PPT ➜ HWPX 양식 융합
                    </button>
                    <button 
                        id="tab-ppt-split"
                        onClick={() => { setActiveTab('ppt_split'); setErrorMsg(null); setSuccessMsg(null); }}
                        className="interactive"
                        style={{
                            padding: '10px 20px', borderRadius: '8px', cursor: 'pointer',
                            background: activeTab === 'ppt_split' ? 'rgba(168, 85, 247, 0.15)' : 'transparent',
                            border: '1px solid ' + (activeTab === 'ppt_split' ? '#a855f7' : 'transparent'),
                            color: activeTab === 'ppt_split' ? '#c084fc' : 'var(--text-secondary)',
                            fontWeight: 600, fontSize: '14px', transition: 'all 0.2s'
                        }}
                    >
                        PPT 목차별 분할
                    </button>
                    <button 
                        id="tab-hwp-to-hwpx"
                        onClick={() => { setActiveTab('hwp_to_hwpx'); setErrorMsg(null); setSuccessMsg(null); }}
                        className="interactive"
                        style={{
                            padding: '10px 20px', borderRadius: '8px', cursor: 'pointer',
                            background: activeTab === 'hwp_to_hwpx' ? 'rgba(16, 185, 129, 0.15)' : 'transparent',
                            border: '1px solid ' + (activeTab === 'hwp_to_hwpx' ? '#10b981' : 'transparent'),
                            color: activeTab === 'hwp_to_hwpx' ? '#10b981' : 'var(--text-secondary)',
                            fontWeight: 600, fontSize: '14px', transition: 'all 0.2s'
                        }}
                    >
                        HWP ➔ HWPX 변환
                    </button>
                    <button 
                        id="tab-md-to-docx"
                        onClick={() => { setActiveTab('md_to_docx'); setErrorMsg(null); setSuccessMsg(null); }}
                        className="interactive"
                        style={{
                            padding: '10px 20px', borderRadius: '8px', cursor: 'pointer',
                            background: activeTab === 'md_to_docx' ? 'rgba(37, 99, 235, 0.15)' : 'transparent',
                            border: '1px solid ' + (activeTab === 'md_to_docx' ? '#2563eb' : 'transparent'),
                            color: activeTab === 'md_to_docx' ? '#3b82f6' : 'var(--text-secondary)',
                            fontWeight: 600, fontSize: '14px', transition: 'all 0.2s'
                        }}
                    >
                        MD ➔ Word 변환
                    </button>
                    <button 
                        id="tab-pdf-convert"
                        onClick={() => { setActiveTab('pdf_convert'); setErrorMsg(null); setSuccessMsg(null); }}
                        className="interactive"
                        style={{
                            padding: '10px 20px', borderRadius: '8px', cursor: 'pointer',
                            background: activeTab === 'pdf_convert' ? 'rgba(16, 185, 129, 0.1)' : 'transparent',
                            border: '1px solid ' + (activeTab === 'pdf_convert' ? 'var(--success-color)' : 'transparent'),
                            color: activeTab === 'pdf_convert' ? 'var(--success-color)' : 'var(--text-secondary)',
                            fontWeight: 600, fontSize: '14px', transition: 'all 0.2s'
                        }}
                    >
                        PPT ➜ PDF 변환
                    </button>
                    <button 
                        id="tab-pdf-to-ppt"
                        onClick={() => { setActiveTab('pdf_to_ppt'); setErrorMsg(null); setSuccessMsg(null); }}
                        className="interactive"
                        style={{
                            padding: '10px 20px', borderRadius: '8px', cursor: 'pointer',
                            background: activeTab === 'pdf_to_ppt' ? 'rgba(251, 146, 60, 0.1)' : 'transparent',
                            border: '1px solid ' + (activeTab === 'pdf_to_ppt' ? '#f97316' : 'transparent'),
                            color: activeTab === 'pdf_to_ppt' ? '#fb923c' : 'var(--text-secondary)',
                            fontWeight: 600, fontSize: '14px', transition: 'all 0.2s'
                        }}
                    >
                        PDF ➞ PPT 변환
                    </button>
                </div>

                {/* 에러/성공 메시지 공통 표시 */}
                {errorMsg && (
                    <div className="animate-fade-in" style={{ padding: '12px', background: 'rgba(239, 68, 68, 0.1)', color: 'var(--danger-color)', borderRadius: '8px', border: '1px solid rgba(239, 68, 68, 0.2)', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <X size={16} /> {errorMsg}
                    </div>
                )}
                {successMsg && (
                    <div className="animate-fade-in" style={{ padding: '12px', background: 'rgba(16, 185, 129, 0.1)', color: 'var(--success-color)', borderRadius: '8px', border: '1px solid rgba(16, 185, 129, 0.2)', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <FileDown size={16} /> {successMsg}
                    </div>
                )}

                {/* 탭 컨텐츠 */}
                {activeTab === 'excel_mapping' ? (
                    <PptExcelMappingTab setErrorMsg={setErrorMsg} setSuccessMsg={setSuccessMsg} />
                ) : activeTab === 'batch_edit' ? (
                    <PptBatchEditTab setErrorMsg={setErrorMsg} setSuccessMsg={setSuccessMsg} />
                ) : activeTab === 'smart_animation' ? (
                    <PptAnimationTab setErrorMsg={setErrorMsg} setSuccessMsg={setSuccessMsg} />
                ) : activeTab === 'pptx_hwpx_fusion' ? (
                    <PptHwpxFusionTab setErrorMsg={setErrorMsg} setSuccessMsg={setSuccessMsg} />
                ) : activeTab === 'pdf_convert' ? (
                    <PptPdfConvertTab setErrorMsg={setErrorMsg} setSuccessMsg={setSuccessMsg} />
                ) : activeTab === 'ppt_split' ? (
                    <div className="animate-fade-in">
                        <PptSplitter />
                    </div>
                ) : activeTab === 'hwp_to_hwpx' ? (
                    <div className="animate-fade-in">
                        <HwpxConverter apiKey={apiKey} llmProvider={llmProvider} omniRouteModel={omniRouteModel} />
                    </div>
                ) : activeTab === 'md_to_docx' ? (
                    <div className="animate-fade-in">
                        <MdToDocxConverter />
                    </div>
                ) : activeTab === 'pdf_to_ppt' ? (
                    <div className="animate-fade-in" style={{
                        padding: '60px 20px',
                        background: 'rgba(251, 146, 60, 0.03)',
                        border: '1px solid rgba(251, 146, 60, 0.15)',
                        borderRadius: '16px',
                        display: 'flex',
                        flexDirection: 'column',
                        alignItems: 'center',
                        justifyContent: 'center',
                        gap: '16px',
                        textAlign: 'center'
                    }}>
                        <Info size={40} color="#f97316" />
                        <h3 style={{
                            margin: 0,
                            fontSize: '22px',
                            fontWeight: 700,
                            color: '#fb923c',
                            letterSpacing: '-0.5px'
                        }}>
                            알PDF를 사용하세요
                        </h3>
                        <p style={{
                            margin: 0,
                            fontSize: '14px',
                            color: 'var(--text-secondary)',
                            maxWidth: '480px',
                            lineHeight: '1.6'
                        }}>
                            PDF 문서를 편집 가능한 파워포인트(PPT) 파일로 변환하시려면 <strong>알PDF(ALPDF)</strong> 프로그램의 PDF ➔ PPT 변환 기능을 이용해 주세요.
                        </p>
                    </div>
                ) : null}
            </div>
        </div>
    );
}


