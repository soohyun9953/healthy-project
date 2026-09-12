import React, { useState, useRef } from 'react';
import { Presentation, Box, Upload, X, Settings, Info, CheckCircle2, Loader2, Sparkles } from 'lucide-react';
import { fusePptToHwpxTemplate, fusePptToHwpxListTemplate } from '../utils/hwpxConverter';
import { saveFileWithLocationPicker } from '../utils/pptExporter';

// PPT ➜ HWPX 표준 양식 융합 탭. PptGenerator.jsx의 pptx_hwpx_fusion 서브탭에서 분리됨.
export default function PptHwpxFusionTab({ setErrorMsg, setSuccessMsg }) {
    const [hwpxFusionPptFile, setHwpxFusionPptFile] = useState(null);
    const [hwpxFusionTemplateFile, setHwpxFusionTemplateFile] = useState(null);
    const [hwpxFusionListTemplateFile, setHwpxFusionListTemplateFile] = useState(null); // 💡 목록표 양식 파일 추가!
    const [isFusingHwpx, setIsFusingHwpx] = useState(false);
    const [isDraggingFusionPpt, setIsDraggingFusionPpt] = useState(false);
    const [isDraggingFusionTemplate, setIsDraggingFusionTemplate] = useState(false);
    const [isDraggingFusionListTemplate, setIsDraggingFusionListTemplate] = useState(false); // 💡 목록표 드래깅 state 추가!
    const [fusionMergeMode, setFusionMergeMode] = useState(true); // true: 하나의 HWPX 병합, false: ZIP 압축 분할
    const [fusionResultStats, setFusionResultStats] = useState(null);
    const fusionPptInputRef = useRef(null);
    const fusionTemplateInputRef = useRef(null);
    const fusionListTemplateInputRef = useRef(null); // 💡 목록표 Ref 추가!

    const handleHwpxFusion = async () => {
        if (!hwpxFusionPptFile || !hwpxFusionTemplateFile) {
            setErrorMsg('PPT 요건기술서 파일과 HWPX 양식 템플릿 파일을 모두 업로드해주세요.');
            return;
        }
        setIsFusingHwpx(true);
        setErrorMsg(null);
        setSuccessMsg(null);
        setFusionResultStats(null);
        try {
            // 1. 상세명세서(표준양식) HWPX 융합 수행
            const resultBlob = await fusePptToHwpxTemplate(hwpxFusionPptFile, hwpxFusionTemplateFile, fusionMergeMode);

            const defaultFileName = fusionMergeMode
                ? `융합_${hwpxFusionPptFile.name.replace(/\.[^/.]+$/, "")}.hwpx`
                : `융합개별_${hwpxFusionPptFile.name.replace(/\.[^/.]+$/, "")}.zip`;

            const saved = await saveFileWithLocationPicker(resultBlob, defaultFileName);

            // 2. 목록표 양식이 업로드되었으면 목록표 HWPX 융합 추가 수행
            let listSaved = false;
            if (hwpxFusionListTemplateFile) {
                try {
                    const listBlob = await fusePptToHwpxListTemplate(hwpxFusionPptFile, hwpxFusionListTemplateFile);
                    const listFileName = `목록표_${hwpxFusionPptFile.name.replace(/\.[^/.]+$/, "")}.hwpx`;
                    listSaved = await saveFileWithLocationPicker(listBlob, listFileName);
                } catch (listErr) {
                    console.error("목록표 HWPX 융합 실패:", listErr);
                    setErrorMsg(`상세명세서는 완료되었으나, 목록표 HWPX 융합에는 실패했습니다: ${listErr.message}`);
                }
            }

            if (saved) {
                let successTxt = fusionMergeMode
                    ? `성공적으로 HWPX 표준 양식 융합 및 다운로드를 마쳤습니다! (1개 HWPX 파일 병합)`
                    : `성공적으로 HWPX 양식 분할 융합 및 ZIP 파일 다운로드를 마쳤습니다! (개별 고유번호별 분할)`;

                if (hwpxFusionListTemplateFile && listSaved) {
                    successTxt += ` + [목록표 HWPX 파일 추가 생성 완료]`;
                }

                setSuccessMsg(successTxt);
                setFusionResultStats({
                    fusedCount: resultBlob.fusedCount,
                    requirementsList: resultBlob.requirementsList,
                    mode: resultBlob.mode
                });
            }
        } catch (err) {
            console.error("HWPX 융합 실패:", err);
            setErrorMsg(`HWPX 양식 융합에 실패했습니다: ${err.message || 'PPT/HWPX 내부 데이터 융합 오류'}`);
        } finally {
            setIsFusingHwpx(false);
        }
    };

    return (
        <div className="animate-fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
            {/* 1. 삼중 파일 업로드 드롭존 (PPTX + HWPX 상세 + HWPX 목록표) */}
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '20px' }}>
                {/* PPTX 업로드 */}
                <div style={{ background: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--panel-border)', display: 'flex', flexDirection: 'column', gap: '16px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <div style={{ padding: '6px', background: 'rgba(59, 130, 246, 0.1)', borderRadius: '6px' }}>
                            <Presentation size={16} color="var(--accent-color)" />
                        </div>
                        <h3 style={{ margin: 0, fontSize: '15px', color: 'var(--text-primary)', fontWeight: 700 }}>1. PPT 요건기술서 파일 (.pptx)</h3>
                    </div>
                    <div
                        onDragOver={(e) => { e.preventDefault(); setIsDraggingFusionPpt(true); }}
                        onDragLeave={() => setIsDraggingFusionPpt(false)}
                        onDrop={(e) => {
                            e.preventDefault();
                            setIsDraggingFusionPpt(false);
                            const file = e.dataTransfer.files[0];
                            if (file && file.name.toLowerCase().endsWith('.pptx')) {
                                setHwpxFusionPptFile(file);
                                setFusionResultStats(null);
                                setErrorMsg(null);
                                setSuccessMsg(null);
                            } else {
                                setErrorMsg('파워포인트 파일(.pptx)만 업로드할 수 있습니다.');
                            }
                        }}
                        onClick={() => fusionPptInputRef.current?.click()}
                        style={{
                            border: `2px dashed ${isDraggingFusionPpt ? 'var(--accent-color)' : 'var(--panel-border)'}`,
                            borderRadius: '10px',
                            padding: '36px 20px',
                            textAlign: 'center',
                            cursor: 'pointer',
                            background: isDraggingFusionPpt ? 'rgba(59, 130, 246, 0.05)' : 'rgba(0,0,0,0.15)',
                            transition: 'all 0.2s',
                            display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '12px'
                        }}
                    >
                        <input
                            type="file"
                            ref={fusionPptInputRef}
                            onChange={(e) => {
                                const file = e.target.files[0];
                                if (file) {
                                    setHwpxFusionPptFile(file);
                                    setFusionResultStats(null);
                                    setErrorMsg(null);
                                    setSuccessMsg(null);
                                }
                            }}
                            accept=".pptx"
                            style={{ display: 'none' }}
                        />
                        <div style={{ padding: '12px', background: 'rgba(255,255,255,0.03)', borderRadius: '50%' }}>
                            <Upload size={30} color={hwpxFusionPptFile ? 'var(--accent-color)' : 'var(--text-muted)'} />
                        </div>
                        <div>
                            <p style={{ margin: 0, fontSize: '13.5px', fontWeight: 600, color: 'var(--text-primary)' }}>
                                {hwpxFusionPptFile ? hwpxFusionPptFile.name : 'PPTX 요건기술서 드래그 또는 클릭'}
                            </p>
                            <p style={{ margin: '4px 0 0', fontSize: '11.5px', color: 'var(--text-muted)' }}>
                                {hwpxFusionPptFile ? `${(hwpxFusionPptFile.size / 1024 / 1024).toFixed(2)} MB` : '변환 데이터 소스가 될 PPTX 파일 등록'}
                            </p>
                        </div>
                        {hwpxFusionPptFile && (
                            <button
                                onClick={(e) => {
                                    e.stopPropagation();
                                    setHwpxFusionPptFile(null);
                                    setFusionResultStats(null);
                                    setErrorMsg(null);
                                    setSuccessMsg(null);
                                }}
                                className="interactive"
                                style={{
                                    padding: '4px 8px', background: 'rgba(239,68,68,0.1)', color: 'var(--danger-color)',
                                    border: '1px solid rgba(239,68,68,0.2)', borderRadius: '4px', fontSize: '11.5px',
                                    cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '4px'
                                }}
                            >
                                <X size={12} /> 파일 제거
                            </button>
                        )}
                    </div>
                </div>

                {/* HWPX 업로드 */}
                <div style={{ background: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--panel-border)', display: 'flex', flexDirection: 'column', gap: '16px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <div style={{ padding: '6px', background: 'rgba(16, 185, 129, 0.1)', borderRadius: '6px' }}>
                            <Box size={16} color="var(--success-color)" />
                        </div>
                        <h3 style={{ margin: 0, fontSize: '15px', color: 'var(--text-primary)', fontWeight: 700 }}>2. HWPX 표준 양식 템플릿 (.hwpx)</h3>
                    </div>
                    <div
                        onDragOver={(e) => { e.preventDefault(); setIsDraggingFusionTemplate(true); }}
                        onDragLeave={() => setIsDraggingFusionTemplate(false)}
                        onDrop={(e) => {
                            e.preventDefault();
                            setIsDraggingFusionTemplate(false);
                            const file = e.dataTransfer.files[0];
                            if (file && file.name.toLowerCase().endsWith('.hwpx')) {
                                setHwpxFusionTemplateFile(file);
                                setFusionResultStats(null);
                                setErrorMsg(null);
                                setSuccessMsg(null);
                            } else {
                                setErrorMsg('한글 표준 문서(.hwpx)만 업로드할 수 있습니다.');
                            }
                        }}
                        onClick={() => fusionTemplateInputRef.current?.click()}
                        style={{
                            border: `2px dashed ${isDraggingFusionTemplate ? 'var(--success-color)' : 'var(--panel-border)'}`,
                            borderRadius: '10px',
                            padding: '36px 20px',
                            textAlign: 'center',
                            cursor: 'pointer',
                            background: isDraggingFusionTemplate ? 'rgba(16, 185, 129, 0.05)' : 'rgba(0,0,0,0.15)',
                            transition: 'all 0.2s',
                            display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '12px'
                        }}
                    >
                        <input
                            type="file"
                            ref={fusionTemplateInputRef}
                            onChange={(e) => {
                                const file = e.target.files[0];
                                if (file) {
                                    setHwpxFusionTemplateFile(file);
                                    setFusionResultStats(null);
                                    setErrorMsg(null);
                                    setSuccessMsg(null);
                                }
                            }}
                            accept=".hwpx"
                            style={{ display: 'none' }}
                        />
                        <div style={{ padding: '12px', background: 'rgba(255,255,255,0.03)', borderRadius: '50%' }}>
                            <Upload size={30} color={hwpxFusionTemplateFile ? 'var(--success-color)' : 'var(--text-muted)'} />
                        </div>
                        <div>
                            <p style={{ margin: 0, fontSize: '13.5px', fontWeight: 600, color: 'var(--text-primary)' }}>
                                {hwpxFusionTemplateFile ? hwpxFusionTemplateFile.name : 'HWPX 템플릿 드래그 또는 클릭'}
                            </p>
                            <p style={{ margin: '4px 0 0', fontSize: '11.5px', color: 'var(--text-muted)' }}>
                                {hwpxFusionTemplateFile ? `${(hwpxFusionTemplateFile.size / 1024).toFixed(2)} KB` : '데이터를 채워넣을 HWPX 양식 등록'}
                            </p>
                        </div>
                        {hwpxFusionTemplateFile && (
                            <button
                                onClick={(e) => {
                                    e.stopPropagation();
                                    setHwpxFusionTemplateFile(null);
                                    setFusionResultStats(null);
                                    setErrorMsg(null);
                                    setSuccessMsg(null);
                                }}
                                className="interactive"
                                style={{
                                    padding: '4px 8px', background: 'rgba(239,68,68,0.1)', color: 'var(--danger-color)',
                                    border: '1px solid rgba(239,68,68,0.2)', borderRadius: '4px', fontSize: '11.5px',
                                    cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '4px'
                                }}
                            >
                                <X size={12} /> 파일 제거
                            </button>
                        )}
                    </div>
                </div>

                {/* HWPX 목록표 양식 업로드 (선택) */}
                <div style={{ background: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--panel-border)', display: 'flex', flexDirection: 'column', gap: '16px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <div style={{ padding: '6px', background: 'rgba(249, 115, 22, 0.1)', borderRadius: '6px' }}>
                            <Box size={16} color="#f97316" />
                        </div>
                        <h3 style={{ margin: 0, fontSize: '15px', color: 'var(--text-primary)', fontWeight: 700 }}>3. HWPX 목록표 양식 (.hwpx) (선택)</h3>
                    </div>
                    <div
                        onDragOver={(e) => { e.preventDefault(); setIsDraggingFusionListTemplate(true); }}
                        onDragLeave={() => setIsDraggingFusionListTemplate(false)}
                        onDrop={(e) => {
                            e.preventDefault();
                            setIsDraggingFusionListTemplate(false);
                            const file = e.dataTransfer.files[0];
                            if (file && file.name.toLowerCase().endsWith('.hwpx')) {
                                setHwpxFusionListTemplateFile(file);
                                setFusionResultStats(null);
                                setErrorMsg(null);
                                setSuccessMsg(null);
                            } else {
                                setErrorMsg('한글 목록표 문서(.hwpx)만 업로드할 수 있습니다.');
                            }
                        }}
                        onClick={() => fusionListTemplateInputRef.current?.click()}
                        style={{
                            border: `2px dashed ${isDraggingFusionListTemplate ? '#f97316' : 'var(--panel-border)'}`,
                            borderRadius: '10px',
                            padding: '36px 20px',
                            textAlign: 'center',
                            cursor: 'pointer',
                            background: isDraggingFusionListTemplate ? 'rgba(249, 115, 22, 0.05)' : 'rgba(0,0,0,0.15)',
                            transition: 'all 0.2s',
                            display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '12px'
                        }}
                    >
                        <input
                            type="file"
                            ref={fusionListTemplateInputRef}
                            onChange={(e) => {
                                const file = e.target.files[0];
                                if (file) {
                                    setHwpxFusionListTemplateFile(file);
                                    setFusionResultStats(null);
                                    setErrorMsg(null);
                                    setSuccessMsg(null);
                                }
                            }}
                            accept=".hwpx"
                            style={{ display: 'none' }}
                        />
                        <div style={{ padding: '12px', background: 'rgba(255,255,255,0.03)', borderRadius: '50%' }}>
                            <Upload size={30} color={hwpxFusionListTemplateFile ? '#f97316' : 'var(--text-muted)'} />
                        </div>
                        <div>
                            <p style={{ margin: 0, fontSize: '13.5px', fontWeight: 600, color: 'var(--text-primary)' }}>
                                {hwpxFusionListTemplateFile ? hwpxFusionListTemplateFile.name : 'HWPX 목록표 드래그 또는 클릭'}
                            </p>
                            <p style={{ margin: '4px 0 0', fontSize: '11.5px', color: 'var(--text-muted)' }}>
                                {hwpxFusionListTemplateFile ? `${(hwpxFusionListTemplateFile.size / 1024).toFixed(2)} KB` : '데이터를 채워넣을 목록표 양식 등록 (선택)'}
                            </p>
                        </div>
                        {hwpxFusionListTemplateFile && (
                            <button
                                onClick={(e) => {
                                    e.stopPropagation();
                                    setHwpxFusionListTemplateFile(null);
                                    setFusionResultStats(null);
                                    setErrorMsg(null);
                                    setSuccessMsg(null);
                                }}
                                className="interactive"
                                style={{
                                    padding: '4px 8px', background: 'rgba(239,68,68,0.1)', color: 'var(--danger-color)',
                                    border: '1px solid rgba(239,68,68,0.2)', borderRadius: '4px', fontSize: '11.5px',
                                    cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '4px'
                                }}
                            >
                                <X size={12} /> 파일 제거
                            </button>
                        )}
                    </div>
                </div>
            </div>

            {/* 2. 생성 제어 방식 설정 패널 */}
            <div style={{ background: 'var(--bg-secondary)', padding: '20px', borderRadius: '12px', border: '1px solid var(--panel-border)', display: 'flex', flexDirection: 'column', gap: '12px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <Settings size={16} color="var(--accent-color)" />
                    <span style={{ fontSize: '14.5px', fontWeight: 700, color: 'var(--text-primary)' }}>HWPX 표준산출물 빌드 옵션</span>
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '12px', paddingLeft: '4px' }}>
                    <label style={{ display: 'flex', alignItems: 'center', gap: '10px', cursor: 'pointer' }}>
                        <input
                            type="radio"
                            name="fusion_merge_mode"
                            checked={fusionMergeMode === true}
                            onChange={() => setFusionMergeMode(true)}
                            style={{ accentColor: 'var(--accent-color)', width: '16px', height: '16px' }}
                        />
                        <div>
                            <span style={{ fontSize: '13.5px', fontWeight: 600, color: 'var(--text-primary)' }}>하나의 완결된 HWPX 문서로 통합 병합 (권장)</span>
                            <p style={{ margin: '2px 0 0', fontSize: '12px', color: 'var(--text-muted)' }}>각 요구사항 건수를 한글 표로 각각 복제 및 쪽 나누기(Page Break)를 가동하여 1개의 완성도 높은 한글 파일로 출력합니다.</p>
                        </div>
                    </label>
                    <div style={{ height: '1px', background: 'var(--panel-border)', margin: '4px 0' }} />
                    <label style={{ display: 'flex', alignItems: 'center', gap: '10px', cursor: 'pointer' }}>
                        <input
                            type="radio"
                            name="fusion_merge_mode"
                            checked={fusionMergeMode === false}
                            onChange={() => setFusionMergeMode(false)}
                            style={{ accentColor: 'var(--accent-color)', width: '16px', height: '16px' }}
                        />
                        <div>
                            <span style={{ fontSize: '13.5px', fontWeight: 600, color: 'var(--text-primary)' }}>요구사항 고유번호별 개별 HWPX 파일들로 저장 (ZIP 압축)</span>
                            <p style={{ margin: '2px 0 0', fontSize: '12px', color: 'var(--text-muted)' }}>고유번호별로 양식이 주입된 다수의 독립된 HWPX 파일을 생성하고 하나의 ZIP 파일로 압축하여 다운로드합니다.</p>
                        </div>
                    </label>
                </div>
            </div>

            {/* 3. 데이터 맵핑 및 스마트 안내 규칙 */}
            <div style={{ background: 'rgba(250,204,21,0.03)', padding: '16px', borderRadius: '8px', border: '1px solid rgba(250,204,21,0.15)', display: 'flex', flexDirection: 'column', gap: '8px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', color: 'var(--accent-gold)' }}>
                    <Info size={16} />
                    <span style={{ fontSize: '13.5px', fontWeight: 700 }}>HWPX 표준산출물 매핑 정보 및 개행 가이드</span>
                </div>
                <ul style={{ margin: 0, paddingLeft: '20px', fontSize: '12.5px', color: 'var(--text-secondary)', lineHeight: '1.7', display: 'flex', flexDirection: 'column', gap: '4px' }}>
                    <li>**매핑 정보**: PPT 표의 **고유번호 ➜ 요구사항 고유번호**, **요구사항 명칭 ➜ 요구사항 명칭**, **정의 ➜ 정의**, **요건 세부내용 ➜ 세부내용**으로 지능 자동 매핑됩니다.</li>
                    <li>**가독성 보정**: 긴 개행 텍스트의 경우 HWPX 셀 내부에 여러 개의 단락(&lt;hp:p&gt;)을 동적 생성하여 한글 오피스 뷰어 상의 완전한 가폭 및 줄바꿈을 완벽히 보장합니다.</li>
                    <li>**스마트 필터**: 첫 번째 열이 `ECR-XXX`처럼 고유번호 패턴을 가지고 있는 4개 열 구조의 표 행만 지능적으로 골라내어 노이즈(헤더나 인트로)를 사전 격리합니다.</li>
                </ul>
            </div>

            {/* 4. 융합 결과 통계 피드백 리포트 */}
            {fusionResultStats && (
                <div className="animate-scale-in" style={{ background: 'rgba(59,130,246,0.03)', padding: '20px', borderRadius: '12px', border: '1px solid rgba(59,130,246,0.2)', display: 'flex', flexDirection: 'column', gap: '12px' }}>
                    <h4 style={{ margin: 0, fontSize: '14.5px', color: 'var(--accent-color)', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '6px' }}>
                        <CheckCircle2 size={16} /> HWPX 양식 융합 완료 통계 리포트
                    </h4>
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
                        <div style={{ background: 'rgba(0,0,0,0.2)', padding: '12px', borderRadius: '8px', textAlign: 'center', border: '1px solid var(--panel-border)' }}>
                            <div style={{ fontSize: '11.5px', color: 'var(--text-muted)' }}>융합된 요구사항 건수</div>
                            <div style={{ fontSize: '20px', fontWeight: 800, color: 'var(--accent-color)', marginTop: '4px' }}>{fusionResultStats.fusedCount}개 건</div>
                        </div>
                        <div style={{ background: 'rgba(0,0,0,0.2)', padding: '12px', borderRadius: '8px', textAlign: 'center', border: '1px solid var(--panel-border)' }}>
                            <div style={{ fontSize: '11.5px', color: 'var(--text-muted)' }}>최종 저장 방식</div>
                            <div style={{ fontSize: '20px', fontWeight: 800, color: 'var(--success-color)', marginTop: '4px' }}>
                                {fusionResultStats.mode === 'merge' ? '통합 HWPX 문서 병합' : '개별 HWPX 분할 ZIP 압축'}
                            </div>
                        </div>
                    </div>
                    <div style={{ marginTop: '8px' }}>
                        <div style={{ fontSize: '12.5px', color: 'var(--text-secondary)', fontWeight: 600, marginBottom: '6px' }}>융합 요구사항 고유번호 목록:</div>
                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px', maxHeight: '100px', overflowY: 'auto', padding: '6px', background: 'rgba(0,0,0,0.3)', borderRadius: '6px' }}>
                            {fusionResultStats.requirementsList.map((req, idx) => (
                                <span key={idx} style={{ padding: '2px 8px', background: 'rgba(59, 130, 246, 0.1)', color: 'var(--accent-color)', borderRadius: '4px', fontSize: '11px', fontWeight: 600 }}>
                                    {req.id}
                                </span>
                            ))}
                        </div>
                    </div>
                </div>
            )}

            {/* 5. 융합 실행 버튼 */}
            <div style={{ marginTop: '10px' }}>
                <button
                    className="interactive"
                    onClick={handleHwpxFusion}
                    disabled={!hwpxFusionPptFile || !hwpxFusionTemplateFile || isFusingHwpx}
                    style={{
                        width: '100%',
                        padding: '18px',
                        background: (!hwpxFusionPptFile || !hwpxFusionTemplateFile || isFusingHwpx) ? 'rgba(255,255,255,0.05)' : 'linear-gradient(135deg, #3b82f6, #6366f1)',
                        color: (!hwpxFusionPptFile || !hwpxFusionTemplateFile || isFusingHwpx) ? 'var(--text-muted)' : 'white',
                        border: 'none',
                        borderRadius: '12px',
                        fontSize: '17px',
                        fontWeight: 800,
                        cursor: (!hwpxFusionPptFile || !hwpxFusionTemplateFile || isFusingHwpx) ? 'not-allowed' : 'pointer',
                        display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '12px',
                        boxShadow: (!hwpxFusionPptFile || !hwpxFusionTemplateFile || isFusingHwpx) ? 'none' : '0 8px 20px rgba(59, 130, 246, 0.2)'
                    }}
                >
                    {isFusingHwpx ? (
                        <><Loader2 size={22} className="animate-spin" /> 양식 딥 복제 및 데이터 융합 중...</>
                    ) : (
                        <><Sparkles size={22} /> HWPX 표준 양식 융합 및 산출물 다운로드</>
                    )}
                </button>
            </div>
        </div>
    );
}
