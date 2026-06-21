import React, { useState, useRef, useEffect } from 'react';
import { FileText, Upload, RefreshCw, CheckCircle2, ChevronRight, HelpCircle, AlertCircle, FileSpreadsheet, Eye, Play } from 'lucide-react';
import { generateReportFromTemplate } from '../utils/hwpxGeneratorService.js';
import { FALLBACK_MODELS } from '../utils/geminiModels.js';

function HwpxGenerator({ apiKey }) {
  const [templateFile, setTemplateFile] = useState(null);
  const [materialFiles, setMaterialFiles] = useState([]);
  const [instruction, setInstruction] = useState(() => localStorage.getItem('hwpx_gen_instruction') || '');
  const [isProcessing, setIsProcessing] = useState(false);
  const [progressMsg, setProgressMsg] = useState('');
  const [isGenerated, setIsGenerated] = useState(false);
  const [generatedBlob, setGeneratedBlob] = useState(null);
  const [selectedModel, setSelectedModel] = useState('auto');

  // 드래그 상태 추가
  const [isTemplateDragActive, setIsTemplateDragActive] = useState(false);
  const [isMaterialsDragActive, setIsMaterialsDragActive] = useState(false);

  const templateInputRef = useRef(null);
  const materialsInputRef = useRef(null);

  // 지시사항 로컬스토리지 보존
  useEffect(() => {
    localStorage.setItem('hwpx_gen_instruction', instruction);
  }, [instruction]);

  const handleTemplateChange = (e) => {
    const file = e.target.files[0];
    if (file) {
      if (!file.name.endsWith('.hwpx')) {
        alert('템플릿 양식은 반드시 HWPX (.hwpx) 파일이어야 합니다.');
        return;
      }
      setTemplateFile(file);
      setIsGenerated(false);
    }
  };

  const handleMaterialsChange = (e) => {
    const files = Array.from(e.target.files);
    const validFiles = files.filter(file => {
      const name = file.name.toLowerCase();
      return name.endsWith('.pptx') || name.endsWith('.hwpx') || name.endsWith('.md') || name.endsWith('.txt');
    });

    if (validFiles.length !== files.length) {
      alert('지원되지 않는 포맷의 파일은 제외되었습니다. (PPTX, HWPX, MD, TXT만 가능)');
    }

    setMaterialFiles(prev => {
      const combined = [...prev];
      validFiles.forEach(file => {
        if (!combined.some(c => c.name === file.name && c.size === file.size)) {
          combined.push(file);
        }
      });
      return combined;
    });
    setIsGenerated(false);
  };

  // 템플릿 드래그앤드롭 핸들러
  const handleTemplateDrag = (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setIsTemplateDragActive(true);
    } else if (e.type === "dragleave") {
      setIsTemplateDragActive(false);
    }
  };

  const handleTemplateDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsTemplateDragActive(false);
    const file = e.dataTransfer.files[0];
    if (file) {
      if (!file.name.endsWith('.hwpx')) {
        alert('템플릿 양식은 반드시 HWPX (.hwpx) 파일이어야 합니다.');
        return;
      }
      setTemplateFile(file);
      setIsGenerated(false);
    }
  };

  // 참고자료 드래그앤드롭 핸들러
  const handleMaterialsDrag = (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setIsMaterialsDragActive(true);
    } else if (e.type === "dragleave") {
      setIsMaterialsDragActive(false);
    }
  };

  const handleMaterialsDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsMaterialsDragActive(false);
    const files = Array.from(e.dataTransfer.files);
    const validFiles = files.filter(file => {
      const name = file.name.toLowerCase();
      return name.endsWith('.pptx') || name.endsWith('.hwpx') || name.endsWith('.md') || name.endsWith('.txt');
    });

    if (validFiles.length !== files.length) {
      alert('지원되지 않는 포맷의 파일은 제외되었습니다. (PPTX, HWPX, MD, TXT만 가능)');
    }

    if (validFiles.length > 0) {
      setMaterialFiles(prev => {
        const combined = [...prev];
        validFiles.forEach(file => {
          if (!combined.some(c => c.name === file.name && c.size === file.size)) {
            combined.push(file);
          }
        });
        return combined;
      });
      setIsGenerated(false);
    }
  };

  const removeMaterialFile = (index) => {
    setMaterialFiles(prev => prev.filter((_, idx) => idx !== index));
    setIsGenerated(false);
  };

  const clearAll = () => {
    setTemplateFile(null);
    setMaterialFiles([]);
    setIsGenerated(false);
    setGeneratedBlob(null);
    setProgressMsg('');
  };

  const handleGenerateReport = async () => {
    if (!templateFile) {
      alert('HWPX 샘플 양식 템플릿 파일을 등록해 주세요.');
      return;
    }
    if (materialFiles.length === 0) {
      alert('보고서 작성에 활용할 참고 자료 파일을 최소 1개 이상 추가해 주세요.');
      return;
    }

    setIsProcessing(true);
    setProgressMsg('자료 취합 시작...');
    setIsGenerated(false);

    try {
      const blob = await generateReportFromTemplate(
        templateFile,
        materialFiles,
        apiKey,
        instruction,
        (msg) => setProgressMsg(msg)
      );

      setGeneratedBlob(blob);
      setIsGenerated(true);
    } catch (err) {
      console.error('HWPX 보고서 생성 에러:', err);
      alert(`보고서 생성 실패: ${err.message}`);
    } finally {
      setIsProcessing(false);
    }
  };

  const handleDownload = () => {
    if (!generatedBlob) return;
    const url = URL.createObjectURL(generatedBlob);
    const link = document.createElement('a');
    link.href = url;
    
    // 파일명 포맷팅: [완성보고서]_원본템플릿파일명
    const origName = templateFile.name;
    link.download = `[완성보고서]_${origName}`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  return (
    <div style={{ padding: '24px', display: 'flex', flexDirection: 'column', gap: '24px', minHeight: 'calc(100vh - 120px)', color: 'var(--text-primary)' }}>
      
      {/* 타이틀 배너 */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: 'linear-gradient(135deg, rgba(16, 185, 129, 0.15), rgba(99, 102, 241, 0.05))', border: '1px solid rgba(16, 185, 129, 0.25)', padding: '24px', borderRadius: '16px', backdropFilter: 'blur(10px)' }}>
        <div>
          <h2 style={{ display: 'flex', alignItems: 'center', gap: '10px', fontSize: '22px', fontWeight: 800, margin: 0, color: 'var(--text-primary)' }}>
            <FileText size={28} color="#10b981" /> HWPX 생성 (표준보고서)
          </h2>
          <p style={{ margin: '8px 0 0 0', fontSize: '14px', color: 'var(--text-secondary)', lineHeight: '1.6' }}>
            HWPX 샘플 양식 보고서와 풍부한 참고 자료(PPTX, HWPX, MD 등)를 바탕으로, 스타일 서식을 그대로 유지한 채 보고서를 자동 완성합니다.<br />
            AI가 템플릿의 문맥을 이해하여 적절한 위치에 지능적으로 본문 내용을 요약 및 이식해 줍니다.
          </p>
        </div>
        <div>
          <button 
            onClick={clearAll}
            disabled={!templateFile && materialFiles.length === 0}
            style={{ padding: '10px 16px', background: 'rgba(255,255,255,0.05)', color: (templateFile || materialFiles.length > 0) ? 'var(--text-primary)' : 'var(--text-muted)', border: '1px solid var(--panel-border)', borderRadius: '8px', cursor: (templateFile || materialFiles.length > 0) ? 'pointer' : 'not-allowed', display: 'flex', alignItems: 'center', gap: '8px', fontSize: '13px', fontWeight: 600 }}
          >
            <RefreshCw size={15} /> 초기화
          </button>
        </div>
      </div>

      {/* 메인 2열 그리드 레이아웃 */}
      <div style={{ display: 'grid', gridTemplateColumns: '1.2fr 1fr', gap: '24px' }}>
        
        {/* 좌측: 파일 드롭존 영역 */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
          
          {/* 1. HWPX 샘플 양식 드롭존 */}
          <div style={{ background: 'var(--panel-bg)', border: '1px solid var(--panel-border)', borderRadius: '16px', padding: '24px', display: 'flex', flexDirection: 'column', gap: '14px' }}>
            <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 800, display: 'flex', alignItems: 'center', gap: '8px' }}>
              <span style={{ width: '20px', height: '20px', borderRadius: '50%', background: 'rgba(16, 185, 129, 0.1)', color: '#10b981', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '11px', fontWeight: 800 }}>1</span>
              HWPX 샘플 양식 템플릿 등록
            </h3>
            
            <div 
              onClick={() => templateInputRef.current?.click()}
              onDragEnter={handleTemplateDrag}
              onDragOver={handleTemplateDrag}
              onDragLeave={handleTemplateDrag}
              onDrop={handleTemplateDrop}
              style={{ 
                border: isTemplateDragActive ? '2px dashed #10b981' : '2px dashed var(--panel-border)', 
                background: isTemplateDragActive ? 'rgba(16, 185, 129, 0.08)' : 'transparent',
                borderRadius: '12px', 
                padding: '30px 20px', 
                textAlign: 'center', 
                cursor: 'pointer',
                transition: 'all 0.3s ease',
                display: 'flex',
                flexDirection: 'column',
                alignItems: 'center',
                justifyContent: 'center',
                gap: '10px',
                transform: isTemplateDragActive ? 'scale(1.01)' : 'none'
              }}
              className="interactive-card"
            >
              <Upload size={32} color="#10b981" />
              <span style={{ fontSize: '13.5px', fontWeight: 600, color: 'var(--text-secondary)' }}>
                {templateFile ? templateFile.name : 'HWPX 보고서 양식 파일 (.hwpx) 선택 또는 드롭'}
              </span>
              <span style={{ fontSize: '11.5px', color: 'var(--text-muted)' }}>
                {templateFile ? `크기: ${(templateFile.size / 1024).toFixed(1)} KB` : '기존 폰트/표 스타일이 보존되는 템플릿 기준 양식'}
              </span>
              <input 
                type="file" 
                ref={templateInputRef} 
                onChange={handleTemplateChange} 
                accept=".hwpx" 
                style={{ display: 'none' }} 
              />
            </div>
          </div>

          {/* 2. 참고 자료 파일들 드롭존 */}
          <div style={{ background: 'var(--panel-bg)', border: '1px solid var(--panel-border)', borderRadius: '16px', padding: '24px', display: 'flex', flexDirection: 'column', gap: '14px' }}>
            <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 800, display: 'flex', alignItems: 'center', gap: '8px' }}>
              <span style={{ width: '20px', height: '20px', borderRadius: '50%', background: 'rgba(16, 185, 129, 0.1)', color: '#10b981', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '11px', fontWeight: 800 }}>2</span>
              작성에 반영할 참고 자료 등록 (복수 파일 가능)
            </h3>
            
            <div 
              onClick={() => materialsInputRef.current?.click()}
              onDragEnter={handleMaterialsDrag}
              onDragOver={handleMaterialsDrag}
              onDragLeave={handleMaterialsDrag}
              onDrop={handleMaterialsDrop}
              style={{ 
                border: isMaterialsDragActive ? '2px dashed var(--accent-blue)' : '2px dashed var(--panel-border)', 
                background: isMaterialsDragActive ? 'rgba(99, 102, 241, 0.08)' : 'transparent',
                borderRadius: '12px', 
                padding: '30px 20px', 
                textAlign: 'center', 
                cursor: 'pointer',
                transition: 'all 0.3s ease',
                display: 'flex',
                flexDirection: 'column',
                alignItems: 'center',
                justifyContent: 'center',
                gap: '10px',
                transform: isMaterialsDragActive ? 'scale(1.01)' : 'none'
              }}
              className="interactive-card"
            >
              <Upload size={32} color="var(--accent-blue)" />
              <span style={{ fontSize: '13.5px', fontWeight: 600, color: 'var(--text-secondary)' }}>
                보고서 자료 파일 (.pptx, .hwpx, .md, .txt) 선택 또는 드롭
              </span>
              <span style={{ fontSize: '11.5px', color: 'var(--text-muted)' }}>
                업로드된 자료들의 핵심 명세를 추출하여 보고서 본문 작성을 돕습니다.
              </span>
              <input 
                type="file" 
                ref={materialsInputRef} 
                onChange={handleMaterialsChange} 
                accept=".pptx,.hwpx,.md,.txt" 
                multiple
                style={{ display: 'none' }} 
              />
            </div>

            {/* 업로드된 파일 리스트 */}
            {materialFiles.length > 0 && (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', marginTop: '8px' }}>
                <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--text-muted)' }}>첨부된 자료 파일 ({materialFiles.length}개)</span>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px', maxHeight: '180px', overflowY: 'auto', paddingRight: '4px' }}>
                  {materialFiles.map((file, idx) => (
                    <div key={idx} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: 'rgba(255,255,255,0.02)', border: '1px solid var(--panel-border)', padding: '10px 14px', borderRadius: '8px' }}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '8px', overflow: 'hidden' }}>
                        <FileText size={16} color="var(--accent-blue)" />
                        <span style={{ fontSize: '13px', fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={file.name}>
                          {file.name}
                        </span>
                        <span style={{ fontSize: '11px', color: 'var(--text-muted)', flexShrink: 0 }}>({(file.size/1024).toFixed(1)} KB)</span>
                      </div>
                      <button 
                        onClick={() => removeMaterialFile(idx)}
                        style={{ background: 'transparent', border: 'none', color: '#ef4444', fontSize: '12px', fontWeight: 700, cursor: 'pointer', padding: '2px 6px' }}
                      >
                        삭제
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        </div>

        {/* 우측: AI 설정 및 결과 실행 영역 */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
          
          {/* AI 보고서 생성 지시 패널 */}
          <div style={{ background: 'var(--panel-bg)', border: '1px solid var(--panel-border)', borderRadius: '16px', padding: '24px', display: 'flex', flexDirection: 'column', gap: '16px' }}>
            <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 800 }}>AI 보고서 작성 설정</h3>
            
            {/* 1) AI 모델 선택 */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
              <label style={{ fontSize: '12px', color: 'var(--text-secondary)', fontWeight: 700 }}>AI 분석 엔진 모델</label>
              <select 
                value={selectedModel}
                onChange={(e) => setSelectedModel(e.target.value)}
                style={{ width: '100%', padding: '10px 12px', background: 'var(--bg-dark)', border: '1px solid var(--panel-border)', borderRadius: '8px', color: 'var(--text-primary)', fontSize: '13px', fontWeight: 600, outline: 'none' }}
              >
                <option value="auto">자동 모델 선택 (Gemini 2.5 Pro 우선)</option>
                {FALLBACK_MODELS.map((model, idx) => (
                  <option key={idx} value={model}>{model.split('/').pop()}</option>
                ))}
              </select>
            </div>

            {/* 2) 추가 지시사항 */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
              <label style={{ fontSize: '12px', color: 'var(--text-secondary)', fontWeight: 700 }}>AI 추가 지시사항 (선택)</label>
              <textarea 
                value={instruction}
                onChange={(e) => setInstruction(e.target.value)}
                placeholder="예: 보고서 본문은 반드시 개조식 표현(~함, ~필요)으로 가공해줘. 작성 시 수치 중심의 정량 데이터를 강조해 줘."
                style={{ width: '100%', height: '120px', padding: '12px', background: 'var(--bg-dark)', border: '1px solid var(--panel-border)', borderRadius: '8px', color: 'var(--text-primary)', fontSize: '13px', outline: 'none', resize: 'none', lineHeight: '1.5' }}
              />
            </div>

            {/* 3) 키 확인 및 실행 버튼 */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px', marginTop: '8px' }}>
              
              {/* 진행 프로그레스 노출 */}
              {(isProcessing || progressMsg) && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px', background: 'rgba(255,255,255,0.02)', border: '1px solid var(--panel-border)', padding: '12px', borderRadius: '8px' }}>
                  <span style={{ fontSize: '12px', color: 'var(--text-secondary)', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '6px' }}>
                    <RefreshCw size={14} className="animate-spin" color="#10b981" />
                    진행 상태
                  </span>
                  <span style={{ fontSize: '13px', fontWeight: 600, color: '#10b981' }}>{progressMsg}</span>
                </div>
              )}

              {/* 완료 결과 노출 */}
              {isGenerated && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', background: 'rgba(16, 185, 129, 0.05)', border: '1px solid rgba(16, 185, 129, 0.25)', padding: '16px', borderRadius: '8px', animation: 'fadeIn 0.3s ease' }}>
                  <span style={{ fontSize: '14px', fontWeight: 800, color: '#10b981', display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <CheckCircle2 size={18} /> 보고서 조립 성공!
                  </span>
                  <p style={{ margin: 0, fontSize: '12.5px', color: 'var(--text-secondary)', lineHeight: '1.4' }}>
                    원본 양식의 폰트 크기, 스타일 규격 및 정렬 테마가 완벽 보존된 상태에서 소스 내용이 성공적으로 작성되었습니다.
                  </p>
                  <button 
                    onClick={handleDownload}
                    style={{ width: '100%', padding: '12px', background: '#10b981', color: 'white', border: 'none', borderRadius: '8px', fontSize: '13.5px', fontWeight: 700, cursor: 'pointer', transition: 'opacity 0.2s', marginTop: '6px' }}
                  >
                    완성 보고서 다운로드 (.hwpx)
                  </button>
                </div>
              )}

              {/* 생성 실행 단추 */}
              {!isGenerated && (
                <button
                  onClick={handleGenerateReport}
                  disabled={!templateFile || materialFiles.length === 0 || isProcessing}
                  style={{ 
                    width: '100%', 
                    padding: '14px', 
                    background: (!templateFile || materialFiles.length === 0 || isProcessing) ? 'rgba(255,255,255,0.05)' : 'linear-gradient(135deg, #10b981, #3b82f6)',
                    color: (!templateFile || materialFiles.length === 0 || isProcessing) ? 'var(--text-muted)' : 'white',
                    border: 'none', 
                    borderRadius: '10px', 
                    fontSize: '15px', 
                    fontWeight: 700, 
                    cursor: (!templateFile || materialFiles.length === 0 || isProcessing) ? 'not-allowed' : 'pointer',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    gap: '10px',
                    transition: 'all 0.3s ease'
                  }}
                >
                  <Play size={18} /> 보고서 생성 및 조립 실행
                </button>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

export default HwpxGenerator;
