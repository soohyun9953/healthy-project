import React, { useState, useCallback, useRef, useEffect } from 'react';
import { ArrowRight, Loader2, PenTool, RotateCcw, History, Trash2, X, BookPlus, Plus, Layers, Upload, CheckCircle2, XCircle, Sparkles } from 'lucide-react';
import InputSection from './InputSection';
import ResultDashboard from './ResultDashboard';
import { analyzeDocumentsWithLLM, apply_typos_to_text } from '../llmAnalyzer';
import { extract_dictionary_typos, extract_dict_pairs_from_glossary } from '../utils/typoDictionary';
import { proofreadHistoryDB } from '../utils/proofreadHistoryDB';
import { getCustomDictionary, addCustomTerm, deleteCustomTerm, clearCustomDictionary } from '../utils/customTypoDictionary';
import { processFile, ALL_ACCEPT } from '../utils/fileExtractor';
import { detect_style_inconsistency } from '../utils/styleConsistency';

let batchIdCounter = 0;

function TypoValidator({ apiKey, llmProvider = 'gemini', omniRouteModel = 'auto' }) {
  const [proofreadMode, setProofreadMode] = useState('typo'); // 'typo' | 'ai_humanize'
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisStage, setAnalysisStage] = useState(0); // 1: 추출, 2: 심층분석
  const [retryStatus, setRetryStatus] = useState(null); // API 재시도 상태 메시지
  const [resultData, setResultData] = useState(null);
  const [historyList, setHistoryList] = useState([]);
  const [showHistory, setShowHistory] = useState(false);
  const [showDictionary, setShowDictionary] = useState(false);
  const [showBatch, setShowBatch] = useState(false);
  const [batchMode, setBatchMode] = useState('typo'); // 일괄 처리 적용 모드
  const [customDict, setCustomDict] = useState(() => getCustomDictionary());
  const [newTermWrong, setNewTermWrong] = useState('');
  const [newTermCorrect, setNewTermCorrect] = useState('');
  const [newTermDesc, setNewTermDesc] = useState('');
  const [dictError, setDictError] = useState('');
  const [batchFiles, setBatchFiles] = useState([]); // { id, file, name, status, result, errorMsg, progress }
  const [batchRunning, setBatchRunning] = useState(false);
  const batchFileInputRef = useRef(null);
  const lastParams = useRef(null);

  const handleAddCustomTerm = (e) => {
    e.preventDefault();
    try {
      const updated = addCustomTerm(newTermWrong, newTermCorrect, newTermDesc);
      setCustomDict(updated);
      setNewTermWrong('');
      setNewTermCorrect('');
      setNewTermDesc('');
      setDictError('');
    } catch (err) {
      setDictError(err.message);
    }
  };

  const handleDeleteCustomTerm = (wrong) => {
    setCustomDict(deleteCustomTerm(wrong));
  };

  const handleClearCustomDict = () => {
    if (window.confirm('등록된 사용자 정의 교정 용어를 모두 삭제하시겠습니까?')) {
      clearCustomDictionary();
      setCustomDict({});
    }
  };

  const refreshHistory = useCallback(() => {
    proofreadHistoryDB.getAll().then(setHistoryList).catch(err => console.error('교정 이력 로드 실패:', err));
  }, []);

  useEffect(() => {
    refreshHistory();
  }, [refreshHistory]);

  const persistToHistory = useCallback((result, mode = 'typo') => {
    const recordWithMode = { ...result, mode };
    proofreadHistoryDB.saveRecord(recordWithMode).then(refreshHistory).catch(err => console.error('교정 이력 저장 실패:', err));
  }, [refreshHistory]);

  const handleLoadHistory = (record) => {
    if (record.mode) {
      setProofreadMode(record.mode);
    }
    setResultData({
      score: record.score,
      inspectionScope: null,
      summary: record.summary,
      rtm: [],
      requirementMapping: [],
      omissions: [],
      typos: record.typos,
      correctedFullText: record.correctedFullText,
      originalFullText: record.originalFullText,
      styleIssues: record.styleIssues || [],
      artifactFileName: record.fileName
    });
    setShowHistory(false);
  };

  const handleDeleteHistory = (id) => {
    proofreadHistoryDB.deleteRecord(id).then(refreshHistory).catch(err => console.error('교정 이력 삭제 실패:', err));
  };

  // 단일 문서에 대한 사전+AI 교정교열 / AI 작성투 탐지 핵심 로직. 단일 분석(handleAnalyze)과 일괄 처리(runBatch)가 공유한다.
  const runCoreAnalysis = useCallback(async (artifact, inspectionScope, glossary, artifactFileName, onProgress, mode = proofreadMode) => {
    const isAiHumanize = (mode === 'ai_humanize');
    
    // 용어 사전 탭에 화살표(→) 형식으로 명시된 지침 용어를 사용자 사전과 병합해 1단계 스캔에 반영
    const glossaryDict = extract_dict_pairs_from_glossary(glossary);
    const effectiveDict = { ...customDict, ...glossaryDict };
    const staticTypos = isAiHumanize ? [] : extract_dictionary_typos(artifact, effectiveDict);
    const styleIssues = isAiHumanize ? [] : detect_style_inconsistency(artifact);

    try {
      if (llmProvider === 'omniroute' || (apiKey && apiKey.match(/^(AIza|AQ\.)/))) {
        const result = await analyzeDocumentsWithLLM(
          '', artifact, inspectionScope, apiKey, glossary,
          onProgress,
          'auto',
          false,
          "",
          llmProvider,
          omniRouteModel,
          effectiveDict,
          mode
        );

        if (!isAiHumanize) {
          // 정적 사전 결과와 AI 결과 병합 (맞춤법 모드일 때만)
          const seenSig = new Set((result.typos || []).map(t => `${t.page}_${t.originalText}_${t.correction}`));
          staticTypos.forEach(st => {
            const sig = `${st.page}_${st.originalText}_${st.correction}`;
            if (!seenSig.has(sig)) {
              seenSig.add(sig);
              (result.typos = result.typos || []).push(st);
            }
          });
        }

        const correctedFullText = apply_typos_to_text(artifact, result.typos || []);
        return { ...result, correctedFullText, originalFullText: artifact, styleIssues, artifactFileName, mode };
      }

      if (isAiHumanize) {
        throw new Error('AI 작성투 탐지 & 인간화 추천 기능은 Gemini API Key 또는 OmniRoute LLM 연결이 필요합니다. 우측 상단 [설정]에서 API 키를 등록해 주세요.');
      }

      // API Key가 등록되지 않은 경우: 사전 기반 전수 검출 결과 우선 반환 (맞춤법 모드)
      const correctedFullText = apply_typos_to_text(artifact, staticTypos);
      return {
        score: staticTypos.length > 0 ? Math.max(60, 100 - staticTypos.length * 5) : 100,
        inspectionScope: inspectionScope || null,
        summary: staticTypos.length > 0
          ? `[사전 기반 100% 전수 검출 완료]\n문서 전체에서 ${staticTypos.length}건의 오탈자, 외래어 표기법 오류 및 순화 대상 단어를 빠짐없이 도출하였습니다. Gemini API Key를 등록하시면 5대 차원 문맥 심층 분석이 추가 적용됩니다.`
          : `[사전 기반 전수 검출 완료]\n기본 내장 사전(1만+ 규칙) 검사 결과 지적할 기계적 오탈자가 발견되지 않았습니다. 문맥상 미세한 결함 점검을 위해 Gemini API Key를 등록해 주세요.`,
        rtm: [],
        requirementMapping: [],
        omissions: [],
        typos: staticTypos,
        correctedFullText,
        originalFullText: artifact,
        styleIssues,
        artifactFileName,
        mode
      };
    } catch (e) {
      console.error('[TypoValidator] 교정교열 오류:', e);
      const correctedFullText = apply_typos_to_text(artifact, staticTypos);
      return {
        score: staticTypos.length > 0 ? Math.max(60, 100 - staticTypos.length * 5) : 0,
        inspectionScope: inspectionScope || null,
        summary: isAiHumanize 
          ? `AI 작성투 탐지 중 오류가 발생하였습니다: ${e?.message || '알 수 없는 오류'}`
          : `교정교열 과정에서 일부 오류가 발생했으나, 사전 기반 전수 검사를 통해 ${staticTypos.length}건의 결함을 도출하였습니다: ${e?.message || '알 수 없는 오류'}`,
        rtm: [],
        requirementMapping: [],
        omissions: [],
        typos: staticTypos,
        correctedFullText,
        originalFullText: artifact,
        styleIssues,
        artifactFileName,
        mode
      };
    }
  }, [apiKey, llmProvider, omniRouteModel, customDict, proofreadMode]);

  const handleAnalyze = useCallback(async (ignoredGuideline, artifact, inspectionScope, glossary, artifactFileName) => {
    lastParams.current = { artifact, inspectionScope, glossary, artifactFileName, mode: proofreadMode };
    setIsAnalyzing(true);
    setResultData(null);
    setRetryStatus(null);
    setAnalysisStage(1);
    setShowHistory(false);
    setShowDictionary(false);
    setShowBatch(false);

    // 시각적 연출을 위한 지연 (UX 목적)
    await new Promise(resolve => setTimeout(resolve, 1500));
    setAnalysisStage(2);

    const finalResult = await runCoreAnalysis(artifact, inspectionScope, glossary, artifactFileName, (status) => setRetryStatus(status), proofreadMode);
    setResultData(finalResult);
    persistToHistory(finalResult, proofreadMode);
    setIsAnalyzing(false);
    setAnalysisStage(0);
    setRetryStatus(null);
  }, [runCoreAnalysis, persistToHistory, proofreadMode]);

  const handleAddBatchFiles = (fileList) => {
    const newEntries = Array.from(fileList || []).map(file => ({
      id: ++batchIdCounter,
      file,
      name: file.name,
      status: 'queued', // queued | processing | done | error
      result: null,
      errorMsg: null,
      progress: ''
    }));
    setBatchFiles(prev => [...prev, ...newEntries]);
  };

  const handleRemoveBatchFile = (id) => {
    setBatchFiles(prev => prev.filter(f => f.id !== id));
  };

  const handleClearBatch = () => {
    setBatchFiles([]);
  };

  const handleRunBatch = async () => {
    if (batchRunning) return;
    setBatchRunning(true);

    // 클로저 문제 방지를 위해 최신 상태를 직접 참조하며 순차 처리
    const queue = batchFiles.filter(f => f.status === 'queued' || f.status === 'error');
    for (const entry of queue) {
      setBatchFiles(prev => prev.map(f => f.id === entry.id ? { ...f, status: 'processing', progress: '파일 텍스트 추출 중...', errorMsg: null } : f));
      try {
        const { text } = await processFile(entry.file);
        const result = await runCoreAnalysis(text, '', '', entry.name, (status) => {
          setBatchFiles(prev => prev.map(f => f.id === entry.id ? { ...f, progress: status } : f));
        }, batchMode);
        persistToHistory(result, batchMode);
        setBatchFiles(prev => prev.map(f => f.id === entry.id ? { ...f, status: 'done', result, progress: '완료' } : f));
      } catch (err) {
        console.error('[TypoValidator] 일괄 처리 오류:', err);
        setBatchFiles(prev => prev.map(f => f.id === entry.id ? { ...f, status: 'error', errorMsg: err?.message || '알 수 없는 오류', progress: '' } : f));
      }
      // Rate Limit 방지를 위한 파일 간 소폭 대기
      await new Promise(resolve => setTimeout(resolve, 800));
    }
    setBatchRunning(false);
  };

  const handleViewBatchResult = (entry) => {
    if (!entry.result) return;
    if (entry.result.mode) {
      setProofreadMode(entry.result.mode);
    }
    setResultData(entry.result);
    setShowBatch(false);
  };

  const handleRetry = () => {
    if (lastParams.current) {
        const { artifact, inspectionScope, glossary, artifactFileName, mode } = lastParams.current;
        if (mode) setProofreadMode(mode);
        handleAnalyze('', artifact, inspectionScope, glossary, artifactFileName);
    }
  };

  const handleReset = () => {
    setResultData(null);
    setIsAnalyzing(false);
    setAnalysisStage(0);
    lastParams.current = null;
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', width: '100%', height: '100%', gap: '16px', minHeight: 0, overflow: 'hidden' }}>
      {/* ── 상단 메인 서브 탭 & 도구 바 ───────────────────────── */}
      <div style={{
        display: 'flex', alignItems: 'center', justifyContent: 'space-between',
        background: 'rgba(255, 255, 255, 0.03)', border: '1px solid var(--glass-border)',
        borderRadius: '14px', padding: '8px 16px', flexShrink: 0, gap: '16px', flexWrap: 'wrap'
      }}>
        {/* 모드 전환 서브 탭 */}
        <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
          <button
            onClick={() => { setProofreadMode('typo'); setResultData(null); }}
            className="interactive"
            style={{
              background: proofreadMode === 'typo' ? 'linear-gradient(135deg, rgba(59, 130, 246, 0.25), rgba(37, 99, 235, 0.35))' : 'rgba(255, 255, 255, 0.04)',
              border: proofreadMode === 'typo' ? '1px solid var(--accent-blue)' : '1px solid var(--glass-border)',
              padding: '8px 18px', borderRadius: '10px',
              color: proofreadMode === 'typo' ? '#fff' : 'var(--text-secondary)',
              fontSize: '13px', fontWeight: proofreadMode === 'typo' ? 700 : 500, cursor: 'pointer',
              display: 'flex', alignItems: 'center', gap: '8px',
              boxShadow: proofreadMode === 'typo' ? '0 0 16px rgba(59, 130, 246, 0.25)' : 'none',
              transition: 'all 0.2s ease'
            }}
          >
            <PenTool size={15} color={proofreadMode === 'typo' ? 'var(--accent-blue)' : 'var(--text-secondary)'} />
            표준 맞춤법·오탈자 교정
          </button>

          <button
            onClick={() => { setProofreadMode('ai_humanize'); setResultData(null); }}
            className="interactive"
            style={{
              background: proofreadMode === 'ai_humanize' ? 'linear-gradient(135deg, rgba(168, 85, 247, 0.25), rgba(147, 51, 234, 0.35))' : 'rgba(255, 255, 255, 0.04)',
              border: proofreadMode === 'ai_humanize' ? '1px solid var(--accent-purple)' : '1px solid var(--glass-border)',
              padding: '8px 18px', borderRadius: '10px',
              color: proofreadMode === 'ai_humanize' ? '#fff' : 'var(--text-secondary)',
              fontSize: '13px', fontWeight: proofreadMode === 'ai_humanize' ? 700 : 500, cursor: 'pointer',
              display: 'flex', alignItems: 'center', gap: '8px',
              boxShadow: proofreadMode === 'ai_humanize' ? '0 0 16px rgba(168, 85, 247, 0.25)' : 'none',
              transition: 'all 0.2s ease'
            }}
          >
            <Sparkles size={15} color={proofreadMode === 'ai_humanize' ? 'var(--accent-purple)' : 'var(--text-secondary)'} />
            AI 작성투 탐지 & 인간화 추천
            <span style={{ fontSize: '10px', background: 'rgba(168, 85, 247, 0.3)', color: '#d8b4fe', padding: '2px 6px', borderRadius: '6px', fontWeight: 700 }}>NEW</span>
          </button>
        </div>

        {/* 우측 보조 도구 버튼들 */}
        <div style={{ display: 'flex', gap: '8px', alignItems: 'center', marginLeft: 'auto' }}>
          <button
            onClick={() => { setShowHistory(v => !v); setShowDictionary(false); setShowBatch(false); }}
            className="interactive"
            style={{
              background: showHistory ? 'rgba(59, 130, 246, 0.18)' : 'rgba(255, 255, 255, 0.06)',
              border: '1px solid var(--glass-border)',
              padding: '7px 14px', borderRadius: '9px',
              color: showHistory ? 'var(--accent-blue)' : 'var(--text-secondary)',
              fontSize: '12px', fontWeight: 600, cursor: 'pointer',
              display: 'flex', alignItems: 'center', gap: '6px'
            }}
          >
            <History size={14} /> 교정 이력{historyList.length > 0 ? ` (${historyList.length})` : ''}
          </button>
          <button
            onClick={() => { setShowDictionary(v => !v); setShowHistory(false); setShowBatch(false); }}
            className="interactive"
            style={{
              background: showDictionary ? 'rgba(168, 85, 247, 0.18)' : 'rgba(255, 255, 255, 0.06)',
              border: '1px solid var(--glass-border)',
              padding: '7px 14px', borderRadius: '9px',
              color: showDictionary ? 'var(--accent-purple)' : 'var(--text-secondary)',
              fontSize: '12px', fontWeight: 600, cursor: 'pointer',
              display: 'flex', alignItems: 'center', gap: '6px'
            }}
          >
            <BookPlus size={14} /> 사용자 사전{Object.keys(customDict).length > 0 ? ` (${Object.keys(customDict).length})` : ''}
          </button>
          <button
            onClick={() => { setShowBatch(v => !v); setShowHistory(false); setShowDictionary(false); }}
            className="interactive"
            style={{
              background: showBatch ? 'rgba(16, 185, 129, 0.18)' : 'rgba(255, 255, 255, 0.06)',
              border: '1px solid var(--glass-border)',
              padding: '7px 14px', borderRadius: '9px',
              color: showBatch ? 'var(--success-color)' : 'var(--text-secondary)',
              fontSize: '12px', fontWeight: 600, cursor: 'pointer',
              display: 'flex', alignItems: 'center', gap: '6px'
            }}
          >
            <Layers size={14} /> 일괄 처리{batchFiles.length > 0 ? ` (${batchFiles.length})` : ''}
          </button>
        </div>
      </div>

      {/* ── 본문 영역: 입력 섹션 + 결과/대기 패널 ────────────────── */}
      <div style={{ display: 'flex', flex: 1, width: '100%', gap: '24px', minHeight: 0, overflow: 'hidden' }}>
        <InputSection onAnalyze={handleAnalyze} isAnalyzing={isAnalyzing} isTypoMode={true} onReset={handleReset} />

        <div style={{ flex: 1, position: 'relative', minHeight: 0, overflow: 'hidden', display: 'flex', flexDirection: 'column' }}>

        {isAnalyzing ? (
          <div className="glass-panel animate-fade-in" style={{ flex: 1, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', color: 'var(--text-secondary)', padding: '40px' }}>
            <div className="scanning-container" style={{ marginBottom: '32px' }}>
              <div className="scanning-line"></div>
              <div style={{ padding: '15px', display: 'flex', flexDirection: 'column', gap: '8px', opacity: 0.4 }}>
                  {[...Array(6)].map((_, i) => (
                      <div key={i} style={{ height: '6px', width: `${Math.random() * 40 + 40}%`, background: 'rgba(255,255,255,0.2)', borderRadius: '3px' }}></div>
                  ))}
              </div>
            </div>

            <h2 className="pulse-text" style={{ margin: '0 0 16px', fontSize: '24px', color: 'var(--text-primary)', fontWeight: 700, letterSpacing: '-0.5px' }}>
              {proofreadMode === 'ai_humanize'
                ? (analysisStage === 1 ? '단락별 문맥 구조 스캐닝 중...' : 'AI 작성투·번역투 탐지 및 자연스러운 인간화 문구 생성 중...')
                : (analysisStage === 1 ? '사전 규칙 및 단락별 전수 스캐닝 중...' : 'AI 5대 차원 심층 분석 & 2-Pass 잔여 검수 중...')}
            </h2>
            
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px', alignItems: 'center' }}>
              <div className={`page-container active`} style={{ position: 'relative', display: 'flex', alignItems: 'center', gap: '10px', fontSize: '15px', opacity: analysisStage >= 1 ? 1 : 0.3 }}>
                  <div style={{ width: '10px', height: '10px', borderRadius: '50%', background: analysisStage > 1 ? 'var(--success-color)' : (proofreadMode === 'ai_humanize' ? 'var(--accent-purple)' : 'var(--accent-blue)'), boxShadow: analysisStage === 1 ? `0 0 10px ${proofreadMode === 'ai_humanize' ? 'var(--accent-purple)' : 'var(--accent-blue)'}` : 'none' }}></div>
                  <span style={{ color: analysisStage === 1 ? 'var(--text-primary)' : 'var(--text-muted)' }}>
                    {proofreadMode === 'ai_humanize' ? '1단계: 문서 단락 및 문장 전수 스캐닝' : '1단계: 사전 기반 규칙 & 단락별 문장 전수 스캐닝'}
                  </span>
                  {analysisStage > 1 && <span style={{ color: 'var(--success-color)', fontSize: '12px', fontWeight: 700 }}>완료</span>}
              </div>
              <div className={`page-container active`} style={{ position: 'relative', display: 'flex', alignItems: 'center', gap: '10px', fontSize: '15px', opacity: analysisStage >= 2 ? 1 : 0.3 }}>
                  <div style={{ width: '10px', height: '10px', borderRadius: '50%', background: analysisStage === 2 ? (proofreadMode === 'ai_humanize' ? 'var(--accent-purple)' : 'var(--accent-blue)') : 'rgba(255,255,255,0.1)', boxShadow: analysisStage === 2 ? `0 0 10px ${proofreadMode === 'ai_humanize' ? 'var(--accent-purple)' : 'var(--accent-blue)'}` : 'none' }}></div>
                  <span style={{ color: analysisStage === 2 ? 'var(--text-primary)' : 'var(--text-muted)' }}>
                    {proofreadMode === 'ai_humanize' ? '2단계: AI 작성투·번역투·공허한 미사여구 정밀 분석 및 인간화 추천' : '2단계: 5대 차원 심층 품질 분석 & 2-Pass 잔여 오류 완벽 도출'}
                  </span>
              </div>
            </div>

            <div className="progress-track" style={{ width: '350px' }}>
              <div className="progress-fill" style={{ width: analysisStage === 1 ? '45%' : '90%', background: proofreadMode === 'ai_humanize' ? 'linear-gradient(90deg, var(--accent-purple), #ec4899)' : undefined }}></div>
            </div>

            {retryStatus && (
              <div className="animate-fade-in" style={{ marginTop: '20px', padding: '10px 20px', background: 'rgba(245, 158, 11, 0.1)', border: '1px solid rgba(245, 158, 11, 0.2)', borderRadius: '10px', color: 'var(--warning-color)', fontSize: '13px', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <Loader2 size={14} className="animate-spin" />
                  {retryStatus}
              </div>
            )}

            <p style={{ marginTop: '24px', fontSize: '14px', color: 'var(--text-muted)', maxWidth: '440px', textAlign: 'center', lineHeight: '1.6' }}>
              {proofreadMode === 'ai_humanize'
                ? 'AI 특유의 기계적인 접속사 남발, 어색한 피동 번역투, 공허한 수식어를 탐지하여 실무 전문가가 쓴 자연스러운 문장으로 변환합니다.'
                : (analysisStage === 1 
                    ? '단락별 분할 스캔을 통해 문서 전체의 문장을 100% 누락 없이 읽어 들이고 있습니다.' 
                    : '1차 오탈자부터 가려져 있던 2차 잔여 문맥·호응 결함까지 한 번에 완벽히 도출합니다.')}
            </p>
          </div>
        ) : resultData ? (
          <div style={{ flex: 1, display: 'flex', flexDirection: 'column', minHeight: 0, overflow: 'hidden', position: 'relative' }}>
            <div style={{ position: 'absolute', top: '24px', right: '24px', zIndex: 10 }}>
              <button 
                onClick={() => setResultData(null)}
                className="interactive"
                style={{
                  background: 'rgba(255, 255, 255, 0.08)',
                  border: '1px solid var(--glass-border)',
                  padding: '8px 16px',
                  borderRadius: '10px',
                  color: 'var(--text-secondary)',
                  fontSize: '13px',
                  fontWeight: 600,
                  cursor: 'pointer',
                  display: 'flex',
                  alignItems: 'center',
                  gap: '8px',
                  backdropFilter: 'blur(8px)'
                }}
              >
                <ArrowRight size={16} style={{ transform: 'rotate(180deg)' }} /> 결과 닫기
              </button>
            </div>
            <ResultDashboard data={resultData} isTypoMode={true} onRetry={handleRetry} />
          </div>
        ) : (
          <div className="glass-panel animate-fade-in" style={{ flex: 1, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', color: 'var(--text-secondary)', padding: '40px' }}>
            <div style={{ padding: '32px', borderRadius: '50%', background: proofreadMode === 'ai_humanize' ? 'rgba(168, 85, 247, 0.08)' : 'rgba(59, 130, 246, 0.08)', border: `1px solid ${proofreadMode === 'ai_humanize' ? 'rgba(168, 85, 247, 0.2)' : 'rgba(59, 130, 246, 0.2)'}`, marginBottom: '32px', boxShadow: `0 0 40px ${proofreadMode === 'ai_humanize' ? 'rgba(168, 85, 247, 0.15)' : 'rgba(59, 130, 246, 0.15)'}` }}>
              {proofreadMode === 'ai_humanize' ? (
                <Sparkles size={56} color="var(--accent-purple)" opacity={0.8} />
              ) : (
                <PenTool size={56} color="var(--accent-blue)" opacity={0.8} />
              )}
            </div>
            <h2 style={{ margin: '0 0 16px', fontSize: '24px', color: 'var(--text-primary)', fontWeight: 700, letterSpacing: '-0.5px' }}>
              {proofreadMode === 'ai_humanize' ? 'AI 작성투 탐지 & 인간화 추천 대기 중' : '표준 맞춤법·오탈자 교정 대기 중'}
            </h2>
            <div style={{ maxWidth: '440px', textAlign: 'center', lineHeight: '1.7', fontSize: '15px', color: 'var(--text-secondary)' }}>
              <p style={{ marginBottom: '16px' }}>
                좌측 영역에 <strong style={{ color: 'var(--text-primary)' }}>&lt;검증 대상 문서&gt;</strong> 내용을 입력하거나 파일을 업로드하세요.
              </p>
              <div style={{ background: proofreadMode === 'ai_humanize' ? 'rgba(168, 85, 247, 0.06)' : 'rgba(59, 130, 246, 0.06)', padding: '16px', borderRadius: '12px', border: `1px solid ${proofreadMode === 'ai_humanize' ? 'rgba(168, 85, 247, 0.15)' : 'rgba(59, 130, 246, 0.15)'}`, color: 'var(--text-muted)', fontSize: '13px' }}>
                💡 <span style={{ color: proofreadMode === 'ai_humanize' ? 'var(--accent-purple)' : 'var(--accent-blue)', fontWeight: 600 }}>Tip:</span> {proofreadMode === 'ai_humanize' ? 'AI 특유의 번역투/기계적 수식어를 찾아 사람이 쓴 자연스럽고 전문적인 문장으로 변환합니다.' : '오탈자, 외래어 표기, 띄어쓰기 및 비즈니스 문맥 오기를 전수 검출하여 완벽한 공문서로 다듬습니다.'}
              </div>
            </div>
          </div>
        )}

        {showHistory && (
          <div className="animate-fade-in" style={{
            position: 'absolute', inset: 0, zIndex: 20, borderRadius: '16px',
            background: 'rgba(15, 15, 22, 0.94)', backdropFilter: 'blur(10px)',
            padding: '24px', display: 'flex', flexDirection: 'column', overflow: 'hidden'
          }}>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '16px' }}>
              <h3 style={{ margin: 0, fontSize: '18px', fontWeight: 700, color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                <History size={20} color="var(--accent-blue)" /> 교정 이력 (최근 {historyList.length}건, 세션 간 유지)
              </h3>
              <div style={{ display: 'flex', gap: '8px' }}>
                {historyList.length > 0 && (
                  <button
                    onClick={() => {
                      if (window.confirm('저장된 교정 이력을 모두 삭제하시겠습니까?')) {
                        proofreadHistoryDB.clearAll().then(refreshHistory);
                      }
                    }}
                    className="interactive"
                    style={{ background: 'rgba(239, 68, 68, 0.1)', border: '1px solid rgba(239, 68, 68, 0.2)', color: 'var(--danger-color)', borderRadius: '10px', padding: '8px 14px', fontSize: '12px', fontWeight: 600, cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '6px' }}
                  >
                    <Trash2 size={14} /> 전체 삭제
                  </button>
                )}
                <button
                  onClick={() => setShowHistory(false)}
                  className="interactive"
                  style={{ background: 'rgba(255,255,255,0.08)', border: '1px solid var(--glass-border)', color: 'var(--text-secondary)', borderRadius: '10px', padding: '8px 10px', cursor: 'pointer', display: 'flex', alignItems: 'center' }}
                >
                  <X size={16} />
                </button>
              </div>
            </div>

            <div style={{ flex: 1, overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '10px' }}>
              {historyList.length === 0 ? (
                <div style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'var(--text-muted)', fontSize: '14px' }}>
                  아직 저장된 교정 이력이 없습니다. 분석을 완료하면 자동으로 이력이 쌓입니다.
                </div>
              ) : (
                historyList.map((record) => (
                  <div key={record.id} className="interactive" style={{
                    display: 'flex', alignItems: 'center', gap: '14px', padding: '14px 16px',
                    background: 'rgba(255,255,255,0.03)', border: '1px solid var(--glass-border)', borderRadius: '12px'
                  }}>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontSize: '14px', fontWeight: 600, color: 'var(--text-primary)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <span>{record.fileName}</span>
                        {record.mode === 'ai_humanize' ? (
                          <span style={{ fontSize: '10px', background: 'rgba(168, 85, 247, 0.2)', color: '#d8b4fe', padding: '2px 6px', borderRadius: '4px', fontWeight: 700 }}>AI 인간화</span>
                        ) : (
                          <span style={{ fontSize: '10px', background: 'rgba(59, 130, 246, 0.2)', color: '#93c5fd', padding: '2px 6px', borderRadius: '4px', fontWeight: 700 }}>맞춤법 교정</span>
                        )}
                      </div>
                      <div style={{ fontSize: '12px', color: 'var(--text-secondary)', marginTop: '4px', display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
                        <span>{new Date(record.createdAt).toLocaleString('ko-KR')}</span>
                        <span>발견 항목 {record.typoCount}건</span>
                        {typeof record.score === 'number' && <span>점수 {record.score}점</span>}
                      </div>
                    </div>
                    <button
                      onClick={() => handleLoadHistory(record)}
                      className="interactive"
                      style={{ background: 'rgba(59, 130, 246, 0.12)', border: '1px solid rgba(59, 130, 246, 0.3)', color: 'var(--accent-blue)', borderRadius: '8px', padding: '8px 14px', fontSize: '12px', fontWeight: 700, cursor: 'pointer', flexShrink: 0 }}
                    >
                      불러오기
                    </button>
                    <button
                      onClick={() => handleDeleteHistory(record.id)}
                      className="interactive"
                      style={{ background: 'transparent', border: 'none', color: 'var(--text-secondary)', cursor: 'pointer', flexShrink: 0, padding: '6px' }}
                    >
                      <Trash2 size={16} />
                    </button>
                  </div>
                ))
              )}
            </div>
          </div>
        )}

        {showDictionary && (
          <div className="animate-fade-in" style={{
            position: 'absolute', inset: 0, zIndex: 20, borderRadius: '16px',
            background: 'rgba(15, 15, 22, 0.94)', backdropFilter: 'blur(10px)',
            padding: '24px', display: 'flex', flexDirection: 'column', overflow: 'hidden'
          }}>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '16px' }}>
              <h3 style={{ margin: 0, fontSize: '18px', fontWeight: 700, color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                <BookPlus size={20} color="var(--accent-purple)" /> 사용자 정의 교정 용어 사전 ({Object.keys(customDict).length}건)
              </h3>
              <div style={{ display: 'flex', gap: '8px' }}>
                {Object.keys(customDict).length > 0 && (
                  <button
                    onClick={handleClearCustomDict}
                    className="interactive"
                    style={{ background: 'rgba(239, 68, 68, 0.1)', border: '1px solid rgba(239, 68, 68, 0.2)', color: 'var(--danger-color)', borderRadius: '10px', padding: '8px 14px', fontSize: '12px', fontWeight: 600, cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '6px' }}
                  >
                    <Trash2 size={14} /> 전체 삭제
                  </button>
                )}
                <button
                  onClick={() => setShowDictionary(false)}
                  className="interactive"
                  style={{ background: 'rgba(255,255,255,0.08)', border: '1px solid var(--glass-border)', color: 'var(--text-secondary)', borderRadius: '10px', padding: '8px 10px', cursor: 'pointer', display: 'flex', alignItems: 'center' }}
                >
                  <X size={16} />
                </button>
              </div>
            </div>

            <form onSubmit={handleAddCustomTerm} style={{ display: 'flex', gap: '8px', marginBottom: '16px', flexWrap: 'wrap' }}>
              <input
                value={newTermWrong}
                onChange={(e) => setNewTermWrong(e.target.value)}
                placeholder="오류 표현 (예: 리스트관리)"
                style={{ flex: '1 1 160px', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--glass-border)', background: 'rgba(0,0,0,0.2)', color: 'var(--text-primary)', outline: 'none', fontSize: '13px' }}
              />
              <input
                value={newTermCorrect}
                onChange={(e) => setNewTermCorrect(e.target.value)}
                placeholder="올바른 표현 (예: 리스크관리)"
                style={{ flex: '1 1 160px', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--glass-border)', background: 'rgba(0,0,0,0.2)', color: 'var(--text-primary)', outline: 'none', fontSize: '13px' }}
              />
              <input
                value={newTermDesc}
                onChange={(e) => setNewTermDesc(e.target.value)}
                placeholder="사유 (선택)"
                style={{ flex: '1 1 160px', padding: '10px 12px', borderRadius: '8px', border: '1px solid var(--glass-border)', background: 'rgba(0,0,0,0.2)', color: 'var(--text-primary)', outline: 'none', fontSize: '13px' }}
              />
              <button
                type="submit"
                className="interactive"
                style={{ background: 'rgba(168, 85, 247, 0.15)', border: '1px solid rgba(168, 85, 247, 0.4)', color: 'var(--accent-purple)', borderRadius: '8px', padding: '10px 16px', fontSize: '13px', fontWeight: 700, cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '6px', flexShrink: 0 }}
              >
                <Plus size={16} /> 추가
              </button>
            </form>
            {dictError && (
              <div style={{ marginBottom: '12px', padding: '8px 12px', fontSize: '12px', background: 'rgba(239, 68, 68, 0.08)', color: 'var(--danger-color)', borderRadius: '8px', border: '1px solid rgba(239, 68, 68, 0.15)' }}>
                {dictError}
              </div>
            )}

            <div style={{ flex: 1, overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '10px' }}>
              {Object.keys(customDict).length === 0 ? (
                <div style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'var(--text-muted)', fontSize: '14px' }}>
                  등록된 사용자 정의 용어가 없습니다. 위 입력창에서 자주 틀리는 기관/사업 용어를 등록해 보세요.
                </div>
              ) : (
                Object.entries(customDict).map(([wrong, info]) => (
                  <div key={wrong} className="interactive" style={{
                    display: 'flex', alignItems: 'center', gap: '14px', padding: '14px 16px',
                    background: 'rgba(255,255,255,0.03)', border: '1px solid var(--glass-border)', borderRadius: '12px'
                  }}>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontSize: '14px', fontWeight: 600, color: 'var(--text-primary)' }}>
                        <span style={{ color: '#f87171' }}>{wrong}</span>
                        <span style={{ margin: '0 8px', color: 'var(--text-muted)' }}>→</span>
                        <span style={{ color: '#34d399' }}>{info.correction}</span>
                      </div>
                      {info.desc && (
                        <div style={{ fontSize: '12px', color: 'var(--text-secondary)', marginTop: '4px' }}>{info.desc}</div>
                      )}
                    </div>
                    <button
                      onClick={() => handleDeleteCustomTerm(wrong)}
                      className="interactive"
                      style={{ background: 'transparent', border: 'none', color: 'var(--text-secondary)', cursor: 'pointer', flexShrink: 0, padding: '6px' }}
                    >
                      <Trash2 size={16} />
                    </button>
                  </div>
                ))
              )}
            </div>
          </div>
        )}

        {showBatch && (
          <div className="animate-fade-in" style={{
            position: 'absolute', inset: 0, zIndex: 20, borderRadius: '16px',
            background: 'rgba(15, 15, 22, 0.94)', backdropFilter: 'blur(10px)',
            padding: '24px', display: 'flex', flexDirection: 'column', overflow: 'hidden'
          }}>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '12px' }}>
              <h3 style={{ margin: 0, fontSize: '18px', fontWeight: 700, color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                <Layers size={20} color="var(--success-color)" /> 다중 문서 일괄 처리 ({batchFiles.length}건)
              </h3>
              <div style={{ display: 'flex', gap: '8px' }}>
                {batchFiles.length > 0 && (
                  <button
                    onClick={handleClearBatch}
                    disabled={batchRunning}
                    className="interactive"
                    style={{ background: 'rgba(239, 68, 68, 0.1)', border: '1px solid rgba(239, 68, 68, 0.2)', color: 'var(--danger-color)', borderRadius: '10px', padding: '8px 14px', fontSize: '12px', fontWeight: 600, cursor: batchRunning ? 'not-allowed' : 'pointer', opacity: batchRunning ? 0.5 : 1, display: 'flex', alignItems: 'center', gap: '6px' }}
                  >
                    <Trash2 size={14} /> 목록 비우기
                  </button>
                )}
                <button
                  onClick={() => setShowBatch(false)}
                  className="interactive"
                  style={{ background: 'rgba(255,255,255,0.08)', border: '1px solid var(--glass-border)', color: 'var(--text-secondary)', borderRadius: '10px', padding: '8px 10px', cursor: 'pointer', display: 'flex', alignItems: 'center' }}
                >
                  <X size={16} />
                </button>
              </div>
            </div>

            <p style={{ margin: '0 0 16px', fontSize: '13px', color: 'var(--text-muted)', lineHeight: '1.6' }}>
              여러 산출물 파일을 한 번에 등록해 순차적으로 검수합니다. 각 파일 결과는 교정 이력에도 자동 저장됩니다.
            </p>

            {/* 일괄 처리 모드 선택 */}
            <div style={{ display: 'flex', gap: '12px', alignItems: 'center', marginBottom: '16px', padding: '10px 14px', background: 'rgba(255,255,255,0.02)', border: '1px solid var(--glass-border)', borderRadius: '10px' }}>
              <span style={{ fontSize: '12px', fontWeight: 600, color: 'var(--text-secondary)' }}>일괄 처리 모드:</span>
              <label style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px', color: 'var(--text-primary)', cursor: 'pointer' }}>
                <input
                  type="radio"
                  name="batchMode"
                  value="typo"
                  checked={batchMode === 'typo'}
                  onChange={() => setBatchMode('typo')}
                  disabled={batchRunning}
                />
                맞춤법·오탈자 교정
              </label>
              <label style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px', color: 'var(--text-primary)', cursor: 'pointer', marginLeft: '12px' }}>
                <input
                  type="radio"
                  name="batchMode"
                  value="ai_humanize"
                  checked={batchMode === 'ai_humanize'}
                  onChange={() => setBatchMode('ai_humanize')}
                  disabled={batchRunning}
                />
                AI 작성투 탐지 & 인간화 추천
              </label>
            </div>

            <div style={{ display: 'flex', gap: '8px', marginBottom: '16px' }}>
              <button
                onClick={() => batchFileInputRef.current?.click()}
                disabled={batchRunning}
                className="interactive"
                style={{ background: 'rgba(59, 130, 246, 0.12)', border: '1px solid rgba(59, 130, 246, 0.3)', color: 'var(--accent-blue)', borderRadius: '8px', padding: '10px 16px', fontSize: '13px', fontWeight: 700, cursor: batchRunning ? 'not-allowed' : 'pointer', opacity: batchRunning ? 0.5 : 1, display: 'flex', alignItems: 'center', gap: '8px' }}
              >
                <Upload size={16} /> 파일 추가
              </button>
              <input
                ref={batchFileInputRef}
                type="file"
                multiple
                accept={ALL_ACCEPT}
                style={{ display: 'none' }}
                onChange={(e) => { handleAddBatchFiles(e.target.files); e.target.value = ''; }}
              />
              <button
                onClick={handleRunBatch}
                disabled={batchRunning || batchFiles.filter(f => f.status === 'queued' || f.status === 'error').length === 0}
                className="interactive"
                style={{
                  background: 'linear-gradient(135deg, #10b981 0%, #059669 100%)', border: 'none', color: '#fff',
                  borderRadius: '8px', padding: '10px 18px', fontSize: '13px', fontWeight: 700,
                  cursor: (batchRunning || batchFiles.filter(f => f.status === 'queued' || f.status === 'error').length === 0) ? 'not-allowed' : 'pointer',
                  opacity: (batchRunning || batchFiles.filter(f => f.status === 'queued' || f.status === 'error').length === 0) ? 0.5 : 1,
                  display: 'flex', alignItems: 'center', gap: '8px'
                }}
              >
                {batchRunning ? <Loader2 size={16} className="animate-spin" /> : <Sparkles size={16} />}
                {batchRunning ? '일괄 처리 중...' : '전체 시작'}
              </button>
            </div>

            <div style={{ flex: 1, overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '10px' }}>
              {batchFiles.length === 0 ? (
                <div style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'var(--text-muted)', fontSize: '14px' }}>
                  파일 추가 버튼으로 여러 산출물을 등록한 뒤 전체 시작을 눌러 일괄 검수하세요.
                </div>
              ) : (
                batchFiles.map((entry) => (
                  <div key={entry.id} className="interactive" style={{
                    display: 'flex', alignItems: 'center', gap: '14px', padding: '14px 16px',
                    background: 'rgba(255,255,255,0.03)', border: '1px solid var(--glass-border)', borderRadius: '12px'
                  }}>
                    <div style={{ flexShrink: 0 }}>
                      {entry.status === 'done' && <CheckCircle2 size={20} color="var(--success-color)" />}
                      {entry.status === 'error' && <XCircle size={20} color="var(--danger-color)" />}
                      {entry.status === 'processing' && <Loader2 size={20} className="animate-spin" color="var(--accent-blue)" />}
                      {entry.status === 'queued' && <div style={{ width: '20px', height: '20px', borderRadius: '50%', border: '2px solid var(--glass-border)' }} />}
                    </div>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontSize: '14px', fontWeight: 600, color: 'var(--text-primary)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                        {entry.name}
                      </div>
                      <div style={{ fontSize: '12px', color: 'var(--text-secondary)', marginTop: '4px' }}>
                        {entry.status === 'queued' && '대기 중'}
                        {entry.status === 'processing' && (entry.progress || '처리 중...')}
                        {entry.status === 'done' && `완료 · 발견 ${entry.result?.typos?.length || 0}건${typeof entry.result?.score === 'number' ? ` · 점수 ${entry.result.score}점` : ''}`}
                        {entry.status === 'error' && `오류: ${entry.errorMsg}`}
                      </div>
                    </div>
                    {entry.status === 'done' && (
                      <button
                        onClick={() => handleViewBatchResult(entry)}
                        className="interactive"
                        style={{ background: 'rgba(59, 130, 246, 0.12)', border: '1px solid rgba(59, 130, 246, 0.3)', color: 'var(--accent-blue)', borderRadius: '8px', padding: '8px 14px', fontSize: '12px', fontWeight: 700, cursor: 'pointer', flexShrink: 0 }}
                      >
                        결과 보기
                      </button>
                    )}
                    {(entry.status === 'queued' || entry.status === 'error') && !batchRunning && (
                      <button
                        onClick={() => handleRemoveBatchFile(entry.id)}
                        className="interactive"
                        style={{ background: 'transparent', border: 'none', color: 'var(--text-secondary)', cursor: 'pointer', flexShrink: 0, padding: '6px' }}
                      >
                        <Trash2 size={16} />
                      </button>
                    )}
                  </div>
                ))
              )}
            </div>
          </div>
        )}
        </div>
      </div>
    </div>
  );
}

export default TypoValidator;
