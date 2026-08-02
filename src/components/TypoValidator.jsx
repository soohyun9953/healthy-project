import React, { useState, useCallback, useRef } from 'react';
import { ArrowRight, Loader2, PenTool, RotateCcw } from 'lucide-react';
import InputSection from './InputSection';
import ResultDashboard from './ResultDashboard';
import { analyzeDocumentsWithLLM } from '../llmAnalyzer';

function TypoValidator({ apiKey, llmProvider = 'gemini', ollamaModel = 'qwen2.5:3b' }) {
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysisStage, setAnalysisStage] = useState(0); // 1: 추출, 2: 심층분석
  const [retryStatus, setRetryStatus] = useState(null); // API 재시도 상태 메시지
  const [resultData, setResultData] = useState(null);
  const lastParams = useRef(null);

  const handleAnalyze = useCallback(async (ignoredGuideline, artifact, inspectionScope, glossary, artifactFileName) => {
    lastParams.current = { artifact, inspectionScope, glossary, artifactFileName };
    setIsAnalyzing(true);
    setResultData(null);
    setRetryStatus(null);
    setAnalysisStage(1);

    // 시각적 연출을 위한 인위적 지연 (UX 목적)
    await new Promise(resolve => setTimeout(resolve, 2500));
    setAnalysisStage(2);

    try {
      if (llmProvider === 'ollama' || (apiKey && apiKey.match(/^(AIza|AQ\.)/))) {
        const result = await analyzeDocumentsWithLLM(
          '', artifact, inspectionScope, apiKey, glossary,
          (status) => setRetryStatus(status),
          'auto',
          false,
          "",
          llmProvider,
          ollamaModel
        );
        setResultData({ ...result, artifactFileName });
      } else {
        throw new Error('유효한 Gemini API Key가 필요합니다. [설정] 메뉴에서 API Key를 입력하거나 로컬 LLM을 선택해 주세요.');
      }
    } catch (e) {
        console.error('[TypoValidator] 교정교열 오류:', e);
        setResultData({
            score: 0,
            inspectionScope: inspectionScope || null,
            summary: `교정교열 과정에서 오류가 발생했습니다: ${e?.message || '알 수 없는 오류'}`,
            rtm: [],
            requirementMapping: [],
            omissions: [],
            typos: [],
        });
    } finally {
        setIsAnalyzing(false);
        setAnalysisStage(0);
        setRetryStatus(null);
    }
  }, [apiKey, llmProvider, ollamaModel]);

  const handleRetry = () => {
    if (lastParams.current) {
        const { artifact, inspectionScope, glossary, artifactFileName } = lastParams.current;
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
    <div style={{ display: 'flex', width: '100%', height: '100%', gap: '24px', minHeight: 0, overflow: 'hidden' }}>
      <InputSection onAnalyze={handleAnalyze} isAnalyzing={isAnalyzing} isTypoMode={true} onReset={handleReset} />

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
            {analysisStage === 1 ? '전수 문장 단위 도출 중...' : 'AI 5대 차원 심층 품질 점검 중...'}
          </h2>
          
          <div style={{ display: 'flex', flexDirection: 'column', gap: '12px', alignItems: 'center' }}>
            <div className={`page-container active`} style={{ position: 'relative', display: 'flex', alignItems: 'center', gap: '10px', fontSize: '15px', opacity: analysisStage >= 1 ? 1 : 0.3 }}>
                <div style={{ width: '10px', height: '10px', borderRadius: '50%', background: analysisStage > 1 ? 'var(--success-color)' : 'var(--accent-blue)', boxShadow: analysisStage === 1 ? '0 0 10px var(--accent-blue)' : 'none' }}></div>
                <span style={{ color: analysisStage === 1 ? 'var(--text-primary)' : 'var(--text-muted)' }}>1단계: 산출물 문장 단위 전수 스캐닝</span>
                {analysisStage > 1 && <span style={{ color: 'var(--success-color)', fontSize: '12px', fontWeight: 700 }}>완료</span>}
            </div>
            <div className={`page-container active`} style={{ position: 'relative', display: 'flex', alignItems: 'center', gap: '10px', fontSize: '15px', opacity: analysisStage >= 2 ? 1 : 0.3 }}>
                <div style={{ width: '10px', height: '10px', borderRadius: '50%', background: analysisStage === 2 ? 'var(--accent-blue)' : 'rgba(255,255,255,0.1)', boxShadow: analysisStage === 2 ? '0 0 10px var(--accent-blue)' : 'none' }}></div>
                <span style={{ color: analysisStage === 2 ? 'var(--text-primary)' : 'var(--text-muted)' }}>2단계: 5대 차원 심층 품질 분석 및 개선 가이드 생성</span>
            </div>
          </div>

          <div className="progress-track" style={{ width: '350px' }}>
            <div className="progress-fill" style={{ width: analysisStage === 1 ? '45%' : '90%' }}></div>
          </div>

          {retryStatus && (
            <div className="animate-fade-in" style={{ marginTop: '20px', padding: '10px 20px', background: 'rgba(245, 158, 11, 0.1)', border: '1px solid rgba(245, 158, 11, 0.2)', borderRadius: '10px', color: 'var(--warning-color)', fontSize: '13px', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '8px' }}>
                <Loader2 size={14} className="animate-spin" />
                {retryStatus}
            </div>
          )}

          <p style={{ marginTop: '24px', fontSize: '14px', color: 'var(--text-muted)', maxWidth: '400px', textAlign: 'center', lineHeight: '1.6' }}>
            {analysisStage === 1 
                ? '분석 대상 문서의 모든 문장을 하나하나 읽어 들이며 분석 대상을 추출하고 있습니다.' 
                : '산출물의 오탈자는 물론 논리 구조, 완결성, 사업 정합성을 5대 차원에서 심층 점검하고 있습니다.'}
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
          <div style={{ padding: '32px', borderRadius: '50%', background: 'rgba(168, 85, 247, 0.05)', border: '1px solid rgba(168, 85, 247, 0.1)', marginBottom: '32px', boxShadow: '0 0 40px rgba(168, 85, 247, 0.1)' }}>
            <PenTool size={56} color="var(--accent-purple)" opacity={0.6} />
          </div>
          <h2 style={{ margin: '0 0 16px', fontSize: '24px', color: 'var(--text-primary)', fontWeight: 700, letterSpacing: '-0.5px' }}>분석 대기 중</h2>
          <div style={{ maxWidth: '420px', textAlign: 'center', lineHeight: '1.7', fontSize: '15px', color: 'var(--text-secondary)' }}>
            <p style={{ marginBottom: '16px' }}>좌측 영역에 <strong style={{ color: 'var(--text-primary)' }}>&lt;검증 대상 문서&gt;</strong> 내용을 입력하거나 파일을 업로드하세요.</p>
            <div style={{ background: 'rgba(168, 85, 247, 0.05)', padding: '16px', borderRadius: '12px', border: '1px solid rgba(168, 85, 247, 0.1)', color: 'var(--text-muted)', fontSize: '13px' }}>
                💡 <span style={{ color: 'var(--accent-purple)', fontWeight: 600 }}>Tip:</span> 전문적인 5대 차원 심층 품질 점검을 위해 우측 상단에 Gemini API Key 입력을 확인해 주세요.
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default TypoValidator;
