import React, { useState, useRef, useCallback } from 'react';
import { 
  ArrowRightLeft, 
  Upload, 
  FileText, 
  Presentation, 
  Database, 
  Download, 
  CheckCircle2, 
  AlertCircle, 
  Trash2, 
  Loader2, 
  Search, 
  Filter, 
  Sparkles, 
  FileCheck, 
  Eye, 
  RefreshCw,
  Layers,
  FileCode,
  Info
} from 'lucide-react';
import { extractStructuredFileContent, getFileExtension, ALL_ACCEPT } from '../utils/fileExtractor';
import { compareDocuments, exportDiffToExcel, generateAiDiffSummary } from '../utils/documentDiffEngine';

/** 파일 형식에 맞는 아이콘 반환 */
function getFormatIcon(ext) {
  const e = (ext || '').toLowerCase();
  if (e === 'pptx' || e === 'ppt') return <Presentation size={24} color="#f97316" />;
  if (e === 'xlsx' || e === 'xls' || e === 'csv') return <Database size={24} color="#10b981" />;
  if (e === 'hwpx' || e === 'hwp') return <FileCode size={24} color="#06b6d4" />;
  if (e === 'pdf') return <FileText size={24} color="#ef4444" />;
  return <FileText size={24} color="#3b82f6" />;
}

export default function DocumentCompareTab({ apiKey, llmProvider = 'gemini', omniRouteModel = 'auto' }) {
  // 파일 A (기준/원본) & 파일 B (비교/수정본) 상태
  const [fileA, setFileA] = useState(null);
  const [docDataA, setDocDataA] = useState(null);
  const [isExtractingA, setIsExtractingA] = useState(false);

  const [fileB, setFileB] = useState(null);
  const [docDataB, setDocDataB] = useState(null);
  const [isExtractingB, setIsExtractingB] = useState(false);

  // 비교 분석 상태
  const [isComparing, setIsComparing] = useState(false);
  const [compareStage, setCompareStage] = useState('');
  const [diffResult, setDiffResult] = useState(null);
  const [aiSummary, setAiSummary] = useState(null);
  const [enableAiSummary, setEnableAiSummary] = useState(true);

  // 결과 필터링 및 검색 상태
  const [filterType, setFilterType] = useState('all'); // 'all' | 'modified' | 'added' | 'deleted'
  const [searchQuery, setSearchQuery] = useState('');
  const [viewMode, setViewMode] = useState('table'); // 'table' | 'cards'

  const inputRefA = useRef(null);
  const inputRefB = useRef(null);

  // 파일 A 파싱 처리
  const handleSelectFileA = useCallback(async (file) => {
    if (!file) return;
    setFileA(file);
    setIsExtractingA(true);
    setDiffResult(null);
    setAiSummary(null);
    try {
      const data = await extractStructuredFileContent(file);
      setDocDataA(data);
    } catch (err) {
      alert(`[원본 문서 파싱 실패] ${err.message || '파일을 읽는 중 오류가 발생했습니다.'}`);
      setFileA(null);
      setDocDataA(null);
    } finally {
      setIsExtractingA(false);
    }
  }, []);

  // 파일 B 파싱 처리
  const handleSelectFileB = useCallback(async (file) => {
    if (!file) return;
    setFileB(file);
    setIsExtractingB(true);
    setDiffResult(null);
    setAiSummary(null);
    try {
      const data = await extractStructuredFileContent(file);
      setDocDataB(data);
    } catch (err) {
      alert(`[비교 문서 파싱 실패] ${err.message || '파일을 읽는 중 오류가 발생했습니다.'}`);
      setFileB(null);
      setDocDataB(null);
    } finally {
      setIsExtractingB(false);
    }
  }, []);

  // 드래그 앤 드롭 핸들러
  const handleDropA = (e) => {
    e.preventDefault();
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleSelectFileA(e.dataTransfer.files[0]);
    }
  };

  const handleDropB = (e) => {
    e.preventDefault();
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleSelectFileB(e.dataTransfer.files[0]);
    }
  };

  // 비교 분석 실행
  const handleRunCompare = async () => {
    if (!docDataA || !docDataB) {
      alert('비교할 원본 문서(A)와 비교 대상 문서(B)를 모두 업로드해 주세요.');
      return;
    }

    setIsComparing(true);
    setCompareStage('페이지별 단락 구조 매핑 및 텍스트 Diff 엔진 가동 중...');

    try {
      // 1단계: 브라우저 고성능 Diff 알고리즘 실행
      const result = compareDocuments(docDataA, docDataB);
      setDiffResult(result);

      // 2단계: AI 요약 생성 (선택 및 API키가 있을 때)
      if (enableAiSummary && apiKey && result.changes.length > 0) {
        setCompareStage('Gemini AI가 주요 변경 내역 요약 브리핑 리포트를 생성하는 중...');
        try {
          const summary = await generateAiDiffSummary(result, apiKey, llmProvider, omniRouteModel);
          setAiSummary(summary);
        } catch (aiErr) {
          console.warn('AI 요약 생성 생략:', aiErr);
        }
      }
    } catch (err) {
      console.error('문서 비교 중 에러:', err);
      alert(`문서 비교 분석 중 오류가 발생했습니다: ${err.message}`);
    } finally {
      setIsComparing(false);
      setCompareStage('');
    }
  };

  // 엑셀 내보내기 핸들러
  const handleExportExcel = async () => {
    if (!diffResult) return;
    try {
      await exportDiffToExcel(diffResult);
    } catch (err) {
      console.error('엑셀 내보내기 실패:', err);
      alert('엑셀 파일 생성 중 오류가 발생했습니다: ' + err.message);
    }
  };

  // 전체 초기화
  const handleResetAll = () => {
    setFileA(null);
    setDocDataA(null);
    setFileB(null);
    setDocDataB(null);
    setDiffResult(null);
    setAiSummary(null);
  };

  // 필터링된 변경 내역 목록
  const filteredChanges = (diffResult?.changes || []).filter((item) => {
    if (filterType !== 'all' && item.type !== filterType) return false;
    if (searchQuery.trim().length > 0) {
      const q = searchQuery.toLowerCase();
      const inA = (item.fileA.text || '').toLowerCase().includes(q) || (item.fileA.pageLabel || '').toLowerCase().includes(q);
      const inB = (item.fileB.text || '').toLowerCase().includes(q) || (item.fileB.pageLabel || '').toLowerCase().includes(q);
      const inSummary = (item.summary || '').toLowerCase().includes(q);
      return inA || inB || inSummary;
    }
    return true;
  });

  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, minHeight: 0, overflowY: 'auto', gap: '20px', padding: '4px' }}>
      
      {/* ── 상단 안내 헤더 ── */}
      <section className="glass-panel" style={{ padding: '20px 24px', borderLeft: '5px solid var(--accent-purple)' }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: '12px' }}>
          <div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
              <ArrowRightLeft size={22} color="var(--accent-purple)" />
              <h2 style={{ margin: 0, fontSize: '19px', fontWeight: 700, color: 'var(--text-primary)' }}>
                문서 버전 간 변경점 추적 및 심층 비교 (Diff Engine)
              </h2>
              <span style={{ fontSize: '11px', background: 'rgba(168, 85, 247, 0.25)', color: '#d8b4fe', padding: '3px 8px', borderRadius: '6px', fontWeight: 700 }}>
                PPT · 엑셀 · HWPX · PDF · Word 지원
              </span>
            </div>
            <p style={{ margin: '6px 0 0', fontSize: '13.5px', color: 'var(--text-secondary)', lineHeight: '1.5' }}>
              서로 다른 2개의 문서(이전 버전 vs 최신 수정본)를 비교하여 <strong>각 파일의 페이지 번호와 수정·추가·삭제된 내용</strong>을 정밀하게 추출하고 <strong>엑셀(.xlsx)로 다운로드</strong>합니다.
            </p>
          </div>

          {(fileA || fileB || diffResult) && (
            <button
              onClick={handleResetAll}
              className="interactive"
              style={{
                display: 'flex', alignItems: 'center', gap: '6px',
                padding: '8px 14px', borderRadius: '8px',
                background: 'rgba(255, 255, 255, 0.06)', border: '1px solid var(--glass-border)',
                color: 'var(--text-secondary)', fontSize: '12.5px', cursor: 'pointer'
              }}
            >
              <RefreshCw size={14} /> 전체 초기화
            </button>
          )}
        </div>
      </section>

      {/* ── 2개 문서 입력 섹션 (나란히 배치) ── */}
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
        
        {/* [A] 원본 / 이전 버전 문서 업로드 카드 */}
        <div 
          className="glass-panel" 
          onDragOver={(e) => e.preventDefault()}
          onDrop={handleDropA}
          style={{ 
            padding: '22px', 
            display: 'flex', 
            flexDirection: 'column', 
            gap: '14px',
            border: fileA ? '1px solid rgba(59, 130, 246, 0.4)' : '1px dashed rgba(255, 255, 255, 0.15)',
            background: fileA ? 'rgba(59, 130, 246, 0.04)' : 'rgba(255, 255, 255, 0.02)',
            transition: 'all 0.2s ease'
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <div style={{ 
                width: '26px', height: '26px', borderRadius: '7px', 
                background: 'rgba(59, 130, 246, 0.2)', color: 'var(--accent-blue)', 
                display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 800, fontSize: '13px' 
              }}>A</div>
              <span style={{ fontSize: '15px', fontWeight: 700, color: 'var(--text-primary)' }}>원본 / 이전 버전 문서</span>
            </div>
            {docDataA && (
              <span style={{ fontSize: '12px', color: 'var(--accent-blue)', fontWeight: 600 }}>
                {docDataA.pages}페이지 / {docDataA.totalItems || Object.keys(docDataA.pageMap || {}).length}섹션
              </span>
            )}
          </div>

          <input 
            type="file" 
            ref={inputRefA} 
            accept={ALL_ACCEPT} 
            style={{ display: 'none' }} 
            onChange={(e) => handleSelectFileA(e.target.files?.[0])}
          />

          {!fileA ? (
            <div 
              onClick={() => inputRefA.current?.click()}
              style={{
                flex: 1, minHeight: '130px', display: 'flex', flexDirection: 'column',
                alignItems: 'center', justifyContent: 'center', gap: '10px',
                cursor: 'pointer', borderRadius: '10px', background: 'rgba(0, 0, 0, 0.15)',
                padding: '20px', textAlign: 'center'
              }}
            >
              <Upload size={30} color="var(--accent-blue)" style={{ opacity: 0.8 }} />
              <div>
                <div style={{ fontSize: '13.5px', fontWeight: 600, color: 'var(--text-primary)' }}>
                  클릭하거나 파일을 드래그하여 업로드
                </div>
                <div style={{ fontSize: '12px', color: 'var(--text-muted)', marginTop: '4px' }}>
                  지원: PPTX, 엑셀(XLSX), HWPX, PDF, DOCX
                </div>
              </div>
            </div>
          ) : isExtractingA ? (
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '10px', padding: '30px', color: 'var(--accent-blue)' }}>
              <Loader2 size={20} className="animate-spin" />
              <span style={{ fontSize: '13px', fontWeight: 600 }}>문서 구조 및 단락 텍스트 분석 중...</span>
            </div>
          ) : (
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '12px 16px', background: 'rgba(0, 0, 0, 0.25)', borderRadius: '10px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '12px', minWidth: 0 }}>
                {getFormatIcon(docDataA?.ext)}
                <div style={{ minWidth: 0 }}>
                  <div style={{ fontSize: '13.5px', fontWeight: 600, color: 'var(--text-primary)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                    {fileA.name}
                  </div>
                  <div style={{ fontSize: '12px', color: 'var(--text-muted)', marginTop: '2px' }}>
                    {(fileA.size / 1024).toFixed(1)} KB · {docDataA?.ext?.toUpperCase()} 포맷
                  </div>
                </div>
              </div>

              <div style={{ display: 'flex', gap: '6px' }}>
                <button
                  onClick={() => inputRefA.current?.click()}
                  title="파일 변경"
                  style={{ padding: '6px 10px', background: 'rgba(255,255,255,0.08)', border: 'none', borderRadius: '6px', color: 'var(--text-secondary)', cursor: 'pointer', fontSize: '12px' }}
                >
                  변경
                </button>
                <button
                  onClick={() => { setFileA(null); setDocDataA(null); setDiffResult(null); }}
                  title="삭제"
                  style={{ padding: '6px 10px', background: 'rgba(239, 68, 68, 0.15)', border: 'none', borderRadius: '6px', color: 'var(--danger-color)', cursor: 'pointer' }}
                >
                  <Trash2 size={14} />
                </button>
              </div>
            </div>
          )}
        </div>

        {/* [B] 비교 / 최신 수정본 문서 업로드 카드 */}
        <div 
          className="glass-panel" 
          onDragOver={(e) => e.preventDefault()}
          onDrop={handleDropB}
          style={{ 
            padding: '22px', 
            display: 'flex', 
            flexDirection: 'column', 
            gap: '14px',
            border: fileB ? '1px solid rgba(168, 85, 247, 0.4)' : '1px dashed rgba(255, 255, 255, 0.15)',
            background: fileB ? 'rgba(168, 85, 247, 0.04)' : 'rgba(255, 255, 255, 0.02)',
            transition: 'all 0.2s ease'
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <div style={{ 
                width: '26px', height: '26px', borderRadius: '7px', 
                background: 'rgba(168, 85, 247, 0.2)', color: 'var(--accent-purple)', 
                display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 800, fontSize: '13px' 
              }}>B</div>
              <span style={{ fontSize: '15px', fontWeight: 700, color: 'var(--text-primary)' }}>비교 대상 / 최신 수정본 문서</span>
            </div>
            {docDataB && (
              <span style={{ fontSize: '12px', color: 'var(--accent-purple)', fontWeight: 600 }}>
                {docDataB.pages}페이지 / {docDataB.totalItems || Object.keys(docDataB.pageMap || {}).length}섹션
              </span>
            )}
          </div>

          <input 
            type="file" 
            ref={inputRefB} 
            accept={ALL_ACCEPT} 
            style={{ display: 'none' }} 
            onChange={(e) => handleSelectFileB(e.target.files?.[0])}
          />

          {!fileB ? (
            <div 
              onClick={() => inputRefB.current?.click()}
              style={{
                flex: 1, minHeight: '130px', display: 'flex', flexDirection: 'column',
                alignItems: 'center', justifyContent: 'center', gap: '10px',
                cursor: 'pointer', borderRadius: '10px', background: 'rgba(0, 0, 0, 0.15)',
                padding: '20px', textAlign: 'center'
              }}
            >
              <Upload size={30} color="var(--accent-purple)" style={{ opacity: 0.8 }} />
              <div>
                <div style={{ fontSize: '13.5px', fontWeight: 600, color: 'var(--text-primary)' }}>
                  클릭하거나 파일을 드래그하여 업로드
                </div>
                <div style={{ fontSize: '12px', color: 'var(--text-muted)', marginTop: '4px' }}>
                  동일 또는 이종 포맷 간 비교 가능 (예: HWPX vs PDF)
                </div>
              </div>
            </div>
          ) : isExtractingB ? (
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '10px', padding: '30px', color: 'var(--accent-purple)' }}>
              <Loader2 size={20} className="animate-spin" />
              <span style={{ fontSize: '13px', fontWeight: 600 }}>문서 구조 및 단락 텍스트 분석 중...</span>
            </div>
          ) : (
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '12px 16px', background: 'rgba(0, 0, 0, 0.25)', borderRadius: '10px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '12px', minWidth: 0 }}>
                {getFormatIcon(docDataB?.ext)}
                <div style={{ minWidth: 0 }}>
                  <div style={{ fontSize: '13.5px', fontWeight: 600, color: 'var(--text-primary)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                    {fileB.name}
                  </div>
                  <div style={{ fontSize: '12px', color: 'var(--text-muted)', marginTop: '2px' }}>
                    {(fileB.size / 1024).toFixed(1)} KB · {docDataB?.ext?.toUpperCase()} 포맷
                  </div>
                </div>
              </div>

              <div style={{ display: 'flex', gap: '6px' }}>
                <button
                  onClick={() => inputRefB.current?.click()}
                  title="파일 변경"
                  style={{ padding: '6px 10px', background: 'rgba(255,255,255,0.08)', border: 'none', borderRadius: '6px', color: 'var(--text-secondary)', cursor: 'pointer', fontSize: '12px' }}
                >
                  변경
                </button>
                <button
                  onClick={() => { setFileB(null); setDocDataB(null); setDiffResult(null); }}
                  title="삭제"
                  style={{ padding: '6px 10px', background: 'rgba(239, 68, 68, 0.15)', border: 'none', borderRadius: '6px', color: 'var(--danger-color)', cursor: 'pointer' }}
                >
                  <Trash2 size={14} />
                </button>
              </div>
            </div>
          )}
        </div>
      </div>

      {/* ── 비교 실행 버튼 및 옵션 ── */}
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '16px', margin: '6px 0', flexWrap: 'wrap' }}>
        <button
          onClick={handleRunCompare}
          disabled={!docDataA || !docDataB || isComparing}
          className="interactive"
          style={{
            padding: '13px 36px',
            fontSize: '15px',
            fontWeight: 700,
            borderRadius: '12px',
            border: 'none',
            background: (!docDataA || !docDataB || isComparing)
              ? 'rgba(255, 255, 255, 0.08)'
              : 'linear-gradient(135deg, #a855f7 0%, #3b82f6 100%)',
            color: (!docDataA || !docDataB || isComparing) ? 'var(--text-muted)' : '#ffffff',
            cursor: (!docDataA || !docDataB || isComparing) ? 'not-allowed' : 'pointer',
            display: 'flex',
            alignItems: 'center',
            gap: '10px',
            boxShadow: (!docDataA || !docDataB || isComparing) ? 'none' : '0 4px 18px rgba(168, 85, 247, 0.4)',
            transition: 'all 0.2s ease'
          }}
        >
          {isComparing ? (
            <>
              <Loader2 size={18} className="animate-spin" />
              <span>{compareStage || '비교 분석 진행 중...'}</span>
            </>
          ) : (
            <>
              <ArrowRightLeft size={18} />
              <span>두 문서 변경점 비교 및 추적 시작</span>
            </>
          )}
        </button>

        {apiKey && (
          <label style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px', color: 'var(--text-secondary)', cursor: 'pointer', userSelect: 'none' }}>
            <input 
              type="checkbox" 
              checked={enableAiSummary} 
              onChange={(e) => setEnableAiSummary(e.target.checked)} 
              style={{ accentColor: 'var(--accent-purple)' }}
            />
            <Sparkles size={14} color="var(--accent-purple)" />
            AI 주요 변경점 브리핑 보고서 함께 생성
          </label>
        )}
      </div>

      {/* ── 비교 결과 대시보드 ── */}
      {diffResult && (
        <section className="glass-panel animate-fade-in" style={{ padding: '24px', display: 'flex', flexDirection: 'column', gap: '20px' }}>
          
          {/* 1. 상단 통계 카드 & 엑셀 다운로드 버튼 */}
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: '16px' }}>
            <div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <CheckCircle2 size={20} color="var(--success-color)" />
                <h3 style={{ margin: 0, fontSize: '18px', fontWeight: 700, color: 'var(--text-primary)' }}>
                  비교 분석 완료: 총 {diffResult.stats.total}건의 변경 사항 감지
                </h3>
              </div>
              <p style={{ margin: '4px 0 0', fontSize: '13px', color: 'var(--text-secondary)' }}>
                [원본] {diffResult.docA.fileName} vs [비교본] {diffResult.docB.fileName}
              </p>
            </div>

            {/* 🔥 엑셀 다운로드 버튼 (사용자 핵심 요구사항) */}
            <button
              onClick={handleExportExcel}
              className="interactive"
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: '8px',
                padding: '11px 22px',
                fontSize: '14px',
                fontWeight: 700,
                borderRadius: '10px',
                border: 'none',
                background: 'linear-gradient(135deg, #10b981 0%, #059669 100%)',
                color: '#ffffff',
                cursor: 'pointer',
                boxShadow: '0 4px 14px rgba(16, 185, 129, 0.35)'
              }}
            >
              <Download size={18} />
              <span>📥 변경 내역 엑셀 다운로드 (.xlsx)</span>
            </button>
          </div>

          {/* 2. 통계 뱃지 바 */}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(130px, 1fr))', gap: '12px' }}>
            <div style={{ padding: '12px', background: 'rgba(255, 255, 255, 0.04)', borderRadius: '8px', border: '1px solid var(--glass-border)', textAlign: 'center' }}>
              <div style={{ fontSize: '11.5px', color: 'var(--text-secondary)' }}>총 변경 건수</div>
              <div style={{ fontSize: '20px', fontWeight: 800, color: 'var(--text-primary)', marginTop: '2px' }}>{diffResult.stats.total}건</div>
            </div>
            <div style={{ padding: '12px', background: 'rgba(245, 158, 11, 0.08)', borderRadius: '8px', border: '1px solid rgba(245, 158, 11, 0.25)', textAlign: 'center' }}>
              <div style={{ fontSize: '11.5px', color: 'var(--warning-color)' }}>✏️ 내용 수정</div>
              <div style={{ fontSize: '20px', fontWeight: 800, color: 'var(--warning-color)', marginTop: '2px' }}>{diffResult.stats.modified}건</div>
            </div>
            <div style={{ padding: '12px', background: 'rgba(16, 185, 129, 0.08)', borderRadius: '8px', border: '1px solid rgba(16, 185, 129, 0.25)', textAlign: 'center' }}>
              <div style={{ fontSize: '11.5px', color: 'var(--success-color)' }}>➕ 신규 추가</div>
              <div style={{ fontSize: '20px', fontWeight: 800, color: 'var(--success-color)', marginTop: '2px' }}>{diffResult.stats.added}건</div>
            </div>
            <div style={{ padding: '12px', background: 'rgba(239, 68, 68, 0.08)', borderRadius: '8px', border: '1px solid rgba(239, 68, 68, 0.25)', textAlign: 'center' }}>
              <div style={{ fontSize: '11.5px', color: 'var(--danger-color)' }}>🗑️ 내용 삭제</div>
              <div style={{ fontSize: '20px', fontWeight: 800, color: 'var(--danger-color)', marginTop: '2px' }}>{diffResult.stats.deleted}건</div>
            </div>
            <div style={{ padding: '12px', background: 'rgba(59, 130, 246, 0.08)', borderRadius: '8px', border: '1px solid rgba(59, 130, 246, 0.25)', textAlign: 'center' }}>
              <div style={{ fontSize: '11.5px', color: 'var(--accent-blue)' }}>영향 페이지 (원본)</div>
              <div style={{ fontSize: '20px', fontWeight: 800, color: 'var(--accent-blue)', marginTop: '2px' }}>{diffResult.stats.affectedPagesA}개</div>
            </div>
            <div style={{ padding: '12px', background: 'rgba(168, 85, 247, 0.08)', borderRadius: '8px', border: '1px solid rgba(168, 85, 247, 0.25)', textAlign: 'center' }}>
              <div style={{ fontSize: '11.5px', color: 'var(--accent-purple)' }}>영향 페이지 (비교본)</div>
              <div style={{ fontSize: '20px', fontWeight: 800, color: 'var(--accent-purple)', marginTop: '2px' }}>{diffResult.stats.affectedPagesB}개</div>
            </div>
          </div>

          {/* 3. AI 주요 변경점 브리핑 보고서 */}
          {aiSummary && (
            <div style={{ padding: '16px 20px', background: 'rgba(168, 85, 247, 0.08)', borderRadius: '10px', border: '1px solid rgba(168, 85, 247, 0.25)' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '8px' }}>
                <Sparkles size={16} color="var(--accent-purple)" />
                <span style={{ fontSize: '13.5px', fontWeight: 700, color: 'var(--accent-purple)' }}>AI 주요 변경점 브리핑 리포트</span>
              </div>
              <div style={{ fontSize: '13.5px', lineHeight: '1.6', color: 'var(--text-primary)', whiteSpace: 'pre-wrap' }}>
                {aiSummary}
              </div>
            </div>
          )}

          {/* 4. 필터 및 검색 컨트롤 */}
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: '12px', borderBottom: '1px solid var(--glass-border)', paddingBottom: '14px' }}>
            <div style={{ display: 'flex', gap: '6px', alignItems: 'center' }}>
              <button
                onClick={() => setFilterType('all')}
                style={{
                  padding: '6px 12px', borderRadius: '7px', fontSize: '12.5px', fontWeight: filterType === 'all' ? 700 : 500,
                  background: filterType === 'all' ? 'rgba(255,255,255,0.15)' : 'transparent',
                  color: filterType === 'all' ? '#fff' : 'var(--text-secondary)',
                  border: '1px solid var(--glass-border)', cursor: 'pointer'
                }}
              >
                전체 ({diffResult.stats.total})
              </button>
              <button
                onClick={() => setFilterType('modified')}
                style={{
                  padding: '6px 12px', borderRadius: '7px', fontSize: '12.5px', fontWeight: filterType === 'modified' ? 700 : 500,
                  background: filterType === 'modified' ? 'rgba(245, 158, 11, 0.2)' : 'transparent',
                  color: filterType === 'modified' ? 'var(--warning-color)' : 'var(--text-secondary)',
                  border: '1px solid var(--glass-border)', cursor: 'pointer'
                }}
              >
                수정 ({diffResult.stats.modified})
              </button>
              <button
                onClick={() => setFilterType('added')}
                style={{
                  padding: '6px 12px', borderRadius: '7px', fontSize: '12.5px', fontWeight: filterType === 'added' ? 700 : 500,
                  background: filterType === 'added' ? 'rgba(16, 185, 129, 0.2)' : 'transparent',
                  color: filterType === 'added' ? 'var(--success-color)' : 'var(--text-secondary)',
                  border: '1px solid var(--glass-border)', cursor: 'pointer'
                }}
              >
                추가 ({diffResult.stats.added})
              </button>
              <button
                onClick={() => setFilterType('deleted')}
                style={{
                  padding: '6px 12px', borderRadius: '7px', fontSize: '12.5px', fontWeight: filterType === 'deleted' ? 700 : 500,
                  background: filterType === 'deleted' ? 'rgba(239, 68, 68, 0.2)' : 'transparent',
                  color: filterType === 'deleted' ? 'var(--danger-color)' : 'var(--text-secondary)',
                  border: '1px solid var(--glass-border)', cursor: 'pointer'
                }}
              >
                삭제 ({diffResult.stats.deleted})
              </button>
            </div>

            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <div style={{ position: 'relative', display: 'flex', alignItems: 'center' }}>
                <Search size={14} color="var(--text-muted)" style={{ position: 'absolute', left: '10px' }} />
                <input
                  type="text"
                  placeholder="페이지 번호 또는 내용 검색..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  style={{
                    padding: '6px 12px 6px 30px',
                    borderRadius: '8px',
                    background: 'rgba(0, 0, 0, 0.2)',
                    border: '1px solid var(--glass-border)',
                    color: 'var(--text-primary)',
                    fontSize: '12.5px',
                    width: '220px'
                  }}
                />
              </div>
            </div>
          </div>

          {/* 5. 변경 내역 상세 테이블 */}
          {filteredChanges.length === 0 ? (
            <div style={{ padding: '40px', textAlign: 'center', color: 'var(--text-muted)' }}>
              {diffResult.stats.total === 0 ? (
                <div>
                  <CheckCircle2 size={36} color="var(--success-color)" style={{ marginBottom: '10px' }} />
                  <div style={{ fontSize: '16px', fontWeight: 700, color: 'var(--text-primary)' }}>텍스트 변경 사항이 없습니다.</div>
                  <div style={{ fontSize: '13px', marginTop: '4px' }}>두 문서의 내용이 100% 동일합니다.</div>
                </div>
              ) : (
                <div>조건에 부합하는 변경 내역이 없습니다.</div>
              )}
            </div>
          ) : (
            <div style={{ overflowX: 'auto', borderRadius: '8px', border: '1px solid var(--glass-border)' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '13px' }}>
                <thead>
                  <tr style={{ background: 'rgba(0, 0, 0, 0.3)', borderBottom: '1px solid var(--glass-border)', color: 'var(--text-secondary)' }}>
                    <th style={{ padding: '12px 14px', width: '55px', textAlign: 'center' }}>순번</th>
                    <th style={{ padding: '12px 14px', width: '85px', textAlign: 'center' }}>구분</th>
                    <th style={{ padding: '12px 14px', width: '130px' }}>원본(A) 위치</th>
                    <th style={{ padding: '12px 14px', width: '28%' }}>원본(A) 내용</th>
                    <th style={{ padding: '12px 14px', width: '130px' }}>비교본(B) 위치</th>
                    <th style={{ padding: '12px 14px', width: '28%' }}>비교본(B) 변경 내용</th>
                    <th style={{ padding: '12px 14px' }}>인라인 변경 하이라이트</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredChanges.map((change) => {
                    const isMod = change.type === 'modified';
                    const isAdd = change.type === 'added';
                    const isDel = change.type === 'deleted';

                    return (
                      <tr 
                        key={change.id}
                        style={{
                          borderBottom: '1px solid rgba(255, 255, 255, 0.05)',
                          background: isMod 
                            ? 'rgba(245, 158, 11, 0.02)' 
                            : (isAdd ? 'rgba(16, 185, 129, 0.02)' : 'rgba(239, 68, 68, 0.02)'),
                          verticalAlign: 'top'
                        }}
                      >
                        {/* 순번 */}
                        <td style={{ padding: '12px 14px', textAlign: 'center', color: 'var(--text-muted)', fontWeight: 600 }}>
                          {change.index}
                        </td>

                        {/* 구분 배지 */}
                        <td style={{ padding: '12px 14px', textAlign: 'center' }}>
                          <span style={{
                            display: 'inline-block',
                            padding: '3px 8px',
                            borderRadius: '5px',
                            fontSize: '11.5px',
                            fontWeight: 700,
                            background: isMod ? 'rgba(245, 158, 11, 0.15)' : (isAdd ? 'rgba(16, 185, 129, 0.15)' : 'rgba(239, 68, 68, 0.15)'),
                            color: isMod ? 'var(--warning-color)' : (isAdd ? 'var(--success-color)' : 'var(--danger-color)'),
                            border: `1px solid ${isMod ? 'rgba(245, 158, 11, 0.3)' : (isAdd ? 'rgba(16, 185, 129, 0.3)' : 'rgba(239, 68, 68, 0.3)')}`
                          }}>
                            {change.typeLabel}
                          </span>
                        </td>

                        {/* 원본 위치 */}
                        <td style={{ padding: '12px 14px', fontWeight: 600, color: 'var(--accent-blue)', whiteSpace: 'nowrap' }}>
                          {change.fileA.pageLabel}
                        </td>

                        {/* 원본 내용 */}
                        <td style={{ padding: '12px 14px', color: isDel ? 'var(--danger-color)' : 'var(--text-primary)', lineHeight: '1.5' }}>
                          {isDel ? (
                            <span style={{ textDecoration: 'line-through', opacity: 0.85 }}>{change.fileA.text}</span>
                          ) : (
                            change.fileA.text
                          )}
                        </td>

                        {/* 비교본 위치 */}
                        <td style={{ padding: '12px 14px', fontWeight: 600, color: 'var(--accent-purple)', whiteSpace: 'nowrap' }}>
                          {change.fileB.pageLabel}
                        </td>

                        {/* 비교본 내용 */}
                        <td style={{ padding: '12px 14px', color: isAdd ? 'var(--success-color)' : 'var(--text-primary)', lineHeight: '1.5' }}>
                          {isAdd ? (
                            <span style={{ fontWeight: 600 }}>{change.fileB.text}</span>
                          ) : (
                            change.fileB.text
                          )}
                        </td>

                        {/* 세부 인라인 하이라이트 */}
                        <td style={{ padding: '12px 14px', fontSize: '12.5px', lineHeight: '1.6' }}>
                          {isMod && change.inlineDiff ? (
                            <div style={{ wordBreak: 'break-word' }}>
                              {change.inlineDiff.map((tok, tIdx) => {
                                if (tok.type === 'delete') {
                                  return (
                                    <span key={tIdx} style={{ background: 'rgba(239, 68, 68, 0.25)', color: '#fca5a5', textDecoration: 'line-through', padding: '1px 3px', borderRadius: '3px', marginRight: '2px' }}>
                                      {tok.text}
                                    </span>
                                  );
                                }
                                if (tok.type === 'insert') {
                                  return (
                                    <span key={tIdx} style={{ background: 'rgba(16, 185, 129, 0.25)', color: '#86efac', fontWeight: 600, padding: '1px 3px', borderRadius: '3px', marginRight: '2px' }}>
                                      {tok.text}
                                    </span>
                                  );
                                }
                                return <span key={tIdx} style={{ color: 'var(--text-muted)' }}>{tok.text}</span>;
                              })}
                            </div>
                          ) : (
                            <span style={{ color: 'var(--text-muted)', fontStyle: 'italic' }}>
                              {change.summary}
                            </span>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </section>
      )}

    </div>
  );
}
