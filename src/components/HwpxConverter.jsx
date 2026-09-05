import React, { useState, useRef, useEffect } from 'react';
import { 
  FileText, 
  Upload, 
  RefreshCw, 
  CheckCircle2, 
  Download, 
  Trash2, 
  FileArchive, 
  AlertCircle, 
  Sparkles, 
  Eye, 
  FileSpreadsheet, 
  FolderArchive,
  ArrowRight,
  ShieldCheck,
  Zap,
  Info,
  Layers,
  ChevronDown,
  ChevronUp
} from 'lucide-react';
import { 
  convertSingleHwpToHwpx, 
  convertBatchHwpToHwpx, 
  downloadAllAsZip,
  saveBlobAs
} from '../utils/hwpToHwpxConverter.js';
import HwpxGenerator from './HwpxGenerator.jsx';

export default function HwpxConverter({ apiKey, llmProvider = 'gemini', omniRouteModel = 'auto' }) {
  // 메인 모드 선택 ('convert' : HWP 일괄 변환 | 'ai_report' : AI 표준보고서 생성)
  const [activeSubMode, setActiveSubMode] = useState('convert');

  // 업로드된 파일 목록 [{ id, file, status: 'idle'|'processing'|'done'|'error', result, error, progress }]
  const [fileList, setFileList] = useState([]);
  const [isConverting, setIsConverting] = useState(false);
  const [overallProgress, setOverallProgress] = useState({ currentIndex: 0, total: 0, percent: 0, currentFileName: '' });
  const [isDragActive, setIsDragActive] = useState(false);
  const [selectedPreview, setSelectedPreview] = useState(null);
  const [isZipping, setIsZipping] = useState(false);

  const fileInputRef = useRef(null);

  // 파일 크기 포맷터
  const formatBytes = (bytes, decimals = 2) => {
    if (!bytes || bytes === 0) return '0 Bytes';
    const k = 1024;
    const dm = decimals < 0 ? 0 : decimals;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
  };

  // 파일 추가 핸들러
  const handleAddFiles = (selectedFiles) => {
    const rawFiles = Array.from(selectedFiles);
    const validHwpFiles = rawFiles.filter(f => f.name.toLowerCase().endsWith('.hwp'));

    if (validHwpFiles.length === 0) {
      alert('한글(.hwp) 확장자 파일만 업로드 가능합니다.\n(HWPX 파일은 이미 변환된 파일입니다)');
      return;
    }

    if (validHwpFiles.length !== rawFiles.length) {
      alert(`선택된 파일 중 .hwp 형식이 아닌 ${rawFiles.length - validHwpFiles.length}개 파일은 제외되었습니다.`);
    }

    setFileList(prev => {
      const newItems = validHwpFiles.filter(f => 
        !prev.some(p => p.file.name === f.name && p.file.size === f.size)
      ).map((f, idx) => ({
        id: `hwp_${Date.now()}_${idx}_${Math.random().toString(36).substr(2, 5)}`,
        file: f,
        name: f.name,
        size: f.size,
        status: 'idle', // 'idle' | 'processing' | 'done' | 'error'
        progress: 0,
        result: null,
        error: null
      }));

      return [...prev, ...newItems];
    });
  };

  // 드래그 앤 드롭 핸들러
  const handleDrag = (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === 'dragenter' || e.type === 'dragover') {
      setIsDragActive(true);
    } else if (e.type === 'dragleave') {
      setIsDragActive(false);
    }
  };

  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleAddFiles(e.dataTransfer.files);
    }
  };

  // 특정 파일 제거
  const handleRemoveFile = (id) => {
    setFileList(prev => prev.filter(item => item.id !== id));
  };

  // 전체 목록 초기화
  const handleClearAll = () => {
    if (isConverting) return;
    if (fileList.length === 0) return;
    if (window.confirm('업로드된 파일 목록을 모두 비우시겠습니까?')) {
      setFileList([]);
      setOverallProgress({ currentIndex: 0, total: 0, percent: 0, currentFileName: '' });
      setSelectedPreview(null);
    }
  };

  // 복수 파일 일괄 변환 시작
  const handleStartBatchConvert = async () => {
    const idleOrErrorItems = fileList.filter(item => item.status === 'idle' || item.status === 'error');
    if (idleOrErrorItems.length === 0) {
      alert('변환할 대기 파일이 없습니다. 새로운 .hwp 파일을 추가해 주세요.');
      return;
    }

    setIsConverting(true);
    const targetTotal = idleOrErrorItems.length;

    for (let i = 0; i < idleOrErrorItems.length; i++) {
      const currentItem = idleOrErrorItems[i];

      // 상태를 'processing'으로 갱신
      setFileList(prev => prev.map(item => 
        item.id === currentItem.id 
          ? { ...item, status: 'processing', progress: 20 } 
          : item
      ));

      setOverallProgress({
        currentIndex: i + 1,
        total: targetTotal,
        percent: Math.round(((i) / targetTotal) * 100),
        currentFileName: currentItem.name
      });

      try {
        const res = await convertSingleHwpToHwpx(currentItem.file, (p) => {
          setFileList(prev => prev.map(item => 
            item.id === currentItem.id 
              ? { ...item, progress: p.percent || 50 } 
              : item
          ));
        });

        // 성공 상태 반영
        setFileList(prev => prev.map(item => 
          item.id === currentItem.id 
            ? { 
                ...item, 
                status: 'done', 
                progress: 100, 
                result: res, 
                error: null 
              } 
            : item
        ));
      } catch (err) {
        console.error(`변환 에러 (${currentItem.name}):`, err);
        // 에러 상태 반영
        setFileList(prev => prev.map(item => 
          item.id === currentItem.id 
            ? { 
                ...item, 
                status: 'error', 
                progress: 0, 
                error: err.message || 'HWP 파일 파싱 실패' 
              } 
            : item
        ));
      }
    }

    setOverallProgress({
      currentIndex: targetTotal,
      total: targetTotal,
      percent: 100,
      currentFileName: '전체 파일 변환 작업이 완료되었습니다.'
    });
    setIsConverting(false);
  };

  // 단일 HWPX 파일 다운로드
  const handleDownloadSingle = (item) => {
    if (!item.result || !item.result.blob) {
      alert('변환된 파일 데이터가 존재하지 않습니다.');
      return;
    }
    saveBlobAs(item.result.blob, item.result.outputName);
  };

  // 전체 HWPX 일괄 ZIP 다운로드
  const handleDownloadAllZip = async () => {
    const doneItems = fileList.filter(item => item.status === 'done' && item.result && item.result.blob);
    if (doneItems.length === 0) {
      alert('다운로드할 수 있는 변환 완료 파일이 없습니다.');
      return;
    }

    try {
      setIsZipping(true);
      const resultsToZip = doneItems.map(item => item.result);
      const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
      await downloadAllAsZip(resultsToZip, `HWPX_일괄변환_${today}.zip`);
    } catch (err) {
      alert(`ZIP 압축 다운로드 중 오류 발생: ${err.message}`);
    } finally {
      setIsZipping(false);
    }
  };

  // 통계 계산
  const totalUploaded = fileList.length;
  const doneCount = fileList.filter(f => f.status === 'done').length;
  const errorCount = fileList.filter(f => f.status === 'error').length;
  const idleCount = fileList.filter(f => f.status === 'idle').length;

  return (
    <div className="tab-container animate-fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
      
      {/* 서브 모드 전환 탭 */}
      <div style={{
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        background: 'rgba(255, 255, 255, 0.03)',
        padding: '6px 10px',
        borderRadius: '12px',
        border: '1px solid var(--glass-border)'
      }}>
        <div style={{ display: 'flex', gap: '8px' }}>
          <button
            onClick={() => setActiveSubMode('convert')}
            style={{
              display: 'flex',
              alignItems: 'center',
              gap: '8px',
              padding: '8px 16px',
              borderRadius: '8px',
              fontSize: '13px',
              fontWeight: 700,
              cursor: 'pointer',
              border: activeSubMode === 'convert' ? '1px solid var(--accent-blue)' : '1px solid transparent',
              background: activeSubMode === 'convert' ? 'rgba(59, 130, 246, 0.15)' : 'transparent',
              color: activeSubMode === 'convert' ? 'var(--accent-blue)' : 'var(--text-secondary)',
              transition: 'all 0.2s ease'
            }}
          >
            <Zap size={16} />
            <span>HWP ➔ HWPX 복수 파일 일괄 변환</span>
            {doneCount > 0 && (
              <span style={{
                background: 'var(--success-color)',
                color: '#fff',
                fontSize: '11px',
                padding: '1px 6px',
                borderRadius: '10px'
              }}>
                {doneCount}
              </span>
            )}
          </button>

          <button
            onClick={() => setActiveSubMode('ai_report')}
            style={{
              display: 'flex',
              alignItems: 'center',
              gap: '8px',
              padding: '8px 16px',
              borderRadius: '8px',
              fontSize: '13px',
              fontWeight: 700,
              cursor: 'pointer',
              border: activeSubMode === 'ai_report' ? '1px solid #10b981' : '1px solid transparent',
              background: activeSubMode === 'ai_report' ? 'rgba(16, 185, 129, 0.15)' : 'transparent',
              color: activeSubMode === 'ai_report' ? '#10b981' : 'var(--text-secondary)',
              transition: 'all 0.2s ease'
            }}
          >
            <Sparkles size={16} />
            <span>AI 표준 HWPX 보고서 자동 생성</span>
          </button>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '12px', color: 'var(--text-muted)' }}>
          <ShieldCheck size={14} color="var(--success-color)" />
          <span>한컴 표준 OWPML 100% 호환</span>
        </div>
      </div>

      {/* 모드 1: HWP -> HWPX 일괄 변환기 */}
      {activeSubMode === 'convert' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
          
          {/* 상단 안내 & 드롭존 카드 */}
          <div className="glass-panel" style={{ padding: '24px', borderRadius: '16px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '16px' }}>
              <div>
                <h3 style={{ margin: '0 0 6px', fontSize: '17px', fontWeight: 800, color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <FileText size={20} color="var(--accent-blue)" />
                  한글 문서(.hwp) ➔ 개방형 문서(.hwpx) 자동 변환
                </h3>
                <p style={{ margin: 0, fontSize: '13px', color: 'var(--text-secondary)', lineHeight: 1.5 }}>
                  수동으로 하나씩 변환하기 힘든 <strong>복수의 HWP 파일들을 일괄 업로드</strong>하여, 한컴오피스 최신 규격인 <strong>HWPX(ZIP+XML) 파일로 즉시 변환</strong>합니다.
                </p>
              </div>

              {fileList.length > 0 && (
                <button
                  onClick={handleClearAll}
                  disabled={isConverting}
                  style={{
                    background: 'rgba(239, 68, 68, 0.1)',
                    border: '1px solid rgba(239, 68, 68, 0.3)',
                    color: '#f87171',
                    padding: '6px 12px',
                    borderRadius: '8px',
                    fontSize: '12px',
                    cursor: isConverting ? 'not-allowed' : 'pointer',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '4px'
                  }}
                >
                  <Trash2 size={14} /> 목록 전체 비우기
                </button>
              )}
            </div>

            {/* 드래그 앤 드롭 영역 */}
            <div
              onDragEnter={handleDrag}
              onDragOver={handleDrag}
              onDragLeave={handleDrag}
              onDrop={handleDrop}
              onClick={() => fileInputRef.current && fileInputRef.current.click()}
              style={{
                border: `2px dashed ${isDragActive ? 'var(--accent-blue)' : 'var(--glass-border)'}`,
                background: isDragActive ? 'rgba(59, 130, 246, 0.08)' : 'rgba(255, 255, 255, 0.02)',
                borderRadius: '14px',
                padding: '36px 20px',
                textAlign: 'center',
                cursor: 'pointer',
                transition: 'all 0.2s ease',
                display: 'flex',
                flexDirection: 'column',
                alignItems: 'center',
                justifyContent: 'center',
                gap: '12px'
              }}
            >
              <input
                ref={fileInputRef}
                type="file"
                multiple
                accept=".hwp"
                onChange={(e) => {
                  if (e.target.files && e.target.files.length > 0) {
                    handleAddFiles(e.target.files);
                  }
                  e.target.value = ''; // 재선택 가능하게 리셋
                }}
                style={{ display: 'none' }}
              />

              <div style={{
                width: '56px',
                height: '56px',
                borderRadius: '16px',
                background: 'linear-gradient(135deg, rgba(59, 130, 246, 0.2), rgba(99, 102, 241, 0.2))',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                color: 'var(--accent-blue)',
                boxShadow: '0 8px 16px rgba(0,0,0,0.2)'
              }}>
                <Upload size={28} />
              </div>

              <div>
                <p style={{ margin: '0 0 4px', fontSize: '15px', fontWeight: 700, color: 'var(--text-primary)' }}>
                  변환할 .hwp 파일들을 여기에 끌어다 놓으세요 (복수 선택 지원)
                </p>
                <p style={{ margin: 0, fontSize: '12px', color: 'var(--text-muted)' }}>
                  또는 클릭하여 파일 탐색기에서 여러 개의 .hwp 파일을 한 번에 선택할 수 있습니다.
                </p>
              </div>

              <div style={{
                display: 'inline-flex',
                alignItems: 'center',
                gap: '6px',
                background: 'rgba(59, 130, 246, 0.1)',
                border: '1px solid rgba(59, 130, 246, 0.25)',
                padding: '4px 12px',
                borderRadius: '20px',
                fontSize: '11px',
                color: '#60a5fa',
                fontWeight: 600
              }}>
                <Sparkles size={12} /> OLE5 바이너리 스트림 정밀 해독 및 표준 OWPML 패키징
              </div>
            </div>

            {/* 작업 제어 및 상태 통계 바 */}
            {fileList.length > 0 && (
              <div style={{
                marginTop: '20px',
                display: 'flex',
                flexWrap: 'wrap',
                alignItems: 'center',
                justifyContent: 'space-between',
                gap: '12px',
                padding: '14px 18px',
                background: 'rgba(0, 0, 0, 0.25)',
                borderRadius: '12px',
                border: '1px solid var(--glass-border)'
              }}>
                {/* 좌측: 통계 배지 */}
                <div style={{ display: 'flex', alignItems: 'center', gap: '12px', flexWrap: 'wrap' }}>
                  <span style={{ fontSize: '13px', fontWeight: 700, color: 'var(--text-primary)' }}>
                    총 {totalUploaded}개 파일
                  </span>
                  {doneCount > 0 && (
                    <span style={{ fontSize: '12px', color: 'var(--success-color)', display: 'flex', alignItems: 'center', gap: '4px', fontWeight: 600 }}>
                      <CheckCircle2 size={14} /> 완료 {doneCount}개
                    </span>
                  )}
                  {idleCount > 0 && (
                    <span style={{ fontSize: '12px', color: 'var(--warning-color)', fontWeight: 600 }}>
                      대기 {idleCount}개
                    </span>
                  )}
                  {errorCount > 0 && (
                    <span style={{ fontSize: '12px', color: 'var(--danger-color)', display: 'flex', alignItems: 'center', gap: '4px', fontWeight: 600 }}>
                      <AlertCircle size={14} /> 오류 {errorCount}개
                    </span>
                  )}
                </div>

                {/* 우측: 액션 버튼들 */}
                <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                  {doneCount > 0 && (
                    <button
                      onClick={handleDownloadAllZip}
                      disabled={isZipping || isConverting}
                      style={{
                        background: 'linear-gradient(135deg, #10b981, #059669)',
                        color: '#ffffff',
                        border: 'none',
                        padding: '9px 16px',
                        borderRadius: '8px',
                        fontSize: '13px',
                        fontWeight: 700,
                        cursor: (isZipping || isConverting) ? 'not-allowed' : 'pointer',
                        display: 'flex',
                        alignItems: 'center',
                        gap: '6px',
                        boxShadow: '0 4px 12px rgba(16, 185, 129, 0.3)'
                      }}
                    >
                      <FolderArchive size={16} />
                      {isZipping ? 'ZIP 묶는 중...' : `전체 HWPX 다운로드 (.ZIP ${doneCount}건)`}
                    </button>
                  )}

                  <button
                    onClick={handleStartBatchConvert}
                    disabled={isConverting || idleCount === 0}
                    style={{
                      background: (isConverting || idleCount === 0)
                        ? 'rgba(255, 255, 255, 0.05)'
                        : 'linear-gradient(135deg, var(--accent-blue), #6366f1)',
                      color: (isConverting || idleCount === 0) ? 'var(--text-muted)' : '#ffffff',
                      border: 'none',
                      padding: '9px 20px',
                      borderRadius: '8px',
                      fontSize: '13px',
                      fontWeight: 700,
                      cursor: (isConverting || idleCount === 0) ? 'not-allowed' : 'pointer',
                      display: 'flex',
                      alignItems: 'center',
                      gap: '8px',
                      boxShadow: (isConverting || idleCount === 0) ? 'none' : '0 4px 14px rgba(59, 130, 246, 0.4)'
                    }}
                  >
                    {isConverting ? (
                      <>
                        <RefreshCw size={16} className="spin" />
                        <span>변환 진행 중... ({overallProgress.currentIndex}/{overallProgress.total})</span>
                      </>
                    ) : (
                      <>
                        <Zap size={16} />
                        <span>{doneCount > 0 ? '남은 파일 일괄 변환' : '일괄 변환 시작'}</span>
                      </>
                    )}
                  </button>
                </div>
              </div>
            )}

            {/* 전체 진행률 게이지 바 */}
            {isConverting && (
              <div style={{ marginTop: '14px' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '12px', marginBottom: '6px', color: 'var(--text-secondary)' }}>
                  <span>{overallProgress.currentFileName}</span>
                  <span style={{ fontWeight: 700, color: 'var(--accent-blue)' }}>{overallProgress.percent}%</span>
                </div>
                <div style={{ width: '100%', height: '6px', background: 'rgba(255,255,255,0.06)', borderRadius: '4px', overflow: 'hidden' }}>
                  <div 
                    style={{ 
                      width: `${overallProgress.percent}%`, 
                      height: '100%', 
                      background: 'linear-gradient(to right, var(--accent-blue), #10b981)', 
                      transition: 'width 0.3s ease' 
                    }} 
                  />
                </div>
              </div>
            )}
          </div>

          {/* 파일 리스트 테이블 */}
          {fileList.length > 0 && (
            <div className="glass-panel" style={{ padding: '20px', borderRadius: '16px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '14px' }}>
                <h4 style={{ margin: 0, fontSize: '15px', fontWeight: 700, color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <Layers size={18} color="var(--accent-blue)" />
                  변환 대상 목록 ({fileList.length}건)
                </h4>
              </div>

              <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                {fileList.map((item, idx) => (
                  <div
                    key={item.id}
                    style={{
                      background: item.status === 'done' 
                        ? 'rgba(16, 185, 129, 0.04)' 
                        : item.status === 'error'
                        ? 'rgba(239, 68, 68, 0.05)'
                        : 'rgba(255, 255, 255, 0.02)',
                      border: `1px solid ${
                        item.status === 'done' 
                          ? 'rgba(16, 185, 129, 0.25)' 
                          : item.status === 'error'
                          ? 'rgba(239, 68, 68, 0.3)'
                          : 'var(--glass-border)'
                      }`,
                      borderRadius: '12px',
                      padding: '14px 16px',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'space-between',
                      gap: '16px',
                      transition: 'all 0.2s ease'
                    }}
                  >
                    {/* 좌측: 파일 정보 */}
                    <div style={{ display: 'flex', alignItems: 'center', gap: '12px', flex: 1, minWidth: 0 }}>
                      <div style={{
                        width: '38px',
                        height: '38px',
                        borderRadius: '10px',
                        background: item.status === 'done'
                          ? 'rgba(16, 185, 129, 0.15)'
                          : item.status === 'error'
                          ? 'rgba(239, 68, 68, 0.15)'
                          : 'rgba(59, 130, 246, 0.1)',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        color: item.status === 'done'
                          ? '#10b981'
                          : item.status === 'error'
                          ? '#ef4444'
                          : 'var(--accent-blue)',
                        flexShrink: 0
                      }}>
                        <FileText size={20} />
                      </div>

                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '2px' }}>
                          <span style={{ fontSize: '13.5px', fontWeight: 700, color: 'var(--text-primary)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                            {item.name}
                          </span>
                          <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>
                            ({formatBytes(item.size)})
                          </span>
                        </div>

                        {item.status === 'done' && item.result && (
                          <div style={{ display: 'flex', alignItems: 'center', gap: '10px', fontSize: '11.5px', color: 'var(--text-secondary)' }}>
                            <span style={{ color: '#10b981', fontWeight: 600 }}>
                              ➔ {item.result.outputName} ({formatBytes(item.result.outputSize)})
                            </span>
                            <span>· 문단 {item.result.paragraphsCount}개</span>
                            {item.result.tablesCount > 0 && <span>· 표 {item.result.tablesCount}개</span>}
                          </div>
                        )}

                        {item.status === 'processing' && (
                          <div style={{ fontSize: '11.5px', color: 'var(--accent-blue)', display: 'flex', alignItems: 'center', gap: '6px' }}>
                            <RefreshCw size={12} className="spin" />
                            <span>바이너리 스트림 분석 및 HWPX 빌드 중...</span>
                          </div>
                        )}

                        {item.status === 'error' && (
                          <div style={{ fontSize: '11.5px', color: '#f87171' }}>
                            ❌ 변환 실패: {item.error}
                          </div>
                        )}

                        {item.status === 'idle' && (
                          <div style={{ fontSize: '11.5px', color: 'var(--text-muted)' }}>
                            변환 대기 중
                          </div>
                        )}
                      </div>
                    </div>

                    {/* 우측: 버튼들 */}
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexShrink: 0 }}>
                      {item.status === 'done' && item.result && (
                        <>
                          <button
                            onClick={() => setSelectedPreview(selectedPreview === item.id ? null : item.id)}
                            style={{
                              background: 'rgba(255, 255, 255, 0.05)',
                              border: '1px solid var(--glass-border)',
                              color: 'var(--text-secondary)',
                              padding: '6px 10px',
                              borderRadius: '7px',
                              fontSize: '12px',
                              cursor: 'pointer',
                              display: 'flex',
                              alignItems: 'center',
                              gap: '4px'
                            }}
                            title="추출 텍스트 미리보기"
                          >
                            <Eye size={14} />
                            <span>미리보기</span>
                            {selectedPreview === item.id ? <ChevronUp size={12} /> : <ChevronDown size={12} />}
                          </button>

                          <button
                            onClick={() => handleDownloadSingle(item)}
                            style={{
                              background: 'rgba(16, 185, 129, 0.15)',
                              border: '1px solid rgba(16, 185, 129, 0.4)',
                              color: '#10b981',
                              padding: '6px 12px',
                              borderRadius: '7px',
                              fontSize: '12px',
                              fontWeight: 700,
                              cursor: 'pointer',
                              display: 'flex',
                              alignItems: 'center',
                              gap: '4px'
                            }}
                          >
                            <Download size={14} /> HWPX 저장
                          </button>
                        </>
                      )}

                      {!isConverting && (
                        <button
                          onClick={() => handleRemoveFile(item.id)}
                          style={{
                            background: 'transparent',
                            border: 'none',
                            color: 'var(--text-muted)',
                            padding: '6px',
                            cursor: 'pointer',
                            borderRadius: '6px'
                          }}
                          title="목록에서 제거"
                        >
                          <Trash2 size={15} />
                        </button>
                      )}
                    </div>
                  </div>
                ))}
              </div>

              {/* 선택된 파일 텍스트 미리보기 드롭다운 */}
              {selectedPreview && (() => {
                const previewItem = fileList.find(f => f.id === selectedPreview);
                if (!previewItem || !previewItem.result) return null;
                return (
                  <div style={{
                    marginTop: '16px',
                    padding: '16px',
                    background: 'rgba(0, 0, 0, 0.3)',
                    borderRadius: '10px',
                    border: '1px solid var(--glass-border)',
                    animation: 'slide-up 0.2s ease'
                  }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
                      <span style={{ fontSize: '13px', fontWeight: 700, color: 'var(--text-primary)' }}>
                        📄 [{previewItem.name}] 텍스트 추출 내용 미리보기
                      </span>
                      <button
                        onClick={() => setSelectedPreview(null)}
                        style={{ background: 'none', border: 'none', color: 'var(--text-muted)', cursor: 'pointer', fontSize: '12px' }}
                      >
                        닫기
                      </button>
                    </div>
                    <pre style={{
                      margin: 0,
                      padding: '12px',
                      background: 'rgba(0,0,0,0.4)',
                      borderRadius: '8px',
                      fontSize: '12px',
                      color: 'var(--text-secondary)',
                      lineHeight: 1.6,
                      maxHeight: '180px',
                      overflowY: 'auto',
                      whiteSpace: 'pre-wrap',
                      wordBreak: 'break-all',
                      fontFamily: 'monospace'
                    }}>
                      {previewItem.result.previewText || '(추출된 텍스트가 없습니다)'}
                    </pre>
                  </div>
                );
              })()}
            </div>
          )}

          {/* 특징 및 안내 카드 */}
          <div style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))',
            gap: '16px'
          }}>
            <div className="glass-panel" style={{ padding: '18px', borderRadius: '14px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '8px' }}>
                <Zap size={18} color="var(--accent-blue)" />
                <h4 style={{ margin: 0, fontSize: '14px', fontWeight: 700, color: 'var(--text-primary)' }}>
                  순수 브라우저 기반 초고속 일괄 변환
                </h4>
              </div>
              <p style={{ margin: 0, fontSize: '12.5px', color: 'var(--text-secondary)', lineHeight: 1.5 }}>
                서버로 파일을 전송하지 않고 브라우저 내에서 즉각 변환되므로 대외비 보고서, 개인정보가 포함된 공문서도 보안 유출 걱정 없이 안전하게 변환됩니다.
              </p>
            </div>

            <div className="glass-panel" style={{ padding: '18px', borderRadius: '14px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '8px' }}>
                <FolderArchive size={18} color="#10b981" />
                <h4 style={{ margin: 0, fontSize: '14px', fontWeight: 700, color: 'var(--text-primary)' }}>
                  원클릭 ZIP 압축 일괄 다운로드
                </h4>
              </div>
              <p style={{ margin: 0, fontSize: '12.5px', color: 'var(--text-secondary)', lineHeight: 1.5 }}>
                수십 개의 .hwp 파일을 하나씩 다운로드할 필요 없이, 변환 완료 후 <strong>[전체 HWPX 다운로드 (.ZIP)]</strong> 버튼 하나로 압축 파일로 일괄 저장할 수 있습니다.
              </p>
            </div>

            <div className="glass-panel" style={{ padding: '18px', borderRadius: '14px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '8px' }}>
                <ShieldCheck size={18} color="#8b5cf6" />
                <h4 style={{ margin: 0, fontSize: '14px', fontWeight: 700, color: 'var(--text-primary)' }}>
                  한컴 표준 OWPML 규격 준수
                </h4>
              </div>
              <p style={{ margin: 0, fontSize: '12.5px', color: 'var(--text-secondary)', lineHeight: 1.5 }}>
                표준 section0.xml, header.xml, content.hpf, mimetype(STORED) 등 공공기관 및 한컴오피스 2024의 개방형 한글 스키마를 완벽히 준수하여 제작됩니다.
              </p>
            </div>
          </div>

        </div>
      )}

      {/* 모드 2: AI 표준 HWPX 보고서 자동 생성기 */}
      {activeSubMode === 'ai_report' && (
        <HwpxGenerator apiKey={apiKey} llmProvider={llmProvider} omniRouteModel={omniRouteModel} />
      )}

    </div>
  );
}
