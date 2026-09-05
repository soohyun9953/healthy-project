import React, { useState, useRef, useEffect } from 'react';
import { 
  FileText, 
  Upload, 
  RefreshCw, 
  CheckCircle2, 
  Download, 
  Trash2, 
  AlertCircle, 
  Sparkles, 
  Eye, 
  FolderArchive,
  FolderOpen,
  ArrowRight,
  ShieldCheck,
  Zap,
  Layers,
  ChevronDown,
  ChevronUp,
  Check
} from 'lucide-react';
import { 
  convertSingleHwpToHwpx, 
  convertBatchHwpToHwpx, 
  downloadAllAsZip,
  saveFilesToDirectory,
  saveBlobAs
} from '../utils/hwpToHwpxConverter.js';
import HwpxGenerator from './HwpxGenerator.jsx';

export default function HwpxConverter({ apiKey, llmProvider = 'gemini', omniRouteModel = 'auto' }) {
  // 메인 모드 선택 ('convert' : HWP 일괄 변환 | 'ai_report' : AI 표준보고서 생성)
  const [activeSubMode, setActiveSubMode] = useState('convert');

  // 업로드된 파일 목록 [{ id, file, name, size, status: 'idle'|'processing'|'done'|'error', result, error, progress }]
  const [fileList, setFileList] = useState([]);
  const [isConverting, setIsConverting] = useState(false);
  const [overallProgress, setOverallProgress] = useState({ currentIndex: 0, total: 0, percent: 0, currentFileName: '' });
  const [isDragActive, setIsDragActive] = useState(false);
  const [selectedPreview, setSelectedPreview] = useState(null);
  const [isZipping, setIsZipping] = useState(false);
  const [isSavingToDir, setIsSavingToDir] = useState(false);
  const [saveSuccessMsg, setSaveSuccessMsg] = useState('');

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

  // 핵심: 파일 추가 즉시 자동으로 전체 일괄 변환을 실행하는 함수
  const processAndConvertFiles = async (selectedFiles) => {
    const rawFiles = Array.from(selectedFiles);
    const validHwpFiles = rawFiles.filter(f => f.name.toLowerCase().endsWith('.hwp'));

    if (validHwpFiles.length === 0) {
      alert('한글(.hwp) 확장자 파일만 업로드 가능합니다.\n(HWPX 파일은 이미 변환된 파일입니다)');
      return;
    }

    if (validHwpFiles.length !== rawFiles.length) {
      alert(`선택된 파일 중 .hwp 형식이 아닌 ${rawFiles.length - validHwpFiles.length}개 파일은 제외되었습니다.`);
    }

    // 새로 들어온 파일 아이템들 생성
    const newItems = validHwpFiles.map((f, idx) => ({
      id: `hwp_${Date.now()}_${idx}_${Math.random().toString(36).substr(2, 5)}`,
      file: f,
      name: f.name,
      size: f.size,
      status: 'processing', // 즉시 자동 변환 시작
      progress: 10,
      result: null,
      error: null
    }));

    // 기존 목록 초기화 후 새 파일들로 세팅 (또는 덮어쓰기)
    setFileList(newItems);
    setIsConverting(true);
    setSaveSuccessMsg('');
    setSelectedPreview(null);

    const targetTotal = newItems.length;

    for (let i = 0; i < targetTotal; i++) {
      const currentItem = newItems[i];

      setOverallProgress({
        currentIndex: i + 1,
        total: targetTotal,
        percent: Math.round(((i) / targetTotal) * 100),
        currentFileName: `[${i + 1}/${targetTotal}] ${currentItem.name} 변환 중...`
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
      currentFileName: '전체 파일 변환이 완료되었습니다! 아래에서 원하는 저장 위치를 선택해 주세요.'
    });
    setIsConverting(false);
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
      processAndConvertFiles(e.dataTransfer.files);
    }
  };

  // 전체 목록 초기화
  const handleClearAll = () => {
    if (isConverting) return;
    setFileList([]);
    setOverallProgress({ currentIndex: 0, total: 0, percent: 0, currentFileName: '' });
    setSelectedPreview(null);
    setSaveSuccessMsg('');
  };

  // 1. [원하는 저장 폴더 선택하여 전체 저장]
  const handleSelectFolderAndSaveAll = async () => {
    const doneItems = fileList.filter(item => item.status === 'done' && item.result && item.result.blob);
    if (doneItems.length === 0) {
      alert('저장할 수 있는 변환 완료 파일이 없습니다.');
      return;
    }

    try {
      setIsSavingToDir(true);
      setSaveSuccessMsg('');

      // File System Access API 시도
      if (window.showDirectoryPicker) {
        const { dirName, savedCount } = await saveFilesToDirectory(doneItems);
        setSaveSuccessMsg(`🎉 선택하신 [${dirName}] 폴더에 총 ${savedCount}개의 '변환_*.hwpx' 파일이 성공적으로 저장되었습니다!`);
      } else {
        // DirectoryPicker 미지원 브라우저는 ZIP 일괄 다운로드로 자동 안내
        const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        await downloadAllAsZip(doneItems, `변환_HWPX_전체문서_${today}.zip`);
        setSaveSuccessMsg(`🎉 ZIP 압축 파일로 일괄 저장되었습니다!`);
      }
    } catch (err) {
      if (err.name === 'AbortError') {
        // 사용자가 폴더 선택창을 취소한 경우
        return;
      }
      console.warn('Directory save fallback to zip:', err);
      try {
        const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        await downloadAllAsZip(doneItems, `변환_HWPX_전체문서_${today}.zip`);
        setSaveSuccessMsg(`🎉 ZIP 압축 파일로 일괄 다운로드되었습니다!`);
      } catch (zipErr) {
        alert(`저장 중 오류 발생: ${zipErr.message}`);
      }
    } finally {
      setIsSavingToDir(false);
    }
  };

  // 2. [전체 ZIP 일괄 다운로드]
  const handleDownloadAllZip = async () => {
    const doneItems = fileList.filter(item => item.status === 'done' && item.result && item.result.blob);
    if (doneItems.length === 0) {
      alert('다운로드할 수 있는 변환 완료 파일이 없습니다.');
      return;
    }

    try {
      setIsZipping(true);
      const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
      await downloadAllAsZip(doneItems, `변환_HWPX_전체문서_${today}.zip`);
      setSaveSuccessMsg(`🎉 ZIP 압축 파일로 일괄 저장되었습니다!`);
    } catch (err) {
      alert(`ZIP 압축 다운로드 중 오류 발생: ${err.message}`);
    } finally {
      setIsZipping(false);
    }
  };

  // 3. 단일 HWPX 파일 개별 다운로드
  const handleDownloadSingle = (item) => {
    if (!item.result || !item.result.blob) {
      alert('변환된 파일 데이터가 존재하지 않습니다.');
      return;
    }
    saveBlobAs(item.result.blob, item.result.outputName);
  };

  // 통계 계산
  const totalUploaded = fileList.length;
  const doneCount = fileList.filter(f => f.status === 'done').length;
  const errorCount = fileList.filter(f => f.status === 'error').length;
  const processingCount = fileList.filter(f => f.status === 'processing').length;

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
            <span>HWP ➔ HWPX 원터치 자동 변환</span>
            {doneCount > 0 && (
              <span style={{
                background: 'var(--success-color)',
                color: '#fff',
                fontSize: '11px',
                padding: '1px 6px',
                borderRadius: '10px'
              }}>
                {doneCount}건 완료
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
          <span>파일명 앞 "변환_" 자동 추가 &amp; 표준 OWPML</span>
        </div>
      </div>

      {/* 모드 1: HWP -> HWPX 자동 변환기 */}
      {activeSubMode === 'convert' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
          
          {/* 상단 안내 & 드롭존 카드 */}
          <div className="glass-panel" style={{ padding: '24px', borderRadius: '16px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '16px' }}>
              <div>
                <h3 style={{ margin: '0 0 6px', fontSize: '17px', fontWeight: 800, color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <FileText size={20} color="var(--accent-blue)" />
                  한글 파일(.hwp) ➔ HWPX 자동 변환 및 원하는 폴더 저장
                </h3>
                <p style={{ margin: 0, fontSize: '13px', color: 'var(--text-secondary)', lineHeight: 1.5 }}>
                  파일을 끌어다 놓으면 <strong>즉시 자동으로 전체 변환</strong>되며, 파일명 앞에 <strong>"변환_"</strong>이 추가되어 <strong>원하는 PC 저장 위치에 한 번에 저장</strong>됩니다.
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
                  <Trash2 size={14} /> 목록 비우기
                </button>
              )}
            </div>

            {/* 원터치 자동 변환 드롭존 */}
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
                padding: '38px 20px',
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
                    processAndConvertFiles(e.target.files);
                  }
                  e.target.value = ''; // 재선택 가능하게 리셋
                }}
                style={{ display: 'none' }}
              />

              <div style={{
                width: '60px',
                height: '60px',
                borderRadius: '18px',
                background: isConverting 
                  ? 'rgba(59, 130, 246, 0.2)' 
                  : 'linear-gradient(135deg, rgba(59, 130, 246, 0.25), rgba(16, 185, 129, 0.25))',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                color: isConverting ? 'var(--accent-blue)' : '#10b981',
                boxShadow: '0 8px 20px rgba(0,0,0,0.2)'
              }}>
                {isConverting ? <RefreshCw size={30} className="spin" /> : <Upload size={30} />}
              </div>

              <div>
                <p style={{ margin: '0 0 4px', fontSize: '15.5px', fontWeight: 800, color: 'var(--text-primary)' }}>
                  {isConverting 
                    ? '⚡ 파일을 실시간 자동 변환하고 있습니다...' 
                    : '변환할 .hwp 파일들을 여기에 끌어다 놓으세요 (복수 선택 지원)'}
                </p>
                <p style={{ margin: 0, fontSize: '12.5px', color: 'var(--text-muted)' }}>
                  파일을 놓는 즉시 <strong>자동으로 전체 변환</strong>이 시작되며, <strong>"변환_*.hwpx"</strong>로 생성됩니다.
                </p>
              </div>

              <div style={{
                display: 'inline-flex',
                alignItems: 'center',
                gap: '6px',
                background: 'rgba(16, 185, 129, 0.1)',
                border: '1px solid rgba(16, 185, 129, 0.3)',
                padding: '5px 14px',
                borderRadius: '20px',
                fontSize: '11.5px',
                color: '#34d399',
                fontWeight: 700
              }}>
                <Zap size={13} /> 추가 클릭 없이 등록 즉시 자동 변환 ➔ 원하는 폴더에 원클릭 일괄 저장
              </div>
            </div>

            {/* 변환 진행률 게이지 바 */}
            {isConverting && (
              <div style={{ marginTop: '16px' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '12.5px', marginBottom: '6px', color: 'var(--text-secondary)' }}>
                  <span style={{ fontWeight: 600 }}>{overallProgress.currentFileName}</span>
                  <span style={{ fontWeight: 800, color: 'var(--accent-blue)' }}>{overallProgress.percent}%</span>
                </div>
                <div style={{ width: '100%', height: '8px', background: 'rgba(255,255,255,0.08)', borderRadius: '6px', overflow: 'hidden' }}>
                  <div 
                    style={{ 
                      width: `${overallProgress.percent}%`, 
                      height: '100%', 
                      background: 'linear-gradient(to right, var(--accent-blue), #10b981)', 
                      transition: 'width 0.25s ease' 
                    }} 
                  />
                </div>
              </div>
            )}

            {/* 저장 성공 안내 알림 배너 */}
            {saveSuccessMsg && (
              <div style={{
                marginTop: '16px',
                padding: '12px 16px',
                background: 'rgba(16, 185, 129, 0.12)',
                border: '1px solid rgba(16, 185, 129, 0.4)',
                borderRadius: '10px',
                color: '#34d399',
                fontSize: '13px',
                fontWeight: 700,
                display: 'flex',
                alignItems: 'center',
                gap: '8px',
                animation: 'scale-in 0.2s ease'
              }}>
                <CheckCircle2 size={18} />
                <span>{saveSuccessMsg}</span>
              </div>
            )}

            {/* 변환 완료 후: 원하는 저장 위치 선택 및 일괄 다운로드 액션 바 */}
            {doneCount > 0 && (
              <div style={{
                marginTop: '20px',
                padding: '16px 20px',
                background: 'linear-gradient(135deg, rgba(16, 185, 129, 0.08), rgba(59, 130, 246, 0.08))',
                borderRadius: '14px',
                border: '1px solid rgba(16, 185, 129, 0.3)',
                display: 'flex',
                flexWrap: 'wrap',
                alignItems: 'center',
                justifyContent: 'space-between',
                gap: '14px'
              }}>
                {/* 좌측 요약 */}
                <div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '2px' }}>
                    <CheckCircle2 size={18} color="#10b981" />
                    <span style={{ fontSize: '15px', fontWeight: 800, color: 'var(--text-primary)' }}>
                      총 {doneCount}개 파일 변환 완료!
                    </span>
                  </div>
                  <p style={{ margin: 0, fontSize: '12px', color: 'var(--text-muted)' }}>
                    원하는 PC 폴더를 지정하여 <strong>"변환_*.hwpx"</strong> 파일들을 한 번에 쏙 저장하세요.
                  </p>
                </div>

                {/* 우측 메인 저장 버튼들 */}
                <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                  {/* 핵심 1: [원하는 폴더 선택하여 전체 저장] */}
                  <button
                    onClick={handleSelectFolderAndSaveAll}
                    disabled={isSavingToDir || isConverting}
                    style={{
                      background: 'linear-gradient(135deg, #2563eb, #1d4ed8)',
                      color: '#ffffff',
                      border: 'none',
                      padding: '11px 20px',
                      borderRadius: '10px',
                      fontSize: '14px',
                      fontWeight: 800,
                      cursor: (isSavingToDir || isConverting) ? 'not-allowed' : 'pointer',
                      display: 'flex',
                      alignItems: 'center',
                      gap: '8px',
                      boxShadow: '0 6px 18px rgba(37, 99, 235, 0.4)',
                      transition: 'all 0.2s ease'
                    }}
                    title="원하는 PC 저장 폴더를 선택하여 변환된 파일들을 일괄 저장합니다"
                  >
                    <FolderOpen size={18} />
                    <span>{isSavingToDir ? '폴더에 저장 중...' : `📁 원하는 폴더 선택하여 전체 저장 (${doneCount}건)`}</span>
                  </button>

                  {/* 핵심 2: [ZIP 압축 파일로 일괄 다운로드] */}
                  <button
                    onClick={handleDownloadAllZip}
                    disabled={isZipping || isConverting}
                    style={{
                      background: 'rgba(255, 255, 255, 0.08)',
                      color: 'var(--text-primary)',
                      border: '1px solid var(--glass-border)',
                      padding: '11px 16px',
                      borderRadius: '10px',
                      fontSize: '13px',
                      fontWeight: 700,
                      cursor: (isZipping || isConverting) ? 'not-allowed' : 'pointer',
                      display: 'flex',
                      alignItems: 'center',
                      gap: '6px'
                    }}
                    title="ZIP 압축 파일 1개로 묶어 다운로드합니다"
                  >
                    <FolderArchive size={16} />
                    <span>{isZipping ? 'ZIP 압축 중...' : '📦 ZIP으로 저장'}</span>
                  </button>
                </div>
              </div>
            )}
          </div>

          {/* 변환된 파일 목록 */}
          {fileList.length > 0 && (
            <div className="glass-panel" style={{ padding: '20px', borderRadius: '16px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '14px' }}>
                <h4 style={{ margin: 0, fontSize: '15px', fontWeight: 700, color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <Layers size={18} color="var(--accent-blue)" />
                  변환 대상 및 결과 파일 목록 ({fileList.length}건)
                </h4>
                <span style={{ fontSize: '12px', color: 'var(--text-muted)' }}>
                  파일명 앞에 <strong>"변환_"</strong>이 자동 부여되었습니다
                </span>
              </div>

              <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                {fileList.map((item) => (
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
                          <span style={{ fontSize: '13px', color: 'var(--text-muted)' }}>
                            [원본] {item.name} ({formatBytes(item.size)})
                          </span>
                        </div>

                        {item.status === 'done' && item.result && (
                          <div style={{ display: 'flex', alignItems: 'center', gap: '10px', fontSize: '13.5px', color: 'var(--text-primary)', fontWeight: 700, flexWrap: 'wrap' }}>
                            <span style={{ color: '#10b981', display: 'flex', alignItems: 'center', gap: '4px' }}>
                              <ArrowRight size={14} /> {item.result.outputName}
                            </span>
                            <span style={{ fontSize: '11.5px', color: 'var(--text-muted)', fontWeight: 400 }}>
                              ({formatBytes(item.result.outputSize)} · 문단 {item.result.paragraphsCount}개)
                            </span>
                          </div>
                        )}

                        {item.status === 'processing' && (
                          <div style={{ fontSize: '12px', color: 'var(--accent-blue)', display: 'flex', alignItems: 'center', gap: '6px' }}>
                            <RefreshCw size={12} className="spin" />
                            <span>바이너리 스트림 분석 및 HWPX 자동 빌드 중...</span>
                          </div>
                        )}

                        {item.status === 'error' && (
                          <div style={{ fontSize: '12px', color: '#f87171' }}>
                            ❌ 변환 실패: {item.error}
                          </div>
                        )}
                      </div>
                    </div>

                    {/* 우측: 개별 다운로드 및 미리보기 */}
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
                            title="이 파일만 개별 저장"
                          >
                            <Download size={14} /> 개별 저장
                          </button>
                        </>
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
                        📄 [{previewItem.result.outputName}] 추출 텍스트 미리보기
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

          {/* 하단 3가지 핵심 안내 */}
          <div style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))',
            gap: '16px'
          }}>
            <div className="glass-panel" style={{ padding: '18px', borderRadius: '14px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '8px' }}>
                <Zap size={18} color="var(--accent-blue)" />
                <h4 style={{ margin: 0, fontSize: '14px', fontWeight: 700, color: 'var(--text-primary)' }}>
                  드래그 즉시 전체 자동 변환
                </h4>
              </div>
              <p style={{ margin: 0, fontSize: '12.5px', color: 'var(--text-secondary)', lineHeight: 1.5 }}>
                파일을 끌어다 놓는 순간 수동 버튼을 누를 필요 없이 즉시 전체 파일이 백그라운드에서 고속 자동 변환됩니다.
              </p>
            </div>

            <div className="glass-panel" style={{ padding: '18px', borderRadius: '14px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '8px' }}>
                <FolderOpen size={18} color="#2563eb" />
                <h4 style={{ margin: 0, fontSize: '14px', fontWeight: 700, color: 'var(--text-primary)' }}>
                  원하는 로컬 폴더 직접 지정 저장
                </h4>
              </div>
              <p style={{ margin: 0, fontSize: '12.5px', color: 'var(--text-secondary)', lineHeight: 1.5 }}>
                <strong>[원하는 폴더 선택하여 전체 저장]</strong>을 클릭하여 내 PC의 원하는 폴더를 선택하면, <strong>변환_*.hwpx</strong> 파일들이 해당 폴더에 한 번에 저장됩니다.
              </p>
            </div>

            <div className="glass-panel" style={{ padding: '18px', borderRadius: '14px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '8px' }}>
                <ShieldCheck size={18} color="#10b981" />
                <h4 style={{ margin: 0, fontSize: '14px', fontWeight: 700, color: 'var(--text-primary)' }}>
                  파일명 '변환_' 추가 및 보안성
                </h4>
              </div>
              <p style={{ margin: 0, fontSize: '12.5px', color: 'var(--text-secondary)', lineHeight: 1.5 }}>
                원본 파일과 구분하기 쉽게 <strong>"변환_원본파일명.hwpx"</strong>로 생성되며, 외부 서버 전송 없이 100% 브라우저 내에서 안전하게 처리됩니다.
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
