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
  FileCheck,
  HardDrive
} from 'lucide-react';
import { 
  convertSingleHwpToHwpx, 
  convertBatchHwpToHwpx, 
  downloadAllAsZip,
  downloadAllIndividually,
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
  const [isDownloadingAll, setIsDownloadingAll] = useState(false);
  const [saveSuccessMsg, setSaveSuccessMsg] = useState('');
  const [showDirTip, setShowDirTip] = useState(false);

  const fileInputRef = useRef(null);

  const [excludedNotice, setExcludedNotice] = useState('');

  // 파일 크기 포맷터
  const formatBytes = (bytes, decimals = 2) => {
    if (!bytes || bytes === 0) return '0 Bytes';
    const k = 1024;
    const dm = decimals < 0 ? 0 : decimals;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
  };

  // 핵심: 파일 추가 즉시 자동으로 전체 일괄 변환을 실행하는 함수 (HWPX 파일은 자동 제외)
  const processAndConvertFiles = async (selectedFiles) => {
    const rawFiles = Array.from(selectedFiles);
    
    // 1. HWPX 파일 및 비 HWP 파일 분리 필터링 (HWPX 파일은 전환 대상에서 자동 제외)
    const validHwpFiles = rawFiles.filter(f => f.name.toLowerCase().endsWith('.hwp'));
    const hwpxFiles = rawFiles.filter(f => f.name.toLowerCase().endsWith('.hwpx'));
    const otherFiles = rawFiles.filter(f => !f.name.toLowerCase().endsWith('.hwp') && !f.name.toLowerCase().endsWith('.hwpx'));

    // 제외된 파일 안내 메시지 구성
    const excludedParts = [];
    if (hwpxFiles.length > 0) {
      excludedParts.push(`이미 HWPX 형식인 파일 ${hwpxFiles.length}건(${hwpxFiles.map(f => f.name).slice(0, 2).join(', ')}${hwpxFiles.length > 2 ? ' 외' : ''})`);
    }
    if (otherFiles.length > 0) {
      excludedParts.push(`지원되지 않는 형식 ${otherFiles.length}건`);
    }

    if (excludedParts.length > 0) {
      setExcludedNotice(`💡 ${excludedParts.join(', ')}은(는) 변환 대상에서 자동으로 제외되었습니다.`);
    } else {
      setExcludedNotice('');
    }

    if (validHwpFiles.length === 0) {
      if (hwpxFiles.length > 0) {
        alert('선택하신 파일들이 이미 HWPX(.hwpx) 파일이므로 변환 대상에서 제외되어 변환할 .hwp 파일이 없습니다.');
      } else {
        alert('변환 대상인 한글(.hwp) 파일이 없습니다. .hwp 파일을 선택해 주세요.');
      }
      return;
    }

    // 새로 들어온 유효 HWP 파일 아이템들 생성
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
      currentFileName: '전체 파일 변환이 완료되었습니다! 아래에서 원하는 방식으로 저장하세요.'
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
    setExcludedNotice('');
  };

  // 1. [원클릭 ZIP 압축 파일 저장 - 추천 & 가장 안전]
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
      setSaveSuccessMsg(`🎉 ZIP 압축 파일(변환_HWPX_전체문서_${today}.zip)로 다운로드 폴더에 안전하게 일괄 저장되었습니다!`);
    } catch (err) {
      alert(`ZIP 압축 다운로드 중 오류 발생: ${err.message}`);
    } finally {
      setIsZipping(false);
    }
  };

  // 2. [원하는 로컬 폴더를 직접 선택하여 저장]
  const handleSelectFolderAndSaveAll = async () => {
    const doneItems = fileList.filter(item => item.status === 'done' && item.result && item.result.blob);
    if (doneItems.length === 0) {
      alert('저장할 수 있는 변환 완료 파일이 없습니다.');
      return;
    }

    try {
      setIsSavingToDir(true);
      setSaveSuccessMsg('');

      if (window.showDirectoryPicker) {
        const { dirName, savedCount } = await saveFilesToDirectory(doneItems);
        setSaveSuccessMsg(`🎉 선택하신 [${dirName}] 폴더에 총 ${savedCount}개의 '변환_*.hwpx' 파일이 성공적으로 저장되었습니다!`);
      } else {
        // DirectoryPicker 미지원 브라우저는 ZIP 다운로드로 자동 전환
        const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        await downloadAllAsZip(doneItems, `변환_HWPX_전체문서_${today}.zip`);
        setSaveSuccessMsg(`🎉 ZIP 압축 파일로 일괄 저장되었습니다!`);
      }
    } catch (err) {
      if (err.name === 'AbortError') {
        // 사용자가 폴더 선택창에서 '취소'를 누른 경우
        return;
      }
      
      // 브라우저 보안 에러 (C: 루트나 Windows 등 시스템 폴더 접근 차단)인 경우
      console.warn('Directory save blocked by browser policy:', err);
      setShowDirTip(true);
      
      // 사용자에게 친절하게 안내하고 ZIP 다운로드로 즉시 안전하게 저장
      const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
      try {
        await downloadAllAsZip(doneItems, `변환_HWPX_전체문서_${today}.zip`);
        setSaveSuccessMsg(`💡 시스템 폴더 보안 제한으로 인해 안전한 ZIP 압축 파일로 다운로드되었습니다. (또는 '새 폴더'를 생성하여 선택해 주세요)`);
      } catch (zipErr) {
        alert(`저장 중 오류 발생: ${zipErr.message}`);
      }
    } finally {
      setIsSavingToDir(false);
    }
  };

  // 3. [모든 파일 개별 순차 다운로드]
  const handleDownloadAllIndividually = async () => {
    const doneItems = fileList.filter(item => item.status === 'done' && item.result && item.result.blob);
    if (doneItems.length === 0) {
      alert('다운로드할 수 있는 변환 완료 파일이 없습니다.');
      return;
    }

    try {
      setIsDownloadingAll(true);
      const count = await downloadAllIndividually(doneItems);
      setSaveSuccessMsg(`🎉 총 ${count}개의 '변환_*.hwpx' 파일이 다운로드 폴더로 개별 저장되었습니다!`);
    } catch (err) {
      alert(`개별 다운로드 중 오류 발생: ${err.message}`);
    } finally {
      setIsDownloadingAll(false);
    }
  };

  // 4. 단일 HWPX 파일 개별 다운로드
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
                accept=".hwp,.hwpx"
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
                <Zap size={13} /> 추가 클릭 없이 등록 즉시 자동 변환 (HWPX 파일은 자동 제외)
              </div>
            </div>

            {/* 제외된 파일(HWPX 등) 안내 알림 */}
            {excludedNotice && (
              <div style={{
                marginTop: '12px',
                padding: '10px 14px',
                background: 'rgba(59, 130, 246, 0.08)',
                border: '1px solid rgba(59, 130, 246, 0.25)',
                borderRadius: '8px',
                color: '#93c5fd',
                fontSize: '12.5px',
                display: 'flex',
                alignItems: 'center',
                gap: '6px'
              }}>
                <Sparkles size={14} color="#60a5fa" />
                <span>{excludedNotice}</span>
              </div>
            )}

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

            {/* 브라우저 폴더 선택 팁 안내 */}
            {showDirTip && (
              <div style={{
                marginTop: '12px',
                padding: '10px 14px',
                background: 'rgba(59, 130, 246, 0.08)',
                border: '1px solid rgba(59, 130, 246, 0.25)',
                borderRadius: '8px',
                color: '#93c5fd',
                fontSize: '12px',
                lineHeight: 1.5
              }}>
                💡 <strong>폴더 선택 팁</strong>: 브라우저 보안 규정상 <code>C:\</code> 드라이브 최상위나 <code>Windows</code> 시스템 폴더는 직접 저장이 제한됩니다. 작업용 <strong>'새 폴더'</strong>(예: 바탕화면 내 폴더, 문서 폴더)를 선택하시거나 <strong>[ZIP으로 일괄 다운로드]</strong>를 이용하시면 가장 편리합니다.
              </div>
            )}

            {/* 변환 완료 후: 3가지 편리한 일괄 저장 옵션 카드 */}
            {doneCount > 0 && (
              <div style={{
                marginTop: '20px',
                padding: '18px 20px',
                background: 'linear-gradient(135deg, rgba(16, 185, 129, 0.08), rgba(59, 130, 246, 0.08))',
                borderRadius: '14px',
                border: '1px solid rgba(16, 185, 129, 0.3)',
                display: 'flex',
                flexDirection: 'column',
                gap: '14px'
              }}>
                {/* 상단 완료 타이틀 */}
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '8px' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <CheckCircle2 size={20} color="#10b981" />
                    <span style={{ fontSize: '15.5px', fontWeight: 800, color: 'var(--text-primary)' }}>
                      총 {doneCount}개 파일 변환 완료! 아래에서 원하는 저장 방식을 선택하세요:
                    </span>
                  </div>
                </div>

                {/* 3가지 일괄 저장 방식 버튼 그룹 */}
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))', gap: '12px' }}>
                  
                  {/* 옵션 1: [📦 ZIP 압축 파일로 다운로드 (추천 - 시스템 보안 제한 없음)] */}
                  <button
                    onClick={handleDownloadAllZip}
                    disabled={isZipping || isConverting}
                    style={{
                      background: 'linear-gradient(135deg, #10b981, #059669)',
                      color: '#ffffff',
                      border: 'none',
                      padding: '14px 16px',
                      borderRadius: '10px',
                      fontSize: '13.5px',
                      fontWeight: 800,
                      cursor: (isZipping || isConverting) ? 'not-allowed' : 'pointer',
                      display: 'flex',
                      flexDirection: 'column',
                      alignItems: 'center',
                      justifyContent: 'center',
                      gap: '4px',
                      boxShadow: '0 6px 16px rgba(16, 185, 129, 0.3)',
                      transition: 'all 0.2s ease'
                    }}
                    title="시스템 보안 제한 없이 1개의 ZIP 파일로 바로 일괄 저장 (추천)"
                  >
                    <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                      <FolderArchive size={18} />
                      <span>{isZipping ? 'ZIP 묶는 중...' : `📦 전체 ZIP 일괄 다운로드 (${doneCount}건)`}</span>
                    </div>
                    <span style={{ fontSize: '11px', fontWeight: 500, opacity: 0.9 }}>
                      ★ 추천: 보안 제한 없이 1초 만에 바로 저장
                    </span>
                  </button>

                  {/* 옵션 2: [📁 원하는 작업 폴더 선택하여 저장] */}
                  <button
                    onClick={handleSelectFolderAndSaveAll}
                    disabled={isSavingToDir || isConverting}
                    style={{
                      background: 'linear-gradient(135deg, #2563eb, #1d4ed8)',
                      color: '#ffffff',
                      border: 'none',
                      padding: '14px 16px',
                      borderRadius: '10px',
                      fontSize: '13.5px',
                      fontWeight: 800,
                      cursor: (isSavingToDir || isConverting) ? 'not-allowed' : 'pointer',
                      display: 'flex',
                      flexDirection: 'column',
                      alignItems: 'center',
                      justifyContent: 'center',
                      gap: '4px',
                      boxShadow: '0 6px 16px rgba(37, 99, 235, 0.3)',
                      transition: 'all 0.2s ease'
                    }}
                    title="원하는 작업 폴더를 선택하여 변환된 파일들을 각각 저장합니다"
                  >
                    <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                      <FolderOpen size={18} />
                      <span>{isSavingToDir ? '폴더에 저장 중...' : `📁 원하는 폴더 선택하여 저장`}</span>
                    </div>
                    <span style={{ fontSize: '11px', fontWeight: 500, opacity: 0.9 }}>
                      지정한 폴더에 '변환_*.hwpx'로 풀어서 저장
                    </span>
                  </button>

                  {/* 옵션 3: [📥 모든 파일 개별 다운로드] */}
                  <button
                    onClick={handleDownloadAllIndividually}
                    disabled={isDownloadingAll || isConverting}
                    style={{
                      background: 'rgba(255, 255, 255, 0.06)',
                      color: 'var(--text-primary)',
                      border: '1px solid var(--glass-border)',
                      padding: '14px 16px',
                      borderRadius: '10px',
                      fontSize: '13.5px',
                      fontWeight: 700,
                      cursor: (isDownloadingAll || isConverting) ? 'not-allowed' : 'pointer',
                      display: 'flex',
                      flexDirection: 'column',
                      alignItems: 'center',
                      justifyContent: 'center',
                      gap: '4px',
                      transition: 'all 0.2s ease'
                    }}
                    title="ZIP 없이 각 파일을 브라우저 다운로드 폴더에 연속 저장합니다"
                  >
                    <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                      <Download size={18} />
                      <span>{isDownloadingAll ? '개별 다운로드 중...' : `📥 모든 파일 개별 다운로드`}</span>
                    </div>
                    <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>
                      브라우저 다운로드 폴더로 각각 저장
                    </span>
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
                <FolderArchive size={18} color="#10b981" />
                <h4 style={{ margin: 0, fontSize: '14px', fontWeight: 700, color: 'var(--text-primary)' }}>
                  원클릭 안전 일괄 저장
                </h4>
              </div>
              <p style={{ margin: 0, fontSize: '12.5px', color: 'var(--text-secondary)', lineHeight: 1.5 }}>
                브라우저 보안 제한 없는 <strong>[전체 ZIP 일괄 다운로드]</strong> 또는 <strong>[원하는 폴더 선택하여 저장]</strong>을 통해 파일들을 한 번에 손쉽게 저장할 수 있습니다.
              </p>
            </div>

            <div className="glass-panel" style={{ padding: '18px', borderRadius: '14px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '8px' }}>
                <ShieldCheck size={18} color="#8b5cf6" />
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
