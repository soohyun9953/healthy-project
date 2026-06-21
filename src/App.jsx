import React, { useState, useEffect, useCallback } from 'react';
import DocumentValidator from './components/DocumentValidator';
import TypoValidator from './components/TypoValidator';
import LawConsultant from './components/LawConsultant';
import ErdGenerator from './components/ErdGenerator';
import ReferenceLibrary from './components/ReferenceLibrary';
import PptGenerator from './components/PptGenerator';
import MeetingMinutes from './components/MeetingMinutes';
import AiPptDesigner from './components/AiPptDesigner';
import RagKnowledgeBase from './components/RagKnowledgeBase';
import IsmpDaDashboard from './components/IsmpDaDashboard';
import PptValidator from './components/PptValidator';
import HwpxGenerator from './components/HwpxGenerator';
import { 
  Shield, 
  ShieldAlert,
  Activity, 
  FileText, 
  Presentation,
  CheckCircle2, 
  MessageSquare, 
  Database, 
  Settings, 
  Key, 
  AlertCircle, 
  Info, 
  Globe, 
  PlusCircle, 
  Menu, 
  X,
  BarChart3,
  Cpu,
  ChevronLeft,
  ChevronRight,
  PanelLeftClose,
  PanelLeftOpen,
  Trash2,
  Mic2,
  Eye,
  EyeOff,
  Sun,
  Moon,
  Sliders
} from 'lucide-react';
import { processFile } from './utils/fileExtractor';

// Google Analytics 이벤트 헬퍼
const gaEvent = (eventName, params = {}) => {
  if (typeof window.gtag === 'function') {
    window.gtag('event', eventName, params);
  }
};

function App() {
  const [activeTab, setActiveTab] = useState('ppt');
  // apiKeys: 문자열 배열로 관리, 기존 localStorage 콤마 구분 값과 호환
  const [apiKeys, setApiKeys] = useState(() => {
    const stored = localStorage.getItem('gemini_api_key') || '';
    const parsed = stored.split(',').map(k => k.trim()).filter(Boolean);
    return parsed.length > 0 ? parsed : [''];
  });
  const [newKeyInput, setNewKeyInput] = useState('');
  // 하위 컴포넌트에 전달할 콤마 구분 문자열
  const apiKey = apiKeys.filter(k => k.trim().startsWith('AIza') || k.trim().startsWith('AQ.')).join(',');
  const [modelUsage, setModelUsage] = useState(() => JSON.parse(localStorage.getItem('gemini_model_usage') || '{}'));
  const [showSettings, setShowSettings] = useState(false);
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState(() => localStorage.getItem('sidebar_collapsed') === 'true');
  const [showKey, setShowKey] = useState(false); // API 키 보이기/숨기기 토글
  const [isLightMode, setIsLightMode] = useState(() => localStorage.getItem('theme_mode') === 'light');

  // 사이드바 및 테마 상태 저장
  useEffect(() => {
    localStorage.setItem('sidebar_collapsed', isSidebarCollapsed);
  }, [isSidebarCollapsed]);

  useEffect(() => {
    localStorage.setItem('theme_mode', isLightMode ? 'light' : 'dark');
  }, [isLightMode]);

  // 탭 변경 시 GA 페이지뷰 전송
  useEffect(() => {
    const tabLabels = {
      main: 'AI 산출물 검증',
      typo: 'AI 교정교열',
      law: 'AI 법률 자문(제미나이)',
      'law-mcp': 'AI 법률 자문(로컬 RAG)',
      erd: 'AI ERD 설계',
      ppt: 'PPT 생성(표준산출물)',
      'hwpx-report': 'HWPX 생성(표준보고서)',
      library: '참고자료 라이브러리',
      meeting: 'AI 회의록 생성',
      aippt: '(작업중)AI PPT 디자이너',
    };
    gaEvent('page_view', {
      page_title: tabLabels[activeTab] || activeTab,
      page_location: window.location.href,
      page_path: `/?tab=${activeTab}`,
    });
  }, [activeTab]);

  // 로컬 스토리지 동기화
  useEffect(() => {
    localStorage.setItem('gemini_api_key', apiKey);
  }, [apiKey]);

  // API 키 추가
  const handleAddKey = () => {
    const trimmed = newKeyInput.trim();
    if (!trimmed) return;
    if (!apiKeys.includes(trimmed)) {
      setApiKeys(prev => {
        if (prev.length === 1 && prev[0].trim() === '') return [trimmed];
        return [...prev, trimmed];
      });
    }
    setNewKeyInput('');
  };

  // API 키 삭제
  const handleRemoveKey = (idx) => {
    setApiKeys(prev => {
        const next = prev.filter((_, i) => i !== idx);
        return next.length > 0 ? next : [''];
    });
  };

  // API 키 개별 수정
  const handleEditKey = (idx, value) => {
    setApiKeys(prev => prev.map((k, i) => i === idx ? value : k));
  };




  // 사용량 업데이트 이벤트 리스너
  useEffect(() => {
    const handleUsageUpdate = () => {
        setModelUsage(JSON.parse(localStorage.getItem('gemini_model_usage') || '{}'));
    };
    window.addEventListener('gemini_usage_updated', handleUsageUpdate);
    return () => window.removeEventListener('gemini_usage_updated', handleUsageUpdate);
  }, []);

  const tabs = [
    { id: 'ppt', label: 'PPT 생성(표준산출물)', icon: FileText, color: '#f97316' },
    { id: 'hwpx-report', label: 'HWPX 생성(표준보고서)', icon: FileText, color: '#10b981' },
    { id: 'ppt-verify', label: 'PPT 검증(표준산출물)', icon: ShieldAlert, color: '#e11d48' },
    { id: 'main', label: 'AI 산출물 검증', icon: Shield, color: 'var(--accent-blue)' },
    { id: 'typo', label: 'AI 교정교열', icon: CheckCircle2, color: 'var(--accent-purple)' },
    { id: 'law', label: 'AI 법률 자문(제미나이)', icon: MessageSquare, color: 'var(--success-color)' },
    { id: 'law-mcp', label: 'AI 법률 자문(로컬 RAG)', icon: MessageSquare, color: 'var(--accent-purple)' },
    { id: 'erd', label: 'AI ERD 설계', icon: Database, color: 'var(--warning-color)' },
    { id: 'meeting', label: 'AI 회의록 생성', icon: Mic2, color: '#8b5cf6' },
    { id: 'library', label: '참고자료 라이브러리', icon: Activity, color: '#64748b' },
    { id: 'aippt', label: '(작업중)AI PPT 디자이너', icon: Presentation, color: '#ec4899' },
    { id: 'rag', label: '프로젝트 RAG 지식베이스', icon: Database, color: 'var(--accent-blue)' },
    { id: 'ismpda', label: '(작업중)ISMP DA 검증 대시보드', icon: Sliders, color: 'var(--accent-purple)' },
  ];

  const activeTabData = tabs.find(t => t.id === activeTab);

  const keyCount = apiKey.split(',').filter(k => k.trim().startsWith('AIza') || k.trim().startsWith('AQ.')).length;

  return (
    <div className={`app-container ${isSidebarCollapsed ? 'sidebar-collapsed' : ''} ${isLightMode ? 'light-mode' : ''}`}>
      {/* Sidebar */}
      <aside className={`sidebar ${isMobileMenuOpen ? 'mobile-open' : ''} ${isSidebarCollapsed ? 'collapsed' : ''}`}>
        <div className="sidebar-header">
          <div className="logo-container">
            <div className="logo-icon">
              <Shield size={24} color="white" />
            </div>
            {!isSidebarCollapsed && (
              <div className="logo-text">
                <h1>건강한 프로젝트</h1>
                <span>AI 산출물 검수 v2.4</span>
              </div>
            )}
          </div>
          
          <button 
            className="sidebar-toggle-btn"
            onClick={() => setIsSidebarCollapsed(!isSidebarCollapsed)}
            title={isSidebarCollapsed ? "펼치기" : "접기"}
          >
            {isSidebarCollapsed ? <PanelLeftOpen size={18} /> : <PanelLeftClose size={18} />}
          </button>

          <button className="mobile-close" onClick={() => setIsMobileMenuOpen(false)}>
            <X size={20} />
          </button>
        </div>

        <nav className="sidebar-nav">
          {tabs.map((tab) => {
            const Icon = tab.icon;
            const isActive = activeTab === tab.id;
            return (
              <button
                key={tab.id}
                id={`sidebar-tab-${tab.id}`}
                className={`nav-item ${isActive ? 'active' : ''}`}
                onClick={() => {
                  setActiveTab(tab.id);
                  setIsMobileMenuOpen(false);
                  gaEvent('tab_click', {
                    tab_id: tab.id,
                    tab_label: tab.label,
                  });
                }}
              >
                <div className="nav-icon-wrapper" style={{ color: isActive ? tab.color : 'inherit' }}>
                  <Icon size={20} />
                </div>
                {!isSidebarCollapsed && <span>{tab.label}</span>}
                {isActive && !isSidebarCollapsed && <div className="active-indicator" style={{ backgroundColor: tab.color }} />}
              </button>
            );
          })}
        </nav>


      </aside>

      {/* Main Content */}
      <main className="main-content">
        <header className="content-header">
          <div className="header-left">
            <button className="mobile-menu-btn" onClick={() => setIsMobileMenuOpen(true)}>
              <Menu size={24} />
            </button>
            <div className="breadcrumb">
              <span className="breadcrumb-parent">Validator Space</span>
              <span className="breadcrumb-separator">/</span>
              <span className="breadcrumb-current">{activeTabData?.label}</span>
            </div>
          </div>

          <div className="header-right">
            {/* API Status Badge */}
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginRight: '8px', padding: '4px 12px', background: 'rgba(255,255,255,0.03)', borderRadius: '10px', border: '1px solid var(--glass-border)' }}>
              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-start' }}>
                <span className="mobile-hide-text" style={{ fontSize: '8px', color: 'var(--text-muted)', fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.3px' }}>API Multi-Key</span>
                <span style={{ fontSize: '11px', fontWeight: 700, color: keyCount > 0 ? 'var(--success-color)' : 'var(--danger-color)' }}>
                  {keyCount} Keys Connected
                </span>
              </div>
              <div style={{ width: '32px', height: '3px', background: 'rgba(255,255,255,0.1)', borderRadius: '2px', overflow: 'hidden' }}>
                <div style={{ width: `${Math.min(keyCount * 25, 100)}%`, height: '100%', backgroundColor: keyCount > 1 ? 'var(--success-color)' : 'var(--warning-color)' }}></div>
              </div>
            </div>

            <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: '2px', marginRight: '8px' }}>
              <span className="mobile-hide-text" style={{ fontSize: '9px', color: 'var(--text-muted)', fontWeight: 600, letterSpacing: '0.5px', textTransform: 'uppercase' }}>Last Update</span>
              <span style={{ fontSize: '12px', color: 'var(--accent-blue)', fontWeight: 700, fontFamily: 'monospace' }}>2026.06.13 v2.4</span>
            </div>
            
            <button 
              className="settings-btn" 
              onClick={() => setIsLightMode(!isLightMode)}
              title={isLightMode ? "다크 모드로 전환" : "라이트 모드로 전환"}
            >
              {isLightMode ? <Moon size={20} /> : <Sun size={20} />}
            </button>

            <button 
              className={`settings-btn ${showSettings ? 'active' : ''}`} 
              onClick={() => setShowSettings(!showSettings)}
              title="API 설정 및 모델 관리"
            >
              <Settings size={20} />
            </button>
          </div>

          {/* Settings Dropdown */}
          {showSettings && (
            <div className="settings-dropdown animate-scale-in">
              <div className="settings-header">
                <Settings size={16} />
                <h3>환경 설정 및 모델 관리</h3>
              </div>
              
              <div className="settings-body">
                <div className="setting-group">
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
                    <label style={{ margin: 0 }}>
                      <Key size={14} /> Gemini API Keys
                      <span className="badge">{keyCount}개 연결됨</span>
                    </label>
                    <button 
                      onClick={() => setShowKey(!showKey)} 
                      style={{ background: 'none', border: 'none', color: 'var(--text-muted)', cursor: 'pointer', display: 'flex', alignItems: 'center', padding: '4px' }}
                      title="키 텍스트 보이기/숨기기"
                    >
                      {showKey ? <EyeOff size={14} /> : <Eye size={14} />}
                    </button>
                  </div>

                  {/* 등록된 키 목록 */}
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '6px', marginBottom: '8px' }}>
                    {apiKeys.map((k, idx) => (
                      <div key={idx} style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                        <div style={{ position: 'relative', flex: 1 }}>
                          <input
                            type={showKey ? "text" : "password"}
                            value={k}
                            onChange={(e) => handleEditKey(idx, e.target.value)}
                            placeholder={`API Key ${idx + 1} (AIza... 또는 AQ...)`}
                            style={{
                              width: '100%',
                              background: (k.trim().startsWith('AIza') || k.trim().startsWith('AQ.'))
                                ? 'rgba(16,185,129,0.08)'
                                : 'rgba(255,255,255,0.04)',
                              border: `1px solid ${(k.trim().startsWith('AIza') || k.trim().startsWith('AQ.')) ? 'rgba(16,185,129,0.4)' : 'var(--glass-border)'}`,
                              borderRadius: '8px',
                              padding: '8px 12px',
                              color: 'var(--text-primary)',
                              fontSize: '12px',
                              outline: 'none',
                              boxSizing: 'border-box',
                            }}
                          />
                        </div>
                        <button
                          onClick={() => handleRemoveKey(idx)}
                          title="키 삭제"
                          style={{
                            background: 'rgba(239,68,68,0.1)',
                            border: '1px solid rgba(239,68,68,0.3)',
                            borderRadius: '7px',
                            padding: '7px',
                            cursor: 'pointer',
                            color: '#ef4444',
                            display: 'flex',
                            alignItems: 'center',
                            flexShrink: 0,
                          }}
                        >
                          <Trash2 size={14} />
                        </button>
                      </div>
                    ))}
                  </div>

                  {/* 새 키 추가 입력창 */}
                  <div style={{ display: 'flex', gap: '6px' }}>
                    <input
                      type={showKey ? "text" : "password"}
                      value={newKeyInput}
                      onChange={(e) => setNewKeyInput(e.target.value)}
                      onKeyDown={(e) => e.key === 'Enter' && handleAddKey()}
                      placeholder="새 API Key 입력 후 + 버튼 또는 Enter"
                      style={{
                        flex: 1,
                        background: 'rgba(255,255,255,0.04)',
                        border: '1px dashed var(--glass-border)',
                        borderRadius: '8px',
                        padding: '8px 12px',
                        color: 'var(--text-primary)',
                        fontSize: '12px',
                        outline: 'none',
                      }}
                    />
                    <button
                      onClick={handleAddKey}
                      title="키 추가"
                      style={{
                        background: 'rgba(99,102,241,0.15)',
                        border: '1px solid rgba(99,102,241,0.4)',
                        borderRadius: '7px',
                        padding: '7px 12px',
                        cursor: 'pointer',
                        color: 'var(--accent-blue)',
                        display: 'flex',
                        alignItems: 'center',
                        gap: '4px',
                        fontSize: '12px',
                        fontWeight: 600,
                        flexShrink: 0,
                      }}
                    >
                      <PlusCircle size={14} /> 추가
                    </button>
                  </div>

                  <p className="helper-text" style={{ marginTop: '6px' }}>할당량 초과 시 자동으로 다음 키로 전환됩니다.</p>
                </div>



                <div className="setting-group usage-section">
                  <label>
                    <BarChart3 size={14} /> 모델별 누적 사용량
                  </label>
                  <div className="usage-stats">
                    {Object.keys(modelUsage).length === 0 ? (
                        <div className="no-usage">사용 기록이 없습니다.</div>
                    ) : (
                        Object.entries(modelUsage).map(([model, count]) => (
                            <div key={model} className="usage-item">
                                <span className="model-name">{model.split('/').pop()}</span>
                                <span className="use-count">{count}회</span>
                            </div>
                        ))
                    )}
                  </div>
                  {Object.keys(modelUsage).length > 0 && (
                      <button 
                        className="reset-usage-btn"
                        onClick={() => {
                            if(window.confirm('사용 통계를 초기화하시겠습니까?')) {
                                localStorage.removeItem('gemini_model_usage');
                                setModelUsage({});
                            }
                        }}
                      >통계 초기화</button>
                  )}
                </div>
              </div>
            </div>
          )}
        </header>

        <div className="content-body">
          {activeTab === 'main' && <DocumentValidator apiKey={apiKey} />}
          {activeTab === 'hwpx-report' && <HwpxGenerator apiKey={apiKey} />}
          {activeTab === 'ismpda' && <IsmpDaDashboard />}
          {activeTab === 'typo' && <TypoValidator apiKey={apiKey} />}
          {activeTab === 'law' && <LawConsultant apiKey={apiKey} isMcpMode={false} />}
          {activeTab === 'law-mcp' && <LawConsultant apiKey={apiKey} isMcpMode={true} />}
          { activeTab === 'erd' && <ErdGenerator apiKey={apiKey} /> }
          { activeTab === 'aippt' && <AiPptDesigner apiKey={apiKey} /> }
          { activeTab === 'ppt' && <PptGenerator apiKey={apiKey} /> }
          { activeTab === 'ppt-verify' && <PptValidator apiKey={apiKey} /> }
          {activeTab === 'meeting' && <MeetingMinutes apiKey={apiKey} />}
          {activeTab === 'library' && <ReferenceLibrary />}
          {activeTab === 'rag' && <RagKnowledgeBase apiKey={apiKey} />}
        </div>
      </main>


    </div>
  );
}

export default App;
