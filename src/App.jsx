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
  Sliders,
  Sparkles,
  HelpCircle
} from 'lucide-react';
import { processFile } from './utils/fileExtractor';

// Google Analytics 이벤트 헬퍼
const gaEvent = (eventName, params = {}) => {
  if (typeof window.gtag === 'function') {
    window.gtag('event', eventName, params);
  }
};

// ── 탭별 상세 가이드 데이터 ─────────────────────────
const tab_guides = {
  ppt: {
    title: 'PPT 생성(표준산출물)',
    desc: '엑셀 데이터 매핑을 통해 파워포인트 슬라이드를 자동으로 생성하거나, 디자인 서식을 일괄 적용합니다.',
    steps: [
      '좌측 탭에서 [엑셀 매핑] 또는 [일괄 편집] 모드를 선택합니다.',
      '매핑 템플릿(PPTX) 파일과 치환할 데이터가 담긴 엑셀(XLSX) 파일을 드래그하여 업로드합니다.',
      '치환할 데이터 매핑 규칙을 설정하고 [자동 매핑 및 파일 변환 시작] 버튼을 누릅니다.',
      '완성된 PPTX 파일을 다운로드할 위치를 지정하여 저장합니다.'
    ],
    tips: '다량의 데이터를 한번에 처리할 때는 배치 청크 사이즈를 조절해 속도를 높일 수 있습니다.'
  },
  'ppt-verify': {
    title: 'PPT 검증(표준산출물)',
    desc: '생성된 파워포인트 문서 내에 동일 단어가 중복으로 표시되거나, 포맷 서식이 틀어지는 등의 품질 이슈를 검증합니다.',
    steps: [
      '검증할 PPTX 파일 또는 HWPX 파일을 업로드 영역에 드롭합니다.',
      '중복 단어 검색 모드(숫자 제외, 대소문자 구분 여부 등)를 선택합니다.',
      '[문서 검증 시작] 버튼을 클릭해 분석을 수행합니다.',
      '검출된 중복 단어와 오탈자 통계를 대시보드로 확인하고 보고서로 저장합니다.'
    ],
    tips: '의미상 스페이스를 포함하지 않는 단어들의 중복(예: xxse sexx 등)은 자동으로 필터링됩니다.'
  },
  main: {
    title: 'AI 산출물 검증',
    desc: '프로젝트 산출물(사업 제안서, 보고서 등)을 제미나이 AI가 심층 분석하여 보완점 및 논리적 오류를 검토합니다.',
    steps: [
      '분석 대상 문서 파일을 업로드합니다.',
      '검토 기준(ISP/ISMP 표준, 보안성, 기술 정합성 등)을 설정합니다.',
      'Gemini AI를 구동하여 문서 품질 진단을 시작합니다.',
      'AI가 도출한 개선 가이드와 정량적인 품질 점수를 리포트로 다운로드합니다.'
    ],
    tips: 'Gemini API Key가 올바르게 설정되어 있어야 AI 진단 기능이 정상 작동합니다.'
  },
  typo: {
    title: 'AI 교정교열',
    desc: '작성된 텍스트나 문서를 맞춤법, 띄어쓰기, 문맥 오류에 맞게 교정하여 공문서 격식에 어울리도록 수정합니다.',
    steps: [
      '텍스트를 입력창에 붙여넣거나 원본 문서를 업로드합니다.',
      '공공 표준/비즈니스 격식 등 원하는 교정 스타일을 선택합니다.',
      '[교정 시작] 버튼을 눌러 실시간 교정본을 확인합니다.',
      '교정이 완료된 텍스트를 클립보드에 복사하거나 덮어씁니다.'
    ],
    tips: '긴 장문의 경우 단락 단위로 자동 분할되어 순차적으로 처리됩니다.'
  },
  law: {
    title: 'AI 법률 자문(제미나이)',
    desc: '프로젝트 수행 및 계약 과정에서 발생할 수 있는 법적 쟁점에 대하여 제미나이 AI 법률 비서가 실시간으로 자문합니다.',
    steps: [
      '법률 자문이 필요한 계약서 조항이나 쟁점 사항을 질문창에 입력합니다.',
      '[자문 요청] 버튼을 눌러 법적 판례 및 가이드라인을 확인합니다.',
      '생성된 답변 내의 주요 핵심 제언을 확인하고 기록으로 저장합니다.'
    ],
    tips: '추가 검증이 필요한 계약서는 RAG 지식베이스에 미리 등록해 두면 더욱 정확합니다.'
  },
  'law-mcp': {
    title: 'AI 법률 자문(로컬 RAG)',
    desc: '로컬 개발 환경의 법률 데이터베이스(RAG)를 기반으로 하여 인터넷 연결 없이 오프라인 자문 및 판례 검색을 지원합니다.',
    steps: [
      '질의할 내용 또는 분석할 계약서 초안을 기입합니다.',
      '검색 범위를 지정하고 로컬 지식베이스 질의를 수행합니다.',
      '검출된 로컬 참조 문헌 조항과 AI의 종합 법률 분석 내용을 검토합니다.'
    ],
    tips: '로컬 데이터 동기화 상태가 최신인지 지식베이스 탭에서 점검해 주세요.'
  },
  erd: {
    title: 'AI ERD 설계',
    desc: '요구사항 정의서나 테이블 스펙 텍스트로부터 논리/물리 ERD 다이어그램을 AI가 자동으로 설계하고 DDL 쿼리를 추출합니다.',
    steps: [
      '설계할 데이터베이스 요구사항 명세서를 기재합니다.',
      '[AI ERD 설계 시작]을 클릭해 테이블 간 관계 및 스키마를 구성합니다.',
      'Mermaid 기반의 관계도 시각화 및 DDL SQL 스크립트를 생성합니다.',
      '수정 요구사항을 입력해 스키마를 실시간 업데이트합니다.'
    ],
    tips: '테이블 관계선이 복잡할 경우 추가 요구사항 필드에 구체적인 관계를 지정해 보정할 수 있습니다.'
  },
  meeting: {
    title: 'AI 회의록 생성',
    desc: '음성 녹음 파일 또는 전사 텍스트를 바탕으로 회의의 핵심 주제, 결정 사항, 액션 아이템을 요약 추출합니다.',
    steps: [
      '회의 오디오 파일(MP3, WAV 등)을 업로드하거나 회의록 텍스트를 기입합니다.',
      '화자 수와 화자 분리 정밀도 옵션을 설정합니다.',
      '[회의록 분석 시작]을 클릭해 회의 주제별 요약을 수행합니다.',
      '자동 완성된 공식 회의록 산출물을 다운로드합니다.'
    ],
    tips: '화자가 여러 명일 경우 [오디오 파일 분석] 탭을 권장합니다.'
  },
  library: {
    title: '참고자료 라이브러리',
    desc: '프로젝트 산출물 작성에 지침이 되는 표준 고시 문서, 템플릿 서식 등을 열람하고 활용할 수 있는 로컬 라이브러리입니다.',
    steps: [
      '참조가 필요한 정부 가이드라인이나 양식 카테고리를 확인합니다.',
      '필요한 참고 문헌 파일명을 클릭해 상세 내용을 브라우저로 뷰잉합니다.',
      '해당 문헌의 핵심 작성 가이드 요약본을 복사해 문서 작성에 인용합니다.'
    ],
    tips: '로컬 파일 시스템에 문서를 추가하면 실시간으로 라이브러리 목록에 추가됩니다.'
  },
  aippt: {
    title: 'AI PPT 디자이너',
    desc: '긴 보고서나 텍스트를 입력하면 AI가 의미 단락별로 슬라이드를 쪼개고 최적의 레이아웃 디자인을 자동 반영해 세련된 PPTX 파일을 빌드합니다.',
    steps: [
      'PPT로 변환할 원본 텍스트를 붙여넣거나 문서를 드롭합니다.',
      '마스터 템플릿 테마 색상 및 슬라이드 수를 지정합니다.',
      '[AI 슬라이드 생성]을 클릭해 레이아웃 생성을 수행합니다.',
      '시각화가 완료된 결과물을 확인하고 저장 경로를 지정해 다운로드합니다.'
    ],
    tips: '템플릿이 되는 PPTX를 함께 업로드하면 기존 마스터 슬라이드의 스타일을 계승할 수 있습니다.'
  },
  rag: {
    title: '프로젝트 RAG 지식베이스',
    desc: '로컬 문서 저장소의 다양한 형식(.pdf, .pptx, .docx, .txt)의 산출물 파일들을 통합 인덱싱하여 벡터 데이터베이스로 관리합니다.',
    steps: [
      '지식베이스로 등록할 다양한 형태의 파일을 드롭존에 업로드합니다.',
      '전체 인덱스를 업데이트하기 위해 [RAG 동기화 및 재빌드]를 수행합니다.',
      '우측 Q&A 채팅창에서 질문을 입력하면 등록된 모든 문서의 내용을 실시간 종합해 답변합니다.'
    ],
    tips: '문서가 업데이트될 때마다 동기화를 실행해야 AI 자문 기능에서 최신 내용을 반영합니다.'
  },
  ismpda: {
    title: 'ISMP DA 검증 대시보드',
    desc: 'ISMP 및 DA 품질 점검에 대한 전체 프로젝트 산출물의 현황과 검수 합격률, 수정 필요 요소를 종합 대시보드로 시각화합니다.',
    steps: [
      '대시보드에서 각 영역별 검수 진행률과 불합격 지표를 모니터링합니다.',
      '품질 경고가 발생한 특정 과제 조항 및 파일명을 클릭합니다.',
      '제시되는 조치 가이드라인에 따라 원본 문서를 수정한 후 재검수를 진행합니다.'
    ],
    tips: '주기적으로 검증을 돌리면 실시간 합격률 추이가 차트로 기록됩니다.'
  },
  'hwpx-report': {
    title: 'HWPX 생성(표준보고서)',
    desc: '제미나이 AI가 사용자 요청 사항에 부합하는 정밀하고 체계적인 한글 표준 보고서 파일(.hwpx)을 스키마 구조에 맞춰 자동 퍼블리싱합니다.',
    steps: [
      '작성하고자 하는 기획서나 보고서의 개요 및 주제를 입력합니다.',
      '보고서에 필수로 포함할 항목이나 스타일 템플릿을 선택합니다.',
      '[HWPX 보고서 자동 생성] 버튼을 클릭해 백그라운드 파싱을 진행합니다.',
      '최종 렌더링된 한글 문서 파일을 로컬에 저장하고 오피스 툴로 확인합니다.'
    ],
    tips: '목차 템플릿을 상세히 입력할수록 보다 깊이 있는 다량의 분량이 생성됩니다.'
  }
};

function App() {
  const [activeTab, setActiveTab] = useState('ppt');
  const [show_guide, set_show_guide] = useState(true);
  const [show_manual_modal, set_show_manual_modal] = useState(false);
  // apiKeys: 문자열 배열로 관리, 기존 localStorage 콤마 구분 값과 호환
  const [apiKeys, setApiKeys] = useState(() => {
    const stored = localStorage.getItem('gemini_api_key') || '';
    const parsed = stored.split(',').map(k => k.trim()).filter(Boolean);
    return parsed.length > 0 ? parsed : [''];
  });
  const [newKeyInput, setNewKeyInput] = useState('');
  const [key_error_msg, set_key_error_msg] = useState(''); // API 키 에러 메시지
  // 키 유효성 검사 함수: AIza 또는 AQ로 시작하고 최소 30자 이상
  const is_valid_key = (k) => {
    const s = k.trim();
    return (s.startsWith('AIza') || s.startsWith('AQ')) && s.length >= 30;
  };
  // 하위 컴포넌트에 전달할 콤마 구분 문자열 (유효한 키만)
  const apiKey = apiKeys.filter(k => is_valid_key(k)).join(',');
  // 로컬 스토리지 보존용 원본 문자열 (빈값 제외)
  const raw_api_key_str = apiKeys.filter(k => k.trim() !== '').join(',');
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
      'hwpx-report': '(작업중)HWPX생성(표준보고서)',
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
    localStorage.setItem('gemini_api_key', raw_api_key_str);
  }, [raw_api_key_str]);

  // API 키 추가 (유효성 검사 및 에러 메시지 제공)
  const handleAddKey = () => {
    const trimmed = newKeyInput.trim();
    if (!trimmed) {
      set_key_error_msg('키를 입력해 주세요.');
      return;
    }
    if (!is_valid_key(trimmed)) {
      set_key_error_msg(`❌ 유효하지 않은 키 형식입니다. Gemini API 키는 'AIza'로 시작하고 39자 이상이어야 합니다. (현재: '${trimmed.substring(0, 10)}...', ${trimmed.length}자)`);
      return;
    }
    if (apiKeys.includes(trimmed)) {
      set_key_error_msg('이미 등록된 키입니다.');
      return;
    }
    set_key_error_msg('');
    setApiKeys(prev => {
      if (prev.length === 1 && prev[0].trim() === '') return [trimmed];
      return [...prev, trimmed];
    });
    setNewKeyInput('');
  };

  // API 키 삭제
  const handleRemoveKey = (idx) => {
    setApiKeys(prev => {
        const next = prev.filter((_, i) => i !== idx);
        return next.length > 0 ? next : [''];
    });
  };

  // API 키 개별 수정 (실시간 replace 제거하여 한글/커서 튐/붙여넣기 꼬임 전면 해소)
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
    { id: 'ppt', label: 'PPT 생성(표준산출물)', icon: FileText, color: '#f97316', useGemini: false },
    { id: 'ppt-verify', label: 'PPT 검증(표준산출물)', icon: ShieldAlert, color: '#e11d48', useGemini: true },
    { id: 'main', label: 'AI 산출물 검증', icon: Shield, color: 'var(--accent-blue)', useGemini: true },
    { id: 'typo', label: 'AI 교정교열', icon: CheckCircle2, color: 'var(--accent-purple)', useGemini: true },
    { id: 'law', label: 'AI 법률 자문(제미나이)', icon: MessageSquare, color: 'var(--success-color)', useGemini: true },
    { id: 'law-mcp', label: 'AI 법률 자문(로컬 RAG)', icon: MessageSquare, color: 'var(--accent-purple)', useGemini: true },
    { id: 'erd', label: 'AI ERD 설계', icon: Database, color: 'var(--warning-color)', useGemini: true },
    { id: 'meeting', label: 'AI 회의록 생성', icon: Mic2, color: '#8b5cf6', useGemini: true },
    { id: 'library', label: '참고자료 라이브러리', icon: Activity, color: '#64748b', useGemini: false },
    { id: 'aippt', label: '(작업중)AI PPT 디자이너', icon: Presentation, color: '#ec4899', useGemini: true },
    { id: 'rag', label: '프로젝트 RAG 지식베이스', icon: Database, color: 'var(--accent-blue)', useGemini: true },
    { id: 'ismpda', label: '(작업중)ISMP DA 검증 대시보드', icon: Sliders, color: 'var(--accent-purple)', useGemini: false },
    { id: 'hwpx-report', label: '(작업중)HWPX생성(표준보고서)', icon: FileText, color: '#10b981', useGemini: true },
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
                <span>AI 산출물 검수 v2.7</span>
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
                <div className="nav-icon-wrapper" style={{ color: isActive ? tab.color : 'inherit', position: 'relative' }}>
                  <Icon size={20} />
                  {tab.useGemini && (
                    <Sparkles 
                      size={10} 
                      style={{ 
                        position: 'absolute', 
                        top: '-4px', 
                        right: '-4px', 
                        color: '#38bdf8',
                        filter: 'drop-shadow(0 0 2px rgba(56, 189, 248, 0.7))'
                      }} 
                    />
                  )}
                </div>
                {!isSidebarCollapsed && (
                  <span style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                    {tab.label}
                    {tab.useGemini && (
                      <Sparkles 
                        size={13} 
                        style={{ 
                          color: '#38bdf8', 
                          filter: 'drop-shadow(0 0 2px rgba(56, 189, 248, 0.5))',
                          flexShrink: 0
                        }} 
                      />
                    )}
                  </span>
                )}
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
            <div className="breadcrumb" style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <span className="breadcrumb-parent">Validator Space</span>
              <span className="breadcrumb-separator">/</span>
              <span className="breadcrumb-current" style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                {activeTabData?.label}
                {activeTabData?.useGemini && (
                  <span style={{ 
                    display: 'inline-flex', 
                    alignItems: 'center', 
                    gap: '4px', 
                    fontSize: '10px', 
                    fontWeight: 700, 
                    color: '#38bdf8', 
                    background: 'rgba(56, 189, 248, 0.1)', 
                    border: '1px solid rgba(56, 189, 248, 0.25)', 
                    padding: '2px 8px', 
                    borderRadius: '12px',
                    marginLeft: '6px',
                    filter: 'drop-shadow(0 0 1px rgba(56, 189, 248, 0.2))'
                  }}>
                    <Sparkles size={10} /> Gemini AI
                  </span>
                )}
              </span>
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
              <span style={{ fontSize: '12px', color: 'var(--accent-blue)', fontWeight: 700, fontFamily: 'monospace' }}>2026.8.1 v2.8</span>
            </div>
            
            <button 
              className={`settings-btn ${show_manual_modal ? 'active' : ''}`} 
              onClick={() => set_show_manual_modal(true)}
              title="통합 사용 매뉴얼 보기"
              style={{ marginRight: '4px' }}
            >
              <HelpCircle size={20} />
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
                              background: is_valid_key(k)
                                ? 'rgba(16,185,129,0.08)'
                                : 'rgba(255,255,255,0.04)',
                              border: `1px solid ${is_valid_key(k) ? 'rgba(16,185,129,0.4)' : 'rgba(239,68,68,0.4)'}`,
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

                  {/* 키 에러 메시지 */}
                  {key_error_msg && (
                    <div style={{
                      marginTop: '6px',
                      padding: '8px 10px',
                      background: 'rgba(239,68,68,0.12)',
                      border: '1px solid rgba(239,68,68,0.35)',
                      borderRadius: '7px',
                      color: '#f87171',
                      fontSize: '11px',
                      lineHeight: 1.5,
                      wordBreak: 'break-all',
                    }}>
                      {key_error_msg}
                    </div>
                  )}
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
          {/* 가이드 배너 */}
          {show_guide && activeTab !== 'rag' && tab_guides[activeTab] && (
            <div className="glass-panel animate-slide-up" style={{ 
              padding: '16px 20px', 
              marginBottom: '20px', 
              border: '1px solid rgba(59, 130, 246, 0.2)', 
              borderRadius: '12px',
              background: 'linear-gradient(to right, rgba(59, 130, 246, 0.05), rgba(99, 102, 241, 0.03))',
              position: 'relative'
            }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '8px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <Info size={16} color="var(--accent-blue)" />
                  <span style={{ fontSize: '14.5px', fontWeight: 700, color: 'var(--text-primary)' }}>
                    {tab_guides[activeTab].title} 사용 가이드
                  </span>
                </div>
                <button 
                  onClick={() => set_show_guide(false)}
                  style={{ background: 'none', border: 'none', color: 'var(--text-muted)', cursor: 'pointer', display: 'flex', alignItems: 'center', fontSize: '12px', gap: '4px' }}
                  title="이 가이드 접기"
                >
                  <X size={14} /> 접기
                </button>
              </div>
              <p style={{ fontSize: '13px', color: 'var(--text-secondary)', margin: '0 0 10px', lineHeight: 1.45 }}>
                {tab_guides[activeTab].desc}
              </p>
              <div style={{ background: 'rgba(0,0,0,0.15)', padding: '10px 12px', borderRadius: '8px', marginBottom: '8px' }}>
                <span style={{ fontSize: '11px', color: 'var(--text-muted)', fontWeight: 700, textTransform: 'uppercase', display: 'block', marginBottom: '4px' }}>진행 단계</span>
                <ul style={{ margin: 0, paddingLeft: '16px', fontSize: '12.5px', color: 'var(--text-secondary)', display: 'flex', flexDirection: 'column', gap: '4px' }}>
                  {tab_guides[activeTab].steps.map((step, idx) => (
                    <li key={idx} style={{ listStyleType: 'decimal' }}>{step}</li>
                  ))}
                </ul>
              </div>
              {tab_guides[activeTab].tips && (
                <div style={{ fontSize: '12px', color: '#38bdf8', display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <Sparkles size={12} /> <strong>Tip:</strong> {tab_guides[activeTab].tips}
                </div>
              )}
            </div>
          )}

          {!show_guide && tab_guides[activeTab] && (
            <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: '12px' }}>
              <button 
                onClick={() => set_show_guide(true)}
                style={{ 
                  background: 'rgba(59, 130, 246, 0.08)', 
                  border: '1px solid rgba(59, 130, 246, 0.2)', 
                  borderRadius: '8px', 
                  padding: '6px 12px', 
                  color: 'var(--accent-blue)', 
                  fontSize: '12px', 
                  fontWeight: 600,
                  cursor: 'pointer',
                  display: 'flex',
                  alignItems: 'center',
                  gap: '6px'
                }}
              >
                <Info size={14} /> 사용 가이드 펼치기
              </button>
            </div>
          )}

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


      {/* 도움말 매뉴얼 모달 */}
      {show_manual_modal && (
        <div style={{ 
          position: 'fixed', top: 0, left: 0, width: '100%', height: '100%', 
          background: 'rgba(0,0,0,0.6)', backdropFilter: 'blur(8px)', 
          zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center',
          padding: '20px'
        }}>
          <div className="glass-panel animate-scale-in" style={{ 
            width: '100%', maxWidth: '800px', maxHeight: '90vh', 
            display: 'flex', flexDirection: 'column',
            border: '1px solid var(--panel-border)', borderRadius: '20px',
            background: 'linear-gradient(to bottom right, rgba(20,20,30,0.95), rgba(10,10,15,0.98))',
            boxShadow: '0 24px 64px rgba(0,0,0,0.5)', overflow: 'hidden'
          }}>
            {/* 모달 헤더 */}
            <div style={{ 
              padding: '20px 24px', borderBottom: '1px solid var(--panel-border)', 
              display: 'flex', justifyContent: 'space-between', alignItems: 'center',
              background: 'rgba(255,255,255,0.02)'
            }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                <HelpCircle size={24} color="var(--accent-blue)" />
                <h3 style={{ margin: 0, fontSize: '18px', fontWeight: 800, color: 'var(--text-primary)' }}>
                  건강한 프로젝트 통합 매뉴얼 (v2.7)
                </h3>
              </div>
              <button 
                onClick={() => set_show_manual_modal(false)}
                style={{ 
                  background: 'none', border: 'none', color: 'var(--text-muted)', 
                  cursor: 'pointer', padding: '6px' 
                }}
              >
                <X size={20} />
              </button>
            </div>
            
            {/* 모달 본문 */}
            <div style={{ padding: '24px', overflowY: 'auto', flex: 1, display: 'flex', flexDirection: 'column', gap: '20px' }}>
              <div style={{ display: 'flex', gap: '12px', background: 'rgba(59, 130, 246, 0.05)', padding: '16px', borderRadius: '12px', border: '1px solid rgba(59, 130, 246, 0.15)' }}>
                <Sparkles size={20} color="#38bdf8" style={{ flexShrink: 0, marginTop: '2px' }} />
                <div>
                  <h4 style={{ margin: '0 0 6px', fontSize: '14px', fontWeight: 700, color: 'var(--text-primary)' }}>제미나이 AI 기반 산출물 종합 플랫폼</h4>
                  <p style={{ margin: 0, fontSize: '13px', color: 'var(--text-secondary)', lineHeight: 1.5 }}>
                    본 어플리케이션은 파워포인트(PPTX), 한글(HWPX) 파일 등 다양한 산출물의 자동 매핑 생성, 오탈자 및 데이터 중복 검증, AI 법률 및 ERD 데이터베이스 설계를 아우르는 올인원 품질 검수 지원 도구입니다.
                  </p>
                </div>
              </div>

              <div>
                <h4 style={{ fontSize: '14.5px', fontWeight: 700, color: 'var(--text-primary)', borderLeft: '3px solid var(--accent-blue)', paddingLeft: '8px', margin: '0 0 12px' }}>
                  🔑 필수 선행 설정 (API 키 연동)
                </h4>
                <p style={{ margin: '0 0 8px', fontSize: '13px', color: 'var(--text-secondary)', lineHeight: 1.5 }}>
                  제미나이 AI를 활용한 컴포넌트들을 정상 이용하려면 상단의 <strong>설정 아이콘(<Settings size={12} style={{ verticalAlign: 'middle' }} />)</strong>을 클릭해 최소 1개 이상의 Google Gemini API Key를 연동해야 합니다.
                </p>
                <div style={{ fontSize: '12px', color: 'var(--text-muted)', background: 'rgba(0,0,0,0.2)', padding: '10px', borderRadius: '8px' }}>
                  * API 키는 <code>AIzaSy</code> 또는 <code>AQ.</code>로 시작하는 정식 Google 키여야 합니다. 다중 키를 등록하면 한도가 초과될 시 자동으로 백업 키로 전환됩니다.
                </div>
              </div>

              <div>
                <h4 style={{ fontSize: '14.5px', fontWeight: 700, color: 'var(--text-primary)', borderLeft: '3px solid var(--accent-blue)', paddingLeft: '8px', margin: '0 0 12px' }}>
                  💡 주요 탭 기능별 요약 매뉴얼
                </h4>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                  {Object.entries(tab_guides).map(([id, guide]) => (
                    <div key={id} style={{ padding: '12px', background: 'rgba(255,255,255,0.02)', borderRadius: '10px', border: '1px solid var(--glass-border)' }}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '6px', marginBottom: '4px' }}>
                        <span style={{ fontSize: '13.5px', fontWeight: 700, color: 'var(--text-primary)' }}>{guide.title}</span>
                        {tabs.find(t => t.id === id)?.useGemini && (
                          <span style={{ fontSize: '10px', color: '#38bdf8', background: 'rgba(56, 189, 248, 0.1)', padding: '1px 5px', borderRadius: '4px', display: 'flex', alignItems: 'center', gap: '2px' }}>
                            <Sparkles size={8} /> AI
                          </span>
                        )}
                      </div>
                      <p style={{ margin: '0 0 6px', fontSize: '12px', color: 'var(--text-secondary)' }}>{guide.desc}</p>
                      <div style={{ fontSize: '11px', color: 'var(--text-muted)' }}>
                        <strong>사용 프로세스:</strong> {guide.steps.join(' ➔ ')}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            {/* 모달 푸터 */}
            <div style={{ 
              padding: '16px 24px', borderTop: '1px solid var(--panel-border)', 
              display: 'flex', justifyContent: 'flex-end',
              background: 'rgba(0,0,0,0.2)'
            }}>
              <button 
                onClick={() => set_show_manual_modal(false)}
                style={{ 
                  padding: '8px 20px', background: 'var(--accent-blue)', 
                  border: 'none', borderRadius: '8px', color: 'white', 
                  fontSize: '13px', fontWeight: 600, cursor: 'pointer' 
                }}
              >
                닫기
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default App;
