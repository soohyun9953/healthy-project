import React from 'react';
import { ShieldCheck, Scale, FileText, Settings, LayoutDashboard, PenTool, Library, ChevronLeft, ChevronRight, Fingerprint, Database, Presentation } from 'lucide-react';

function Sidebar({ activeMenu, setActiveMenu }) {
  const menuItems = [
    { id: 'validator', label: 'AI 산출물 검증', icon: <FileText size={20} /> },
    { id: 'typo', label: 'AI 문서 품질/오탈자 점검', icon: <PenTool size={20} /> },
    { id: 'law_general', label: 'AI 법률 자문 (Gemini)', icon: <Scale size={20} /> },
    { id: 'law', label: 'AI 법률 자문 (로컬 RAG)', icon: <Fingerprint size={20} /> },
    { id: 'erd', label: 'ERD 자동 설계', icon: <Database size={20} /> },
    { id: 'ppt', label: 'PPT 자동 생성', icon: <Presentation size={20} /> },
    { id: 'reference', label: '참고 자료 관리', icon: <Library size={20} /> },
  ];

  return (
    <aside className="animate-fade-in" style={{ 
      width: '280px', 
      flexShrink: 0, 
      display: 'flex', 
      flexDirection: 'column', 
      background: 'var(--bg-dark)', 
      borderRight: '1px solid var(--glass-border)',
      position: 'relative',
      zIndex: 10
    }}>
      <div style={{ padding: '32px 24px', borderBottom: '1px solid var(--glass-border)' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
          <div style={{ 
            width: '36px', 
            height: '36px', 
            borderRadius: '10px', 
            background: 'linear-gradient(135deg, var(--accent-purple), var(--accent-blue))', 
            display: 'flex', 
            alignItems: 'center', 
            justifyContent: 'center',
            boxShadow: '0 8px 24px rgba(168, 85, 247, 0.2)'
          }}>
            <LayoutDashboard size={22} color="white" />
          </div>
          <div>
            <h2 style={{ margin: 0, fontSize: '18px', fontWeight: 800, color: 'var(--text-primary)', letterSpacing: '-0.5px' }}>PM Dashboard</h2>
            <p style={{ margin: 0, fontSize: '10px', color: 'var(--text-muted)', fontWeight: 700, letterSpacing: '0.5px' }}>ENTERPRISE AI HUB</p>
          </div>
        </div>
      </div>

      <nav style={{ flex: 1, padding: '24px 12px', display: 'flex', flexDirection: 'column', gap: '4px' }}>
        {menuItems.map((item) => {
          const isActive = activeMenu === item.id;
          return (
            <button
              key={item.id}
              onClick={() => setActiveMenu(item.id)}
              className="interactive"
              style={{
                width: '100%',
                display: 'flex',
                alignItems: 'center',
                gap: '12px',
                padding: '12px 16px',
                borderRadius: '12px',
                background: isActive ? 'rgba(168, 85, 247, 0.1)' : 'transparent',
                border: 'none',
                color: isActive ? 'var(--text-primary)' : 'var(--text-secondary)',
                cursor: 'pointer',
                transition: 'all 0.2s',
                textAlign: 'left',
                fontWeight: isActive ? 700 : 500
              }}
            >
              <div style={{ color: isActive ? 'var(--accent-purple)' : 'var(--text-muted)' }}>
                {item.icon}
              </div>
              <span style={{ fontSize: '14px' }}>{item.label}</span>
            </button>
          );
        })}
      </nav>

      <div style={{ padding: '24px', borderTop: '1px solid var(--glass-border)' }}>
        <button
          className="interactive"
          style={{ width: '100%', display: 'flex', alignItems: 'center', gap: '12px', padding: '12px', borderRadius: '10px', background: 'rgba(255,255,255,0.03)', border: '1px solid var(--glass-border)', color: 'var(--text-muted)', cursor: 'not-allowed', fontSize: '13px', opacity: 0.6 }}
          disabled
        >
          <Settings size={18} /> 환경 설정
        </button>
      </div>
    </aside>
  );
}

export default Sidebar;
