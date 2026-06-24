import React, { useState, useEffect, useRef, useCallback } from 'react';
import { Database, Download, FileText, Upload, Loader2, Play, CheckCircle2, AlertCircle, Info, Trash2, X, ChevronDown, Code2, BarChart3, BookOpen, Eye, Copy, Check, FileSpreadsheet, Maximize2, Sparkles } from 'lucide-react';
import * as XLSX from 'xlsx';
import mermaid from 'mermaid';
import { analyzeERDWithLLM } from '../erdAnalyzer';
import { processFile, ALL_ACCEPT } from '../utils/fileExtractor';

// ── JSON 원문 뷰어 모달 ────────────────────────────────────────
const JsonViewerModal = ({ data, onClose }) => {
  const [copied, setCopied] = useState(false);
  const jsonStr = JSON.stringify(data, null, 2);

  const handleCopy = () => {
    navigator.clipboard.writeText(jsonStr);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <div style={{
      position: 'fixed', top: 0, left: 0, right: 0, bottom: 0,
      background: 'rgba(0,0,0,0.85)', backdropFilter: 'blur(12px)',
      display: 'flex', alignItems: 'center', justifyContent: 'center',
      zIndex: 10000
    }}>
      <div className="glass-panel animate-fade-in" style={{
        width: '90vw', maxWidth: '1000px', height: '85vh',
        display: 'flex', flexDirection: 'column',
        background: 'var(--bg-secondary)', border: '1px solid var(--panel-border)',
        borderRadius: '20px', overflow: 'hidden',
        boxShadow: '0 30px 60px rgba(0,0,0,0.8)'
      }}>
        {/* 모달 헤더 */}
        <div style={{
          padding: '20px 28px', borderBottom: '1px solid var(--panel-border)',
          display: 'flex', alignItems: 'center', justifyContent: 'space-between',
          background: 'rgba(0,0,0,0.3)'
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <Code2 size={22} color="var(--accent-purple)" />
            <div>
              <h3 style={{ margin: 0, fontSize: '18px', fontWeight: 700 }}>ERD 설계 결과 원문 (JSON)</h3>
              <p style={{ margin: '2px 0 0', fontSize: '12px', color: 'var(--text-muted)' }}>
                mermaidCode · entities · relationships · normalizationNotes 포함
              </p>
            </div>
          </div>
          <div style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
            <button onClick={handleCopy} className="interactive" style={{
              padding: '8px 16px', borderRadius: '8px', border: '1px solid var(--panel-border)',
              background: copied ? 'rgba(16,185,129,0.1)' : 'rgba(255,255,255,0.05)',
              color: copied ? 'var(--success-color)' : 'var(--text-primary)',
              fontSize: '13px', fontWeight: 600, cursor: 'pointer',
              display: 'flex', alignItems: 'center', gap: '6px', transition: 'all 0.2s'
            }}>
              {copied ? <Check size={14} /> : <Copy size={14} />}
              {copied ? '복사 완료!' : '전체 복사'}
            </button>
            <button onClick={() => {
              const blob = new Blob([jsonStr], { type: 'application/json' });
              const url = URL.createObjectURL(blob);
              const a = document.createElement('a');
              a.href = url;
              a.download = `ERD_Design_${new Date().toISOString().split('T')[0]}.json`;
              a.click();
            }} className="interactive" style={{
              padding: '8px 16px', borderRadius: '8px',
              background: 'rgba(168,85,247,0.1)', border: '1px solid rgba(168,85,247,0.2)',
              color: 'var(--accent-purple)', fontSize: '13px', fontWeight: 600,
              cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '6px'
            }}>
              <Download size={14} /> JSON 다운로드
            </button>
            <button onClick={onClose} style={{
              background: 'rgba(255,255,255,0.05)', border: '1px solid var(--panel-border)',
              color: 'var(--text-secondary)', cursor: 'pointer', padding: '8px',
              borderRadius: '8px', display: 'flex', alignItems: 'center'
            }}>
              <X size={20} />
            </button>
          </div>
        </div>

        {/* JSON 본문 */}
        <div style={{ flex: 1, overflowY: 'auto', padding: '0' }}>
          <pre style={{
            margin: 0, padding: '24px 28px', fontFamily: "'Fira Code', 'Consolas', monospace",
            fontSize: '13px', lineHeight: '1.7', color: '#e2e8f0',
            background: '#0d1117', height: '100%', overflowY: 'auto'
          }}>
            {/* 간단한 구문 강조를 위해 파트별로 분리 */}
            {jsonStr.split('\n').map((line, i) => {
              const keyMatch = line.match(/^(\s*)"([^"]+)":/);
              const strMatch = line.match(/:\s*"(.+)"[,]?$/);
              const numMatch = line.match(/:\s*(\d+)[,]?$/);

              if (keyMatch) {
                return (
                  <div key={i}>
                    <span style={{ color: '#6b8cff' }}>{line.match(/^\s*/)[0]}</span>
                    <span style={{ color: '#79c0ff' }}>"{keyMatch[2]}"</span>
                    <span style={{ color: '#c9d1d9' }}>: </span>
                    <span style={{ color: strMatch ? '#a5d6ff' : numMatch ? '#f8a261' : '#c9d1d9' }}>
                      {line.slice(keyMatch[0].length)}
                    </span>
                  </div>
                );
              }
              return <div key={i}><span style={{ color: '#c9d1d9' }}>{line}</span></div>;
            })}
          </pre>
        </div>
      </div>
    </div>
  );
};

// ── Mermaid 다이어그램 렌더링 컴포넌트 ─────────────────────
const MermaidDiagram = ({ chart }) => {
  const ref = useRef(null);
  const [error, setError] = useState(null);

  useEffect(() => {
    if (ref.current && chart) {
      const renderChart = async () => {
        try {
          setError(null);
          ref.current.innerHTML = '';
          
          // 데이터 정제: 불필요한 이스케이프 문자 및 마크다운 태그 제거
          let cleanChart = chart
            .replace(/```mermaid/g, '')
            .replace(/```/g, '')
            .replace(/\\n/g, '\n') // literal \n -> 실제 개행
            .replace(/\\"/g, '"')  // literal \" -> "
            .replace(/\\/g, '')    // 남은 모든 단일 역슬래시(\) 제거
            .trim();

          // erDiagram 키워드가 누락된 경우 자동 추가 시도
          if (!cleanChart.startsWith('erDiagram')) {
             if (cleanChart.includes('erDiagram')) {
                cleanChart = cleanChart.substring(cleanChart.indexOf('erDiagram'));
             } else {
                // 정말 Mermaid 형식이 아니면 간단한 변환 시도 (필요시)
                console.warn("Mermaid format might be incorrect, attempting to fix prefix.");
                cleanChart = 'erDiagram\n' + cleanChart;
             }
          }
          
          mermaid.initialize({
            startOnLoad: false,
            theme: 'dark',
            securityLevel: 'loose',
            fontFamily: 'Inter, system-ui, sans-serif',
            er: {
              useMaxWidth: true,
              fontSize: 14
            }
          });

          const { svg } = await mermaid.render(`mermaid-${Math.random().toString(36).substr(2, 9)}`, cleanChart);
          ref.current.innerHTML = svg;
          
          // SVG 스타일 조정
          const svgEl = ref.current.querySelector('svg');
          if (svgEl) {
            svgEl.style.maxWidth = '100%';
            svgEl.style.height = 'auto';
            svgEl.style.display = 'block';
            svgEl.style.margin = '0 auto';
          }
        } catch (e) {
          console.error("Mermaid Render Error:", e);
          setError(e.message);
        }
      };
      renderChart();
    }
  }, [chart]);

  if (error) {
    return (
      <div style={{
        padding: '24px', borderRadius: '12px', background: 'rgba(239,68,68,0.05)',
        border: '1px solid rgba(239,68,68,0.2)', color: 'var(--danger-color)',
        fontSize: '13px', display: 'flex', flexDirection: 'column', gap: '10px'
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', fontWeight: 700 }}>
          <AlertCircle size={16} /> 다이어그램 생성 실패
        </div>
        <p style={{ margin: 0, opacity: 0.8 }}>AI가 생성한 Mermaid 코드의 문법이 올바르지 않습니다.</p>
        <pre style={{
          margin: 0, padding: '12px', background: 'rgba(0,0,0,0.2)',
          borderRadius: '8px', fontSize: '11px', overflowX: 'auto', color: 'var(--text-muted)'
        }}>
          {chart}
        </pre>
      </div>
    );
  }

  return (
    <div className="glass-panel" style={{
      padding: '24px', background: 'rgba(255,255,255,0.02)',
      borderRadius: '16px', border: '1px solid var(--panel-border)',
      minHeight: '200px', display: 'flex', alignItems: 'center', justifyContent: 'center',
      overflow: 'auto'
    }}>
      <div ref={ref} style={{ width: '100%' }} />
    </div>
  );
};

// ── 결과 탭 뷰 ────────────────────────────────────────────────
const ResultSection = ({ result, onOpenJsonViewer }) => {
  const [activeTab, setActiveTab] = useState('entities');
  const [copied, setCopied] = useState(false);

  const tabs = [
    { id: 'diagram', label: '다이어그램 시각화', icon: Eye },
    { id: 'entities', label: '엔티티 & 속성 명세', icon: Database },
    { id: 'relations', label: '관계 & 정규화', icon: BookOpen },
    { id: 'dbml', label: 'DBML 소스 (dbdiagram)', icon: Code2 },
  ];

  const getDbmlCode = useCallback(() => {
    if (!result || !result.entities) return '';
    let dbml = '';
    result.entities.forEach(entity => {
      dbml += `Table "${entity.description || entity.name}" {\n`;
      if (entity.attributes) {
        entity.attributes.forEach(attr => {
          let rawType = attr.type || '';
          let pureType = rawType;
          const match = rawType.match(/^([a-zA-Z_]+)(?:\\((\\d+)\\))?/);
          if (match) pureType = match[1];

          let mods = [];
          if (attr.key === 'PK') mods.push('pk');
          if (attr.desc) {
            const safeDesc = attr.desc.replace(/\\\\n/g, ' ').replace(/\\n/g, ' ').replace(/'/g, "''").replace(/\\\\/g, '');
            mods.push(`note: '${safeDesc}'`);
          }
          let modsStr = mods.length > 0 ? ` [${mods.join(', ')}]` : '';
          dbml += `  "${attr.name}" ${pureType.toUpperCase()}${modsStr}\n`;
        });
      }
      if (entity.reason) {
        const safeReason = entity.reason.replace(/\\n/g, ' ').replace(/'/g, "''").replace(/\\/g, '');
        dbml += `  Note: '${safeReason}'\n`;
      }
      dbml += `}\n\n`;
    });

    if (result.relationships) {
      result.relationships.forEach(rel => {
        const fromEnt = result.entities.find(e => e.name === rel.from);
        const toEnt = result.entities.find(e => e.name === rel.to);
        const fromName = fromEnt ? (fromEnt.description || fromEnt.name) : rel.from;
        const toName = toEnt ? (toEnt.description || toEnt.name) : rel.to;
        
        let fromCol = fromEnt?.attributes?.find(a => a.key === 'PK')?.name || 'id';
        let toCol = toEnt?.attributes?.find(a => a.key === 'FK')?.name || toEnt?.attributes?.[0]?.name || 'id';

        let link = '-';
        if (rel.type.includes('1:N')) link = '<';
        if (rel.type.includes('N:M') || rel.type.includes('M:N')) link = '<>';

        const safeRelDesc = (rel.desc || '').replace(/\\n/g, ' ').replace(/\\/g, '');
        dbml += `Ref: "${fromName}"."${fromCol}" ${link} "${toName}"."${toCol}" // ${safeRelDesc}\n`;
      });
    }
    return dbml;
  }, [result]);

  const handleCopyCode = (codeText) => {
    navigator.clipboard.writeText(codeText);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const handleExportExcel = () => {
    if (!result || !result.entities) return;
    
    // 1. 엔티티 정의서 시트 데이터 (양식 적용)
    const entityData = result.entities.map((entity, idx) => {
      // 속성 배열을 개행문자(\n)로 연결, PK나 FK가 있으면 접미사 추가
      let attrString = '';
      if (entity.attributes && entity.attributes.length > 0) {
        attrString = entity.attributes.map(attr => {
          let suffix = '';
          if (attr.key === 'PK') suffix = '(PK)';
          else if (attr.key === 'FK') suffix = '(FK)';
          return attr.name + suffix;
        }).join('\n');
      }

      return {
        'No': idx + 1,
        '엔티티(한글)': entity.description || '',
        '엔티티(영문)': entity.name || '',
        '설명': entity.reason || '',
        '주요 속성': attrString
      };
    });

    // 2. 속성 정의서 시트 데이터 (양식 적용)
    const attributeData = [];
    let attrIdx = 1;
    result.entities.forEach(entity => {
      if (entity.attributes && entity.attributes.length > 0) {
        entity.attributes.forEach(attr => {
          // 데이터타입에서 길이(숫자) 분리. ex: varchar(50) -> 타입: varchar, 길이: 50
          let rawType = attr.type || '';
          let pureType = rawType;
          let length = '';
          const match = rawType.match(/^([a-zA-Z_]+)(?:\\((\\d+)\\))?/);
          if (match) {
            pureType = match[1];
            if (match[2]) length = match[2];
          }

          attributeData.push({
            'No': attrIdx++,
            '엔터티': entity.description || entity.name,
            '속성명(한)': attr.name || '',
            '속성명(영)': '',
            'PK/FK': attr.key === 'PK' || attr.key === 'FK' ? attr.key : '',
            '참조엔터티': attr.key === 'FK' ? '(확인필요)' : '',
            '데이터타입': pureType.toUpperCase(),
            '길이': length,
            'NULL': attr.key === 'PK' ? 'N' : 'Y',
            '기본값': '',
            '도메인/허용값': '',
            '정의': attr.desc || ''
          });
        });
      }
    });

    // 3. 관계 및 정규화 시트 데이터 (양식 적용)
    const relationData = [];
    if (result.relationships && result.relationships.length > 0) {
      result.relationships.forEach(rel => {
        const fromEnt = result.entities.find(e => e.name === rel.from);
        const toEnt = result.entities.find(e => e.name === rel.to);
        const fromName = fromEnt ? (fromEnt.description || fromEnt.name) : rel.from;
        const toName = toEnt ? (toEnt.description || toEnt.name) : rel.to;

        relationData.push({
          'FROM 엔터티': fromName,
          '관계': rel.type,
          'TO 엔터티': toName,
          '설명': rel.desc || ''
        });
      });
    }
    
    // 정규화 노트를 같은 시트 아래쪽에 추가
    relationData.push({}); // 빈 줄
    relationData.push({
      'FROM 엔터티': '[정규화 준수 논거]',
      '관계': result.normalizationNotes || ''
    });

    // 3. DBML 소스코드 시트 데이터
    const dbmlData = getDbmlCode().split('\n').map(line => ({
      'DBML 다이어그램 소스코드 (dbdiagram.io 에서 사용 가능)': line
    }));

    const workbook = XLSX.utils.book_new();
    
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(entityData), "엔티티정의서");
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(attributeData), "속성정의서");
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(relationData), "관계및정규화");
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(dbmlData), "DBML코드");
    
    XLSX.writeFile(workbook, `ERD_상세설계서_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  return (
    <div className="animate-slide-up glass-panel" style={{ display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
      {/* 탭 헤더 */}
      <div style={{
        display: 'flex', alignItems: 'center', justifyContent: 'space-between',
        padding: '12px 24px', borderBottom: '1px solid var(--panel-border)',
        background: 'rgba(0,0,0,0.2)', flexShrink: 0, gap: '16px', flexWrap: 'wrap'
      }}>
        <div style={{ 
          display: 'flex', overflowX: 'auto', flex: 1, minWidth: '200px',
          msOverflowStyle: 'none', scrollbarWidth: 'none' 
        }} className="hide-scrollbar">
          <style>{`.hide-scrollbar::-webkit-scrollbar { display: none; }`}</style>
          {tabs.map(tab => {
            const Icon = tab.icon;
            const isActive = activeTab === tab.id;
            return (
              <button key={tab.id} onClick={() => setActiveTab(tab.id)} className="interactive" style={{
                padding: '12px 16px', background: 'transparent',
                border: 'none', borderBottom: isActive ? '2px solid var(--accent-purple)' : '2px solid transparent',
                color: isActive ? 'var(--text-primary)' : 'var(--text-muted)',
                fontWeight: isActive ? 700 : 500, fontSize: '13px', cursor: 'pointer',
                display: 'flex', alignItems: 'center', gap: '7px',
                transition: 'all 0.2s', marginBottom: '-1px', whiteSpace: 'nowrap', flexShrink: 0
              }}>
                <Icon size={14} color={isActive ? 'var(--accent-purple)' : 'inherit'} />
                {tab.label}
              </button>
            );
          })}
        </div>
        <div style={{ display: 'flex', gap: '8px', flexShrink: 0, flexWrap: 'wrap', justifyContent: 'flex-end' }}>
          <button onClick={onOpenJsonViewer} className="interactive" style={{
            padding: '6px 12px', borderRadius: '8px', border: '1px solid rgba(168,85,247,0.3)',
            background: 'rgba(168,85,247,0.08)', color: 'var(--accent-purple)',
            fontSize: '12px', fontWeight: 600, cursor: 'pointer',
            display: 'flex', alignItems: 'center', gap: '6px', whiteSpace: 'nowrap'
          }}>
            <Eye size={13} /> <span className="mobile-hide-text">JSON 원문 보기</span><span style={{ display: 'none' }} className="mobile-only-show">JSON</span>
          </button>
          <button onClick={handleExportExcel} className="interactive" style={{
            padding: '6px 12px', borderRadius: '8px', border: '1px solid rgba(16,185,129,0.3)',
            background: 'rgba(16,185,129,0.08)', color: 'var(--success-color)',
            fontSize: '12px', fontWeight: 600, cursor: 'pointer',
            display: 'flex', alignItems: 'center', gap: '6px', whiteSpace: 'nowrap'
          }}>
            <FileSpreadsheet size={13} /> <span className="mobile-hide-text">엑셀 저장</span><span style={{ display: 'none' }} className="mobile-only-show">Excel</span>
          </button>
          <button onClick={() => {
            const blob = new Blob([JSON.stringify(result, null, 2)], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `ERD_Design_${new Date().toISOString().split('T')[0]}.json`;
            a.click();
          }} className="interactive" style={{
            padding: '6px 12px', borderRadius: '8px', border: '1px solid var(--panel-border)',
            background: 'rgba(255,255,255,0.04)', color: 'var(--text-secondary)',
            fontSize: '12px', fontWeight: 600, cursor: 'pointer',
            display: 'flex', alignItems: 'center', gap: '6px', whiteSpace: 'nowrap'
          }}>
            <Download size={13} /> <span className="mobile-hide-text">JSON 저장</span>
          </button>
        </div>
      </div>

      {/* 탭 콘텐츠 */}
      <div style={{ flex: 1, padding: '24px', overflowY: 'auto' }}>

        {/* ── 탭 1: 상세 요약 ── */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: '16px', marginBottom: '24px' }}>
          <div style={{
            padding: '16px 20px', background: 'rgba(168,85,247,0.05)',
            borderRadius: '12px', border: '1px solid rgba(168,85,247,0.15)',
            fontSize: '14px', color: 'var(--text-secondary)', lineHeight: '1.7'
          }}>
            <Info size={15} style={{ marginRight: '6px', verticalAlign: 'middle', color: 'var(--accent-purple)' }} />
            <strong style={{ color: 'var(--text-primary)' }}>설계 요약:</strong> {result.summary}
          </div>
        </div>

        {/* ── 탭 0: 다이어그램 시각화 ── */}
        {activeTab === 'diagram' && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '4px' }}>
               <h4 style={{ margin: 0, fontSize: '15px', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '8px' }}>
                <Eye size={16} color="var(--accent-purple)" /> ERD 다이어그램 (Mermaid)
              </h4>
              <p style={{ margin: 0, fontSize: '12px', color: 'var(--text-muted)' }}>
                * AI가 생성한 코드로 자동 렌더링한 논리 모델입니다.
              </p>
            </div>
            <MermaidDiagram chart={result.mermaidCode} />
          </div>
        )}

        {/* ── 탭 2: 엔티티 & 속성 명세 ── */}
        {activeTab === 'entities' && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
            {result.entities.map((entity, i) => (
              <details key={i} open style={{
                background: 'rgba(255,255,255,0.02)', borderRadius: '12px',
                border: '1px solid var(--panel-border)', overflow: 'hidden'
              }}>
                <summary style={{
                  padding: '14px 18px', cursor: 'pointer', fontWeight: 700,
                  fontSize: '14px', color: 'var(--text-primary)', listStyle: 'none',
                  display: 'flex', justifyContent: 'space-between', alignItems: 'center',
                  background: 'rgba(168,85,247,0.04)'
                }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                    <span style={{
                      background: 'rgba(168,85,247,0.15)', color: 'var(--accent-purple)',
                      padding: '2px 8px', borderRadius: '5px', fontSize: '12px', fontWeight: 700
                    }}>#{i + 1}</span>
                    <span>{entity.name}</span>
                    <span style={{ color: 'var(--text-muted)', fontWeight: 400, fontSize: '13px' }}>
                      {entity.description}
                    </span>
                  </div>
                  <span style={{ fontSize: '12px', color: 'var(--accent-blue)', fontWeight: 600 }}>
                    속성 {entity.attributes?.length ?? 0}개
                  </span>
                </summary>
                <div style={{ padding: '16px 18px', borderTop: '1px solid var(--panel-border)' }}>
                  <p style={{
                    margin: '0 0 14px', padding: '10px 14px', fontSize: '13px',
                    color: 'var(--text-secondary)', fontStyle: 'italic',
                    background: 'rgba(255,255,255,0.02)', borderRadius: '8px',
                    borderLeft: '3px solid var(--accent-purple)'
                  }}>
                    💡 {entity.reason}
                  </p>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                    <thead>
                      <tr style={{ background: 'rgba(255,255,255,0.03)' }}>
                        <th style={{ padding: '10px 12px', textAlign: 'left', color: 'var(--text-muted)', fontWeight: 600, borderBottom: '1px solid var(--panel-border)' }}>속성명</th>
                        <th style={{ padding: '10px 12px', textAlign: 'left', color: 'var(--text-muted)', fontWeight: 600, borderBottom: '1px solid var(--panel-border)' }}>데이터 타입</th>
                        <th style={{ padding: '10px 12px', textAlign: 'center', color: 'var(--text-muted)', fontWeight: 600, borderBottom: '1px solid var(--panel-border)', width: '60px' }}>Key</th>
                        <th style={{ padding: '10px 12px', textAlign: 'left', color: 'var(--text-muted)', fontWeight: 600, borderBottom: '1px solid var(--panel-border)' }}>설명</th>
                      </tr>
                    </thead>
                    <tbody>
                      {entity.attributes?.map((attr, j) => (
                        <tr key={j} style={{ borderBottom: '1px solid rgba(255,255,255,0.04)' }}>
                          <td style={{ padding: '10px 12px', fontWeight: attr.key ? 700 : 400, color: attr.key === 'PK' ? '#fbbf24' : attr.key === 'FK' ? 'var(--accent-blue)' : 'var(--text-primary)' }}>
                            {attr.name}
                          </td>
                          <td style={{ padding: '10px 12px', color: '#a78bfa', fontSize: '12px', fontFamily: 'monospace' }}>
                            {attr.type}
                          </td>
                          <td style={{ padding: '10px 12px', textAlign: 'center' }}>
                            {attr.key === 'PK'
                              ? <span style={{ background: 'rgba(251,191,36,0.15)', color: '#fbbf24', padding: '2px 8px', borderRadius: '4px', fontSize: '11px', fontWeight: 800 }}>PK</span>
                              : attr.key === 'FK'
                              ? <span style={{ background: 'rgba(59,130,246,0.15)', color: 'var(--accent-blue)', padding: '2px 8px', borderRadius: '4px', fontSize: '11px', fontWeight: 800 }}>FK</span>
                              : <span style={{ color: 'var(--text-muted)', fontSize: '12px' }}>-</span>}
                          </td>
                          <td style={{ padding: '10px 12px', color: 'var(--text-secondary)', fontSize: '13px' }}>
                            {attr.desc || '-'}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </details>
            ))}
          </div>
        )}

        {/* ── 탭 3: 관계 & 정규화 ── */}
        {activeTab === 'relations' && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
            <div>
              <h4 style={{ margin: '0 0 14px', fontSize: '15px', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '8px' }}>
                <Database size={16} color="var(--accent-blue)" /> 주요 관계 정의 ({result.relationships?.length ?? 0}개)
              </h4>
              <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                {result.relationships?.map((rel, i) => (
                  <div key={i} style={{
                    padding: '16px 18px', background: 'rgba(255,255,255,0.02)',
                    borderRadius: '10px', borderLeft: '3px solid var(--accent-purple)',
                    border: '1px solid var(--panel-border)'
                  }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '6px' }}>
                      <span style={{ fontWeight: 700, color: 'var(--text-primary)' }}>{rel.from}</span>
                      <span style={{ color: 'var(--text-muted)', fontSize: '13px' }}>→</span>
                      <span style={{ fontWeight: 700, color: 'var(--text-primary)' }}>{rel.to}</span>
                      <span style={{
                        background: 'rgba(168,85,247,0.1)', color: 'var(--accent-purple)',
                        padding: '2px 10px', borderRadius: '20px', fontSize: '12px', fontWeight: 700
                      }}>{rel.type}</span>
                    </div>
                    <div style={{ color: 'var(--text-secondary)', fontSize: '13px', lineHeight: '1.6' }}>
                      {rel.desc}
                    </div>
                  </div>
                ))}
              </div>
            </div>

            <div>
              <h4 style={{ margin: '0 0 14px', fontSize: '15px', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '8px' }}>
                <AlertCircle size={16} color="var(--warning-color)" /> 정규화 준수 논거
              </h4>
              <div style={{
                padding: '18px 20px', background: 'rgba(245,158,11,0.04)',
                border: '1px solid rgba(245,158,11,0.15)', borderRadius: '12px',
                fontSize: '14px', lineHeight: '1.8', color: 'var(--text-secondary)', whiteSpace: 'pre-wrap'
              }}>
                {result.normalizationNotes}
              </div>
            </div>
          </div>
        )}

        {/* ── 탭 3: DBML 소스코드 ── */}
        {activeTab === 'dbml' && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '14px' }}>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
              <p style={{ margin: 0, fontSize: '13px', color: 'var(--text-muted)' }}>
                아래 코드를 <a href="https://dbdiagram.io" target="_blank" rel="noreferrer" style={{ color: 'var(--accent-blue)', textDecoration: 'underline' }}>dbdiagram.io</a>에 붙여넣기하면 편집할 수 있습니다.
              </p>
              <button 
                onClick={() => handleCopyCode(getDbmlCode())} 
                className="interactive" style={{
                padding: '8px 14px', borderRadius: '8px', border: '1px solid var(--panel-border)',
                background: copied ? 'rgba(16,185,129,0.1)' : 'rgba(255,255,255,0.05)',
                color: copied ? 'var(--success-color)' : 'var(--text-primary)',
                fontSize: '13px', fontWeight: 600, cursor: 'pointer',
                display: 'flex', alignItems: 'center', gap: '6px', transition: 'all 0.2s', whiteSpace: 'nowrap'
              }}>
                {copied ? <Check size={14} /> : <Copy size={14} />}
                {copied ? '복사 완료!' : '소스 복사'}
              </button>
            </div>
            <pre style={{
              margin: 0, padding: '20px 24px',
              background: '#0d1117', borderRadius: '12px',
              border: '1px solid var(--panel-border)',
              fontFamily: "'Fira Code', 'Consolas', monospace",
              fontSize: '13px', lineHeight: '1.7',
              color: '#a5d6ff', overflowX: 'auto', whiteSpace: 'pre'
            }}>
              {getDbmlCode()}
            </pre>
          </div>
        )}
      </div>
    </div>
  );
};


// ── 메인 ErdGenerator 컴포넌트 ────────────────────────────────
const ErdGenerator = ({ apiKey }) => {
  const [inputText, setInputText] = useState('');
  const [feedback, setFeedback] = useState(''); // 추가 요청사항 상태
  const [fileName, setFileName] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [progressMsg, setProgressMsg] = useState('');
  const [error, setError] = useState(null);
  const [result, setResult] = useState(null);
  const [showJsonModal, setShowJsonModal] = useState(false);
  const [isDragging, setIsDragging] = useState(false);
  const [cooldown, setCooldown] = useState(0); 
  const fileInputRef = useRef(null);

  // ── 쿨다운 로직 (로컬 스토리지 연동 및 타이머) ──
  useEffect(() => {
    // 마운트 시 로컬 스토리지에서 마지막 차단 시간 확인
    const lastBlock = localStorage.getItem('erd_cooldown_expiry');
    if (lastBlock) {
      const remaining = Math.ceil((parseInt(lastBlock) - Date.now()) / 1000);
      if (remaining > 0) setCooldown(remaining);
    }
  }, []);

  useEffect(() => {
    let timer;
    if (cooldown > 0) {
      timer = setInterval(() => {
        setCooldown(prev => {
          if (prev <= 1) {
            localStorage.removeItem('erd_cooldown_expiry');
            return 0;
          }
          return prev - 1;
        });
      }, 1000);
    }
    return () => clearInterval(timer);
  }, [cooldown]);

  const startCooldown = (seconds) => {
    const expiry = Date.now() + (seconds * 1000);
    localStorage.setItem('erd_cooldown_expiry', expiry.toString());
    setCooldown(seconds);
  };

  // ── 드래그 앤 드롭 핸들러 ──
  const handleDragEnter = (e) => { e.preventDefault(); e.stopPropagation(); setIsDragging(true); };
  const handleDragLeave = (e) => { e.preventDefault(); e.stopPropagation(); setIsDragging(false); };
  const handleDragOver  = (e) => { e.preventDefault(); e.stopPropagation(); };
  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
    const file = e.dataTransfer.files[0];
    if (file) handleFileSelect(file);
  };

  const handleFileSelect = useCallback(async (file) => {
    if (!file) return;
    setFileName(file.name);
    setIsLoading(true);
    setError(null);
    try {
      const result = await processFile(file);
      const text = result.text;
      setInputText(text);
    } catch (err) {
      setError(err.message);
    } finally {
      setIsLoading(false);
    }
  }, []);

  const handleAnalyze = async () => {
    if (!apiKey || !apiKey.match(/^(AIza|AQ\.)/)) {
      alert('Gemini API 키가 필요합니다. 상단 설정 메뉴에서 API 키를 입력해 주세요.');
      return;
    }
    if (!inputText.trim()) {
      setError("분석할 요구사항 내용을 입력하거나 문서를 업로드해 주세요.");
      return;
    }
    setError(null);
    setIsAnalyzing(true);
    setResult(null); // 신규 분석 시 초기화
    setProgressMsg(feedback.trim() ? "요청사항을 반영하여 요구사항 분석 중..." : "데이터 요구사항 분석 중...");
    
    try {
      const data = await analyzeERDWithLLM(inputText, apiKey, (msg) => setProgressMsg(msg), 'auto', null, feedback);
      setResult(data);
    } catch (err) {
      setError(err.message);
      const lowerErr = err.message.toLowerCase();
      if (lowerErr.includes("quota") || lowerErr.includes("limit") || lowerErr.includes("exhausted") || lowerErr.includes("429")) {
        startCooldown(60);
      }
    } finally {
      setIsAnalyzing(false);
      setProgressMsg('');
    }
  };

  const handleRefine = async () => {
    if (!apiKey || !apiKey.match(/^(AIza|AQ\.)/)) {
      alert('Gemini API 키가 필요합니다.');
      return;
    }
    setError(null);
    setIsAnalyzing(true);
    setProgressMsg("추가 요청사항 반영 중...");
    
    try {
      const data = await analyzeERDWithLLM(inputText, apiKey, (msg) => setProgressMsg(msg), 'auto', result, feedback);
      setResult(data);
    } catch (err) {
      setError(err.message);
      const lowerErr = err.message.toLowerCase();
      if (lowerErr.includes("quota") || lowerErr.includes("limit") || lowerErr.includes("exhausted") || lowerErr.includes("429")) {
        startCooldown(60);
      }
    } finally {
      setIsAnalyzing(false);
      setProgressMsg('');
    }
  };

  const handleReset = () => {
    setInputText('');
    setFeedback('');
    setFileName('');
    setResult(null);
    setError(null);
  };

  return (
    <>
      <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>

        {/* ── 입력 섹션 ── */}
        <div className="glass-panel" style={{ padding: '24px', display: 'flex', flexDirection: 'column', gap: '16px', flexShrink: 0 }}>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
              <Database size={20} color="var(--accent-purple)" />
              <h3 style={{ margin: 0, fontSize: '18px', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '8px' }}>
                AI 데이터베이스(ERD) 설계
                <Sparkles size={16} style={{ color: '#38bdf8', filter: 'drop-shadow(0 0 2px rgba(56, 189, 248, 0.5))' }} />
              </h3>
            </div>
            {(inputText || fileName) && (
              <button onClick={handleReset} className="interactive"
                style={{ background: 'transparent', border: 'none', color: 'var(--text-muted)', fontSize: '13px', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '4px' }}>
                <Trash2 size={14} /> 초기화
              </button>
            )}
          </div>

          <div style={{ display: 'flex', gap: '12px', alignItems: 'center' }}>
            <button onClick={() => fileInputRef.current?.click()} disabled={isLoading || isAnalyzing} className="interactive"
              style={{ padding: '10px 18px', borderRadius: '10px', background: 'rgba(168,85,247,0.1)', color: 'var(--accent-purple)', border: '1px solid rgba(168,85,247,0.2)', fontSize: '14px', fontWeight: 600, cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '8px' }}>
              {isLoading ? <Loader2 size={16} className="animate-spin" /> : <Upload size={16} />}
              {isLoading ? '문서 읽는 중...' : '요구사항 문서 업로드'}
            </button>
            {fileName && (
              <div style={{ fontSize: '13px', color: 'var(--success-color)', background: 'rgba(16,185,129,0.1)', padding: '6px 12px', borderRadius: '8px', border: '1px solid rgba(16,185,129,0.2)', display: 'flex', alignItems: 'center', gap: '6px' }}>
                <FileText size={14} /> {fileName}
              </div>
            )}
            <input type="file" ref={fileInputRef} hidden accept={ALL_ACCEPT} onChange={(e) => handleFileSelect(e.target.files[0])} />
          </div>

          {/* ── 추가 요청사항 (상단 배치) ── */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
              <Play size={14} color="var(--accent-purple)" fill="var(--accent-purple)" />
              <span style={{ fontSize: '13px', fontWeight: 700, color: 'var(--text-secondary)' }}>강조사항/추가요청사항:</span>
            </div>
            <textarea 
              placeholder="분석 시 특별히 고려할 점(예: '특정 엔티티 추가', 'PK 형식 지정' 등)을 입력하세요. 공백 시 일반 요구사항 기반으로 설계됩니다."
              value={feedback}
              onChange={(e) => setFeedback(e.target.value)}
              disabled={isAnalyzing}
              style={{ 
                width: '100%', height: '50px', 
                background: 'rgba(0,0,0,0.25)', border: '1px solid var(--panel-border)', 
                borderRadius: '10px', padding: '12px 14px', 
                color: 'var(--text-primary)', fontSize: '13px', 
                resize: 'none', transition: 'border-color 0.2s',
                outline: 'none'
              }}
            />
          </div>

          {/* 메인 요구사항 텍스트 영역 */}
          <div
            onDragEnter={handleDragEnter}
            onDragLeave={handleDragLeave}
            onDragOver={handleDragOver}
            onDrop={handleDrop}
            style={{ position: 'relative' }}
          >
            <textarea placeholder="데이터베이스 설계의 근거가 될 비즈니스 로직, 요구사항, RFP 내용 등을 자유롭게 입력하세요. 상세할수록 정확한 모델이 도출됩니다. (파일을 이 영역에 드래그해도 됩니다)"
              value={inputText} onChange={(e) => setInputText(e.target.value)}
              style={{ width: '100%', height: '140px', background: isDragging ? 'rgba(59,130,246,0.06)' : 'rgba(0,0,0,0.2)', border: `1px solid ${isDragging ? 'var(--accent-blue)' : 'var(--panel-border)'}`, borderRadius: '12px', padding: '16px', color: 'var(--text-primary)', fontSize: '14px', lineHeight: '1.6', resize: 'none', transition: 'border-color 0.2s, background 0.2s', boxSizing: 'border-box' }}
            />
            {isDragging && (
              <div style={{
                position: 'absolute', inset: 0,
                background: 'rgba(59,130,246,0.1)',
                border: '2px dashed var(--accent-blue)',
                borderRadius: '12px',
                display: 'flex', flexDirection: 'column',
                alignItems: 'center', justifyContent: 'center', gap: '10px',
                pointerEvents: 'none', backdropFilter: 'blur(2px)'
              }}>
                <Upload size={36} color="var(--accent-blue)" />
                <span style={{ fontSize: '15px', fontWeight: 700, color: 'var(--accent-blue)' }}>파일을 여기에 놓으세요</span>
                <span style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>PDF, Excel, PPTX, HWPX, TXT 지원</span>
              </div>
            )}
          </div>

          {/* 분석 시작/반영 버튼 통합 */}
          <div style={{ display: 'flex', gap: '12px' }}>
            <button className="interactive" onClick={handleAnalyze}
              disabled={isAnalyzing || isLoading || !inputText.trim() || cooldown > 0}
              style={{ 
                flex: result ? 1 : 2, padding: '16px', borderRadius: '12px', border: 'none', 
                background: (isAnalyzing || isLoading || !inputText.trim() || cooldown > 0) 
                  ? 'rgba(255,255,255,0.05)' 
                  : result ? 'rgba(255,255,255,0.08)' : 'linear-gradient(135deg, var(--accent-purple), var(--accent-blue))', 
                color: (isAnalyzing || isLoading || !inputText.trim() || cooldown > 0) ? 'var(--text-muted)' : 'white', 
                fontWeight: 700, fontSize: '16px', 
                cursor: (isAnalyzing || isLoading || !inputText.trim() || cooldown > 0) ? 'not-allowed' : 'pointer', 
                display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '10px', 
                transition: 'all 0.3s ease'
              }}>
              {isAnalyzing && !result ? (
                <Loader2 size={20} className="animate-spin" />
              ) : cooldown > 0 ? (
                <AlertCircle size={20} />
              ) : (
                <Sparkles size={20} />
              )}
              {isAnalyzing && !result
                ? progressMsg || 'AI 분석 중...' 
                : cooldown > 0 
                  ? `대기 중 (${cooldown}초)` 
                  : result ? '새로 분석하기' : 'ERD 논리 모델 설계 시작'}
            </button>

            {result && (
              <button className="interactive" onClick={handleRefine}
                disabled={isAnalyzing || !feedback.trim() || cooldown > 0}
                style={{ 
                  flex: 2, padding: '16px', borderRadius: '12px', border: 'none', 
                  background: (isAnalyzing || !feedback.trim() || cooldown > 0) 
                    ? 'rgba(255,255,255,0.05)' 
                    : 'linear-gradient(135deg, var(--accent-purple), var(--accent-blue))', 
                  color: (isAnalyzing || !feedback.trim() || cooldown > 0) ? 'var(--text-muted)' : 'white', 
                  fontWeight: 700, fontSize: '16px', 
                  cursor: (isAnalyzing || !feedback.trim() || cooldown > 0) ? 'not-allowed' : 'pointer', 
                  display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '10px', 
                  transition: 'all 0.3s ease',
                  boxShadow: (isAnalyzing || !feedback.trim() || cooldown > 0) ? 'none' : '0 8px 20px rgba(168,85,247,0.2)'
                }}>
                {isAnalyzing ? <Loader2 size={20} className="animate-spin" /> : <Sparkles size={20} />}
                {isAnalyzing ? progressMsg || '요청 반영 중...' : '추가 요청사항 반영하기'}
              </button>
            )}
          </div>
        </div>

        {/* ── 에러 표시 ── */}
        {error && (
          <div className="animate-fade-in" style={{ padding: '16px', borderRadius: '12px', background: 'rgba(239,68,68,0.1)', border: '1px solid rgba(239,68,68,0.2)', color: 'var(--danger-color)', display: 'flex', alignItems: 'center', gap: '10px', fontSize: '14px', flexShrink: 0 }}>
            <AlertCircle size={18} /> {error}
          </div>
        )}

        {/* ── 결과 탭 뷰 (① 탭 UI) ── */}
        {result && (
          <ResultSection 
            result={result} 
            onOpenJsonViewer={() => setShowJsonModal(true)} 
            onRefine={handleRefine}
            isRefining={isAnalyzing}
            progressMsg={progressMsg}
          />
        )}

        {/* 하단 여백 */}
        <div style={{ height: '20px', flexShrink: 0 }} />
      </div>

      {/* ── JSON 원문 뷰어 모달 (② 팝업 JSON 뷰어) ── */}
      {showJsonModal && result && (
        <JsonViewerModal data={result} onClose={() => setShowJsonModal(false)} />
      )}
    </>
  );
};

export default ErdGenerator;
