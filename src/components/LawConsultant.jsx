import React, { useState, useRef, useEffect, useCallback } from 'react';
import { Send, Scale, User, Bot, Loader2, Sparkles, X, Copy, Check } from 'lucide-react';
import { askLocalRagAssistant, askGeneralLawAssistant } from '../lawAnalyzer';
import { refDB } from '../utils/db';

function LawConsultant({ apiKey, isMcpMode = false }) {
  const [messages, setMessages] = useState([
    { role: 'model', text: isMcpMode 
        ? '안녕하세요! 완전한 로컬(오프라인) 환경에서 구동되는 [로컬 RAG 기반] AI 법률 자문입니다.\n\n지능형 검색 도구를 사용하여 자체 벡터 DB에 내장된 법령을 조회하고 답변해 드립니다.\n(현재 포함 법령: 전자정부법, 지능정보화기본법, 국가계약법, 소프트웨어진흥법, 민법 등)'
        : '안녕하세요! [일반 지식 기반] AI 법률 자문입니다.\n\n실시간 검색 없이 Gemini의 내부 지식만으로 빠르게 답변해 드립니다. 가벼운 규정 확인에 적합하며 토큰 사용이 경제적입니다.' }
  ]);
  const [input, setInput] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [mcpQueryStatus, setMcpQueryStatus] = useState(null);
  const [copiedId, setCopiedId] = useState(null);
  const [refDocs, setRefDocs] = useState([]); // IndexedDB에서 불러온 참고자료
  const messagesEndRef = useRef(null);

  // 참고자료 로드 (성능을 위해 메타데이터만)
  useEffect(() => {
    const loadRefDocs = async () => {
        try {
            const docs = await refDB.getDocsMetadata();
            setRefDocs(docs);
        } catch (err) {
            console.error('Failed to load refDocs in LawConsultant:', err);
        }
    };
    loadRefDocs();
    
    // 타 컴포넌트에서 문서 리스트 변경 시 동기화를 위한 이벤트 리스너 추가
    window.addEventListener('refdocs_changed', loadRefDocs);
    return () => window.removeEventListener('refdocs_changed', loadRefDocs);
  }, []);

  const renderMessageWithLawHighlight = useCallback((text) => {
    if (!text) return null;
    // 「법령명」 또는 영문/숫자/한글이 섞인 가이드/지침 명칭 하이라이트 가능하도록 정규식 확장
    const regex = /(「[^」]+」|[a-zA-Z0-9가-힣\s]{2,40}(?:기본법|진흥법|보호법|계약법|촉진법|이용법|관리법|처리법|지원법|특별법|처벌법|에관한법률|에관한특별법|하도급법|조례|훈령|규칙|지침|고시|예규|가이드|매뉴얼|안내서|지침서)|제\d+조(?:의\d+)?(?:\s?제\d+항)?(?:의\d+)?)/g;
    const parts = text.split(regex);
    
    return parts.map((part, i) => {
        const isLawName = part.match(/^「[^」]+」$/) || part.match(/^[a-zA-Z0-9가-힣\s]{2,40}(?:기본법|진흥법|보호법|계약법|촉진법|이용법|관리법|처리법|지원법|특별법|처벌법|에관한법률|에관한특별법|하도급법|조례|훈령|규칙|지침|고시|예규|가이드|매뉴얼|안내서|지침서)$/);
        const isClause = part.match(/^제\d+조(?:의\d+)?(?:\s?제\d+항)?(?:의\d+)?$/);

        if (isLawName || isClause) {
            const isClickable = !!isLawName;
            
            return (
                <span key={i} style={{
                    background: isClickable ? 'rgba(59, 130, 246, 0.15)' : 'rgba(16, 185, 129, 0.15)',
                    color: isClickable ? '#60a5fa' : '#34d399',
                    padding: '2px 6px',
                    borderRadius: '4px',
                    fontWeight: 'bold',
                    margin: '0 2px',
                    cursor: isClickable ? 'pointer' : 'default',
                    border: `1px solid ${isClickable ? 'rgba(59, 130, 246, 0.3)' : 'rgba(16, 185, 129, 0.3)'}`,
                    transition: 'all 0.2s',
                }} 
                title={isClickable ? `${part} 상세 정보 보기` : '해당 법령 내 조항 번호'}
                onClick={async () => {
                    if (isClickable) {
                        const rawName = part.replace(/[「」]/g, '').trim();
                        // 1. 사용자 등록 문서 선행 확인 (IndexedDB에서 로드된 refDocs 사용)
                        const cleanName = rawName.replace(/\s/g, '').toLowerCase();
                        const matchedDocMeta = refDocs.find(d => 
                            d.title.replace(/\s/g, '').toLowerCase() === cleanName
                        );

                        if (matchedDocMeta) {
                            try {
                                setIsLoading(true);
                                const fullDoc = await refDB.getDocById(matchedDocMeta.id);
                                if (fullDoc) {
                                    if (fullDoc.type === 'link') {
                                        window.open(fullDoc.content, '_blank');
                                    } else {
                                        window.dispatchEvent(new CustomEvent('open_reference_modal', { 
                                            detail: { 
                                                title: fullDoc.title, 
                                                content: fullDoc.content,
                                                blob: fullDoc.blob,
                                                ext: fullDoc.ext,
                                                filename: fullDoc.filename
                                            } 
                                        }));
                                    }
                                }
                            } catch (err) {
                                console.error('Failed to load full doc:', err);
                            } finally {
                                setIsLoading(false);
                            }
                            return;
                        }

                        // 2. 가이드라인 등 비법령 문서 확인 -> 구글 검색
                        if (rawName.includes('가이드') || rawName.includes('매뉴얼') || rawName.includes('안내서') || rawName.includes('지침서')) {
                            window.open(`https://www.google.com/search?q=${encodeURIComponent(rawName)}`, '_blank');
                            return;
                        }

                        // 3. 국가법령정보센터 직접 링크
                        let cleanNameForLink = rawName;
                        
                        // LLM이 실수로 「법령명 제O조」 처럼 조문 번호까지 「」 안에 넣었을 경우를 대비하여 법령명만 파싱
                        const clauseMatch = rawName.match(/^(.*?)(?:\s*제?\s*\d+\s*조)/);
                        if (clauseMatch) {
                            cleanNameForLink = clauseMatch[1];
                        }
                        
                        cleanNameForLink = cleanNameForLink.replace(/\s/g, '').trim();
                        
                        let path = '법령';
                        if (cleanNameForLink.endsWith('지침') || cleanNameForLink.endsWith('고시') || cleanNameForLink.endsWith('훈령') || cleanNameForLink.endsWith('예규') || cleanNameForLink.endsWith('규정')) {
                            path = '행정규칙';
                        } else if (cleanNameForLink.endsWith('조례')) {
                            path = '조례';
                        }
                        window.open(`https://www.law.go.kr/${path}/${encodeURIComponent(cleanNameForLink)}`, '_blank');
                    }
                }}
                onMouseOver={(e) => { 
                    if (isClickable) { e.target.style.background = 'rgba(59, 130, 246, 0.3)'; e.target.style.textDecoration = 'underline'; }
                }}
                onMouseOut={(e) => { 
                    if (isClickable) { e.target.style.background = 'rgba(59, 130, 246, 0.15)'; e.target.style.textDecoration = 'none'; }
                }}
                >
                    {part}
                </span>
            );
        }
        return <span key={i}>{part}</span>;
    });
  }, [refDocs]);

  const handleCopy = (text, id) => {
    navigator.clipboard.writeText(text).then(() => {
        setCopiedId(id);
        setTimeout(() => setCopiedId(null), 2000);
    });
  };

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  };

  useEffect(() => {
    scrollToBottom();
  }, [messages]);

  const handleSend = async () => {
    if (!input.trim()) return;
    
    if (!apiKey || !apiKey.match(/^(AIza|AQ\.)/)) {
        alert("유효한 Gemini API Key를 우측 상단에 입력해 주세요.");
        return;
    }

    const userMessage = { role: 'user', text: input };
    setMessages(prev => [...prev, userMessage]);
    setInput('');
    setIsLoading(true);
    setMcpQueryStatus(null);

    try {
      let responseText = "";
      if (isMcpMode) {
        responseText = await askLocalRagAssistant(userMessage.text, apiKey, messages, (keyword) => {
          setMcpQueryStatus(`Local RAG에서 '${keyword}' 관련 법령 검색 중...`);
        });
      } else {
        responseText = await askGeneralLawAssistant(userMessage.text, apiKey, messages);
      }
      const botMessage = { role: 'model', text: responseText };
      setMessages(prev => [...prev, botMessage]);
    } catch (error) {
      setMessages(prev => [...prev, { role: 'model', text: `오류가 발생했습니다: ${error.message}` }]);
    } finally {
      setIsLoading(false);
      setMcpQueryStatus(null);
    }
  };

  return (
    <div className="glass-panel animate-fade-in" style={{ flex: 1, display: 'flex', flexDirection: 'column', height: '100%', overflow: 'hidden', border: '1px solid var(--panel-border)', borderRadius: '12px' }}>
      {/* Header */}
      <div style={{ padding: '20px 24px', borderBottom: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)', display: 'flex', alignItems: 'center', gap: '12px', flexShrink: 0 }}>
        {isMcpMode ? <Scale size={24} color="var(--accent-purple)" /> : <Sparkles size={24} color="var(--success-color)" />}
        <div>
          <h2 style={{ margin: 0, fontSize: '18px', color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '8px' }}>
            {isMcpMode ? 'AI 법률/규정 자문 (로컬 RAG)' : 'AI 법률 자문(제미나이)'}
            <Sparkles size={16} style={{ color: '#38bdf8', filter: 'drop-shadow(0 0 2px rgba(56, 189, 248, 0.5))' }} />
          </h2>
          <p style={{ margin: '4px 0 0', fontSize: '13px', color: 'var(--text-secondary)' }}>
            {isMcpMode ? '내부 벡터 DB 기반 실무 검토 및 오프라인 해설 지원' : '실시간 검색 없는 Gemini 내부 지식 기반 빠른 해설'}
            {isMcpMode && <span style={{ marginLeft: '10px', color: '#10b981', fontWeight: 600, fontSize: '11px' }}>※ 완전한 로컬 환경에서 안전하게 구동됩니다.</span>}
          </p>
        </div>
        <button
          onClick={() => {
            if (window.confirm('대화 내역을 모두 삭제하고 초기화하시겠습니까?')) {
              const initialText = isMcpMode
                ? '안녕하세요! 완전한 로컬(오프라인) 환경에서 구동되는 [로컬 RAG 기반] AI 법률 자문입니다.\n\n지능형 검색 도구를 사용하여 자체 벡터 DB에 내장된 법령을 조회하고 답변해 드립니다.\n(예시: "전자정부법상 시스템 감리 대상은 무엇인가요?")'
                : '안녕하세요! [일반 지식 기반] AI 법률 자문입니다.\n\n실시간 검색 없이 Gemini의 내부 지식만으로 빠르게 답변해 드립니다. 가벼운 규정 확인에 적합합니다.';
              setMessages([{ role: 'model', text: initialText }]);
              setInput(''); // 입력창도 초기화
            }
          }}
          style={{
            marginLeft: 'auto',
            background: 'rgba(239, 68, 68, 0.1)',
            border: '1px solid rgba(239, 68, 68, 0.2)',
            padding: '6px 12px',
            borderRadius: '6px',
            color: 'var(--danger-color)',
            fontSize: '12px',
            fontWeight: 600,
            cursor: 'pointer',
            display: 'flex',
            alignItems: 'center',
            gap: '6px',
            transition: 'all 0.2s'
          }}
          onMouseEnter={e => e.currentTarget.style.background = 'rgba(239, 68, 68, 0.2)'}
          onMouseLeave={e => e.currentTarget.style.background = 'rgba(239, 68, 68, 0.1)'}
        >
          <X size={14} /> 대화 초기화
        </button>
      </div>

      {/* Message List */}
      <div style={{ flex: 1, overflowY: 'auto', padding: '24px', display: 'flex', flexDirection: 'column', gap: '20px' }}>
        {messages.map((msg, idx) => (
          <div key={idx} style={{ display: 'flex', gap: '16px', flexDirection: msg.role === 'user' ? 'row-reverse' : 'row' }}>
            <div style={{ 
                width: '36px', height: '36px', borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0,
                background: msg.role === 'user' ? 'rgba(59, 130, 246, 0.2)' : 'rgba(255, 255, 255, 0.1)',
                color: msg.role === 'user' ? 'var(--accent-color)' : 'var(--text-primary)'
             }}>
                {msg.role === 'user' ? <User size={20} /> : <Bot size={20} color="var(--success-color)" />}
            </div>
            <div style={{
                background: msg.role === 'user' ? 'rgba(59, 130, 246, 0.15)' : 'rgba(0, 0, 0, 0.2)',
                border: `1px solid ${msg.role === 'user' ? 'rgba(59, 130, 246, 0.3)' : 'var(--panel-border)'}`,
                padding: '16px', borderRadius: '12px', maxWidth: '80%', lineHeight: '1.6', fontSize: '14px', color: 'var(--text-secondary)',
                whiteSpace: 'pre-wrap', position: 'relative', group: 'true'
            }}>
                {msg.role === 'model' ? renderMessageWithLawHighlight(msg.text) : msg.text}
                
                {msg.role === 'model' && (
                  <button
                    onClick={() => handleCopy(msg.text, idx)}
                    style={{
                      position: 'absolute',
                      bottom: '8px',
                      right: '8px',
                      background: 'rgba(255, 255, 255, 0.15)',
                      border: '1px solid rgba(255, 255, 255, 0.2)',
                      borderRadius: '6px',
                      padding: '4px 8px',
                      cursor: 'pointer',
                      color: copiedId === idx ? 'var(--success-color)' : 'var(--text-primary)',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      gap: '4px',
                      transition: 'all 0.2s',
                      opacity: 1,
                      boxShadow: '0 2px 8px rgba(0,0,0,0.2)',
                      zIndex: 10,
                      fontSize: '11px',
                      fontWeight: 600
                    }}
                    title="답변 복사"
                    onMouseEnter={e => {
                      e.currentTarget.style.background = 'rgba(255, 255, 255, 0.25)';
                      e.currentTarget.style.transform = 'translateY(-1px)';
                    }}
                    onMouseLeave={e => {
                      e.currentTarget.style.background = 'rgba(255, 255, 255, 0.15)';
                      e.currentTarget.style.transform = 'translateY(0)';
                    }}
                  >
                    {copiedId === idx ? <Check size={12} /> : <Copy size={12} />}
                    {copiedId === idx ? '복사됨' : '복사'}
                  </button>
                )}
            </div>
          </div>
        ))}
        {isLoading && (
            <div style={{ display: 'flex', gap: '16px' }}>
                <div style={{ width: '36px', height: '36px', borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'rgba(255, 255, 255, 0.1)' }}>
                    <Bot size={20} color="var(--success-color)" />
                </div>
                <div style={{ background: 'rgba(0,0,0,0.2)', padding: '16px', borderRadius: '12px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <Loader2 size={16} color="var(--accent-color)" style={{ animation: 'spin 1.2s linear infinite', flexShrink: 0 }} />
                    <span style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>{mcpQueryStatus || 'Gemini가 답변을 생성하는 중...'}</span>
                </div>
            </div>
        )}
        <div ref={messagesEndRef} />
      </div>

      {/* Input Area */}
      <div style={{ padding: '20px', borderTop: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.1)', flexShrink: 0 }}>
        <div style={{ display: 'flex', gap: '12px', background: 'var(--bg-primary)', padding: '8px 12px', borderRadius: '8px', border: '1px solid var(--panel-border)' }}>
            <input 
                type="text" 
                value={input}
                onChange={e => setInput(e.target.value)}
                onKeyDown={e => e.key === 'Enter' && handleSend()}
                placeholder="어떤 지침에 대해 궁금하신가요? (예: 과업변경심의위원회를 개최하려면 조건이 있나요?)"
                style={{ flex: 1, background: 'transparent', border: 'none', color: 'var(--text-primary)', outline: 'none', fontSize: '14px' }}
                disabled={isLoading}
            />
            <button 
                onClick={handleSend}
                disabled={isLoading || !input.trim()}
                style={{ 
                    background: 'var(--accent-color)', color: 'white', border: 'none', padding: '8px 16px', borderRadius: '6px', 
                    cursor: (isLoading || !input.trim()) ? 'not-allowed' : 'pointer', opacity: (isLoading || !input.trim()) ? 0.6 : 1,
                    display: 'flex', alignItems: 'center', gap: '8px', fontWeight: 600, fontSize: '13px', transition: 'all 0.2s'
                }}
            >
                <Sparkles size={16} /> 질문하기
            </button>
        </div>
        <p style={{ margin: '8px 0 0 4px', fontSize: '11px', color: 'var(--text-secondary)' }}>
            * AI를 통해 실무 보조용 법적 정보를 제공하며, 실제 법적 효력을 갖는 공식 유권해석은 아닙니다.
        </p>
      </div>
    </div>
  );
}

export default LawConsultant;
