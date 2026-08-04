import React, { useState, useEffect, useRef } from 'react';
import { Search, FileText, Database, Info, ExternalLink, ChevronRight, FileType, Clock, HardDrive, Filter, RefreshCw, Loader2, Send, Bot, User, MessageSquare, Upload, Plus, Trash2, Copy, Check, Sparkles, BookOpen } from 'lucide-react';
import { loadRagData, searchRag } from '../utils/ragService';
import { askRagQuestion, askTotalRagQuestion } from '../llmAnalyzer';
import { processFile } from '../utils/fileExtractor';

const RagKnowledgeBase = ({ apiKey, llmProvider = 'gemini', omniRouteModel = 'auto' }) => {
    const [allDocs, setAllDocs] = useState([]);
    const [filteredDocs, setFilteredDocs] = useState([]);
    const [searchTerm, setSearchTerm] = useState('');
    const [isLoading, setIsLoading] = useState(true);
    const [selectedDoc, setSelectedDoc] = useState(null);
    const [stats, setStats] = useState({ count: 0, size: 0 });
    const [isReindexing, setIsReindexing] = useState(false);
    const [isUploading, setIsUploading] = useState(false);
    const fileInputRef = useRef(null);

    const handleFileUpload = async (e) => {
        const file = e.target.files[0];
        if (!file) return;

        // 운영 환경(Vercel 등) 체크
        if (window.location.hostname !== 'localhost' && window.location.hostname !== '127.0.0.1') {
            alert("문서 등록/삭제 기능은 로컬 개발 환경에서만 지원됩니다. 로컬에서 데이터를 관리한 후 GitHub에 푸시하여 배포해 주세요.");
            return;
        }

        setIsUploading(true);
        try {
            // 1. 파일 텍스트 추출
            const extractedData = await processFile(file);
            
            // 2. RAG 데이터 형식으로 변환
            const newRagDoc = {
                id: `manual_${Date.now()}`,
                title: file.name,
                content: extractedData.text,
                pages: extractedData.pages || 1, 
                metadata: {
                    type: file.name.split('.').pop().toUpperCase(),
                    size: (file.size / 1024).toFixed(1) + ' KB',
                    lastModified: new Date(file.lastModified).toLocaleDateString(),
                    path: 'Directly Uploaded'
                },
                tags: ['Manual', file.name.split('.').pop().toUpperCase()]
            };

            // 3. 서버에 영구 저장 요청
            const response = await fetch('/api/save-rag', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(newRagDoc)
            });

            const result = await response.json();
            if (result.success) {
                alert(result.duplicated ? "이미 등록된 문서입니다." : "문서가 지식베이스에 성공적으로 등록되었습니다.");
                init(true); 
            } else {
                throw new Error(result.error);
            }
        } catch (error) {
            console.error("Upload failed:", error);
            alert("문서 등록 중 오류가 발생했습니다: " + error.message);
        } finally {
            setIsUploading(false);
            if (fileInputRef.current) fileInputRef.current.value = '';
        }
    };

    const handleDeleteDoc = async (docId, docTitle) => {
        if (!window.confirm(`[${docTitle}] 문서를 지식베이스에서 삭제하시겠습니까?`)) return;

        // 운영 환경(Vercel 등) 체크
        if (window.location.hostname !== 'localhost' && window.location.hostname !== '127.0.0.1') {
            alert("운영 환경(Vercel)에서는 직접적인 파일 삭제가 제한됩니다. 로컬 환경에서 삭제 후 배포해 주세요.");
            return;
        }

        try {
            const response = await fetch('/api/delete-rag', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ id: docId })
            });

            const result = await response.json();
            if (result.success) {
                alert("문서가 성공적으로 삭제되었습니다.");
                if (selectedDoc?.id === docId) setSelectedDoc(null);
                init(true); 
            } else {
                throw new Error(result.error);
            }
        } catch (error) {
            console.error("Delete failed:", error);
            alert("문서 삭제 중 오류가 발생했습니다: " + error.message);
        }
    };
    
    // 개별 문서 Q&A 상태
    const [chatMessages, setChatMessages] = useState([]);
    const [userQuestion, setUserQuestion] = useState('');
    const [isAnswering, setIsAnswering] = useState(false);
    const [answerStatus, setAnswerStatus] = useState('');
    
    // 전체 지식베이스 Q&A 상태
    const [totalChatMessages, setTotalChatMessages] = useState([]);
    const [totalUserQuestion, setTotalUserQuestion] = useState('');
    const [isTotalAnswering, setIsTotalAnswering] = useState(false);
    const [totalAnswerStatus, setTotalAnswerStatus] = useState('');
    
    // 다중 선택 RAG Q&A 상태
    const [selected_doc_ids, set_selected_doc_ids] = useState([]);
    const [selection_chat_messages, set_selection_chat_messages] = useState([]);
    const [selection_user_question, set_selection_user_question] = useState('');
    const [is_selection_answering, set_is_selection_answering] = useState(false);
    const [selection_answer_status, set_selection_answer_status] = useState('');
    
    const [copiedId, setCopiedId] = useState(null);
    const [isDocContentExpanded, setIsDocContentExpanded] = useState(false);
    const chatContainerRef = useRef(null);
    const totalChatContainerRef = useRef(null);
    const selectionChatContainerRef = useRef(null);

    const handleCopy = (text, id) => {
        navigator.clipboard.writeText(text).then(() => {
            setCopiedId(id);
            setTimeout(() => setCopiedId(null), 2000);
        });
    };

    // 전체 지식베이스 질의 응답 핸들러
    const handleAskTotalQuestion = async () => {
        if (!totalUserQuestion.trim() || isTotalAnswering) return;
        if (!apiKey && llmProvider !== 'omniroute') {
            alert('Gemini API Key가 설정되지 않았습니다. 상단 설정 메뉴에서 키를 입력해 주세요.');
            return;
        }

        const question = totalUserQuestion.trim();
        setTotalUserQuestion('');
        
        const newMsg = { role: 'user', text: question, timestamp: new Date().toLocaleTimeString() };
        setTotalChatMessages(prev => [...prev, newMsg]);
        
        setIsTotalAnswering(true);
        setTotalAnswerStatus('관련 문서 검색 중...');
        
        try {
            // 전체 지식베이스에서 연관된 조각 3개 검색
            const searchResults = await searchRag(question, 3);
            setTotalAnswerStatus('Gemini AI 답변 분석 중...');
            
            const answer = await askTotalRagQuestion(
                question,
                searchResults,
                apiKey,
                (status) => setTotalAnswerStatus(status),
                llmProvider,
                omniRouteModel
            );
            
            setTotalChatMessages(prev => [...prev, { 
                role: 'ai', 
                text: answer, 
                timestamp: new Date().toLocaleTimeString(),
                sources: searchResults.length > 0 ? searchResults.map(r => r.title) : null
            }]);
        } catch (error) {
            console.error('Total Q&A Error:', error);
            setTotalChatMessages(prev => [...prev, { 
                role: 'ai', 
                text: `오류가 발생했습니다: ${error.message}`, 
                timestamp: new Date().toLocaleTimeString() 
            }]);
        } finally {
            setIsTotalAnswering(false);
            setTotalAnswerStatus('');
        }
    };

    // 다중 선택 문서 질의 응답 관련 핸들러 및 훅
    const handle_toggle_doc_selection = (doc_id) => {
        set_selected_doc_ids(prev => {
            if (prev.includes(doc_id)) {
                return prev.filter(id => id !== doc_id);
            } else {
                return [...prev, doc_id];
            }
        });
    };

    const handle_clear_doc_selection = () => {
        set_selected_doc_ids([]);
        set_selection_chat_messages([]);
    };

    const handle_ask_selection_question = async () => {
        if (!selection_user_question.trim() || is_selection_answering) return;
        if (!apiKey && llmProvider !== 'omniroute') {
            alert('Gemini API Key가 설정되지 않았습니다. 상단 설정 메뉴에서 키를 입력해 주세요.');
            return;
        }
        if (selected_doc_ids.length === 0) {
            alert('질의할 문서를 하나 이상 선택해 주세요.');
            return;
        }

        const question = selection_user_question.trim();
        set_selection_user_question('');

        const new_msg = { role: 'user', text: question, timestamp: new Date().toLocaleTimeString() };
        set_selection_chat_messages(prev => [...prev, new_msg]);

        set_is_selection_answering(true);
        set_selection_answer_status('선택된 문서 분석 중...');

        try {
            const selected_docs = allDocs.filter(d => selected_doc_ids.includes(d.id));
            set_selection_answer_status('Gemini AI 답변 도출 중...');

            const answer = await askTotalRagQuestion(
                question,
                selected_docs,
                apiKey,
                (status) => set_selection_answer_status(status),
                llmProvider,
                omniRouteModel
            );

            set_selection_chat_messages(prev => [...prev, {
                role: 'ai',
                text: answer,
                timestamp: new Date().toLocaleTimeString(),
                sources: selected_docs.map(d => d.title)
            }]);
        } catch (error) {
            console.error('Selection Q&A Error:', error);
            set_selection_chat_messages(prev => [...prev, {
                role: 'ai',
                text: `오류가 발생했습니다: ${error.message}`,
                timestamp: new Date().toLocaleTimeString()
            }]);
        } finally {
            set_is_selection_answering(false);
            set_selection_answer_status('');
        }
    };

    useEffect(() => {
        if (selection_chat_messages.length > 0) {
            const lastMsg = selection_chat_messages[selection_chat_messages.length - 1];
            if (lastMsg.role === 'ai') {
                setTimeout(() => {
                    selectionChatContainerRef.current?.scrollTo({
                        top: selectionChatContainerRef.current.scrollHeight,
                        behavior: 'smooth'
                    });
                }, 100);
            }
        }
    }, [selection_chat_messages]);

    useEffect(() => {
        if (chatMessages.length > 0) {
            const lastMsg = chatMessages[chatMessages.length - 1];
            if (lastMsg.role === 'ai') {
                setTimeout(() => {
                    chatContainerRef.current?.scrollTo({
                        top: chatContainerRef.current.scrollHeight,
                        behavior: 'smooth'
                    });
                }, 100);
            }
        }
    }, [chatMessages]);

    useEffect(() => {
        if (totalChatMessages.length > 0) {
            const lastMsg = totalChatMessages[totalChatMessages.length - 1];
            if (lastMsg.role === 'ai') {
                setTimeout(() => {
                    totalChatContainerRef.current?.scrollTo({
                        top: totalChatContainerRef.current.scrollHeight,
                        behavior: 'smooth'
                    });
                }, 100);
            }
        }
    }, [totalChatMessages]);

    const init = async (refresh = false) => {
        setIsLoading(true);
        const data = await loadRagData(refresh);
        setAllDocs(data);
        setFilteredDocs(data);
        
        const totalSize = data.reduce((acc, doc) => acc + (doc.size || 0), 0);
        setStats({
            count: data.length,
            size: (totalSize / (1024 * 1024)).toFixed(1)
        });
        setIsLoading(false);
    };

    useEffect(() => {
        init();
    }, []);

    const handleReindex = async () => {
        if (!window.confirm('로컬 산출물 폴더를 다시 스캔하여 지식베이스를 갱신하시겠습니까?\n(약 수초~수십초가 소요될 수 있습니다)')) return;
        
        // 운영 환경(Vercel 등) 체크
        if (window.location.hostname !== 'localhost' && window.location.hostname !== '127.0.0.1') {
            alert("지식베이스 갱신은 로컬 환경에서만 가능합니다.");
            return;
        }

        setIsReindexing(true);
        try {
            const response = await fetch('/api/reindex', { method: 'POST' });
            const result = await response.json();
            if (result.success) {
                await init(true);
                alert('지식베이스 인덱싱이 완료되었습니다.');
            } else {
                alert('인덱싱 실패: ' + result.error);
            }
        } catch (error) {
            console.error('Reindex error:', error);
            alert('인덱싱 중 오류가 발생했습니다. 개발 서버 상태를 확인하세요.');
        } finally {
            setIsReindexing(false);
        }
    };

    const handleAskQuestion = async () => {
        if (!userQuestion.trim() || isAnswering || !selectedDoc) return;
        if (!apiKey && llmProvider !== 'omniroute') {
            alert('Gemini API Key가 설정되지 않았습니다. 상단 설정 메뉴에서 키를 입력해 주세요.');
            return;
        }

        const question = userQuestion.trim();
        setUserQuestion('');
        
        const newMsg = { role: 'user', text: question, timestamp: new Date().toLocaleTimeString() };
        setChatMessages(prev => [...prev, newMsg]);
        
        setIsAnswering(true);
        setAnswerStatus('분석 중...');
        
        try {
            const answer = await askRagQuestion(
                selectedDoc.title,
                selectedDoc.content,
                question,
                apiKey,
                (status) => setAnswerStatus(status),
                llmProvider,
                omniRouteModel
            );
            
            setChatMessages(prev => [...prev, { 
                role: 'ai', 
                text: answer, 
                timestamp: new Date().toLocaleTimeString() 
            }]);
        } catch (error) {
            console.error('Q&A Error:', error);
            setChatMessages(prev => [...prev, { 
                role: 'ai', 
                text: `오류가 발생했습니다: ${error.message}`, 
                timestamp: new Date().toLocaleTimeString() 
            }]);
        } finally {
            setIsAnswering(false);
            setAnswerStatus('');
        }
    };

    useEffect(() => {
        setChatMessages([]);
        setUserQuestion('');
        setIsDocContentExpanded(false);
    }, [selectedDoc]);

    const handleSearch = async (e) => {
        const value = e.target.value;
        setSearchTerm(value);
        
        if (!value.trim()) {
            setFilteredDocs(allDocs);
            return;
        }

        const results = await searchRag(value, 50);
        setFilteredDocs(results);
    };

    const formatSize = (bytes) => {
        if (!bytes) return '0 B';
        const k = 1024;
        const sizes = ['B', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
    };

    return (
        <div style={{ flex: 1, display: 'flex', flexDirection: 'column', height: '100%', overflow: 'hidden' }}>
            {/* Header Area */}
            <div className="glass-panel animate-slide-up" style={{ padding: '24px 32px', marginBottom: '24px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', border: '1px solid var(--panel-border)', borderRadius: '16px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '20px' }}>
                    <div style={{ width: '48px', height: '48px', borderRadius: '14px', background: 'rgba(59, 130, 246, 0.1)', display: 'flex', alignItems: 'center', justifyContent: 'center', border: '1px solid rgba(59, 130, 246, 0.2)' }}>
                        <Database size={28} color="var(--accent-blue)" />
                    </div>
                    <div>
                        <h2 style={{ margin: 0, fontSize: '22px', fontWeight: 700, color: 'var(--text-primary)', letterSpacing: '-0.5px', display: 'flex', alignItems: 'center', gap: '6px' }}>
                            프로젝트 RAG 지식베이스
                            <Sparkles size={16} style={{ color: '#38bdf8', filter: 'drop-shadow(0 0 2px rgba(56, 189, 248, 0.5))' }} />
                        </h2>
                        <p style={{ margin: '4px 0 0', fontSize: '14px', color: 'var(--text-secondary)' }}>데스크탑 산출물 폴더 기반 인덱싱 데이터 ({stats.count}개 파일, {stats.size}MB)</p>
                    </div>
                </div>

                <div style={{ display: 'flex', gap: '12px' }}>
                    <input 
                        type="file" 
                        ref={fileInputRef} 
                        onChange={handleFileUpload} 
                        style={{ display: 'none' }} 
                        accept=".pdf,.pptx,.docx,.txt"
                    />
                    <button 
                        onClick={() => fileInputRef.current?.click()}
                        disabled={isUploading}
                        className="interactive"
                        style={{
                            display: 'flex', alignItems: 'center', gap: '8px',
                            padding: '10px 20px', borderRadius: '10px',
                            background: 'var(--accent-blue)', color: 'white',
                            border: 'none', cursor: isUploading ? 'not-allowed' : 'pointer',
                            fontSize: '14px', fontWeight: '500', transition: 'all 0.2s',
                            opacity: isUploading ? 0.7 : 1,
                            boxShadow: '0 4px 12px rgba(59, 130, 246, 0.2)'
                        }}
                    >
                        {isUploading ? <Loader2 size={18} className="animate-spin" /> : <Plus size={18} />}
                        {isUploading ? '등록 중...' : (
                            <>
                                문서직접등록(<span style={{ color: '#ef4444', fontWeight: 'bold' }}>로컬만</span>)
                            </>
                        )}
                    </button>
                    <button 
                        onClick={handleReindex}
                        disabled={isReindexing}
                        className="interactive"
                        style={{
                            display: 'flex', alignItems: 'center', gap: '8px',
                            padding: '10px 20px', borderRadius: '10px',
                            background: 'rgba(255,255,255,0.05)', color: 'var(--text-primary)',
                            border: '1px solid var(--glass-border)', cursor: isReindexing ? 'not-allowed' : 'pointer',
                            fontSize: '14px', fontWeight: '500', transition: 'all 0.2s'
                        }}
                    >
                        <RefreshCw size={18} className={isReindexing ? "animate-spin" : ""} />
                        {isReindexing ? '인덱싱 중...' : '지식베이스 갱신'}
                    </button>
                </div>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: '1fr 3.2fr', gap: '24px', flex: 1, overflow: 'hidden' }}>
                {/* Search and List Column */}
                <div style={{ display: 'flex', flexDirection: 'column', overflow: 'hidden', gap: '16px' }}>
                    <div style={{ position: 'relative' }}>
                        <Search size={20} style={{ position: 'absolute', left: '16px', top: '50%', transform: 'translateY(-50%)', color: 'var(--text-muted)' }} />
                        <input 
                            type="text" 
                            placeholder="전체 지식 베이스 키워드 검색..." 
                            value={searchTerm} 
                            onChange={handleSearch}
                            style={{ 
                                width: '100%', 
                                padding: '14px 14px 14px 52px', 
                                background: 'rgba(255,255,255,0.03)', 
                                borderRadius: '14px', 
                                fontSize: '15px', 
                                border: '1px solid var(--glass-border)',
                                color: 'var(--text-primary)',
                                outline: 'none',
                                transition: 'all 0.2s'
                            }} 
                        />
                    </div>

                    <div className="glass-panel" style={{ flex: 1, overflowY: 'auto', padding: '12px', borderRadius: '16px' }}>
                        {selected_doc_ids.length > 0 && (
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '10px 14px', background: 'rgba(59, 130, 246, 0.08)', borderRadius: '10px', border: '1px solid rgba(59, 130, 246, 0.15)', marginBottom: '12px' }}>
                                <span style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>
                                    선택한 지식: <strong style={{ color: 'var(--accent-blue)', fontSize: '14px' }}>{selected_doc_ids.length}</strong>개
                                </span>
                                <button 
                                    onClick={handle_clear_doc_selection}
                                    className="interactive"
                                    style={{ background: 'none', border: 'none', color: '#ef4444', fontSize: '12px', cursor: 'pointer', fontWeight: 600, padding: '4px 8px', borderRadius: '6px' }}
                                >
                                    선택 초기화
                                </button>
                            </div>
                        )}
                        {isLoading ? (
                            <div style={{ padding: '40px', textAlign: 'center', color: 'var(--text-muted)' }}>로딩 중...</div>
                        ) : filteredDocs.length > 0 ? (
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                                {filteredDocs.map(doc => (
                                    <div 
                                        key={doc.id}
                                        style={{ display: 'flex', alignItems: 'center', gap: '8px', width: '100%' }}
                                    >
                                        <input 
                                            type="checkbox"
                                            checked={selected_doc_ids.includes(doc.id)}
                                            onChange={() => handle_toggle_doc_selection(doc.id)}
                                            style={{ 
                                                width: '18px', 
                                                height: '18px', 
                                                cursor: 'pointer',
                                                accentColor: 'var(--accent-blue)',
                                                marginLeft: '6px',
                                                flexShrink: 0
                                            }}
                                            title="질의 대상 지식으로 선택"
                                        />
                                        <button
                                            onClick={() => setSelectedDoc(doc)}
                                            className="interactive"
                                            style={{
                                                display: 'flex',
                                                alignItems: 'center',
                                                padding: '12px 8px',
                                                borderRadius: '12px',
                                                background: selectedDoc?.id === doc.id ? 'rgba(59, 130, 246, 0.1)' : 'transparent',
                                                border: `1px solid ${selectedDoc?.id === doc.id ? 'rgba(59, 130, 246, 0.3)' : 'transparent'}`,
                                                textAlign: 'left',
                                                flex: 1,
                                                minWidth: 0,
                                                transition: 'all 0.2s'
                                            }}
                                        >
                                            <div style={{ flex: 1, overflow: 'hidden' }}>
                                                <div style={{ fontSize: '14px', fontWeight: 600, color: 'var(--text-primary)', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                                                    {doc.title}
                                                </div>
                                                <div style={{ fontSize: '12px', color: 'var(--text-muted)', display: 'flex', gap: '8px', marginTop: '2px' }}>
                                                    <span>{(doc.type || doc.metadata?.type || 'DOC').toUpperCase()}</span>
                                                    <span>•</span>
                                                    <span>{typeof (doc.size || doc.metadata?.size) === 'string' ? (doc.size || doc.metadata?.size) : formatSize(doc.size || doc.metadata?.size)}</span>
                                                    {doc.score > 0 && (
                                                        <>
                                                            <span>•</span>
                                                            <span style={{ color: 'var(--accent-blue)' }}>정확도 {doc.score}</span>
                                                        </>
                                                    )}
                                                </div>
                                            </div>
                                            <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                                                <button 
                                                    onClick={(e) => {
                                                        e.stopPropagation();
                                                        handleDeleteDoc(doc.id, doc.title);
                                                    }}
                                                    className="interactive-red"
                                                    style={{
                                                        padding: '8px', borderRadius: '8px',
                                                        background: 'transparent', border: 'none',
                                                        color: 'var(--text-muted)', cursor: 'pointer',
                                                        transition: 'all 0.2s'
                                                    }}
                                                    title="문서 삭제"
                                                >
                                                    <Trash2 size={14} />
                                                </button>
                                                <ChevronRight size={16} color="var(--text-muted)" />
                                            </div>
                                        </button>
                                    </div>
                                ))}
                            </div>
                        ) : (
                            <div style={{ padding: '40px', textAlign: 'center', color: 'var(--text-muted)' }}>검색 결과가 없습니다.</div>
                        )}
                    </div>
                </div>

                {/* Detail View Column */}
                <div style={{ display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
                    {selectedDoc ? (
                        <div className="glass-panel animate-scale-in" style={{ display: 'flex', flexDirection: 'column', height: '100%', overflow: 'hidden', borderRadius: '16px', border: '1px solid var(--panel-border)' }}>
                            <div style={{ padding: '24px', borderBottom: '1px solid var(--glass-border)', display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                                <div style={{ flex: 1 }}>
                                    <div style={{ fontSize: '11px', fontWeight: 800, color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '1px', marginBottom: '4px' }}>Document Details</div>
                                    <h3 style={{ margin: 0, fontSize: '18px', fontWeight: 700, color: 'var(--text-primary)' }}>{selectedDoc.title}</h3>
                                    <div style={{ fontSize: '13px', color: 'var(--accent-blue)', marginTop: '6px', wordBreak: 'break-all' }}>
                                        {selectedDoc.path}
                                    </div>
                                </div>
                                <button onClick={() => setSelectedDoc(null)} style={{ padding: '6px 14px', background: 'rgba(255,255,255,0.05)', borderRadius: '8px', border: '1px solid var(--glass-border)', color: 'var(--text-muted)', cursor: 'pointer', fontSize: '12px' }}>
                                    선택 해제
                                </button>
                            </div>

                            <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden', padding: '24px' }}>
                                <button 
                                    onClick={() => setIsDocContentExpanded(!isDocContentExpanded)}
                                    className="interactive"
                                    style={{
                                        display: 'flex', alignItems: 'center', justifyContent: 'space-between',
                                        width: '100%', padding: '12px 16px', background: 'rgba(255,255,255,0.03)',
                                        border: '1px solid var(--glass-border)', borderRadius: '12px',
                                        color: 'var(--text-primary)', cursor: 'pointer', fontSize: '14px', fontWeight: 600,
                                        marginBottom: '16px', transition: 'all 0.2s', outline: 'none'
                                    }}
                                >
                                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                        <BookOpen size={16} color="var(--accent-blue)" />
                                        <span>원본 문서 정보 및 요약 {isDocContentExpanded ? '접기' : '펼쳐보기'}</span>
                                    </div>
                                    <span style={{ fontSize: '12px', color: 'var(--text-muted)' }}>
                                        {isDocContentExpanded ? '▲ 요약/텍스트 접기' : '▼ 펼쳐서 내용 보기'}
                                    </span>
                                </button>

                                {isDocContentExpanded && (
                                    <div className="animate-slide-down" style={{ display: 'flex', flexDirection: 'column', gap: '16px', maxHeight: '280px', overflowY: 'auto', marginBottom: '20px', paddingRight: '4px' }}>
                                        <div style={{ padding: '16px', background: 'rgba(255,255,255,0.02)', borderRadius: '12px', border: '1px solid var(--glass-border)' }}>
                                            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '12px', color: 'var(--text-primary)', fontWeight: 600, fontSize: '14px' }}>
                                                <Info size={16} /> 요약 정보
                                            </div>
                                            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
                                                <div style={{ padding: '10px', background: 'rgba(0,0,0,0.2)', borderRadius: '8px' }}>
                                                    <div style={{ fontSize: '11px', color: 'var(--text-muted)' }}>파일 형식</div>
                                                    <div style={{ fontSize: '13px', fontWeight: 600 }}>{(selectedDoc.type || selectedDoc.metadata?.type || 'DOC').toUpperCase()} Document</div>
                                                </div>
                                                <div style={{ padding: '10px', background: 'rgba(0,0,0,0.2)', borderRadius: '8px' }}>
                                                    <div style={{ fontSize: '11px', color: 'var(--text-muted)' }}>페이지/슬라이드</div>
                                                    <div style={{ fontSize: '13px', fontWeight: 600 }}>
                                                        {(() => {
                                                            const p = selectedDoc.pages || selectedDoc.metadata?.pages;
                                                            return Array.isArray(p) ? p.length : (p || 0);
                                                        })()} Pages
                                                    </div>
                                                </div>
                                            </div>
                                        </div>

                                        <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                                            <div style={{ fontSize: '14px', fontWeight: 700, color: 'var(--text-primary)', borderLeft: '3px solid var(--accent-blue)', paddingLeft: '10px' }}>
                                                추출된 텍스트 내용
                                            </div>
                                            <div style={{ 
                                                fontSize: '13px', color: 'var(--text-secondary)', lineHeight: '1.7', 
                                                whiteSpace: 'pre-wrap', background: 'rgba(0,0,0,0.15)', padding: '16px', 
                                                borderRadius: '12px', border: '1px solid var(--glass-border)',
                                                maxHeight: '150px', overflowY: 'auto'
                                            }}>
                                                {selectedDoc.content}
                                            </div>
                                        </div>
                                    </div>
                                )}

                                {/* Q&A Chat Section */}
                                <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden', borderTop: '1px solid var(--glass-border)', paddingTop: '20px' }}>
                                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '14px', color: 'var(--text-primary)', fontWeight: 700, fontSize: '15px' }}>
                                        <MessageSquare size={18} color="var(--accent-blue)" /> AI 문서 Q&A 대화
                                    </div>

                                    <div style={{ 
                                        flex: 1, background: 'rgba(255,255,255,0.02)', border: '1px solid var(--glass-border)', 
                                        borderRadius: '16px', display: 'flex', flexDirection: 'column', overflow: 'hidden' 
                                    }}>
                                        <div 
                                            ref={chatContainerRef}
                                            style={{ 
                                                flex: 1, overflowY: 'auto', padding: '16px 8px', display: 'flex', flexDirection: 'column', gap: '16px',
                                                scrollBehavior: 'smooth'
                                            }}
                                        >
                                            {chatMessages.length === 0 ? (
                                                <div style={{ height: '100%', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', color: 'var(--text-muted)', textAlign: 'center', gap: '12px' }}>
                                                    <Bot size={40} opacity={0.2} />
                                                    <p style={{ margin: 0, fontSize: '14px' }}>선택한 문서의 내용에 대해 궁금한 점을 물어보세요.<br/>예: "이 문서의 주요 요건 3가지를 알려줘"</p>
                                                </div>
                                            ) : (
                                                chatMessages.map((msg, i) => (
                                                    <div key={i} style={{ display: 'flex', gap: '12px', alignSelf: msg.role === 'user' ? 'flex-end' : 'flex-start', maxWidth: '97%' }}>
                                                        {msg.role === 'ai' && (
                                                            <div style={{ width: '32px', height: '32px', borderRadius: '8px', background: 'var(--accent-blue)', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                                                                <Bot size={18} color="white" />
                                                            </div>
                                                        )}
                                                            <div style={{ position: 'relative', display: 'flex', flexDirection: 'column', gap: '4px', alignItems: msg.role === 'user' ? 'flex-end' : 'flex-start' }}>
                                                                <div style={{ 
                                                                    padding: '14px 18px', borderRadius: '14px', fontSize: '15px', lineHeight: '1.6',
                                                                    background: msg.role === 'user' ? 'var(--accent-blue)' : 'rgba(255,255,255,0.08)',
                                                                    color: msg.role === 'user' ? 'white' : 'var(--text-primary)',
                                                                    border: msg.role === 'user' ? 'none' : '1px solid var(--glass-border)',
                                                                    whiteSpace: 'pre-wrap',
                                                                    minHeight: '44px',
                                                                    display: 'flex',
                                                                    alignItems: 'center',
                                                                    position: 'relative'
                                                                }}>
                                                                    {msg.text}
                                                                    {msg.role === 'ai' && (
                                                                        <button 
                                                                            onClick={() => handleCopy(msg.text, i)}
                                                                            style={{
                                                                                position: 'absolute', right: '-32px', bottom: '0',
                                                                                background: 'none', border: 'none', color: 'var(--text-muted)',
                                                                                cursor: 'pointer', padding: '4px', display: 'flex',
                                                                                alignItems: 'center', justifyContent: 'center',
                                                                                transition: 'all 0.2s'
                                                                            }}
                                                                            title="복사하기"
                                                                        >
                                                                            {copiedId === i ? <Check size={14} color="#10b981" /> : <Copy size={14} />}
                                                                        </button>
                                                                    )}
                                                                </div>
                                                                <span style={{ fontSize: '10px', color: 'var(--text-muted)' }}>{msg.timestamp}</span>
                                                            </div>
                                                        {msg.role === 'user' && (
                                                            <div style={{ width: '32px', height: '32px', borderRadius: '8px', background: 'rgba(255,255,255,0.1)', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                                                                <User size={18} color="var(--text-secondary)" />
                                                            </div>
                                                        )}
                                                    </div>
                                                ))
                                            )}
                                            {isAnswering && (
                                                <div style={{ display: 'flex', gap: '12px', alignSelf: 'flex-start' }}>
                                                    <div style={{ width: '32px', height: '32px', borderRadius: '8px', background: 'var(--accent-blue)', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                                                        <Loader2 size={18} color="white" className="animate-spin" />
                                                    </div>
                                                    <div style={{ padding: '12px 16px', borderRadius: '14px', background: 'rgba(255,255,255,0.05)', border: '1px solid var(--glass-border)', fontSize: '13px', color: 'var(--text-muted)', fontStyle: 'italic' }}>
                                                        {answerStatus}
                                                    </div>
                                                </div>
                                            )}
                                        </div>

                                        <div style={{ padding: '16px', background: 'rgba(0,0,0,0.2)', borderTop: '1px solid var(--glass-border)', display: 'flex', gap: '10px' }}>
                                            <input 
                                                type="text"
                                                value={userQuestion}
                                                onChange={(e) => setUserQuestion(e.target.value)}
                                                onKeyDown={(e) => e.key === 'Enter' && handleAskQuestion()}
                                                placeholder="문서 내용에 대해 질문하세요..."
                                                disabled={isAnswering}
                                                style={{ 
                                                    flex: 1, padding: '12px 16px', background: 'rgba(255,255,255,0.03)', 
                                                    border: '1px solid var(--glass-border)', borderRadius: '10px', 
                                                    color: 'var(--text-primary)', outline: 'none', fontSize: '14px' 
                                                }}
                                            />
                                            <button 
                                                onClick={handleAskQuestion}
                                                disabled={isAnswering || !userQuestion.trim()}
                                                style={{ 
                                                    width: '44px', height: '44px', borderRadius: '10px', 
                                                    background: (isAnswering || !userQuestion.trim()) ? 'rgba(255,255,255,0.05)' : 'var(--accent-blue)',
                                                    border: 'none', color: 'white', cursor: 'pointer',
                                                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                                                    transition: 'all 0.2s'
                                                }}
                                            >
                                                <Sparkles size={20} />
                                            </button>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    ) : selected_doc_ids.length > 0 ? (
                        <div className="glass-panel animate-scale-in" style={{ display: 'flex', flexDirection: 'column', height: '100%', overflow: 'hidden', borderRadius: '16px', border: '1px solid var(--panel-border)' }}>
                            <div style={{ padding: '24px', borderBottom: '1px solid var(--glass-border)', display: 'flex', alignItems: 'center', gap: '14px', background: 'rgba(59, 130, 246, 0.03)' }}>
                                <div style={{ width: '40px', height: '40px', borderRadius: '10px', background: 'rgba(59, 130, 246, 0.15)', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                                    <Database size={20} color="var(--accent-blue)" />
                                </div>
                                <div style={{ flex: 1, minWidth: 0 }}>
                                    <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 700, color: 'var(--text-primary)' }}>선택한 지식 기반 Q&A</h3>
                                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px', marginTop: '4px', maxHeight: '40px', overflowY: 'auto' }}>
                                        {allDocs.filter(d => selected_doc_ids.includes(d.id)).map(d => (
                                            <div key={d.id} style={{ display: 'flex', alignItems: 'center', gap: '4px', padding: '2px 8px', background: 'rgba(59, 130, 246, 0.08)', border: '1px solid rgba(59, 130, 246, 0.15)', borderRadius: '6px', fontSize: '11px', color: 'var(--text-secondary)' }}>
                                                <FileText size={10} color="var(--accent-blue)" />
                                                <span style={{ maxWidth: '120px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{d.title}</span>
                                            </div>
                                        ))}
                                    </div>
                                </div>
                                <button 
                                    onClick={handle_clear_doc_selection}
                                    className="interactive"
                                    style={{ padding: '6px 14px', background: 'rgba(255,255,255,0.05)', borderRadius: '8px', border: '1px solid var(--glass-border)', color: 'var(--text-muted)', cursor: 'pointer', fontSize: '12px' }}
                                >
                                    선택 해제
                                </button>
                            </div>

                            <div 
                                ref={selectionChatContainerRef}
                                style={{ 
                                    flex: 1, overflowY: 'auto', padding: '24px 10px', display: 'flex', flexDirection: 'column', gap: '16px',
                                    scrollBehavior: 'smooth'
                                }}
                            >
                                {selection_chat_messages.length === 0 ? (
                                    <div style={{ height: '100%', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', color: 'var(--text-muted)', textAlign: 'center', gap: '16px' }}>
                                        <Bot size={48} style={{ color: 'var(--accent-blue)', opacity: 0.3 }} />
                                        <div>
                                            <p style={{ margin: 0, fontSize: '15px', fontWeight: 600, color: 'var(--text-secondary)' }}>선택한 지식 문서({selected_doc_ids.length}개) 대상 질문하기</p>
                                            <p style={{ margin: '4px 0 0', fontSize: '13px', color: 'var(--text-muted)', lineHeight: '1.5' }}>
                                                선택된 지식 문서들 내부 내용으로 검색 대상을 제한하여 질문에 답합니다.
                                            </p>
                                        </div>
                                    </div>
                                ) : (
                                    selection_chat_messages.map((msg, i) => (
                                        <div key={i} style={{ display: 'flex', gap: '12px', alignSelf: msg.role === 'user' ? 'flex-end' : 'flex-start', maxWidth: '97%' }}>
                                            {msg.role === 'ai' && (
                                                <div style={{ width: '32px', height: '32px', borderRadius: '8px', background: 'var(--accent-blue)', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                                                    <Bot size={18} color="white" />
                                                </div>
                                            )}
                                            <div style={{ position: 'relative', display: 'flex', flexDirection: 'column', gap: '4px', alignItems: msg.role === 'user' ? 'flex-end' : 'flex-start' }}>
                                                <div style={{ 
                                                    padding: '14px 18px', borderRadius: '14px', fontSize: '15px', lineHeight: '1.6',
                                                    background: msg.role === 'user' ? 'var(--accent-blue)' : 'rgba(255,255,255,0.08)',
                                                    color: msg.role === 'user' ? 'white' : 'var(--text-primary)',
                                                    border: msg.role === 'user' ? 'none' : '1px solid var(--glass-border)',
                                                    whiteSpace: 'pre-wrap',
                                                    minHeight: '44px',
                                                    display: 'flex',
                                                    alignItems: 'center',
                                                    position: 'relative'
                                                }}>
                                                    {msg.text}
                                                    {msg.role === 'ai' && (
                                                        <button 
                                                            onClick={() => handleCopy(msg.text, `select_${i}`)}
                                                            style={{
                                                                position: 'absolute', right: '-32px', bottom: '0',
                                                                background: 'none', border: 'none', color: 'var(--text-muted)',
                                                                cursor: 'pointer', padding: '4px', display: 'flex',
                                                                alignItems: 'center', justifyContent: 'center',
                                                                transition: 'all 0.2s'
                                                            }}
                                                            title="복사하기"
                                                        >
                                                            {copiedId === `select_${i}` ? <Check size={14} color="#10b981" /> : <Copy size={14} />}
                                                        </button>
                                                    )}
                                                </div>
                                                {msg.role === 'ai' && msg.sources && (
                                                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px', marginTop: '6px' }}>
                                                        {msg.sources.map((src, idx) => (
                                                            <div key={idx} style={{ display: 'flex', alignItems: 'center', gap: '4px', padding: '4px 8px', background: 'rgba(59, 130, 246, 0.1)', border: '1px solid rgba(59, 130, 246, 0.2)', borderRadius: '6px', fontSize: '11px', color: '#93c5fd' }}>
                                                                <BookOpen size={10} />
                                                                {src}
                                                            </div>
                                                        ))}
                                                    </div>
                                                )}
                                                <span style={{ fontSize: '10px', color: 'var(--text-muted)' }}>{msg.timestamp}</span>
                                            </div>
                                            {msg.role === 'user' && (
                                                <div style={{ width: '32px', height: '32px', borderRadius: '8px', background: 'rgba(255,255,255,0.1)', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                                                    <User size={18} color="var(--text-secondary)" />
                                                </div>
                                            )}
                                        </div>
                                    ))
                                )}
                                {is_selection_answering && (
                                    <div style={{ display: 'flex', gap: '12px', alignSelf: 'flex-start' }}>
                                        <div style={{ width: '32px', height: '32px', borderRadius: '8px', background: 'var(--accent-blue)', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                                            <Loader2 size={18} color="white" className="animate-spin" />
                                        </div>
                                        <div style={{ padding: '12px 16px', borderRadius: '14px', background: 'rgba(255,255,255,0.05)', border: '1px solid var(--glass-border)', fontSize: '13px', color: 'var(--text-muted)', fontStyle: 'italic' }}>
                                            {selection_answer_status}
                                        </div>
                                    </div>
                                )}
                            </div>

                            <div style={{ padding: '16px', background: 'rgba(0,0,0,0.2)', borderTop: '1px solid var(--glass-border)', display: 'flex', gap: '10px' }}>
                                <input 
                                    type="text"
                                    value={selection_user_question}
                                    onChange={(e) => set_selection_user_question(e.target.value)}
                                    onKeyDown={(e) => e.key === 'Enter' && handle_ask_selection_question()}
                                    placeholder="선택한 지식 문서 대상 질문을 입력하세요..."
                                    disabled={is_selection_answering}
                                    style={{ 
                                        flex: 1, padding: '12px 16px', background: 'rgba(255,255,255,0.03)', 
                                        border: '1px solid var(--glass-border)', borderRadius: '10px', 
                                        color: 'var(--text-primary)', outline: 'none', fontSize: '14px' 
                                    }}
                                />
                                <button 
                                    onClick={handle_ask_selection_question}
                                    disabled={is_selection_answering || !selection_user_question.trim()}
                                    style={{ 
                                        width: '44px', height: '44px', borderRadius: '10px', 
                                        background: (is_selection_answering || !selection_user_question.trim()) ? 'rgba(255,255,255,0.05)' : 'var(--accent-blue)',
                                        border: 'none', color: 'white', cursor: 'pointer',
                                        display: 'flex', alignItems: 'center', justifyContent: 'center',
                                        transition: 'all 0.2s'
                                    }}
                                >
                                    <Send size={20} />
                                </button>
                            </div>
                        </div>
                    ) : (
                        <div className="glass-panel animate-scale-in" style={{ display: 'flex', flexDirection: 'column', height: '100%', overflow: 'hidden', borderRadius: '16px', border: '1px solid var(--panel-border)' }}>
                            <div style={{ padding: '24px', borderBottom: '1px solid var(--glass-border)', display: 'flex', alignItems: 'center', gap: '14px', background: 'rgba(59, 130, 246, 0.03)' }}>
                                <div style={{ width: '40px', height: '40px', borderRadius: '10px', background: 'rgba(59, 130, 246, 0.15)', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                                    <Sparkles size={20} color="var(--accent-blue)" />
                                </div>
                                <div>
                                    <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 700, color: 'var(--text-primary)' }}>전체 지식베이스 통합 Q&A</h3>
                                    <p style={{ margin: '2px 0 0', fontSize: '12px', color: 'var(--text-muted)' }}>등록된 모든 문서를 검색하여 AI가 종합 답변을 도출합니다.</p>
                                </div>
                            </div>

                            <div 
                                ref={totalChatContainerRef}
                                style={{ 
                                    flex: 1, overflowY: 'auto', padding: '24px 10px', display: 'flex', flexDirection: 'column', gap: '16px',
                                    scrollBehavior: 'smooth'
                                }}
                            >
                                {totalChatMessages.length === 0 ? (
                                    <div style={{ height: '100%', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', color: 'var(--text-muted)', textAlign: 'center', gap: '16px' }}>
                                        <Bot size={48} style={{ color: 'var(--accent-blue)', opacity: 0.3 }} />
                                        <div>
                                            <p style={{ margin: 0, fontSize: '15px', fontWeight: 600, color: 'var(--text-secondary)' }}>전체 지식 기반 질문하기</p>
                                            <p style={{ margin: '4px 0 0', fontSize: '13px', color: 'var(--text-muted)', lineHeight: '1.5' }}>
                                                개별 문서를 지정하지 않고 전체 산출물 데이터에서 지식을 찾아 답변합니다.<br/>
                                                예: "클라우드 마이그레이션 이행을 위한 핵심 고려사항이 정리된 문서는 뭐야?"
                                            </p>
                                        </div>
                                    </div>
                                ) : (
                                    totalChatMessages.map((msg, i) => (
                                        <div key={i} style={{ display: 'flex', gap: '12px', alignSelf: msg.role === 'user' ? 'flex-end' : 'flex-start', maxWidth: '97%' }}>
                                            {msg.role === 'ai' && (
                                                <div style={{ width: '32px', height: '32px', borderRadius: '8px', background: 'var(--accent-blue)', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                                                    <Bot size={18} color="white" />
                                                </div>
                                            )}
                                            <div style={{ position: 'relative', display: 'flex', flexDirection: 'column', gap: '4px', alignItems: msg.role === 'user' ? 'flex-end' : 'flex-start' }}>
                                                <div style={{ 
                                                    padding: '14px 18px', borderRadius: '14px', fontSize: '15px', lineHeight: '1.6',
                                                    background: msg.role === 'user' ? 'var(--accent-blue)' : 'rgba(255,255,255,0.08)',
                                                    color: msg.role === 'user' ? 'white' : 'var(--text-primary)',
                                                    border: msg.role === 'user' ? 'none' : '1px solid var(--glass-border)',
                                                    whiteSpace: 'pre-wrap',
                                                    minHeight: '44px',
                                                    display: 'flex',
                                                    alignItems: 'center',
                                                    position: 'relative'
                                                }}>
                                                    {msg.text}
                                                    {msg.role === 'ai' && (
                                                        <button 
                                                            onClick={() => handleCopy(msg.text, `total_${i}`)}
                                                            style={{
                                                                position: 'absolute', right: '-32px', bottom: '0',
                                                                background: 'none', border: 'none', color: 'var(--text-muted)',
                                                                cursor: 'pointer', padding: '4px', display: 'flex',
                                                                alignItems: 'center', justifyContent: 'center',
                                                                transition: 'all 0.2s'
                                                            }}
                                                            title="복사하기"
                                                        >
                                                            {copiedId === `total_${i}` ? <Check size={14} color="#10b981" /> : <Copy size={14} />}
                                                        </button>
                                                    )}
                                                </div>
                                                {msg.role === 'ai' && msg.sources && (
                                                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px', marginTop: '6px' }}>
                                                        {msg.sources.map((src, idx) => (
                                                            <div key={idx} style={{ display: 'flex', alignItems: 'center', gap: '4px', padding: '4px 8px', background: 'rgba(59, 130, 246, 0.1)', border: '1px solid rgba(59, 130, 246, 0.2)', borderRadius: '6px', fontSize: '11px', color: '#93c5fd' }}>
                                                                <BookOpen size={10} />
                                                                {src}
                                                            </div>
                                                        ))}
                                                    </div>
                                                )}
                                                <span style={{ fontSize: '10px', color: 'var(--text-muted)' }}>{msg.timestamp}</span>
                                            </div>
                                            {msg.role === 'user' && (
                                                <div style={{ width: '32px', height: '32px', borderRadius: '8px', background: 'rgba(255,255,255,0.1)', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                                                    <User size={18} color="var(--text-secondary)" />
                                                </div>
                                            )}
                                        </div>
                                    ))
                                )}
                                {isTotalAnswering && (
                                    <div style={{ display: 'flex', gap: '12px', alignSelf: 'flex-start' }}>
                                        <div style={{ width: '32px', height: '32px', borderRadius: '8px', background: 'var(--accent-blue)', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                                            <Loader2 size={18} color="white" className="animate-spin" />
                                        </div>
                                        <div style={{ padding: '12px 16px', borderRadius: '14px', background: 'rgba(255,255,255,0.05)', border: '1px solid var(--glass-border)', fontSize: '13px', color: 'var(--text-muted)', fontStyle: 'italic' }}>
                                            {totalAnswerStatus}
                                        </div>
                                    </div>
                                )}
                            </div>

                            <div style={{ padding: '16px', background: 'rgba(0,0,0,0.2)', borderTop: '1px solid var(--glass-border)', display: 'flex', gap: '10px' }}>
                                <input 
                                    type="text"
                                    value={totalUserQuestion}
                                    onChange={(e) => setTotalUserQuestion(e.target.value)}
                                    onKeyDown={(e) => e.key === 'Enter' && handleAskTotalQuestion()}
                                    placeholder="전체 지식베이스 대상 질문을 입력하세요..."
                                    disabled={isTotalAnswering}
                                    style={{ 
                                        flex: 1, padding: '12px 16px', background: 'rgba(255,255,255,0.03)', 
                                        border: '1px solid var(--glass-border)', borderRadius: '10px', 
                                        color: 'var(--text-primary)', outline: 'none', fontSize: '14px' 
                                    }}
                                />
                                <button 
                                    onClick={handleAskTotalQuestion}
                                    disabled={isTotalAnswering || !totalUserQuestion.trim()}
                                    style={{ 
                                        width: '44px', height: '44px', borderRadius: '10px', 
                                        background: (isTotalAnswering || !totalUserQuestion.trim()) ? 'rgba(255,255,255,0.05)' : 'var(--accent-blue)',
                                        border: 'none', color: 'white', cursor: 'pointer',
                                        display: 'flex', alignItems: 'center', justifyContent: 'center',
                                        transition: 'all 0.2s'
                                    }}
                                >
                                    <Send size={20} />
                                </button>
                            </div>
                        </div>
                    )}
                </div>
            </div>

            <div style={{ marginTop: '20px', padding: '16px', background: 'rgba(59, 130, 246, 0.05)', borderRadius: '12px', border: '1px solid rgba(59, 130, 246, 0.1)', display: 'flex', gap: '12px', alignItems: 'center' }}>
                <Info size={20} color="var(--accent-blue)" />
                <div style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>
                    프로젝트 RAG 지식베이스는 등록된 문서를 기반으로 AI 분석을 수행합니다. 새로운 문서를 추가하려면 <b>'문서 직접 등록'</b>을 이용하거나, 로컬 폴더의 파일 변경 시 <b>'지식베이스 갱신'</b> 버튼을 클릭하여 인덱스를 업데이트해 주세요.
                </div>
            </div>
        </div>
    );
};

export default RagKnowledgeBase;
