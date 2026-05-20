import React, { useState, useEffect, useRef } from 'react';
import { Plus, Trash2, ExternalLink, FileText, Info, Search, AlertCircle, Upload, Loader2, X as CloseIcon, FileCode, FileType, FileLineChart, FileBox, Library } from 'lucide-react';
import { processFile, ALL_ACCEPT, getFileExtension, classifyFile } from '../utils/fileExtractor';
import { refDB } from '../utils/db';

const ReferenceLibrary = () => {
    const [docs, setDocs] = useState([]);
    const [isAdding, setIsAdding] = useState(false);
    const [newDoc, setNewDoc] = useState({ title: '', filename: '', content: '', type: 'file', ext: '', size: 0, blob: null });
    const [searchTerm, setSearchTerm] = useState('');
    const [isProcessing, setIsProcessing] = useState(false);
    const [uploadError, setUploadError] = useState(null);
    const [isDragging, setIsDragging] = useState(false);
    const [isClearing, setIsClearing] = useState(false);
    const [deletingId, setDeletingId] = useState(null);
    const [confirmModal, setConfirmModal] = useState({ open: false, type: 'single', id: null, title: '' });
    const fileInputRef = useRef(null);

    const openClearAllConfirm = (e) => {
        if (e && e.stopPropagation) e.stopPropagation();
        setConfirmModal({ open: true, type: 'all', id: null, title: '모든 보관 문서' });
    };

    const handleConfirmAction = async () => {
        const { type, id } = confirmModal;
        setIsClearing(true);
        const targetId = id; // 로컬 변수로 보존
        setConfirmModal({ ...confirmModal, open: false });
        
        try {
            if (type === 'all') {
                console.log("Starting Clear All transaction...");
                await refDB.clearAllDocs();
                setDocs([]);
            } else {
                console.log("Starting Delete transaction for ID:", targetId);
                setDeletingId(targetId);
                await refDB.deleteDoc(targetId);
                setDocs(prev => prev.filter(d => String(d.id) !== String(targetId)));
            }
            window.dispatchEvent(new CustomEvent('refdocs_changed'));
        } catch (err) {
            console.error('Operation failed:', err);
            alert('작업 중 오류가 발생했습니다: ' + err.message);
        } finally {
            setIsClearing(false);
            setDeletingId(null);
        }
    };

    // Initial Load & Migration
    useEffect(() => {
        const loadDocs = async () => {
            try {
                // 성능 최적화: 메타데이터만 우선 로드
                let dbDocs = await refDB.getDocsMetadata();
                
                // localStorage 마이그레이션 (기존 사용자를 위한 1회성 로직)
                const legacyData = localStorage.getItem('rfp_reference_docs');
                if (legacyData && dbDocs.length === 0) {
                    const parsed = JSON.parse(legacyData);
                    for (const item of parsed) {
                        await refDB.saveDoc({ ...item, size: 0, ext: 'txt' });
                    }
                    dbDocs = await refDB.getDocsMetadata();
                    localStorage.removeItem('rfp_reference_docs'); // 마이그레이션 후 제거
                }
                
                setDocs(dbDocs.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt)));
            } catch (err) {
                console.error('Failed to load docs:', err);
            }
        };
        loadDocs();
    }, []);

    const handleAdd = async (e) => {
        e.preventDefault();
        if (!newDoc.title.trim() || !newDoc.content.trim()) return;

        const docToSave = { 
            ...newDoc, 
            id: Date.now(), 
            createdAt: new Date().toISOString() 
        };

        try {
            setIsProcessing(true);
            await refDB.saveDoc(docToSave);
            // 목록용으로는 메타데이터만 필요
            const { content, blob, ...meta } = docToSave;
            setDocs(current => [meta, ...current]);
            window.dispatchEvent(new CustomEvent('refdocs_changed'));
            setNewDoc({ title: '', filename: '', content: '', type: 'file', ext: '', size: 0, blob: null });
            setIsAdding(false);
            setUploadError(null);
        } catch (err) {
            setUploadError('저장에 실패했습니다: ' + err.message);
        } finally {
            setIsProcessing(false);
        }
    };

    const handleFileChange = async (e) => {
        const file = e.target ? e.target.files[0] : e;
        if (!file) return;

        setIsProcessing(true);
        setUploadError(null);
        try {
            const result = await processFile(file);
            const text = result.text;
            const title = file.name.replace(/\.[^/.]+$/, "");
            const ext = getFileExtension(file.name);
            
            setNewDoc({ 
                title: title, 
                filename: file.name, 
                content: text, 
                type: 'file',
                ext: ext,
                size: file.size,
                blob: file // 원본 파일(Blob) 보존
            });
            setIsAdding(true);
        } catch (err) {
            console.error(err);
            setUploadError(err.message || '파일 처리에 실패했습니다.');
        } finally {
            setIsProcessing(false);
            if (fileInputRef.current) fileInputRef.current.value = '';
        }
    };

    const handleDragOver = (e) => { e.preventDefault(); e.stopPropagation(); setIsDragging(true); };
    const handleDragLeave = (e) => { e.preventDefault(); e.stopPropagation(); setIsDragging(false); };
    const handleDrop = (e) => {
        e.preventDefault(); e.stopPropagation(); setIsDragging(false);
        const files = e.dataTransfer.files;
        if (files && files.length > 0) handleFileChange(files[0]);
    };

    const openDeleteConfirm = (e, id, title) => {
        if (e) {
            e.preventDefault();
            e.stopPropagation();
        }
        console.log("Opening delete confirm for ID:", id, "(", title, ")");
        setConfirmModal({ open: true, type: 'single', id, title });
    };

    const handleDocClick = async (doc, inNewTab = false) => {
        if (doc.type === 'link') {
            window.open(doc.content, '_blank', 'noopener,noreferrer');
        } else {
            setIsClearing(true); 
            try {
                const fullDoc = await refDB.getDocById(doc.id);
                if (!fullDoc) throw new Error('문서를 찾을 수 없습니다.');

                if (inNewTab && fullDoc.blob) {
                    const url = URL.createObjectURL(fullDoc.blob);
                    window.open(url, '_blank');
                } else if (inNewTab && !fullDoc.blob) {
                    const newWindow = window.open('', '_blank');
                    if (newWindow) {
                        newWindow.document.write(`<pre style="white-space: pre-wrap; font-family: sans-serif; padding: 20px;">${fullDoc.content}</pre>`);
                        newWindow.document.title = fullDoc.title;
                        newWindow.document.close();
                    }
                } else {
                    const event = new CustomEvent('open_reference_modal', { 
                        detail: { 
                            title: fullDoc.title, 
                            content: fullDoc.content,
                            blob: fullDoc.blob,
                            ext: fullDoc.ext,
                            filename: fullDoc.filename
                        } 
                    });
                    window.dispatchEvent(event);
                }
            } catch (err) {
                alert('문서를 불러오지 못했습니다: ' + err.message);
            } finally {
                setIsClearing(false);
            }
        }
    };


    const formatSize = (bytes) => {
        if (bytes === 0) return '0 B';
        const k = 1024;
        const sizes = ['B', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
    };

    const getFileIcon = (ext) => {
        const type = classifyFile(ext);
        switch(type) {
            case 'pdf': return <FileText color="#ef4444" />;
            case 'excel': return <FileLineChart color="#10b981" />;
            case 'pptx': return <FileType color="#f59e0b" />;
            case 'hwpx': return <FileCode color="#3b82f6" />;
            default: return <FileBox color="var(--text-secondary)" />;
        }
    };

    const filteredDocs = docs.filter(doc => 
        doc.title.toLowerCase().includes(searchTerm.toLowerCase()) ||
        (doc.filename && doc.filename.toLowerCase().includes(searchTerm.toLowerCase()))
    );

    return (
        <div style={{ flex: 1, display: 'flex', flexDirection: 'column', height: '100%', overflow: 'hidden' }}>
            <div className="glass-panel animate-slide-up" style={{ padding: '24px 32px', marginBottom: '32px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', border: '1px solid var(--panel-border)', borderRadius: '16px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '20px' }}>
                    <div style={{ width: '48px', height: '48px', borderRadius: '14px', background: 'rgba(168, 85, 247, 0.1)', display: 'flex', alignItems: 'center', justifyContent: 'center', border: '1px solid rgba(168, 85, 247, 0.2)' }}>
                        <Library size={28} color="var(--accent-purple)" />
                    </div>
                    <div>
                        <h2 style={{ margin: 0, fontSize: '22px', fontWeight: 700, color: 'var(--text-primary)', letterSpacing: '-0.5px' }}>참고자료 보관소</h2>
                        <p style={{ margin: '4px 0 0', fontSize: '14px', color: 'var(--text-secondary)' }}>AI 분석 및 자문의 근거가 되는 사용자 등록 지침 보관소</p>
                    </div>
                </div>
                
                <div style={{ display: 'flex', gap: '12px' }}>
                    <button 
                        type="button"
                        onClick={(e) => openClearAllConfirm(e)}
                        disabled={isClearing}
                        className="interactive"
                        style={{ 
                            padding: '10px 18px', borderRadius: '10px', fontWeight: 600, 
                            display: 'flex', alignItems: 'center', gap: '8px',
                            background: 'rgba(239, 68, 68, 0.08)', border: '1px solid rgba(239, 68, 68, 0.2)',
                            color: 'var(--danger-color)', fontSize: '13px', cursor: isClearing ? 'wait' : 'pointer'
                        }}
                    >
                        {isClearing ? <Loader2 className="animate-spin" size={16} /> : <Trash2 size={16} />}
                        전체 초기화
                    </button>

                    <button 
                        onClick={() => setIsAdding(!isAdding)}
                        className="primary interactive"
                        style={{ 
                            padding: '10px 20px', borderRadius: '10px', fontWeight: 700, 
                            display: 'flex', alignItems: 'center', gap: '8px',
                            background: isAdding ? 'rgba(255,255,255,0.05)' : 'linear-gradient(135deg, var(--accent-purple), var(--accent-blue))',
                            border: isAdding ? '1px solid var(--glass-border)' : 'none',
                            color: 'white', fontSize: '13px',
                            boxShadow: isAdding ? 'none' : '0 8px 24px rgba(168, 85, 247, 0.2)'
                        }}
                    >
                        {isAdding ? <CloseIcon size={18} /> : <Plus size={18} />} {isAdding ? '닫기' : '새 파일 등록'}
                    </button>
                </div>
            </div>

            <div style={{ flex: 1, overflowY: 'auto', padding: '0 4px 40px' }}>

            {isAdding && (
                <div role="form" className="glass-panel animate-slide-up" style={{ marginBottom: '32px', padding: '32px', border: '1px solid var(--glass-border)' }}>
                    <div 
                        className="interactive"
                        style={{ 
                            marginBottom: '32px', 
                            padding: '40px', 
                            border: `2px dashed ${isDragging ? 'var(--accent-blue)' : 'var(--glass-border)'}`, 
                            borderRadius: '20px', 
                            textAlign: 'center', 
                            background: isDragging ? 'rgba(59, 130, 246, 0.05)' : 'rgba(0,0,0,0.2)', 
                        }}
                        onClick={() => fileInputRef.current?.click()}
                        onDragOver={handleDragOver}
                        onDragLeave={handleDragLeave}
                        onDrop={handleDrop}
                    >
                        <input type="file" ref={fileInputRef} style={{ display: 'none' }} accept={ALL_ACCEPT} onChange={handleFileChange} />
                        {isProcessing ? (
                            <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '16px' }}>
                                <Loader2 className="animate-spin" size={40} color="var(--accent-blue)" />
                                <p style={{ margin: 0, fontSize: '15px', fontWeight: 600 }}>데이터 분석 중...</p>
                            </div>
                        ) : (
                            <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '16px' }}>
                                <Upload size={48} color="var(--accent-blue)" style={{ opacity: 0.8 }} />
                                <div>
                                    <p style={{ margin: '0 0 6px', fontSize: '16px', color: 'var(--text-primary)', fontWeight: 700 }}>파일을 드래그하거나 클릭하여 업로드</p>
                                    <p style={{ margin: 0, fontSize: '13px', color: 'var(--text-secondary)' }}>PDF, Excel, PPTX, HWPX 등의 텍스트를 자동 보관합니다.</p>
                                </div>
                            </div>
                        )}
                    </div>

                    {uploadError && (
                        <div style={{ marginBottom: '24px', padding: '16px', background: 'rgba(239, 68, 68, 0.08)', border: '1px solid rgba(239, 68, 68, 0.2)', borderRadius: '12px', color: 'var(--danger-color)', fontSize: '14px', display: 'flex', alignItems: 'center', gap: '10px', fontWeight: 500 }}>
                            <AlertCircle size={18} /> {uploadError}
                        </div>
                    )}

                    <form onSubmit={handleAdd} style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
                        <div style={{ display: 'grid', gridTemplateColumns: '1.5fr 1fr', gap: '20px' }}>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                                <label style={{ fontSize: '13px', fontWeight: 600, color: 'var(--text-secondary)', marginLeft: '4px' }}>문서 제목 (매칭 키워드)</label>
                                <input type="text" value={newDoc.title} onChange={e => setNewDoc({...newDoc, title: e.target.value})} placeholder="파일을 업로드하면 제목이 자동 입력됩니다." style={{ width: '100%' }} required />
                            </div>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                                <label style={{ fontSize: '13px', fontWeight: 600, color: 'var(--text-secondary)', marginLeft: '4px' }}>구분</label>
                                <div style={{ display: 'flex', gap: '10px' }}>
                                    {['file', 'link'].map(t => (
                                        <button key={t} type="button" onClick={() => setNewDoc({...newDoc, type: t})} className="interactive" style={{ flex: 1, padding: '12px', borderRadius: '12px', background: newDoc.type === t ? 'rgba(59, 130, 246, 0.1)' : 'rgba(0,0,0,0.2)', border: `1px solid ${newDoc.type === t ? 'var(--accent-blue)' : 'var(--glass-border)'}`, color: newDoc.type === t ? 'var(--accent-blue)' : 'var(--text-secondary)', fontWeight: 600, fontSize: '13px' }}>
                                            {t.toUpperCase()}
                                        </button>
                                    ))}
                                </div>
                            </div>
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                            <label style={{ fontSize: '13px', fontWeight: 600, color: 'var(--text-secondary)', marginLeft: '4px' }}>{newDoc.type === 'file' ? '추출된 본문 데이터' : '외부 링크 URL'}</label>
                            {newDoc.type === 'file' ? (
                                <textarea value={newDoc.content} onChange={e => setNewDoc({...newDoc, content: e.target.value})} placeholder="직접 입력하거나 파일을 업로드하세요." style={{ height: '180px' }} required />
                            ) : (
                                <input type="url" value={newDoc.content} onChange={e => setNewDoc({...newDoc, content: e.target.value})} placeholder="https://법령정보.com/..." style={{ width: '100%' }} required />
                            )}
                        </div>
                        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '16px', marginTop: '8px' }}>
                            <button type="submit" className="primary interactive" style={{ padding: '14px 40px', borderRadius: '12px', background: 'linear-gradient(135deg, var(--accent-blue), var(--accent-violet))', color: 'white', fontWeight: 700, border: 'none', boxShadow: '0 8px 24px rgba(59, 130, 246, 0.2)' }}>이 파일 보관하기</button>
                        </div>
                    </form>
                </div>
            )}

            <div style={{ position: 'sticky', top: 0, zIndex: 10, background: 'var(--bg-dark)', paddingBottom: '24px' }}>
                <div style={{ position: 'relative' }}>
                    <Search size={20} style={{ position: 'absolute', left: '16px', top: '50%', transform: 'translateY(-50%)', color: 'var(--text-muted)' }} />
                    <input type="text" placeholder="보관소 내 문서 명칭 또는 파일명 검색..." value={searchTerm} onChange={e => setSearchTerm(e.target.value)} style={{ width: '100%', padding: '16px 16px 16px 52px', background: 'rgba(255,255,255,0.03)', borderRadius: '16px', fontSize: '15px' }} />
                </div>
            </div>

            <div style={{ flex: 1 }}>
                {filteredDocs.length > 0 ? (
                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(360px, 1fr))', gap: '20px' }}>
                        {filteredDocs.map(doc => (
                            <div key={doc.id} className="glass-panel" style={{ padding: '24px', borderRadius: '20px', display: 'flex', flexDirection: 'column', gap: '16px', border: '1px solid var(--glass-border)', position: 'relative' }}>
                                {/* Delete Button - Posited Absolutely for independent hit target */}
                                <button 
                                    type="button"
                                    onClick={(e) => openDeleteConfirm(e, doc.id, doc.title)} 
                                    disabled={deletingId === doc.id}
                                    className="interactive" 
                                    style={{ 
                                        position: 'absolute', top: '24px', right: '24px', zIndex: 10,
                                        background: 'rgba(239, 68, 68, 0.08)', border: 'none', 
                                        color: 'var(--danger-color)', 
                                        cursor: deletingId === doc.id ? 'wait' : 'pointer', 
                                        padding: '8px', borderRadius: '10px' 
                                    }}
                                >
                                    {deletingId === doc.id ? <Loader2 className="animate-spin" size={18} /> : <Trash2 size={18} />}
                                </button>

                                {/* Clickable Content Region (Modal trigger) */}
                                <div 
                                    onClick={() => handleDocClick(doc)}
                                    className="interactive"
                                    style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: '16px', cursor: 'pointer' }}
                                >
                                    <div style={{ display: 'flex', alignItems: 'center', gap: '10px', paddingRight: '40px' }}>
                                        <div style={{ padding: '10px', background: 'rgba(255,255,255,0.03)', borderRadius: '12px', border: '1px solid var(--glass-border)', boxShadow: 'inset 0 0 10px rgba(255,255,255,0.02)' }}>
                                            {doc.type === 'link' ? <ExternalLink size={20} color="var(--accent-blue)" /> : getFileIcon(doc.ext)}
                                        </div>
                                        <div>
                                            <div style={{ fontSize: '10px', fontWeight: 800, color: 'var(--text-muted)', letterSpacing: '1px', textTransform: 'uppercase', marginBottom: '2px' }}>{doc.type}</div>
                                            <h3 style={{ margin: 0, fontSize: '17px', fontWeight: 700, color: 'var(--text-primary)', lineHeight: '1.4' }}>{doc.title}</h3>
                                        </div>
                                    </div>

                                    <div style={{ flex: 1 }}>
                                        {doc.filename && (
                                            <div style={{ fontSize: '12px', color: 'var(--accent-blue)', marginBottom: '10px', display: 'flex', alignItems: 'center', gap: '6px', fontWeight: 600, opacity: 0.9 }}>
                                                <FileText size={14} /> {doc.filename} {doc.size > 0 && <span style={{ opacity: 0.5 }}>({formatSize(doc.size)})</span>}
                                            </div>
                                        )}
                                        <p style={{ margin: 0, fontSize: '14px', color: 'var(--text-secondary)', display: '-webkit-box', WebkitLineClamp: 3, WebkitBoxOrient: 'vertical', overflow: 'hidden', lineHeight: '1.6' }}>
                                            {doc.content}
                                        </p>
                                    </div>
                                </div>

                                <div style={{ padding: '14px 0 0', borderTop: '1px solid var(--glass-border)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', zIndex: 5 }}>
                                    <span style={{ fontSize: '12px', color: 'var(--text-muted)', fontWeight: 500 }}>{new Date(doc.createdAt).toLocaleDateString()} 보관</span>
                                    <button 
                                        type="button"
                                        onClick={(e) => {
                                            e.stopPropagation();
                                            handleDocClick(doc, true);
                                        }}
                                        className="interactive"
                                        style={{ 
                                            background: 'transparent', border: 'none', cursor: 'pointer',
                                            display: 'flex', alignItems: 'center', gap: '4px', 
                                            fontSize: '13px', color: 'var(--accent-blue)', fontWeight: 700 
                                        }}
                                    >
                                        새 탭에서 보기 <ExternalLink size={14} />
                                    </button>
                                </div>
                            </div>
                        ))}
                    </div>
                ) : (
                    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', height: '300px', color: 'var(--text-muted)', gap: '20px' }}>
                        <div style={{ padding: '30px', background: 'rgba(255,255,255,0.02)', borderRadius: '50%', border: '1px solid var(--glass-border)' }}>
                            <Info size={48} strokeWidth={1.5} />
                        </div>
                        <p style={{ fontSize: '16px', fontWeight: 500 }}>검색 결과가 없거나 저장된 문서가 없습니다.</p>
                    </div>
                )}
            </div>

            <div style={{ marginTop: '40px', padding: '24px', background: 'rgba(59, 130, 246, 0.05)', borderRadius: '20px', border: '1px solid rgba(59, 130, 246, 0.1)', display: 'flex', gap: '16px' }}>
                <Info size={24} color="var(--accent-blue)" style={{ flexShrink: 0 }} />
                <div style={{ fontSize: '14px', color: 'var(--text-secondary)', lineHeight: '1.7' }}>
                    <strong style={{ color: 'var(--text-primary)' }}>💡 보관 문서 활용 팁:</strong> AI 분석 결과에서 특정 문서 명칭이 언급될 때, 이곳에 등록된 문서의 내용을 즉시 팝업으로 확인할 수 있습니다. 명칭을 <strong>가이드/지표의 정확한 이름</strong>으로 저장해 주세요.
                </div>
            </div>
            {/* Custom Confirmation Modal */}
            {confirmModal.open && (
                <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(0,0,0,0.7)', backdropFilter: 'blur(8px)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 20000 }}>
                    <div className="glass-panel animate-slide-up" style={{ width: '400px', padding: '32px', textAlign: 'center' }}>
                        <div style={{ marginBottom: '20px', display: 'flex', justifyContent: 'center' }}>
                            <div style={{ padding: '16px', background: 'rgba(239, 68, 68, 0.1)', borderRadius: '50%', color: 'var(--danger-color)' }}>
                                <Trash2 size={32} />
                            </div>
                        </div>
                        <h3 style={{ margin: '0 0 12px', fontSize: '18px', fontWeight: 700 }}>정말 삭제하시겠습니까?</h3>
                        <p style={{ margin: '0 0 32px', fontSize: '14px', color: 'var(--text-secondary)', lineHeight: '1.6' }}>
                            <strong style={{ color: 'var(--text-primary)' }}>{confirmModal.title}</strong> 문서를 보관소에서 영구적으로 삭제합니다. 이 작업은 되돌릴 수 없습니다.
                        </p>
                        <div style={{ display: 'flex', gap: '12px' }}>
                            <button 
                                onClick={() => setConfirmModal({ ...confirmModal, open: false })}
                                className="interactive"
                                style={{ flex: 1, padding: '12px', borderRadius: '12px', background: 'rgba(255,255,255,0.05)', border: '1px solid var(--glass-border)', color: 'var(--text-primary)', fontWeight: 600, fontSize: '14px' }}
                            >
                                취소
                            </button>
                            <button 
                                onClick={handleConfirmAction}
                                className="interactive"
                                style={{ flex: 1, padding: '12px', borderRadius: '12px', background: 'var(--danger-color)', border: 'none', color: 'white', fontWeight: 700, fontSize: '14px', boxShadow: '0 8px 16px rgba(239, 68, 68, 0.25)' }}
                            >
                                삭제 승인
                            </button>
                        </div>
                    </div>
                </div>
            )}
            </div>
        </div>
    );
};

export default ReferenceLibrary;
