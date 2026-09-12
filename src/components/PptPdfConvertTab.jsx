import React, { useState, useRef, useEffect } from 'react';
import { Presentation, Upload, Play, Loader2, CheckCircle2 } from 'lucide-react';

// PPT ➜ PDF 변환(로컬 헬퍼 서버 연동) 탭. PptGenerator.jsx의 pdf_convert 서브탭에서 분리됨.
export default function PptPdfConvertTab({ setErrorMsg, setSuccessMsg }) {
    const [local_file_path, set_local_file_path] = useState('');
    const [is_server_connected, set_is_server_connected] = useState(false);
    const [is_converting, set_is_converting] = useState(false);
    const [pdf_convert_file, set_pdf_convert_file] = useState(null);
    const [is_dragging_pdf, set_is_dragging_pdf] = useState(false);
    const pdf_input_ref = useRef(null);

    // 로컬 헬퍼 서버 상태 실시간 감지 (3초 주기 폴링)
    useEffect(() => {
        const check_server_status = async () => {
            try {
                const response = await fetch('http://127.0.0.1:5000/status', { method: 'GET' });
                if (response.ok) {
                    const data = await response.json();
                    if (data.status === 'online') {
                        set_is_server_connected(true);
                        return;
                    }
                }
                set_is_server_connected(false);
            } catch {
                set_is_server_connected(false);
            }
        };

        check_server_status();
        const interval_id = setInterval(check_server_status, 3000);
        return () => clearInterval(interval_id);
    }, []);

    // 동일 디렉토리에 PDF 생성 (로컬 절대 경로 방식)
    const handle_local_pdf_convert = async () => {
        if (!local_file_path.trim()) {
            setErrorMsg('로컬 파일의 절대 경로를 입력해주세요.');
            return;
        }
        if (!is_server_connected) {
            setErrorMsg('로컬 변환 헬퍼 서버가 오프라인 상태입니다. 서버가 구동 중인지 확인해 주세요.');
            return;
        }

        setErrorMsg(null);
        setSuccessMsg(null);
        set_is_converting(true);

        try {
            const response = await fetch('http://127.0.0.1:5000/convert-local', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    local_file_path: local_file_path.trim()
                })
            });

            const data = await response.json();
            if (response.ok && data.success) {
                setSuccessMsg(`성공적으로 변환되었습니다! 동일 디렉토리에 PDF 파일이 생성되었습니다:\n${data.pdf_path}`);
                set_local_file_path('');
            } else {
                setErrorMsg(data.message || 'PDF 변환 중 오류가 발생했습니다.');
            }
        } catch (err) {
            console.error(err);
            setErrorMsg('로컬 서버와의 통신에 실패했습니다. 서버가 구동 중인지 확인해주세요.');
        } finally {
            set_is_converting(false);
        }
    };

    // 브라우저 드래그앤드롭 업로드 변환 후 다운로드 방식
    const handle_upload_pdf_convert = async (file_to_convert) => {
        const file = file_to_convert || pdf_convert_file;
        if (!file) {
            setErrorMsg('변환할 PPTX 파일을 선택하거나 드래그해 주세요.');
            return;
        }
        if (!is_server_connected) {
            setErrorMsg('로컬 변환 헬퍼 서버가 오프라인 상태입니다. 서버가 구동 중인지 확인해 주세요.');
            return;
        }

        setErrorMsg(null);
        setSuccessMsg(null);
        set_is_converting(true);

        const form_data = new FormData();
        form_data.append('file', file);

        try {
            const response = await fetch('http://127.0.0.1:5000/convert-upload', {
                method: 'POST',
                body: form_data
            });

            if (response.ok) {
                const blob = await response.blob();
                const download_url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = download_url;
                a.download = file.name.replace(/\.[^/.]+$/, "") + ".pdf";
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                window.URL.revokeObjectURL(download_url);

                setSuccessMsg(`성공적으로 변환되어 PDF 파일이 다운로드되었습니다: ${file.name.replace(/\.[^/.]+$/, "")}.pdf`);
                set_pdf_convert_file(null);
            } else {
                const data = await response.json();
                setErrorMsg(data.message || 'PDF 변환 중 오류가 발생했습니다.');
            }
        } catch (err) {
            console.error(err);
            setErrorMsg('로컬 서버와의 통신에 실패했습니다. 서버가 구동 중인지 확인해주세요.');
        } finally {
            set_is_converting(false);
        }
    };

    return (
        <div className="animate-fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
            {/* 서버 연결 상태 위젯 */}
            <div style={{
                display: 'flex', alignItems: 'center', justifyContent: 'space-between',
                padding: '16px 20px', background: 'rgba(255,255,255,0.02)',
                border: '1px solid var(--panel-border)', borderRadius: '12px'
            }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                    <div style={{
                        width: '10px', height: '10px', borderRadius: '50%',
                        background: is_server_connected ? 'var(--success-color)' : 'var(--danger-color)',
                        boxShadow: is_server_connected ? '0 0 10px var(--success-color)' : '0 0 10px var(--danger-color)',
                        transition: 'all 0.3s'
                    }} />
                    <span style={{ fontSize: '14px', fontWeight: 600, color: 'var(--text-primary)' }}>
                        변환 헬퍼 서버 상태: {is_server_connected ? '온라인 (연결됨)' : '오프라인 (연결 해제)'}
                    </span>
                </div>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
                {/* 왼쪽: 동일 디렉토리에 PDF 생성 (절대 경로 모드) */}
                <div style={{
                    background: 'rgba(255,255,255,0.02)', padding: '24px',
                    borderRadius: '12px', border: '1px solid var(--panel-border)',
                    display: 'flex', flexDirection: 'column', gap: '16px',
                    opacity: is_server_connected ? 1 : 0.6,
                    pointerEvents: is_server_connected ? 'auto' : 'none',
                    transition: 'all 0.3s'
                }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <Presentation size={18} color="var(--accent-blue)" />
                        <h3 style={{ margin: 0, fontSize: '15.5px', color: 'var(--text-primary)', fontWeight: 700 }}>방법 A: 로컬 절대 경로 변환 (권장)</h3>
                    </div>
                    <div style={{ fontSize: '13px', color: 'var(--text-secondary)', lineHeight: '1.5' }}>
                        PPTX 파일의 전체 절대 경로를 입력하면, <strong>해당 파일이 위치한 디렉토리에 동일한 이름으로 무손실 PDF가 생성</strong>됩니다.
                    </div>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                        <input
                            type="text"
                            placeholder="예: C:\Users\Desktop\Project\산출물.pptx"
                            value={local_file_path}
                            onChange={(e) => set_local_file_path(e.target.value)}
                            style={{
                                width: '100%', padding: '12px 14px', borderRadius: '8px',
                                background: 'rgba(0,0,0,0.2)', border: '1px solid var(--panel-border)',
                                color: 'white', fontSize: '13.5px', outline: 'none'
                            }}
                        />
                        <button
                            onClick={handle_local_pdf_convert}
                            disabled={is_converting || !local_file_path.trim()}
                            style={{
                                padding: '12px', borderRadius: '8px', border: 'none',
                                background: is_converting ? 'rgba(255,255,255,0.05)' : 'linear-gradient(135deg, #0284c7, #3b82f6)',
                                color: 'white', fontWeight: 700, fontSize: '13.5px',
                                cursor: (is_converting || !local_file_path.trim()) ? 'not-allowed' : 'pointer',
                                display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '6px',
                                transition: 'all 0.2s'
                            }}
                        >
                            {is_converting ? <Loader2 size={16} className="animate-spin" /> : <Play size={16} />}
                            동일 디렉토리에 PDF 생성
                        </button>
                    </div>
                    <div style={{
                        padding: '10px 12px', background: 'rgba(59, 130, 246, 0.05)',
                        borderRadius: '6px', borderLeft: '3px solid var(--accent-blue)',
                        fontSize: '11.5px', color: 'var(--text-secondary)', lineHeight: '1.4'
                    }}>
                        💡 절대 경로는 파일 우클릭 후 <strong>'경로로 복사'</strong>를 눌러 간편하게 붙여넣을 수 있습니다.
                    </div>
                </div>

                {/* 오른쪽: 드래그 앤 드롭 업로드 & 변환 (다운로드 폴더 저장 모드) */}
                <div style={{
                    background: 'rgba(255,255,255,0.02)', padding: '24px',
                    borderRadius: '12px', border: '1px solid var(--panel-border)',
                    display: 'flex', flexDirection: 'column', gap: '16px',
                    opacity: is_server_connected ? 1 : 0.6,
                    pointerEvents: is_server_connected ? 'auto' : 'none',
                    transition: 'all 0.3s'
                }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <Upload size={18} color="var(--success-color)" />
                        <h3 style={{ margin: 0, fontSize: '15.5px', color: 'var(--text-primary)', fontWeight: 700 }}>방법 B: 드래그 앤 드롭 업로드 변환</h3>
                    </div>
                    <div style={{ fontSize: '13px', color: 'var(--text-secondary)', lineHeight: '1.5' }}>
                        파일을 직접 업로드하여 PDF로 변환한 후 브라우저의 <strong>다운로드 폴더에 자동 저장</strong>합니다.
                    </div>
                    <div
                        onDragOver={(e) => { e.preventDefault(); e.stopPropagation(); }}
                        onDragEnter={(e) => { e.preventDefault(); e.stopPropagation(); set_is_dragging_pdf(true); }}
                        onDragLeave={(e) => { e.preventDefault(); e.stopPropagation(); set_is_dragging_pdf(false); }}
                        onDrop={(e) => {
                            e.preventDefault(); e.stopPropagation(); set_is_dragging_pdf(false);
                            const file = e.dataTransfer.files[0];
                            if (file && file.name.toLowerCase().endsWith('.pptx')) {
                                set_pdf_convert_file(file);
                                handle_upload_pdf_convert(file);
                            } else {
                                setErrorMsg('파워포인트 파일(.pptx)만 지원합니다.');
                            }
                        }}
                        onClick={() => pdf_input_ref.current?.click()}
                        style={{
                            border: `2px dashed ${is_dragging_pdf ? 'var(--success-color)' : 'rgba(34, 197, 94, 0.3)'}`,
                            borderRadius: '10px',
                            padding: '24px 16px',
                            textAlign: 'center',
                            cursor: 'pointer',
                            background: is_dragging_pdf ? 'rgba(34, 197, 94, 0.05)' : 'rgba(0,0,0,0.15)',
                            transition: 'all 0.2s',
                            display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '8px'
                        }}
                    >
                        <input
                            type="file"
                            ref={pdf_input_ref}
                            onChange={(e) => {
                                const file = e.target.files[0];
                                if (file) {
                                    set_pdf_convert_file(file);
                                    handle_upload_pdf_convert(file);
                                }
                            }}
                            accept=".pptx"
                            style={{ display: 'none' }}
                        />
                        <Upload size={24} color={pdf_convert_file ? 'var(--success-color)' : 'var(--text-muted)'} />
                        <span style={{ fontSize: '13px', fontWeight: 600, color: 'var(--text-primary)', wordBreak: 'break-all' }}>
                            {pdf_convert_file ? pdf_convert_file.name : '이곳에 PPTX 파일 드래그 또는 클릭'}
                        </span>
                        {pdf_convert_file && (
                            <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>
                                {(pdf_convert_file.size / 1024 / 1024).toFixed(2)} MB - 변환 처리 중...
                            </span>
                        )}
                    </div>
                </div>
            </div>

            {/* OLE 렌더링 보장 안내 */}
            <div style={{
                padding: '16px 20px', background: 'rgba(255,255,255,0.01)',
                border: '1px solid var(--panel-border)', borderRadius: '12px',
                display: 'flex', flexDirection: 'column', gap: '6px'
            }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '6px', color: 'var(--text-primary)', fontWeight: 700, fontSize: '13.5px' }}>
                    <CheckCircle2 size={15} color="var(--success-color)" /> OLE 무손실 엔진 보장
                </div>
                <div style={{ fontSize: '12.5px', color: 'var(--text-secondary)', lineHeight: '1.6' }}>
                    이 변환 기술은 단순 텍스트 추출이 아닌, <strong>실제 로컬 파워포인트 OLE 렌더링 엔진(PowerPoint Object Model)을 사용하여 직접 백그라운드 렌더링을 지시</strong>합니다.
                    이에 따라 원본 PPTX 슬라이드에 적용된 <strong>특수 폰트, 표 스타일, 스마트아트, 이미지 투명도 및 배치 레이아웃이 100% 동일하게 완벽 보존된 PDF</strong>가 출력됩니다.
                </div>
            </div>
        </div>
    );
}
