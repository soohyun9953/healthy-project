import React from 'react';
import {
    FileSearch, FileText, GitCompare, BarChart3,
    CheckCircle2, Loader2, Circle, AlertTriangle
} from 'lucide-react';

const STEPS = [
    { id: 1, label: '기준 문서 파싱', desc: '요구사항 ID 및 내용 추출 중', icon: FileSearch },
    { id: 2, label: '산출물 텍스트 분석', desc: '산출물 구조 및 키워드 색인 중', icon: FileText },
    { id: 3, label: '요구사항 ↔ 산출물 매핑', desc: '키워드 매칭 및 충족률 산출 중', icon: GitCompare },
    { id: 4, label: '검증 결과 종합', desc: 'RTM, 누락 분석, 점수 산출 중', icon: BarChart3 },
];

/**
 * @param {{ currentStep: number, stepStatus: string, error: string|null }} props
 *   currentStep: 현재 진행 중인 단계 (1~4), 0=시작전, 5=완료
 *   stepStatus: 현재 단계의 상세 상태 메시지
 *   error: 에러 메시지 (있을 경우)
 */
export default function AnalysisProgress({ currentStep, stepStatus, error }) {
    const progress = Math.min(100, Math.round(((currentStep - 1) / STEPS.length) * 100));

    return (
        <div className="glass-panel animate-fade-in" style={{
            flex: 1, display: 'flex', flexDirection: 'column',
            padding: '40px 36px', justifyContent: 'center',
        }}>
            {/* 헤더 */}
            <div style={{ textAlign: 'center', marginBottom: '36px' }}>
                <div style={{
                    display: 'inline-flex', padding: '16px', borderRadius: '50%',
                    background: error ? 'rgba(239,68,68,0.1)' : 'rgba(59,130,246,0.1)',
                    marginBottom: '16px',
                }}>
                    {error
                        ? <AlertTriangle size={40} color="var(--danger-color)" />
                        : <Loader2 size={40} color="var(--accent-color)" style={{ animation: 'spin 1.2s linear infinite' }} />
                    }
                </div>
                <h2 style={{ margin: '0 0 8px', fontSize: '22px', color: 'var(--text-primary)' }}>
                    {error ? '분석 중 오류 발생' : 'AI 검증 진행 중'}
                </h2>
                <p style={{ margin: 0, fontSize: '14px', color: 'var(--text-secondary)' }}>
                    {error || stepStatus || '분석을 시작합니다...'}
                </p>
            </div>

            {/* 전체 진행률 바 */}
            <div style={{ marginBottom: '32px' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '8px' }}>
                    <span style={{ fontSize: '13px', color: 'var(--text-secondary)' }}>전체 진행률</span>
                    <span style={{ fontSize: '13px', fontWeight: 600, color: 'var(--accent-color)' }}>{progress}%</span>
                </div>
                <div style={{
                    height: '8px', background: 'rgba(255,255,255,0.08)',
                    borderRadius: '4px', overflow: 'hidden',
                }}>
                    <div style={{
                        height: '100%', width: `${progress}%`, borderRadius: '4px',
                        background: 'linear-gradient(90deg, var(--accent-color), var(--success-color))',
                        transition: 'width 0.6s ease',
                    }} />
                </div>
            </div>

            {/* 단계별 상태 */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: '0' }}>
                {STEPS.map((step, idx) => {
                    const StepIcon = step.icon;
                    const isComplete = currentStep > step.id;
                    const isCurrent = currentStep === step.id;
                    const isPending = currentStep < step.id;
                    const isErrorStep = error && isCurrent;
                    const isLast = idx === STEPS.length - 1;

                    let statusColor = 'var(--text-secondary)';
                    let bgColor = 'transparent';
                    if (isComplete) { statusColor = 'var(--success-color)'; bgColor = 'rgba(16,185,129,0.06)'; }
                    if (isCurrent && !error) { statusColor = 'var(--accent-color)'; bgColor = 'rgba(59,130,246,0.08)'; }
                    if (isErrorStep) { statusColor = 'var(--danger-color)'; bgColor = 'rgba(239,68,68,0.08)'; }

                    return (
                        <div key={step.id}>
                            <div style={{
                                display: 'flex', alignItems: 'center', gap: '14px',
                                padding: '14px 16px', borderRadius: '10px',
                                background: bgColor,
                                transition: 'all 0.3s ease',
                            }}>
                                {/* 상태 아이콘 */}
                                <div style={{
                                    width: '36px', height: '36px', borderRadius: '50%',
                                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                                    flexShrink: 0,
                                    background: isComplete ? 'rgba(16,185,129,0.15)'
                                        : isCurrent ? (error ? 'rgba(239,68,68,0.15)' : 'rgba(59,130,246,0.15)')
                                            : 'rgba(255,255,255,0.05)',
                                    border: `2px solid ${statusColor}`,
                                    transition: 'all 0.3s ease',
                                }}>
                                    {isComplete ? (
                                        <CheckCircle2 size={18} color="var(--success-color)" />
                                    ) : isCurrent && !error ? (
                                        <Loader2 size={18} color="var(--accent-color)" style={{ animation: 'spin 1s linear infinite' }} />
                                    ) : isErrorStep ? (
                                        <AlertTriangle size={18} color="var(--danger-color)" />
                                    ) : (
                                        <Circle size={18} color="var(--text-secondary)" opacity={0.4} />
                                    )}
                                </div>

                                {/* 단계 정보 */}
                                <div style={{ flex: 1, minWidth: 0 }}>
                                    <div style={{
                                        display: 'flex', alignItems: 'center', gap: '8px',
                                        marginBottom: '2px',
                                    }}>
                                        <span style={{
                                            fontSize: '13px', fontWeight: 600,
                                            color: isPending ? 'var(--text-secondary)' : 'var(--text-primary)',
                                            opacity: isPending ? 0.5 : 1,
                                        }}>
                                            {step.label}
                                        </span>
                                        {isComplete && (
                                            <span style={{
                                                fontSize: '11px', padding: '1px 6px', borderRadius: '3px',
                                                background: 'rgba(16,185,129,0.15)', color: 'var(--success-color)',
                                                fontWeight: 500,
                                            }}>완료</span>
                                        )}
                                        {isCurrent && !error && (
                                            <span style={{
                                                fontSize: '11px', padding: '1px 6px', borderRadius: '3px',
                                                background: 'rgba(59,130,246,0.15)', color: 'var(--accent-color)',
                                                fontWeight: 500,
                                            }}>진행 중</span>
                                        )}
                                        {isErrorStep && (
                                            <span style={{
                                                fontSize: '11px', padding: '1px 6px', borderRadius: '3px',
                                                background: 'rgba(239,68,68,0.15)', color: 'var(--danger-color)',
                                                fontWeight: 500,
                                            }}>오류</span>
                                        )}
                                    </div>
                                    <span style={{
                                        fontSize: '12px', color: 'var(--text-secondary)',
                                        opacity: isPending ? 0.4 : 0.8,
                                    }}>
                                        {isCurrent && stepStatus ? stepStatus : step.desc}
                                    </span>
                                </div>

                                {/* 오른쪽 단계 아이콘 */}
                                <StepIcon size={18}
                                    color={statusColor}
                                    style={{ flexShrink: 0, opacity: isPending ? 0.3 : 0.8 }}
                                />
                            </div>

                            {/* 연결선 */}
                            {!isLast && (
                                <div style={{
                                    margin: '0 0 0 33px',
                                    width: '2px', height: '12px',
                                    background: isComplete ? 'var(--success-color)' : 'rgba(255,255,255,0.1)',
                                    transition: 'background 0.3s ease',
                                }} />
                            )}
                        </div>
                    );
                })}
            </div>

            {/* 경과 시간 표시 */}
            <ElapsedTimer />

            <style>{`
        @keyframes spin {
          from { transform: rotate(0deg); }
          to { transform: rotate(360deg); }
        }
      `}</style>
        </div>
    );
}

/** 경과 시간 표시 컴포넌트 */
function ElapsedTimer() {
    const [elapsed, setElapsed] = React.useState(0);

    React.useEffect(() => {
        const timer = setInterval(() => setElapsed(e => e + 1), 1000);
        return () => clearInterval(timer);
    }, []);

    const min = Math.floor(elapsed / 60);
    const sec = elapsed % 60;

    return (
        <div style={{
            textAlign: 'center', marginTop: '24px',
            fontSize: '13px', color: 'var(--text-secondary)',
        }}>
            경과 시간: <span style={{ fontFamily: 'monospace', fontWeight: 600, color: 'var(--text-primary)' }}>
                {min > 0 ? `${min}분 ` : ''}{sec}초
            </span>
        </div>
    );
}
