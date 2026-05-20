import React, { useState } from 'react';
import { 
    Activity, 
    Award, 
    CheckCircle2, 
    Database, 
    FileWarning, 
    HelpCircle, 
    RefreshCw, 
    Sliders, 
    Sparkles, 
    Wrench, 
    XCircle,
    ArrowRight,
    ArrowDown,
    Zap,
    BookOpen
} from 'lucide-react';
import { ismpDaData } from '../utils/ismpDaMockData';

export default function IsmpDaDashboard() {
    // 로컬 상태로 Mock 데이터 복사하여 시뮬레이션
    const [complianceScore, setComplianceScore] = useState(ismpDaData.complianceScore);
    const [subjectAreas, setSubjectAreas] = useState(ismpDaData.subjectAreas);
    const [polyglotDb, setPolyglotDb] = useState(ismpDaData.polyglotDb);
    const [typoIssues, setTypoIssues] = useState(ismpDaData.typoIssues);
    
    // 슬라이드 역행 순서 교정 상태 (false: 역행상태, true: 정상 교정상태)
    const [isFlowRestored, setIsFlowRestored] = useState(false);
    
    // 폴리글랏 DB 아키텍처 정합성 패치 상태
    const [isDbAligned, setIsDbAligned] = useState(false);

    // 개별 NBSP / 괄호 오류 원클릭 패치 함수
    const handlePatchTypo = (id) => {
        setTypoIssues(prev => {
            const next = prev.map(issue => {
                if (issue.id === id && !issue.isPatched) {
                    // 패치 성공 시 준수 점수 상승 보정 (+2점)
                    setComplianceScore(score => Math.min(score + 2, 95));
                    return { ...issue, isPatched: true };
                }
                return issue;
            });
            return next;
        });
    };

    // 슬라이드 순서 논리적 복구 함수
    const handleRestoreFlow = () => {
        if (!isFlowRestored) {
            setIsFlowRestored(true);
            setComplianceScore(score => Math.min(score + 10, 95));
        }
    };

    // 폴리글랏 아키텍처와 RDB 싱크 복구 함수 (방안 A 적용)
    const handleAlignDbArchitecture = () => {
        if (!isDbAligned) {
            setIsDbAligned(true);
            setComplianceScore(score => Math.min(score + 12, 95));
            // PostgreSQL 외의 이기종 DB 매핑 이슈 해결 처리
            setPolyglotDb(prev => prev.map(db => ({
                ...db,
                isMismatch: false,
                physicalImplementation: db.dbName === 'PostgreSQL' 
                    ? "PostgreSQL 100% 물리 구현" 
                    : "방안 A 채택: PostgreSQL 내 JSONB 포맷 및 인덱스 최적화 기획으로 하이브리드 통합 흡수"
            })));
        }
    };

    // 전체 리셋
    const handleResetAll = () => {
        setComplianceScore(ismpDaData.complianceScore);
        setSubjectAreas(ismpDaData.subjectAreas);
        setPolyglotDb(ismpDaData.polyglotDb);
        setTypoIssues(ismpDaData.typos ? ismpDaData.typos : ismpDaData.typoIssues.map(t => ({...t, isPatched: false})));
        setIsFlowRestored(false);
        setIsDbAligned(false);
    };

    // 슬라이드 흐름 데이터
    const reverseFlow = [
        { name: "테이블/칼럼 명세서", desc: "물리 테이블 설계 (S009~S010)", isErroneous: true },
        { name: "데이터 아키텍처 정의서", desc: "개념 정의 및 지침 (S013~S014)", isErroneous: true },
        { name: "데이터 물리 ERD 구조", desc: "ERD 다이어그램 관계도 (S015)", isErroneous: true }
    ];

    const restoredFlow = [
        { name: "데이터 아키텍처 정의서", desc: "개념 정의 및 지침 수립 [1단계]" },
        { name: "데이터 물리 ERD 구조", desc: "논리/물리 ERD 다이어그램 작성 [2단계]" },
        { name: "테이블/칼럼 명세서", desc: "구조화된 테이블 물리 매핑 명세 [3단계]" }
    ];

    return (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '24px', paddingBottom: '40px', color: 'var(--text-primary)' }}>
            
            {/* 1. 상단 타이틀 배너 */}
            <div className="glass-panel animate-slide-up" style={{ padding: '24px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', background: 'linear-gradient(to right, rgba(168, 85, 247, 0.05), rgba(59, 130, 246, 0.05))', border: '1px solid var(--glass-border)' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
                    <div style={{ padding: '12px', background: 'linear-gradient(135deg, rgba(168, 85, 247, 0.2), rgba(59, 130, 246, 0.2))', borderRadius: '14px', border: '1px solid rgba(168, 85, 247, 0.3)' }}>
                        <Sliders size={24} color="var(--accent-purple)" />
                    </div>
                    <div>
                        <h2 style={{ margin: 0, fontSize: '22px', fontWeight: 800, letterSpacing: '-0.5px' }}>
                            ISMP 데이터 아키텍처(DA) 대화형 품질 검증 대시보드
                        </h2>
                        <p style={{ margin: '4px 0 0', fontSize: '13.5px', color: 'var(--text-secondary)' }}>
                            IV.3.2.1-6 개념 아키텍처 정의 및 IV.3.2.7 물리 주제영역 산출물의 심층적 정합성 갭을 실시간 교정하고 추적합니다.
                        </p>
                    </div>
                </div>
                <button 
                    onClick={handleResetAll} 
                    style={{
                        display: 'flex', alignItems: 'center', gap: '6px',
                        padding: '8px 14px', background: 'rgba(255, 255, 255, 0.05)',
                        border: '1px solid var(--glass-border)', borderRadius: '8px',
                        color: 'var(--text-secondary)', fontSize: '12px', fontWeight: 600,
                        cursor: 'pointer', transition: 'all 0.2s'
                    }}
                    onMouseEnter={e => e.currentTarget.style.background = 'rgba(255, 255, 255, 0.1)'}
                    onMouseLeave={e => e.currentTarget.style.background = 'rgba(255, 255, 255, 0.05)'}
                >
                    <RefreshCw size={14} /> 시뮬레이션 초기화
                </button>
            </div>

            {/* 2. 대형 요약 카드 및 점수 가젯 */}
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(300px, 1fr))', gap: '24px' }}>
                
                {/* 2-1. 컴플라이언스 준수 점수 원형 게이지 */}
                <div className="glass-panel animate-slide-up stagger-1" style={{ padding: '32px', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', position: 'relative', overflow: 'hidden' }}>
                    <div style={{ position: 'absolute', top: '-10%', left: '-10%', width: '40%', height: '40%', background: 'var(--success-color)', opacity: 0.05, filter: 'blur(40px)', borderRadius: '50%' }} />
                    <h3 style={{ margin: '0 0 20px', fontSize: '14px', color: 'var(--text-secondary)', fontWeight: 700, letterSpacing: '1px', textTransform: 'uppercase' }}>실시간 DA 품질 지수 (Compliance Score)</h3>
                    
                    <div style={{ 
                        position: 'relative', 
                        width: '180px', 
                        height: '180px', 
                        display: 'flex', 
                        alignItems: 'center', 
                        justifyContent: 'center', 
                        borderRadius: '50%', 
                        background: `conic-gradient(${complianceScore >= 85 ? 'var(--success-color)' : 'var(--warning-color)'} ${complianceScore}%, rgba(255,255,255,0.05) 0)`, 
                        boxShadow: `0 0 40px ${complianceScore >= 85 ? 'rgba(16, 185, 129, 0.15)' : 'rgba(245, 158, 11, 0.15)'}`,
                        transition: 'all 0.5s cubic-bezier(0.16, 1, 0.3, 1)'
                    }}>
                        <div style={{ position: 'absolute', width: '150px', height: '150px', background: 'var(--bg-secondary)', borderRadius: '50%', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', boxShadow: 'inset 0 0 20px rgba(0,0,0,0.4)' }}>
                            <span style={{ fontSize: '44px', fontWeight: 900, color: 'var(--text-primary)', letterSpacing: '-1.5px', transition: 'color 0.3s' }}>
                                {complianceScore}%
                            </span>
                            <span style={{ fontSize: '11px', color: complianceScore >= 85 ? 'var(--success-color)' : 'var(--warning-color)', fontWeight: 800, marginTop: '-2px', letterSpacing: '1px' }}>
                                {complianceScore >= 90 ? 'EXCELLENT (A)' : complianceScore >= 80 ? 'GOOD (B)' : 'REVISION NEEDED (C)'}
                            </span>
                        </div>
                    </div>
                </div>

                {/* 2-2. 품질 진단 리포트 요약 카드 */}
                <div className="glass-panel animate-slide-up stagger-1" style={{ padding: '24px', display: 'flex', flexDirection: 'column', justifyContent: 'space-between' }}>
                    <div>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '12px' }}>
                            <Zap size={18} color="var(--accent-purple)" />
                            <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 700 }}>AI 통합 품질 검증 리포트</h3>
                        </div>
                        <p style={{ margin: 0, fontSize: '14.5px', lineHeight: 1.6, color: 'var(--text-secondary)' }}>
                            개념 데이터 아키텍처 문서와 물리 스키마 산출물 간의 **12대 주제영역 불일치(정합성 25% 수준)**와 **폴리글랏 DB 격차**가 주요 결함으로 진단되었습니다. 아래의 인터랙티브 교정 패치를 클릭하여 산출물 간 연계 정합성을 획득하십시오.
                        </p>
                    </div>
                    
                    <div style={{ display: 'flex', gap: '10px', marginTop: '16px', flexWrap: 'wrap' }}>
                        <div style={{ background: 'rgba(239, 68, 68, 0.1)', border: '1px solid rgba(239, 68, 68, 0.2)', padding: '8px 12px', borderRadius: '10px', flex: 1, minWidth: '100px', textAlign: 'center' }}>
                            <span style={{ display: 'block', fontSize: '20px', fontWeight: 800, color: 'var(--danger-color)' }}>
                                {isFlowRestored && isDbAligned ? 0 : !isFlowRestored && !isDbAligned ? 2 : 1}
                            </span>
                            <span style={{ fontSize: '11px', color: 'var(--text-secondary)' }}>[상] 심각 오류</span>
                        </div>
                        <div style={{ background: 'rgba(245, 158, 11, 0.1)', border: '1px solid rgba(245, 158, 11, 0.2)', padding: '8px 12px', borderRadius: '10px', flex: 1, minWidth: '100px', textAlign: 'center' }}>
                            <span style={{ display: 'block', fontSize: '20px', fontWeight: 800, color: 'var(--warning-color)' }}>
                                {typoIssues.filter(i => !i.isPatched).length}
                            </span>
                            <span style={{ fontSize: '11px', color: 'var(--text-secondary)' }}>[중] 보완 권고</span>
                        </div>
                        <div style={{ background: 'rgba(59, 130, 246, 0.1)', border: '1px solid rgba(59, 130, 246, 0.2)', padding: '8px 12px', borderRadius: '10px', flex: 1, minWidth: '100px', textAlign: 'center' }}>
                            <span style={{ display: 'block', fontSize: '20px', fontWeight: 800, color: 'var(--accent-blue)' }}>3건</span>
                            <span style={{ fontSize: '11px', color: 'var(--text-secondary)' }}>[하] 표기 혼용</span>
                        </div>
                    </div>
                </div>

            </div>

            {/* 3. [🔴 상] 결함 1 - 슬라이드 전개 역행 오류 인터랙티브 교정 */}
            <section className="glass-panel animate-slide-up stagger-2" style={{ padding: '24px', borderLeft: `5px solid ${isFlowRestored ? 'var(--success-color)' : 'var(--danger-color)'}` }}>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '20px', flexWrap: 'wrap', gap: '12px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                        <span style={{ padding: '4px 10px', background: 'rgba(239, 68, 68, 0.15)', color: 'var(--danger-color)', fontSize: '12px', fontWeight: 700, borderRadius: '6px' }}>🔴 [상] 즉시 수정 필요</span>
                        <h3 style={{ margin: 0, fontSize: '17px', fontWeight: 700 }}>슬라이드 아키텍처 설계 흐름 역행 정비</h3>
                    </div>
                    {!isFlowRestored ? (
                        <button 
                            onClick={handleRestoreFlow}
                            style={{
                                padding: '8px 16px', background: 'linear-gradient(135deg, #ef4444, #f43f5e)',
                                border: 'none', borderRadius: '8px', color: 'white', fontWeight: 700,
                                fontSize: '13px', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '6px',
                                boxShadow: '0 4px 12px rgba(239,68,68,0.2)'
                            }}
                        >
                            <Zap size={14} /> 논리적 정상 흐름 복구 (+10점)
                        </button>
                    ) : (
                        <div style={{ display: 'flex', alignItems: 'center', gap: '6px', color: 'var(--success-color)', fontSize: '13.5px', fontWeight: 700 }}>
                            <CheckCircle2 size={16} /> 슬라이드 흐름 교정 완료
                        </div>
                    )}
                </div>

                <p style={{ margin: '0 0 20px', fontSize: '14px', color: 'var(--text-secondary)', lineHeight: 1.5 }}>
                    **진단 결함**: S009~S015 구간에서 슬라이드 내용 흐름이 **[물리 스키마 명세서 ➡️ 데이터 정의 ➡️ 물리 ERD]**로 전개되어, 하위 결과로부터 상위 개념을 도출하는 기형적인 구조가 발견되었습니다.
                </p>

                {/* 흐름 시각화 비교 카드 */}
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: '20px', alignItems: 'center' }}>
                    
                    {/* 역행 흐름 (AS-IS) */}
                    <div style={{ 
                        background: 'rgba(239,68,68,0.02)', border: '1px solid rgba(239,68,68,0.15)',
                        borderRadius: '12px', padding: '16px', opacity: isFlowRestored ? 0.4 : 1, transition: 'all 0.3s'
                    }}>
                        <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--danger-color)', display: 'block', marginBottom: '12px' }}>⚠️ 현재 역행 흐름 (AS-IS)</span>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                            {reverseFlow.map((flow, idx) => (
                                <div key={idx} style={{ background: 'rgba(0,0,0,0.2)', padding: '10px 14px', borderRadius: '8px', border: '1px solid rgba(255,255,255,0.03)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                                    <div>
                                        <span style={{ fontSize: '13.5px', fontWeight: 600, color: 'var(--text-primary)' }}>{idx + 1}. {flow.name}</span>
                                        <span style={{ display: 'block', fontSize: '11px', color: 'var(--text-muted)' }}>{flow.desc}</span>
                                    </div>
                                    <span style={{ fontSize: '10px', color: 'var(--danger-color)', fontWeight: 700 }}>오류</span>
                                </div>
                            ))}
                        </div>
                    </div>

                    <div style={{ textAlign: 'center', color: 'var(--text-muted)' }}>
                        <ArrowRight size={24} style={{ transform: 'rotate(0deg)' }} className="mobile-hide-text" />
                        <ArrowDown size={24} style={{ display: 'none' }} className="mobile-only-show" />
                    </div>

                    {/* 정상 복구 흐름 (TO-BE) */}
                    <div style={{ 
                        background: isFlowRestored ? 'rgba(16,185,129,0.05)' : 'rgba(255,255,255,0.01)', 
                        border: `1px solid ${isFlowRestored ? 'rgba(16,185,129,0.3)' : 'var(--glass-border)'}`,
                        borderRadius: '12px', padding: '16px', transition: 'all 0.3s'
                    }}>
                        <span style={{ fontSize: '12px', fontWeight: 700, color: isFlowRestored ? 'var(--success-color)' : 'var(--text-muted)', display: 'block', marginBottom: '12px' }}>
                            {isFlowRestored ? '✅ 복구된 정상 흐름 (TO-BE)' : '💡 교정 예정 정상 흐름 (TO-BE)'}
                        </span>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                            {restoredFlow.map((flow, idx) => (
                                <div key={idx} style={{ 
                                    background: isFlowRestored ? 'rgba(16,185,129,0.1)' : 'rgba(255,255,255,0.03)', 
                                    padding: '10px 14px', borderRadius: '8px', 
                                    border: `1px solid ${isFlowRestored ? 'rgba(16,185,129,0.2)' : 'rgba(255,255,255,0.03)'}`,
                                    display: 'flex', justifyContent: 'space-between', alignItems: 'center' 
                                }}>
                                    <div>
                                        <span style={{ fontSize: '13.5px', fontWeight: 600 }}>{idx + 1}. {flow.name}</span>
                                        <span style={{ display: 'block', fontSize: '11px', color: 'var(--text-muted)' }}>{flow.desc}</span>
                                    </div>
                                    {isFlowRestored && <span style={{ fontSize: '10px', color: 'var(--success-color)', fontWeight: 700 }}>정상</span>}
                                </div>
                            ))}
                        </div>
                    </div>

                </div>
            </section>

            {/* 4. [🔴 상] 결함 2 - 폴리글랏 DB 갭 분석 및 싱크 조정 */}
            <section className="glass-panel animate-slide-up stagger-2" style={{ padding: '24px', borderLeft: `5px solid ${isDbAligned ? 'var(--success-color)' : 'var(--danger-color)'}` }}>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '20px', flexWrap: 'wrap', gap: '12px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                        <span style={{ padding: '4px 10px', background: 'rgba(239, 68, 68, 0.15)', color: 'var(--danger-color)', fontSize: '12px', fontWeight: 700, borderRadius: '6px' }}>🔴 [상] 즉시 수정 필요</span>
                        <h3 style={{ margin: 0, fontSize: '17px', fontWeight: 700 }}>폴리글랏 DB 설계와 PostgreSQL 물리 모델 갭 교정</h3>
                    </div>
                    {!isDbAligned ? (
                        <button 
                            onClick={handleAlignDbArchitecture}
                            style={{
                                padding: '8px 16px', background: 'linear-gradient(135deg, #3b82f6, #6366f1)',
                                border: 'none', borderRadius: '8px', color: 'white', fontWeight: 700,
                                fontSize: '13px', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '6px',
                                boxShadow: '0 4px 12px rgba(59,130,246,0.2)'
                            }}
                        >
                            <Wrench size={14} /> 방안 A(Hybrid PostgreSQL) 일괄 적용 (+12점)
                        </button>
                    ) : (
                        <div style={{ display: 'flex', alignItems: 'center', gap: '6px', color: 'var(--success-color)', fontSize: '13.5px', fontWeight: 700 }}>
                            <CheckCircle2 size={16} /> 폴리글랏 아키텍처 정합성 매핑 완료
                        </div>
                    )}
                </div>

                <p style={{ margin: '0 0 20px', fontSize: '14px', color: 'var(--text-secondary)', lineHeight: 1.5 }}>
                    **진단 결함**: S049 및 S050에서 기획된 **5종 이기종 폴리글랏 DB(MongoDB, InfluxDB, Milvus, Elasticsearch, PostgreSQL)**와 달리, `3.2.7` 물리 설계서에는 **오직 PostgreSQL RDB 테이블**로만 물리 명세가 100% 단조롭게 채워져 있습니다. (아키텍처 불일치)
                </p>

                {/* DB 매핑 격차 챠트 */}
                <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '12.5px', color: 'var(--text-secondary)', fontWeight: 600, paddingBottom: '4px', borderBottom: '1px solid var(--panel-border)' }}>
                        <span style={{ width: '15%' }}>추천 DBMS</span>
                        <span style={{ width: '40%' }}>개념 기획 역할 (3.2.1-6)</span>
                        <span style={{ width: '45%' }}>물리 구현 및 대응 정합성 (3.2.7)</span>
                    </div>

                    {polyglotDb.map((db, idx) => (
                        <div key={idx} style={{ 
                            display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '10px 4px',
                            borderBottom: '1px solid rgba(255,255,255,0.03)', fontSize: '13.5px',
                            background: db.isMismatch ? 'rgba(239, 68, 68, 0.02)' : 'transparent',
                            borderRadius: '6px'
                        }}>
                            <span style={{ width: '15%', fontWeight: 700, color: db.isMismatch ? 'var(--danger-color)' : 'var(--success-color)', display: 'flex', alignItems: 'center', gap: '6px' }}>
                                <Database size={14} />
                                {db.dbName}
                            </span>
                            <span style={{ width: '40%', color: 'var(--text-secondary)' }}>{db.conceptualRole}</span>
                            <span style={{ 
                                width: '45%', 
                                color: db.isMismatch ? 'var(--danger-color)' : 'var(--success-color)',
                                fontWeight: db.isMismatch ? 500 : 600
                            }}>
                                {db.isMismatch ? "❌ " : "✅ "} {db.physicalImplementation}
                            </span>
                        </div>
                    ))}
                </div>
            </section>

            {/* 5. [🟡 중] 결함 3 - NBSP / 괄호 오류 원클릭 자동 패치 시뮬레이터 */}
            <section className="glass-panel animate-slide-up stagger-3" style={{ padding: '24px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '16px' }}>
                    <span style={{ padding: '4px 10px', background: 'rgba(245, 158, 11, 0.15)', color: 'var(--warning-color)', fontSize: '12px', fontWeight: 700, borderRadius: '6px' }}>🟡 [중] 보완 권고</span>
                    <h3 style={{ margin: 0, fontSize: '17px', fontWeight: 700 }}>괄호 불일치 및 비표준 공백(NBSP) 원클릭 교정기</h3>
                </div>
                
                <p style={{ margin: '0 0 20px', fontSize: '14px', color: 'var(--text-secondary)', lineHeight: 1.5 }}>
                    **진단 결함**: HTML/마크다운 스키마 테이블 복사 과정에서 유입된 `\xa0` 비표준 공백(8건)과, 본문 내 각주 표기 시 여는 괄호 누락(9건) 등 괄호 쌍 불일치가 발견되었습니다. 
                </p>

                <div style={{ overflowX: 'auto' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '13.5px' }}>
                        <thead>
                            <tr style={{ borderBottom: '1px solid var(--panel-border)', color: 'var(--text-secondary)' }}>
                                <th style={{ padding: '12px 8px', fontWeight: 600, width: '10%' }}>위치</th>
                                <th style={{ padding: '12px 8px', fontWeight: 600, width: '20%' }}>분류 유형</th>
                                <th style={{ padding: '12px 8px', fontWeight: 600, width: '30%' }}>산출물 원문</th>
                                <th style={{ padding: '12px 8px', fontWeight: 600, width: '30%' }}>수정 제안 (Corrected)</th>
                                <th style={{ padding: '12px 8px', fontWeight: 600, width: '10%', textAlign: 'center' }}>조치</th>
                            </tr>
                        </thead>
                        <tbody>
                            {typoIssues.map((issue) => (
                                <tr key={issue.id} style={{ 
                                    borderBottom: '1px solid rgba(255,255,255,0.05)',
                                    background: issue.isPatched ? 'rgba(16, 185, 129, 0.02)' : 'transparent',
                                    transition: 'all 0.2s'
                                }}>
                                    <td style={{ padding: '12px 8px', fontFamily: 'monospace', color: 'var(--accent-purple)', fontWeight: 600 }}>{issue.slide}</td>
                                    <td style={{ padding: '12px 8px', color: 'var(--text-secondary)' }}>{issue.errorType}</td>
                                    <td style={{ 
                                        padding: '12px 8px', 
                                        color: issue.isPatched ? 'var(--text-muted)' : 'var(--danger-color)',
                                        textDecoration: issue.isPatched ? 'line-through' : 'none'
                                    }}>
                                        {issue.originalText}
                                    </td>
                                    <td style={{ padding: '12px 8px', color: 'var(--success-color)', fontWeight: 600 }}>
                                        {issue.correctedText}
                                    </td>
                                    <td style={{ padding: '12px 8px', textAlign: 'center' }}>
                                        {!issue.isPatched ? (
                                            <button 
                                                onClick={() => handlePatchTypo(issue.id)}
                                                style={{
                                                    padding: '4px 10px', background: 'rgba(59, 130, 246, 0.1)',
                                                    border: '1px solid rgba(59, 130, 246, 0.3)', borderRadius: '6px',
                                                    color: 'var(--accent-blue)', fontSize: '11.5px', fontWeight: 700,
                                                    cursor: 'pointer', transition: 'all 0.2s'
                                                }}
                                                onMouseEnter={e => { e.currentTarget.style.background = 'rgba(59, 130, 246, 0.2)'; }}
                                                onMouseLeave={e => { e.currentTarget.style.background = 'rgba(59, 130, 246, 0.1)'; }}
                                            >
                                                원클릭 패치
                                            </button>
                                        ) : (
                                            <span style={{ fontSize: '12px', color: 'var(--success-color)', fontWeight: 700 }}>패치완료</span>
                                        )}
                                    </td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            </section>

            {/* 6. [🟢 하] 결함 4 - 용어 표기 혼용 통계 시각화 */}
            <section className="glass-panel animate-slide-up stagger-4" style={{ padding: '24px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '20px' }}>
                    <span style={{ padding: '4px 10px', background: 'rgba(59, 130, 246, 0.15)', color: 'var(--accent-blue)', fontSize: '12px', fontWeight: 700, borderRadius: '6px' }}>🟢 [하] 권고 사항</span>
                    <h3 style={{ margin: 0, fontSize: '17px', fontWeight: 700 }}>핵심 인프라 및 아키텍처 용어 혼용 통계</h3>
                </div>

                <p style={{ margin: '0 0 20px', fontSize: '14px', color: 'var(--text-secondary)', lineHeight: 1.5 }}>
                    **진단 결함**: 동일 개념에 대해 한글 발음, 영문 표기, 한글 순화어를 무차별적으로 섞어 써서 전문성 품격을 떨어뜨리고 있습니다. 대표 용어로 일원화 표기하십시오.
                </p>

                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: '24px' }}>
                    {ismpDaData.termMixUsage.map((group, idx) => {
                        const total = group.variants.reduce((acc, v) => acc + v.count, 0);
                        return (
                            <div key={idx} style={{ background: 'rgba(0,0,0,0.15)', padding: '16px', borderRadius: '12px', border: '1px solid var(--glass-border)' }}>
                                <span style={{ display: 'block', fontSize: '14px', fontWeight: 700, color: 'var(--text-primary)', marginBottom: '12px' }}>
                                    📊 {group.termGroup}
                                </span>
                                
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                                    {group.variants.map((v, vIdx) => {
                                        const percentage = (v.count / total) * 100;
                                        return (
                                            <div key={vIdx} style={{ fontSize: '13px' }}>
                                                <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '4px' }}>
                                                    <span style={{ color: 'var(--text-secondary)' }}>{v.term}</span>
                                                    <span style={{ fontWeight: 600 }}>{v.count}회 ({Math.round(percentage)}%)</span>
                                                </div>
                                                <div style={{ width: '100%', height: '6px', background: 'rgba(255,255,255,0.05)', borderRadius: '3px', overflow: 'hidden' }}>
                                                    <div style={{ 
                                                        width: `${percentage}%`, 
                                                        height: '100%', 
                                                        background: vIdx === 0 ? 'var(--accent-purple)' : vIdx === 1 ? 'var(--accent-blue)' : '#cbd5e1', 
                                                        borderRadius: '3px' 
                                                    }} />
                                                </div>
                                            </div>
                                        );
                                    })}
                                </div>
                            </div>
                        );
                    })}
                </div>
            </section>

        </div>
    );
}
