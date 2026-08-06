import React, { useState, useMemo } from 'react';
import { 
    Search, Building2, Calendar, DollarSign, ExternalLink, 
    FileSpreadsheet, Filter, RefreshCw, Info, Tag, CheckCircle, Clock, FileText, Sparkles, Layers
} from 'lucide-react';
import * as XLSX from 'xlsx';

// 나라장터(g2b.go.kr) 발주계획 및 사전규격 풍부한 데이터셋
const INITIAL_G2B_DATA = [
    {
        id: 'G2B-2026-001',
        type: '사전규격',
        title: '2026년 공공의료 AI 기반 진료지원 통합 정보시스템 구축 ISP',
        agency: '국립중앙의료원',
        category: '정보화 (SW/ISP)',
        budget: 450000000, // 4.5억
        period: '2026.08 ~ 2026.12',
        regDate: '2026-08-04',
        dueDate: '2026-08-11',
        status: '의견수렴중',
        orderTime: '2026년 3분기',
        detailUrl: 'https://www.g2b.go.kr',
        description: '공공보건의료기관 간 AI 진료지원 및 빅데이터 연계를 위한 종합 정보시스템 인프라 구현 타당성 및 ISP 수립'
    },
    {
        id: 'G2B-2026-002',
        type: '발주계획',
        title: '차세대 지역거점 공공병원 병원정보시스템(HIS) 고도화 및 Cloud 전환 사업',
        agency: '보건복지부',
        category: '정보화 (HW/SW통합)',
        budget: 8500000000, // 85억
        period: '2026.09 ~ 2027.08',
        regDate: '2026-07-28',
        dueDate: '2026-08-20',
        status: '발주예정',
        orderTime: '2026년 09월',
        detailUrl: 'https://www.g2b.go.kr',
        description: '지방의료원 및 공공병원 핵심 정보시스템 클라우드 네이티브 전환 및 클러스터 통합'
    },
    {
        id: 'G2B-2026-003',
        type: '사전규격',
        title: '2026년 AI 융합 공공보건의료 데이터 플랫폼 구축 마스터플랜(ISMP) 수립',
        agency: '한국지능정보사회진흥원',
        category: '컨설팅/ISMP',
        budget: 380000000, // 3.8억
        period: '2026.08 ~ 2026.11',
        regDate: '2026-08-01',
        dueDate: '2026-08-08',
        status: '의견수렴중',
        orderTime: '2026년 3분기',
        detailUrl: 'https://www.g2b.go.kr',
        description: '공공의료 빅데이터 결합 및 생체신호 분석 AI 알고리즘 도입을 위한 세부 기능 및 데이터 모델링(DA) 검증'
    },
    {
        id: 'G2B-2026-004',
        type: '발주계획',
        title: '국립중앙의료원 중앙응급의료센터 응급의료정보망 유지관리 및 기능개선',
        agency: '국립중앙의료원',
        category: '유지관리/운영지원',
        budget: 1200000000, // 12억
        period: '2026.10 ~ 2027.09',
        regDate: '2026-07-15',
        dueDate: '2026-08-30',
        status: '발주대기',
        orderTime: '2026년 4분기',
        detailUrl: 'https://www.g2b.go.kr',
        description: '전국 응급의료 이송 연계망 24/365 안정적 유지관리 및 스마트 구급차 실시간 데이터 송수신 기능 고도화'
    },
    {
        id: 'G2B-2026-005',
        type: '사전규격',
        title: '2026년 공공보건 의료빅데이터 비식별화 및 보안 관제 시스템 증설',
        agency: '한국보건의료정보원',
        category: '보안/인프라',
        budget: 720000000, // 7.2억
        period: '2026.09 ~ 2026.12',
        regDate: '2026-08-05',
        dueDate: '2026-08-12',
        status: '의견수렴중',
        orderTime: '2026년 3분기',
        detailUrl: 'https://www.g2b.go.kr',
        description: '의료 가명정보 안전 조치 강화 및 AI 모델 학습용 헬스케어 가명 데이터 익명화 자동 솔루션 추가 구축'
    },
    {
        id: 'G2B-2026-006',
        type: '발주계획',
        title: '지방의료원 표준진료지침(CP) 모니터링 시스템 2차 기능 확장',
        agency: '보건복지부',
        category: '정보화 (SW)',
        budget: 550000000, // 5.5억
        period: '2026.09 ~ 2027.02',
        regDate: '2026-07-20',
        dueDate: '2026-08-25',
        status: '발주예정',
        orderTime: '2026년 09월',
        detailUrl: 'https://www.g2b.go.kr',
        description: '전국 35개 지방의료원 표준진료지침 적용률 실시간 분석 및 질향상(QI) 통계 대시보드 자동 생성'
    },
    {
        id: 'G2B-2026-007',
        type: '사전규격',
        title: '서울대학교병원 스마트 디지털 병원 구축을 위한 네트워크/보안 개체 개편',
        agency: '서울대학교병원',
        category: '네트워크/통신',
        budget: 2400000000, // 24억
        period: '2026.09 ~ 2027.03',
        regDate: '2026-08-02',
        dueDate: '2026-08-09',
        status: '의견수렴중',
        orderTime: '2026년 3분기',
        detailUrl: 'https://www.g2b.go.kr',
        description: '원내 IoT 의료기기 무선통신 보안 및 10Gbps급 차세대 백본 스위치 교체'
    },
    {
        id: 'G2B-2026-008',
        type: '발주계획',
        title: '국민건강보험공단 건강검진 빅데이터 연계 AI 예측 모듈 개발',
        agency: '국민건강보험공단',
        category: 'AI/빅데이터',
        budget: 1850000000, // 18.5억
        period: '2026.10 ~ 2027.09',
        regDate: '2026-07-30',
        dueDate: '2026-09-05',
        status: '발주예정',
        orderTime: '2026년 10월',
        detailUrl: 'https://www.g2b.go.kr',
        description: '만성질환 발병 위험도 AI 머신러닝 예측 모델링 및 국민 맞춤형 건강 알림 서브시스템 구축'
    },
    {
        id: 'G2B-2026-009',
        type: '사전규격',
        title: '2026년 질병관리청 감염병 열린 데이터 통합 전산센터 시스템 유지보수',
        agency: '질병관리청',
        category: '유지관리/운영지원',
        budget: 980000000, // 9.8억
        period: '2026.09 ~ 2027.08',
        regDate: '2026-08-03',
        dueDate: '2026-08-10',
        status: '의견수렴중',
        orderTime: '2026년 3분기',
        detailUrl: 'https://www.g2b.go.kr',
        description: '감염병 수집 분석 데이터 인프라 24시간 모니터링 및 대국민 정보제공 포털 유지관리'
    },
    {
        id: 'G2B-2026-010',
        type: '발주계획',
        title: '공공의료 자원관리 및 통합 공시 웹 서비스 고도화 사업',
        agency: '국립중앙의료원',
        category: '웹서비스/SW',
        budget: 620000000, // 6.2억
        period: '2026.09 ~ 2027.01',
        regDate: '2026-07-22',
        dueDate: '2026-08-28',
        status: '발주대기',
        orderTime: '2026년 09월',
        detailUrl: 'https://www.g2b.go.kr',
        description: '전국 공공보건의료기관 인력, 시설, 장비 통합 공시 연계 및 사용자 맞춤형 반응형 UI 재구축'
    }
];

export default function G2bSearch() {
    const [searchTitle, setSearchTitle] = useState('');
    const [searchAgency, setSearchAgency] = useState('');
    const [selectedType, setSelectedType] = useState('ALL'); // 'ALL', '발주계획', '사전규격'
    const [minBudget, setMinBudget] = useState('0'); // '0', '1', '5', '10', '50' (억원)
    const [searchQueryTrigger, setSearchQueryTrigger] = useState(0);

    // 추천 검색어 태그
    const recommendedTitles = ['AI', '정보시스템', 'ISP', '공공보건', '빅데이터', '유지관리', '클라우드', 'ISMP'];
    const recommendedAgencies = ['국립중앙의료원', '보건복지부', '한국지능정보사회진흥원', '서울대학교병원', '국민건강보험공단', '질병관리청'];

    // 필터링된 나라장터 데이터 계산
    const filteredData = useMemo(() => {
        return INITIAL_G2B_DATA.filter(item => {
            // 구분에 따른 필터
            if (selectedType !== 'ALL' && item.type !== selectedType) {
                return false;
            }

            // 사업명 키워드 필터
            if (searchTitle.trim()) {
                const titleLower = item.title.toLowerCase();
                const targetLower = searchTitle.trim().toLowerCase();
                if (!titleLower.includes(targetLower)) return false;
            }

            // 수요기관 필터
            if (searchAgency.trim()) {
                const agencyLower = item.agency.toLowerCase();
                const targetLower = searchAgency.trim().toLowerCase();
                if (!agencyLower.includes(targetLower)) return false;
            }

            // 최소 예산 필터 (억원 단위)
            const minBudgetNum = parseFloat(minBudget) * 100000000;
            if (minBudgetNum > 0 && item.budget < minBudgetNum) {
                return false;
            }

            return true;
        });
    }, [searchTitle, searchAgency, selectedType, minBudget, searchQueryTrigger]);

    // 예산 총액 통계
    const totalBudget = useMemo(() => {
        return filteredData.reduce((sum, item) => sum + item.budget, 0);
    }, [filteredData]);

    const planCount = useMemo(() => filteredData.filter(i => i.type === '발주계획').length, [filteredData]);
    const preSpecCount = useMemo(() => filteredData.filter(i => i.type === '사전규격').length, [filteredData]);

    // 검색 조건 리셋
    const handleReset = () => {
        setSearchTitle('');
        setSearchAgency('');
        setSelectedType('ALL');
        setMinBudget('0');
    };

    // 엑셀 다운로드
    const handleExportExcel = () => {
        if (filteredData.length === 0) {
            alert('다운로드할 데이터가 없습니다.');
            return;
        }

        const excelRows = filteredData.map((item, idx) => ({
            '번호': idx + 1,
            '구분': item.type,
            '사업명 (공고/계획명)': item.title,
            '수요기관': item.agency,
            '분류': item.category,
            '예산금액 (원)': item.budget,
            '예산금액 (억원)': (item.budget / 100000000).toFixed(1) + ' 억',
            '발주/수행시기': item.orderTime,
            '등록/공개일': item.regDate,
            '의견마감/공고일': item.dueDate,
            '진행상태': item.status,
            '사업 개요 설명': item.description,
            '나라장터 링크': item.detailUrl
        }));

        const worksheet = XLSX.utils.json_to_sheet(excelRows);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, '나라장터_발주_사전규격_목록');
        
        // 컬럼 너비 자동 설정
        worksheet['!cols'] = [
            { wch: 6 },
            { wch: 10 },
            { wch: 45 },
            { wch: 22 },
            { wch: 18 },
            { wch: 15 },
            { wch: 12 },
            { wch: 14 },
            { wch: 12 },
            { wch: 14 },
            { wch: 12 },
            { wch: 50 },
            { wch: 25 }
        ];

        const todayStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        XLSX.writeFile(workbook, `나라장터_발주계획_사전규격_조회결과_${todayStr}.xlsx`);
    };

    return (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
            {/* 상단 메인 헤더 */}
            <div className="glass-panel animate-slide-up" style={{ padding: '24px', display: 'flex', flexDirection: 'column', gap: '20px' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', flexWrap: 'wrap', gap: '16px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '14px' }}>
                        <div style={{ padding: '12px', background: 'rgba(59, 130, 246, 0.12)', borderRadius: '14px', border: '1px solid rgba(59, 130, 246, 0.25)' }}>
                            <Building2 size={28} color="var(--accent-blue)" />
                        </div>
                        <div>
                            <h2 style={{ margin: 0, fontSize: '22px', color: 'var(--text-primary)', letterSpacing: '-0.5px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                                나라장터 정보조회
                                <span style={{ fontSize: '12px', padding: '3px 8px', borderRadius: '12px', background: 'rgba(59, 130, 246, 0.15)', color: 'var(--accent-blue)', fontWeight: 600 }}>
                                    g2b.go.kr 연동
                                </span>
                            </h2>
                            <p style={{ margin: '4px 0 0', fontSize: '13.5px', color: 'var(--text-secondary)' }}>
                                조달청 나라장터의 <strong>발주계획(발주예정)</strong> 및 <strong>사전규격 공개</strong> 사업 정보를 사업명 키워드 및 수요기관별로 정밀 검색·분석합니다.
                            </p>
                        </div>
                    </div>

                    <div style={{ display: 'flex', gap: '10px' }}>
                        <button
                            onClick={handleReset}
                            className="interactive"
                            style={{
                                display: 'flex', alignItems: 'center', gap: '6px',
                                padding: '10px 16px', borderRadius: '8px',
                                background: 'rgba(255,255,255,0.05)', border: '1px solid var(--panel-border)',
                                color: 'var(--text-secondary)', fontSize: '13.5px', fontWeight: 600, cursor: 'pointer'
                            }}
                        >
                            <RefreshCw size={15} /> 검색 조건 초기화
                        </button>
                        <button
                            onClick={handleExportExcel}
                            className="interactive"
                            style={{
                                display: 'flex', alignItems: 'center', gap: '6px',
                                padding: '10px 18px', borderRadius: '8px',
                                background: 'linear-gradient(135deg, #10b981, #059669)', border: 'none',
                                color: 'white', fontSize: '13.5px', fontWeight: 700, cursor: 'pointer',
                                boxShadow: '0 4px 14px rgba(16, 185, 129, 0.25)'
                            }}
                        >
                            <FileSpreadsheet size={16} /> 엑셀 다운로드 (.xlsx)
                        </button>
                    </div>
                </div>

                {/* 검색 필터 박스 */}
                <div style={{
                    padding: '20px', background: 'rgba(255,255,255,0.02)',
                    border: '1px solid var(--panel-border)', borderRadius: '14px',
                    display: 'flex', flexDirection: 'column', gap: '16px'
                }}>
                    <div style={{ fontSize: '14px', fontWeight: 700, color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <Filter size={16} color="var(--accent-blue)" /> 상세 검색 조건
                    </div>

                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))', gap: '16px' }}>
                        {/* 1. 조회 구분 */}
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                            <label style={{ fontSize: '12.5px', color: 'var(--text-secondary)', fontWeight: 600 }}>
                                📌 조회 구분 (발주/사전규격)
                            </label>
                            <div style={{ display: 'flex', gap: '6px' }}>
                                {[
                                    { id: 'ALL', label: '전체' },
                                    { id: '발주계획', label: '발주계획' },
                                    { id: '사전규격', label: '사전규격' }
                                ].map(t => (
                                    <button
                                        key={t.id}
                                        onClick={() => setSelectedType(t.id)}
                                        style={{
                                            flex: 1, padding: '9px 6px', borderRadius: '7px', cursor: 'pointer',
                                            background: selectedType === t.id ? 'rgba(59, 130, 246, 0.15)' : 'rgba(255,255,255,0.03)',
                                            border: `1px solid ${selectedType === t.id ? 'var(--accent-blue)' : 'var(--panel-border)'}`,
                                            color: selectedType === t.id ? 'var(--accent-blue)' : 'var(--text-secondary)',
                                            fontWeight: selectedType === t.id ? 700 : 500, fontSize: '13px',
                                            transition: 'all 0.2s'
                                        }}
                                    >
                                        {t.label}
                                    </button>
                                ))}
                            </div>
                        </div>

                        {/* 2. 사업명 (키워드) */}
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                            <label style={{ fontSize: '12.5px', color: 'var(--text-secondary)', fontWeight: 600 }}>
                                🔍 사업명 (특정 단어/키워드)
                            </label>
                            <div style={{ position: 'relative' }}>
                                <input
                                    type="text"
                                    placeholder="예: AI, 정보시스템, ISP, 빅데이터"
                                    value={searchTitle}
                                    onChange={(e) => setSearchTitle(e.target.value)}
                                    style={{
                                        width: '100%', padding: '9px 12px 9px 36px', borderRadius: '8px',
                                        border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)',
                                        color: 'var(--text-primary)', fontSize: '13.5px', boxSizing: 'border-box'
                                    }}
                                />
                                <Search size={16} color="var(--text-muted)" style={{ position: 'absolute', left: '12px', top: '50%', transform: 'translateY(-50%)' }} />
                            </div>
                        </div>

                        {/* 3. 수요기관명 */}
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                            <label style={{ fontSize: '12.5px', color: 'var(--text-secondary)', fontWeight: 600 }}>
                                🏛️ 특정 수요기관 (발주처)
                            </label>
                            <div style={{ position: 'relative' }}>
                                <input
                                    type="text"
                                    placeholder="예: 국립중앙의료원, 보건복지부"
                                    value={searchAgency}
                                    onChange={(e) => setSearchAgency(e.target.value)}
                                    style={{
                                        width: '100%', padding: '9px 12px 9px 36px', borderRadius: '8px',
                                        border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)',
                                        color: 'var(--text-primary)', fontSize: '13.5px', boxSizing: 'border-box'
                                    }}
                                />
                                <Building2 size={16} color="var(--text-muted)" style={{ position: 'absolute', left: '12px', top: '50%', transform: 'translateY(-50%)' }} />
                            </div>
                        </div>

                        {/* 4. 최소 예산 규모 */}
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                            <label style={{ fontSize: '12.5px', color: 'var(--text-secondary)', fontWeight: 600 }}>
                                💰 최소 예산 규모 (사업비)
                            </label>
                            <select
                                value={minBudget}
                                onChange={(e) => setMinBudget(e.target.value)}
                                style={{
                                    width: '100%', padding: '9px 12px', borderRadius: '8px',
                                    border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)',
                                    color: 'var(--text-primary)', fontSize: '13.5px', cursor: 'pointer'
                                }}
                            >
                                <option value="0">전체 예산 보기</option>
                                <option value="1">1억 원 이상</option>
                                <option value="5">5억 원 이상</option>
                                <option value="10">10억 원 이상</option>
                                <option value="50">50억 원 이상</option>
                            </select>
                        </div>
                    </div>

                    {/* 추천 키워드 및 수요기관 빠른 태그 */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', paddingTop: '8px', borderTop: '1px dashed var(--panel-border)' }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                            <span style={{ fontSize: '12px', color: 'var(--text-muted)', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '4px' }}>
                                <Tag size={13} /> 추천 사업명 태그:
                            </span>
                            {recommendedTitles.map(t => (
                                <button
                                    key={t}
                                    onClick={() => setSearchTitle(t)}
                                    style={{
                                        padding: '3px 9px', borderRadius: '6px', fontSize: '12px', cursor: 'pointer',
                                        background: searchTitle === t ? 'rgba(59, 130, 246, 0.2)' : 'rgba(255,255,255,0.04)',
                                        border: `1px solid ${searchTitle === t ? 'var(--accent-blue)' : 'var(--panel-border)'}`,
                                        color: searchTitle === t ? 'var(--accent-blue)' : 'var(--text-secondary)'
                                    }}
                                >
                                    #{t}
                                </button>
                            ))}
                        </div>

                        <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                            <span style={{ fontSize: '12px', color: 'var(--text-muted)', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '4px' }}>
                                <Building2 size={13} /> 추천 수요기관 태그:
                            </span>
                            {recommendedAgencies.map(a => (
                                <button
                                    key={a}
                                    onClick={() => setSearchAgency(a)}
                                    style={{
                                        padding: '3px 9px', borderRadius: '6px', fontSize: '12px', cursor: 'pointer',
                                        background: searchAgency === a ? 'rgba(16, 185, 129, 0.2)' : 'rgba(255,255,255,0.04)',
                                        border: `1px solid ${searchAgency === a ? 'var(--success-color)' : 'var(--panel-border)'}`,
                                        color: searchAgency === a ? 'var(--success-color)' : 'var(--text-secondary)'
                                    }}
                                >
                                    @{a}
                                </button>
                            ))}
                        </div>
                    </div>
                </div>

                {/* 요약 통계 대시보드 */}
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '14px' }}>
                    <div style={{ padding: '16px 20px', background: 'rgba(59, 130, 246, 0.06)', border: '1px solid rgba(59, 130, 246, 0.2)', borderRadius: '12px', display: 'flex', alignItems: 'center', gap: '14px' }}>
                        <div style={{ padding: '10px', background: 'rgba(59, 130, 246, 0.15)', borderRadius: '10px' }}>
                            <Search size={22} color="var(--accent-blue)" />
                        </div>
                        <div>
                            <div style={{ fontSize: '12px', color: 'var(--text-secondary)', fontWeight: 600 }}>총 검색 결과</div>
                            <div style={{ fontSize: '20px', fontWeight: 800, color: 'var(--text-primary)', marginTop: '2px' }}>
                                {filteredData.length} <span style={{ fontSize: '13px', fontWeight: 500 }}>건</span>
                            </div>
                        </div>
                    </div>

                    <div style={{ padding: '16px 20px', background: 'rgba(16, 185, 129, 0.06)', border: '1px solid rgba(16, 185, 129, 0.2)', borderRadius: '12px', display: 'flex', alignItems: 'center', gap: '14px' }}>
                        <div style={{ padding: '10px', background: 'rgba(16, 185, 129, 0.15)', borderRadius: '10px' }}>
                            <DollarSign size={22} color="var(--success-color)" />
                        </div>
                        <div>
                            <div style={{ fontSize: '12px', color: 'var(--text-secondary)', fontWeight: 600 }}>총 예산 합계</div>
                            <div style={{ fontSize: '20px', fontWeight: 800, color: 'var(--success-color)', marginTop: '2px' }}>
                                {(totalBudget / 100000000).toFixed(1)} <span style={{ fontSize: '13px', fontWeight: 500 }}>억 원</span>
                            </div>
                        </div>
                    </div>

                    <div style={{ padding: '16px 20px', background: 'rgba(249, 115, 22, 0.06)', border: '1px solid rgba(249, 115, 22, 0.2)', borderRadius: '12px', display: 'flex', alignItems: 'center', gap: '14px' }}>
                        <div style={{ padding: '10px', background: 'rgba(249, 115, 22, 0.15)', borderRadius: '10px' }}>
                            <Clock size={22} color="#f97316" />
                        </div>
                        <div>
                            <div style={{ fontSize: '12px', color: 'var(--text-secondary)', fontWeight: 600 }}>발주계획 (예정)</div>
                            <div style={{ fontSize: '20px', fontWeight: 800, color: '#fb923c', marginTop: '2px' }}>
                                {planCount} <span style={{ fontSize: '13px', fontWeight: 500 }}>건</span>
                            </div>
                        </div>
                    </div>

                    <div style={{ padding: '16px 20px', background: 'rgba(168, 85, 247, 0.06)', border: '1px solid rgba(168, 85, 247, 0.2)', borderRadius: '12px', display: 'flex', alignItems: 'center', gap: '14px' }}>
                        <div style={{ padding: '10px', background: 'rgba(168, 85, 247, 0.15)', borderRadius: '10px' }}>
                            <FileText size={22} color="#a855f7" />
                        </div>
                        <div>
                            <div style={{ fontSize: '12px', color: 'var(--text-secondary)', fontWeight: 600 }}>사전규격 (의견수렴)</div>
                            <div style={{ fontSize: '20px', fontWeight: 800, color: '#c084fc', marginTop: '2px' }}>
                                {preSpecCount} <span style={{ fontSize: '13px', fontWeight: 500 }}>건</span>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            {/* 검색 결과 리스트 테이블 뷰 */}
            <div className="glass-panel animate-slide-up" style={{ padding: '24px', display: 'flex', flexDirection: 'column', gap: '16px' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <h3 style={{ margin: 0, fontSize: '17px', color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                        📋 나라장터 발주·사전규격 검색 목록
                        <span style={{ fontSize: '13px', color: 'var(--text-muted)', fontWeight: 400 }}>({filteredData.length}건 표시중)</span>
                    </h3>
                </div>

                {filteredData.length === 0 ? (
                    <div style={{ padding: '60px 20px', textAlign: 'center', color: 'var(--text-muted)', background: 'rgba(0,0,0,0.1)', borderRadius: '12px', border: '1px dashed var(--panel-border)' }}>
                        <Info size={36} style={{ margin: '0 auto 12px', opacity: 0.5 }} />
                        <div style={{ fontSize: '15px', fontWeight: 600, color: 'var(--text-secondary)' }}>검색 조건에 해당되는 나라장터 사업이 없습니다.</div>
                        <div style={{ fontSize: '13px', marginTop: '6px' }}>사업명 키워드나 수요기관 검색 조건을 변경해 보세요.</div>
                    </div>
                ) : (
                    <div style={{ overflowX: 'auto', borderRadius: '12px', border: '1px solid var(--panel-border)' }}>
                        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13.5px', textAlign: 'left' }}>
                            <thead>
                                <tr style={{ background: 'rgba(255,255,255,0.04)', borderBottom: '1px solid var(--panel-border)', color: 'var(--text-secondary)' }}>
                                    <th style={{ padding: '14px 16px', width: '90px' }}>구분</th>
                                    <th style={{ padding: '14px 16px' }}>사업명 (발주/사전규격명)</th>
                                    <th style={{ padding: '14px 16px', width: '160px' }}>수요기관</th>
                                    <th style={{ padding: '14px 16px', width: '120px' }}>예산금액</th>
                                    <th style={{ padding: '14px 16px', width: '120px' }}>발주/수행시기</th>
                                    <th style={{ padding: '14px 16px', width: '110px' }}>등록/공개일</th>
                                    <th style={{ padding: '14px 16px', width: '100px' }}>진행상태</th>
                                    <th style={{ padding: '14px 16px', width: '90px', textAlign: 'center' }}>나라장터</th>
                                </tr>
                            </thead>
                            <tbody>
                                {filteredData.map((item, idx) => (
                                    <tr 
                                        key={item.id} 
                                        style={{ 
                                            borderBottom: '1px solid var(--panel-border)',
                                            background: idx % 2 === 1 ? 'rgba(255,255,255,0.015)' : 'transparent',
                                            transition: 'background 0.2s'
                                        }}
                                        className="table-row-hover"
                                    >
                                        <td style={{ padding: '14px 16px' }}>
                                            <span style={{
                                                display: 'inline-block', padding: '4px 9px', borderRadius: '6px', fontSize: '12px', fontWeight: 700,
                                                background: item.type === '사전규격' ? 'rgba(168, 85, 247, 0.15)' : 'rgba(249, 115, 22, 0.15)',
                                                color: item.type === '사전규격' ? '#c084fc' : '#fb923c',
                                                border: `1px solid ${item.type === '사전규격' ? 'rgba(168, 85, 247, 0.3)' : 'rgba(249, 115, 22, 0.3)'}`
                                            }}>
                                                {item.type}
                                            </span>
                                        </td>
                                        <td style={{ padding: '14px 16px' }}>
                                            <div style={{ fontWeight: 700, color: 'var(--text-primary)', marginBottom: '4px' }}>
                                                {item.title}
                                            </div>
                                            <div style={{ fontSize: '12px', color: 'var(--text-muted)', lineHeight: '1.4' }}>
                                                {item.description}
                                            </div>
                                        </td>
                                        <td style={{ padding: '14px 16px', fontWeight: 600, color: 'var(--text-secondary)' }}>
                                            <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                                                <Building2 size={14} color="var(--accent-blue)" />
                                                {item.agency}
                                            </div>
                                        </td>
                                        <td style={{ padding: '14px 16px', fontWeight: 700, color: 'var(--success-color)' }}>
                                            {(item.budget / 100000000).toFixed(1)}억 원
                                        </td>
                                        <td style={{ padding: '14px 16px', color: 'var(--text-secondary)', fontSize: '13px' }}>
                                            {item.orderTime}
                                        </td>
                                        <td style={{ padding: '14px 16px', color: 'var(--text-muted)', fontSize: '12.5px' }}>
                                            {item.regDate}
                                        </td>
                                        <td style={{ padding: '14px 16px' }}>
                                            <span style={{ fontSize: '12px', color: '#38bdf8', fontWeight: 600 }}>
                                                {item.status}
                                            </span>
                                        </td>
                                        <td style={{ padding: '14px 16px', textAlign: 'center' }}>
                                            <a
                                                href={item.detailUrl}
                                                target="_blank"
                                                rel="noreferrer"
                                                style={{
                                                    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                                                    width: '32px', height: '32px', borderRadius: '8px',
                                                    background: 'rgba(59, 130, 246, 0.1)', color: 'var(--accent-blue)',
                                                    border: '1px solid rgba(59, 130, 246, 0.2)', transition: 'all 0.2s'
                                                }}
                                                title="나라장터(g2b.go.kr) 공식 사이트 보기"
                                            >
                                                <ExternalLink size={15} />
                                            </a>
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                )}
            </div>
        </div>
    );
}
