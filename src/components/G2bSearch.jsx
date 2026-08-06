import React, { useState, useEffect, useMemo } from 'react';
import { 
    Search, Building2, Calendar, DollarSign, ExternalLink, 
    FileSpreadsheet, Filter, RefreshCw, Info, Tag, CheckCircle, Clock, FileText, Sparkles, Key, Globe, Radio
} from 'lucide-react';
import * as XLSX from 'xlsx';

export default function G2bSearch() {
    const [searchTitle, setSearchTitle] = useState('ISP');
    const [searchAgency, setSearchAgency] = useState('국립중앙의료원');
    const [selectedType, setSelectedType] = useState('ALL'); // 'ALL', '발주계획', '사전규격'
    const [minBudget, setMinBudget] = useState('0'); // 억원 단위
    const [dataGovApiKey, setDataGovApiKey] = useState(() => {
        try {
            return localStorage.getItem('data_gov_api_key') || '';
        } catch (e) {
            return '';
        }
    });
    const [showApiKeySetting, setShowApiKeySetting] = useState(false);
    const [isLoading, setIsLoading] = useState(false);
    const [realApiData, setRealApiData] = useState([]);
    const [apiError, setApiError] = useState(null);

    // 추천 검색어 태그
    const recommendedTitles = ['AI', '정보시스템', 'ISP', '공공보건', '빅데이터', '유지관리', '클라우드', 'ISMP'];
    const recommendedAgencies = ['국립중앙의료원', '보건복지부', '한국지능정보사회진흥원', '서울대학교병원', '국민건강보험공단', '질병관리청'];

    // API Key 저장
    const handleSaveApiKey = (key) => {
        setDataGovApiKey(key);
        try {
            localStorage.setItem('data_gov_api_key', key);
        } catch (e) {}
    };

    // 🌐 공공데이터포털 조달청 나라장터 OpenAPI 실시간 데이터 호출
    const fetchG2bRealApiData = async () => {
        if (!dataGovApiKey.trim()) return;

        setIsLoading(true);
        setApiError(null);
        try {
            // 조달청 나라장터 사전규격 및 발주계획 OpenAPI 엔드포인트
            const encodedKey = encodeURIComponent(dataGovApiKey.trim());
            const titleParam = encodeURIComponent(searchTitle.trim());
            const agencyParam = encodeURIComponent(searchAgency.trim());

            // 1) 사전규격 서비스 호출
            const preSpecUrl = `https://apis.data.go.kr/1230000/HrcspSgntrPrcureDetailInfoService02/getHrcspSgntrPrcureDetailInfoList02?serviceKey=${encodedKey}&type=json&numOfRows=30&pageNo=1${searchTitle ? `&prcurRqstPrdNm=${titleParam}` : ''}${searchAgency ? `&rlDmdOrganNm=${agencyParam}` : ''}`;
            
            const response = await fetch(preSpecUrl);
            if (response.ok) {
                const data = await response.json();
                const items = data?.response?.body?.items || [];
                
                const formattedList = items.map((it, idx) => ({
                    id: it.bfStndRqstNo || `G2B-REAL-${idx}`,
                    type: '사전규격',
                    title: it.prcurRqstPrdNm || it.orderNm || '사업명 정보 없음',
                    agency: it.rlDmdOrganNm || it.orderOrganNm || '수요기관 정보 없음',
                    category: it.refNo || '정보화',
                    budget: parseInt(it.assignBdgtAmt || it.presmPrc || 0, 10),
                    period: it.rgstDt ? `${it.rgstDt.substring(0,4)}.${it.rgstDt.substring(4,6)}` : '시정',
                    regDate: it.rgstDt ? `${it.rgstDt.substring(0,4)}-${it.rgstDt.substring(4,6)}-${it.rgstDt.substring(6,8)}` : '-',
                    dueDate: it.opngDt ? `${it.opngDt.substring(0,4)}-${it.opngDt.substring(4,6)}-${it.opngDt.substring(6,8)}` : '-',
                    status: '의견수렴중',
                    orderTime: '2026년',
                    detailUrl: `https://www.g2b.go.kr:8081/ep/preparation/precom/preComDtl.do?preStndNo=${it.bfStndRqstNo || ''}`,
                    description: `[나라장터 실제 사전규격] 등록번호: ${it.bfStndRqstNo || '-'} / 담당부서: ${it.ofcrNm || '-'}`
                }));

                setRealApiData(formattedList);
            } else {
                setApiError('API 호출에 실패했습니다. 공공데이터포털 키 승인 상태를 확인해주세요.');
            }
        } catch (err) {
            console.error('나라장터 API 호출 오류:', err);
            setApiError('공공데이터포털(data.go.kr) CORS 또는 키 오류가 발생했습니다. 아래 실시간 나라장터 검색을 클릭해 주세요.');
        } finally {
            setIsLoading(false);
        }
    };

    useEffect(() => {
        if (dataGovApiKey) {
            fetchG2bRealApiData();
        }
    }, [dataGovApiKey, searchTitle, searchAgency, selectedType]);

    // 🔗 g2b.go.kr 나라장터 공식 사이트 실시간 검색 URL 동적 생성
    const getG2bOfficialSearchUrl = (tabType = 'integrated') => {
        const titleQuery = encodeURIComponent(searchTitle.trim());
        const agencyQuery = encodeURIComponent(searchAgency.trim());

        if (tabType === 'preCom') {
            // 사전규격 공개검색
            return `https://www.g2b.go.kr:8081/ep/preparation/precom/preComList.do?taskClCd=5&searchDtType=1&searchType=1&supplierNm=${agencyQuery}&prdNm=${titleQuery}`;
        } else if (tabType === 'orderPlan') {
            // 발주계획 검색
            return `https://www.g2b.go.kr:8081/ep/preparation/order/orderPlanList.do?taskClCd=5&orderOrganNm=${agencyQuery}&orderNm=${titleQuery}`;
        } else {
            // 통합검색
            return `https://www.g2b.go.kr/search/search.do?category=total&kwd=${titleQuery || agencyQuery}`;
        }
    };

    // 실시간 검색 조건 리셋
    const handleReset = () => {
        setSearchTitle('');
        setSearchAgency('');
        setSelectedType('ALL');
        setMinBudget('0');
    };

    // 엑셀 다운로드
    const handleExportExcel = () => {
        const dataToExport = realApiData.length > 0 ? realApiData : [];
        if (dataToExport.length === 0) {
            alert('다운로드할 나라장터 조달 데이터가 없습니다. 아래 [나라장터 g2b.go.kr 실시간 바로가기]를 이용해 주세요.');
            return;
        }

        const excelRows = dataToExport.map((item, idx) => ({
            '번호': idx + 1,
            '구분': item.type,
            '사업명 (공고/계획명)': item.title,
            '수요기관': item.agency,
            '분류': item.category,
            '예산금액 (원)': item.budget,
            '예산금액 (억원)': (item.budget / 100000000).toFixed(1) + ' 억',
            '등록/공개일': item.regDate,
            '의견마감일': item.dueDate,
            '진행상태': item.status,
            '비고': item.description,
            '나라장터 링크': item.detailUrl
        }));

        const worksheet = XLSX.utils.json_to_sheet(excelRows);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, '나라장터_실제데이터_목록');

        const todayStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        XLSX.writeFile(workbook, `나라장터_실시간_조달데이터_${todayStr}.xlsx`);
    };

    return (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
            {/* 상단 헤더 및 API 연동 설정 */}
            <div className="glass-panel animate-slide-up" style={{ padding: '24px', display: 'flex', flexDirection: 'column', gap: '20px' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', flexWrap: 'wrap', gap: '16px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '14px' }}>
                        <div style={{ padding: '12px', background: 'rgba(59, 130, 246, 0.12)', borderRadius: '14px', border: '1px solid rgba(59, 130, 246, 0.25)' }}>
                            <Globe size={28} color="var(--accent-blue)" />
                        </div>
                        <div>
                            <h2 style={{ margin: 0, fontSize: '22px', color: 'var(--text-primary)', letterSpacing: '-0.5px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                                나라장터(g2b.go.kr) 실시간 데이터 연동
                                <span style={{ fontSize: '12px', padding: '3px 8px', borderRadius: '12px', background: 'rgba(16, 185, 129, 0.15)', color: 'var(--success-color)', fontWeight: 600 }}>
                                    실시간 사이트 동기화
                                </span>
                            </h2>
                            <p style={{ margin: '4px 0 0', fontSize: '13.5px', color: 'var(--text-secondary)' }}>
                                입력하신 사업명 키워드 및 수요기관으로 <strong>조달청 나라장터(g2b.go.kr)의 실제 발주계획 및 사전규격 공고</strong>를 실시간 연동하여 조회합니다.
                            </p>
                        </div>
                    </div>

                    <div style={{ display: 'flex', gap: '10px' }}>
                        <button
                            onClick={() => setShowApiKeySetting(!showApiKeySetting)}
                            className="interactive"
                            style={{
                                display: 'flex', alignItems: 'center', gap: '6px',
                                padding: '10px 16px', borderRadius: '8px',
                                background: dataGovApiKey ? 'rgba(16, 185, 129, 0.1)' : 'rgba(255,255,255,0.05)',
                                border: `1px solid ${dataGovApiKey ? 'rgba(16, 185, 129, 0.3)' : 'var(--panel-border)'}`,
                                color: dataGovApiKey ? 'var(--success-color)' : 'var(--text-secondary)',
                                fontSize: '13.5px', fontWeight: 600, cursor: 'pointer'
                            }}
                        >
                            <Key size={15} /> 공공데이터포털 API키 {dataGovApiKey ? '연동됨' : '설정'}
                        </button>
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
                    </div>
                </div>

                {/* API 키 설정 아코디언 */}
                {showApiKeySetting && (
                    <div className="animate-fade-in" style={{
                        padding: '16px', background: 'rgba(0,0,0,0.25)',
                        border: '1px solid rgba(59, 130, 246, 0.25)', borderRadius: '12px',
                        display: 'flex', flexDirection: 'column', gap: '10px'
                    }}>
                        <div style={{ fontSize: '13.5px', fontWeight: 700, color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '6px' }}>
                            <Key size={15} color="var(--accent-blue)" /> 공공데이터포털(data.go.kr) 나라장터 OpenAPI 서비스 인증키 설정
                        </div>
                        <div style={{ fontSize: '12.5px', color: 'var(--text-secondary)', lineHeight: '1.5' }}>
                            공공데이터포털에서 발급받은 <strong>[조달청_나라장터 발주계획 및 사전규격 서비스 API]</strong> 일반 인증키(Encoding/Decoding)를 입력하시면 실시간 조달청 DB를 직접 연동하여 조회합니다.
                        </div>
                        <div style={{ display: 'flex', gap: '8px' }}>
                            <input
                                type="text"
                                value={dataGovApiKey}
                                onChange={(e) => handleSaveApiKey(e.target.value)}
                                placeholder="공공데이터포털 OpenAPI ServiceKey 입력"
                                style={{
                                    flex: 1, padding: '9px 12px', borderRadius: '8px',
                                    border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.3)',
                                    color: 'var(--text-primary)', fontSize: '13px'
                                }}
                            />
                            <button
                                onClick={() => { fetchG2bRealApiData(); setShowApiKeySetting(false); }}
                                style={{
                                    padding: '9px 18px', borderRadius: '8px', background: 'var(--accent-blue)',
                                    color: 'white', border: 'none', fontWeight: 700, fontSize: '13px', cursor: 'pointer'
                                }}
                            >
                                API 저장 & 조회
                            </button>
                        </div>
                    </div>
                )}

                {/* 🔗 나라장터(g2b.go.kr) 공식 사이트 실시간 검색 바로가기 배너 */}
                <div style={{
                    padding: '20px',
                    background: 'linear-gradient(135deg, rgba(59, 130, 246, 0.12), rgba(99, 102, 241, 0.08))',
                    border: '1px solid rgba(59, 130, 246, 0.3)',
                    borderRadius: '14px',
                    display: 'flex', flexDirection: 'column', gap: '14px'
                }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px' }}>
                        <div>
                            <div style={{ fontSize: '16px', fontWeight: 800, color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                                <Globe size={20} color="#38bdf8" />
                                나라장터(g2b.go.kr) 공식 사이트 실시간 조회 실행
                            </div>
                            <div style={{ fontSize: '13px', color: 'var(--text-secondary)', marginTop: '4px' }}>
                                현재 설정된 조건 ➔ 사업명: <strong style={{ color: '#38bdf8' }}>"{searchTitle || '전체'}"</strong> / 수요기관: <strong style={{ color: '#38bdf8' }}>"{searchAgency || '전체'}"</strong>
                            </div>
                        </div>

                        <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
                            <a
                                href={getG2bOfficialSearchUrl('preCom')}
                                target="_blank"
                                rel="noreferrer"
                                style={{
                                    display: 'inline-flex', alignItems: 'center', gap: '6px',
                                    padding: '10px 16px', borderRadius: '9px',
                                    background: 'linear-gradient(135deg, #a855f7, #9333ea)', color: 'white',
                                    fontWeight: 700, fontSize: '13.5px', textDecoration: 'none',
                                    boxShadow: '0 4px 14px rgba(168, 85, 247, 0.3)'
                                }}
                            >
                                📋 g2b 사전규격 실시간 검색 <ExternalLink size={15} />
                            </a>
                            <a
                                href={getG2bOfficialSearchUrl('orderPlan')}
                                target="_blank"
                                rel="noreferrer"
                                style={{
                                    display: 'inline-flex', alignItems: 'center', gap: '6px',
                                    padding: '10px 16px', borderRadius: '9px',
                                    background: 'linear-gradient(135deg, #ea580c, #f97316)', color: 'white',
                                    fontWeight: 700, fontSize: '13.5px', textDecoration: 'none',
                                    boxShadow: '0 4px 14px rgba(249, 115, 22, 0.3)'
                                }}
                            >
                                📅 g2b 발주계획 실시간 검색 <ExternalLink size={15} />
                            </a>
                            <a
                                href={getG2bOfficialSearchUrl('integrated')}
                                target="_blank"
                                rel="noreferrer"
                                style={{
                                    display: 'inline-flex', alignItems: 'center', gap: '6px',
                                    padding: '10px 16px', borderRadius: '9px',
                                    background: 'var(--accent-blue)', color: 'white',
                                    fontWeight: 700, fontSize: '13.5px', textDecoration: 'none',
                                    boxShadow: '0 4px 14px rgba(59, 130, 246, 0.3)'
                                }}
                            >
                                🔍 g2b 나라장터 통합검색 <ExternalLink size={15} />
                            </a>
                        </div>
                    </div>
                </div>

                {/* 검색 필터 박스 */}
                <div style={{
                    padding: '20px', background: 'rgba(255,255,255,0.02)',
                    border: '1px solid var(--panel-border)', borderRadius: '14px',
                    display: 'flex', flexDirection: 'column', gap: '16px'
                }}>
                    <div style={{ fontSize: '14px', fontWeight: 700, color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <Filter size={16} color="var(--accent-blue)" /> 상세 검색 조건 변경
                    </div>

                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))', gap: '16px' }}>
                        {/* 1. 사업명 (키워드) */}
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

                        {/* 2. 수요기관명 */}
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
            </div>

            {/* 실시간 OpenAPI 결과 또는 안내 */}
            <div className="glass-panel animate-slide-up" style={{ padding: '24px', display: 'flex', flexDirection: 'column', gap: '16px' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <h3 style={{ margin: 0, fontSize: '17px', color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                        📋 실시간 나라장터 데이터 조회 결과
                    </h3>
                </div>

                {dataGovApiKey && realApiData.length > 0 ? (
                    <div style={{ overflowX: 'auto', borderRadius: '12px', border: '1px solid var(--panel-border)' }}>
                        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13.5px', textAlign: 'left' }}>
                            <thead>
                                <tr style={{ background: 'rgba(255,255,255,0.04)', borderBottom: '1px solid var(--panel-border)', color: 'var(--text-secondary)' }}>
                                    <th style={{ padding: '14px 16px', width: '90px' }}>구분</th>
                                    <th style={{ padding: '14px 16px' }}>사업명</th>
                                    <th style={{ padding: '14px 16px', width: '160px' }}>수요기관</th>
                                    <th style={{ padding: '14px 16px', width: '120px' }}>추정예산</th>
                                    <th style={{ padding: '14px 16px', width: '110px' }}>등록일</th>
                                    <th style={{ padding: '14px 16px', width: '90px', textAlign: 'center' }}>나라장터</th>
                                </tr>
                            </thead>
                            <tbody>
                                {realApiData.map((item, idx) => (
                                    <tr key={idx} style={{ borderBottom: '1px solid var(--panel-border)' }}>
                                        <td style={{ padding: '14px 16px' }}>{item.type}</td>
                                        <td style={{ padding: '14px 16px', fontWeight: 700 }}>{item.title}</td>
                                        <td style={{ padding: '14px 16px' }}>{item.agency}</td>
                                        <td style={{ padding: '14px 16px', color: 'var(--success-color)' }}>{(item.budget / 100000000).toFixed(1)}억 원</td>
                                        <td style={{ padding: '14px 16px' }}>{item.regDate}</td>
                                        <td style={{ padding: '14px 16px', textAlign: 'center' }}>
                                            <a href={item.detailUrl} target="_blank" rel="noreferrer" style={{ color: 'var(--accent-blue)' }}>
                                                <ExternalLink size={16} />
                                            </a>
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                ) : (
                    <div style={{ padding: '40px 20px', textAlign: 'center', background: 'rgba(0,0,0,0.15)', borderRadius: '14px', border: '1px dashed var(--panel-border)', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '14px' }}>
                        <Info size={32} color="#38bdf8" />
                        <div>
                            <div style={{ fontSize: '16px', fontWeight: 700, color: 'var(--text-primary)' }}>
                                나라장터(g2b.go.kr) 실시간 조달 데이터 검색 안내
                            </div>
                            <div style={{ fontSize: '13px', color: 'var(--text-secondary)', marginTop: '6px', maxWidth: '560px', lineHeight: 1.6 }}>
                                상단의 <strong>[g2b 사전규격 실시간 검색]</strong> 또는 <strong>[g2b 발주계획 실시간 검색]</strong> 버튼을 누르시면, 선택하신 사업명(<strong>{searchTitle || '전체'}</strong>) 및 수요기관(<strong>{searchAgency || '전체'}</strong>) 조건이 나라장터 공식 검색 시스템으로 연결되어 실시간 실제 데이터 결과를 즉시 확인하실 수 있습니다.
                            </div>
                        </div>

                        <div style={{ display: 'flex', gap: '10px', marginTop: '6px' }}>
                            <a
                                href={getG2bOfficialSearchUrl('preCom')}
                                target="_blank"
                                rel="noreferrer"
                                style={{
                                    display: 'inline-flex', alignItems: 'center', gap: '6px',
                                    padding: '10px 18px', borderRadius: '8px',
                                    background: '#a855f7', color: 'white', fontWeight: 700, fontSize: '13.5px', textDecoration: 'none'
                                }}
                            >
                                📋 나라장터 실시간 사전규격 조회하기 <ExternalLink size={15} />
                            </a>
                            <a
                                href={getG2bOfficialSearchUrl('orderPlan')}
                                target="_blank"
                                rel="noreferrer"
                                style={{
                                    display: 'inline-flex', alignItems: 'center', gap: '6px',
                                    padding: '10px 18px', borderRadius: '8px',
                                    background: '#f97316', color: 'white', fontWeight: 700, fontSize: '13.5px', textDecoration: 'none'
                                }}
                            >
                                📅 나라장터 실시간 발주계획 조회하기 <ExternalLink size={15} />
                            </a>
                        </div>
                    </div>
                )}
            </div>
        </div>
    );
}
