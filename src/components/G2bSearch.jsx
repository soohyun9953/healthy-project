import React, { useState, useEffect, useMemo } from 'react';
import { 
    Search, Building2, Calendar, DollarSign, ExternalLink, 
    FileSpreadsheet, Filter, RefreshCw, Info, Tag, CheckCircle, Clock, FileText, Sparkles, Key, Globe, Loader2, AlertTriangle
} from 'lucide-react';
import * as XLSX from 'xlsx';

export default function G2bSearch() {
    const [searchTitle, setSearchTitle] = useState('ISP');
    const [searchAgency, setSearchAgency] = useState('');
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
    const [apiSuccessCount, setApiSuccessCount] = useState(null);

    // 추천 검색어 태그
    const recommendedTitles = ['AI', '정보시스템', 'ISP', '공공보건', '빅데이터', '유지관리', '클라우드', 'ISMP'];
    const recommendedAgencies = ['국립중앙의료원', '보건복지부', '한국지능정보사회진흥원', '서울대학교병원', '국민건강보험공단', '질병관리청'];

    // API Key 저장
    const handleSaveApiKey = (key) => {
        setDataGovApiKey(key.trim());
        try {
            localStorage.setItem('data_gov_api_key', key.trim());
        } catch (e) {}
    };

    // YYYYMMDD0000 날짜 포맷 헬퍼 (최근 6개월)
    const getFormattedDates = () => {
        const now = new Date();
        const endYear = now.getFullYear();
        const endMonth = String(now.getMonth() + 1).padStart(2, '0');
        const endDate = String(now.getDate()).padStart(2, '0');
        const inqryEndDt = `${endYear}${endMonth}${endDate}2359`;

        // 6개월 전
        const past = new Date();
        past.setMonth(past.getMonth() - 6);
        const bgnYear = past.getFullYear();
        const bgnMonth = String(past.getMonth() + 1).padStart(2, '0');
        const bgnDate = String(past.getDate()).padStart(2, '0');
        const inqryBgnDt = `${bgnYear}${bgnMonth}${bgnDate}0000`;

        return { inqryBgnDt, inqryEndDt };
    };

    // 🌐 공공데이터포털 조달청 사전규격 OpenAPI 실시간 호출 (HrcspSgntrPrcureDetailInfoService02)
    const fetchG2bRealApiData = async () => {
        if (!dataGovApiKey.trim()) {
            setApiError('공공데이터포털(data.go.kr) API 키를 먼저 입력하고 [API 저장 & 조회]를 눌러주세요.');
            setShowApiKeySetting(true);
            return;
        }

        setIsLoading(true);
        setApiError(null);
        setApiSuccessCount(null);

        try {
            const rawKey = dataGovApiKey.trim();
            // 이미 인코딩되어 있는지 검사
            const serviceKeyParam = rawKey.includes('%') ? rawKey : encodeURIComponent(rawKey);
            const { inqryBgnDt, inqryEndDt } = getFormattedDates();

            let queryUrl = `https://apis.data.go.kr/1230000/HrcspSgntrPrcureDetailInfoService02/getHrcspSgntrPrcureDetailInfoList02?serviceKey=${serviceKeyParam}&type=json&numOfRows=100&pageNo=1&inqryBgnDt=${inqryBgnDt}&inqryEndDt=${inqryEndDt}`;

            if (searchTitle.trim()) {
                queryUrl += `&prcurRqstPrdNm=${encodeURIComponent(searchTitle.trim())}`;
            }
            if (searchAgency.trim()) {
                queryUrl += `&rlDmdOrganNm=${encodeURIComponent(searchAgency.trim())}`;
            }

            console.log('나라장터 OpenAPI 호출:', queryUrl);

            // 직접 fetch 시도 (CORS 우회 대응 포함)
            let response;
            try {
                response = await fetch(queryUrl);
            } catch (corsErr) {
                // CORS 대치 프록시 fallback
                const proxyUrl = `https://api.allorigins.win/raw?url=${encodeURIComponent(queryUrl)}`;
                response = await fetch(proxyUrl);
            }

            if (response && response.ok) {
                const data = await response.json();
                
                // 공공데이터포털 응답 구조 파싱
                const bodyObj = data?.response?.body;
                const headerObj = data?.response?.header;

                if (headerObj && headerObj.resultCode !== '00') {
                    setApiError(`API 오류: [${headerObj.resultCode}] ${headerObj.resultMsg}`);
                    setIsLoading(false);
                    return;
                }

                let rawItems = bodyObj?.items;
                // items가 객체인 경우 배열화
                if (rawItems && !Array.isArray(rawItems) && typeof rawItems === 'object') {
                    rawItems = rawItems.item ? (Array.isArray(rawItems.item) ? rawItems.item : [rawItems.item]) : [];
                }

                if (!Array.isArray(rawItems)) rawItems = [];

                const formattedList = rawItems.map((it, idx) => {
                    const budgetNum = parseInt(it.assignBdgtAmt || it.presmPrc || it.bdgtAmt || 0, 10);
                    const regDtStr = it.rgstDt ? `${it.rgstDt.substring(0,4)}-${it.rgstDt.substring(4,6)}-${it.rgstDt.substring(6,8)}` : '-';
                    const opngDtStr = it.opngDt ? `${it.opngDt.substring(0,4)}-${it.opngDt.substring(4,6)}-${it.opngDt.substring(6,8)}` : '-';
                    const preNo = it.bfStndRqstNo || it.preStndNo || '';

                    return {
                        id: preNo || `G2B-SPEC-${idx}`,
                        type: '사전규격',
                        title: it.prcurRqstPrdNm || it.orderNm || it.bidNtceNm || '사업명 미표시',
                        agency: it.rlDmdOrganNm || it.orderOrganNm || '수요기관 미표시',
                        category: it.refNo || it.taskClNm || '정보화사업',
                        budget: budgetNum,
                        regDate: regDtStr,
                        dueDate: opngDtStr,
                        status: '의견수렴중',
                        orderTime: it.rgstDt ? `${it.rgstDt.substring(0,4)}년 ${it.rgstDt.substring(4,6)}월` : '2026년',
                        detailUrl: preNo 
                            ? `https://www.g2b.go.kr:8081/ep/preparation/precom/preComDtl.do?preStndNo=${preNo}`
                            : `https://www.g2b.go.kr/search/search.do?kwd=${encodeURIComponent(it.prcurRqstPrdNm || '')}`,
                        description: `[조달청 사전규격 승인 데이터] 사전규격번호: ${preNo || '-'} / 담당자: ${it.ofcrNm || '조달청 담당부서'}`
                    };
                });

                setRealApiData(formattedList);
                setApiSuccessCount(formattedList.length);
            } else {
                setApiError('조달청 API 서버로부터 응답을 받지 못했습니다. 서비스키 승인 상태(승인완료 여부)를 확인해 주세요.');
            }
        } catch (err) {
            console.error('나라장터 OpenAPI 연동 오류:', err);
            setApiError(`API 통신 실패: ${err.message || '인증키 확인 및 브라우저 통신 상태 확인'}`);
        } finally {
            setIsLoading(false);
        }
    };

    // 초기 API키가 존재할 시 최초 1회 자동으로 실제 조달청 DB 조회
    useEffect(() => {
        if (dataGovApiKey) {
            fetchG2bRealApiData();
        }
    }, []);

    // 🔗 g2b.go.kr 나라장터 공식 사이트 실시간 검색 URL 동적 생성
    const getG2bOfficialSearchUrl = (tabType = 'integrated') => {
        const titleQuery = encodeURIComponent(searchTitle.trim());
        const agencyQuery = encodeURIComponent(searchAgency.trim());

        if (tabType === 'preCom') {
            return `https://www.g2b.go.kr:8081/ep/preparation/precom/preComList.do?taskClCd=5&searchDtType=1&searchType=1&supplierNm=${agencyQuery}&prdNm=${titleQuery}`;
        } else if (tabType === 'orderPlan') {
            return `https://www.g2b.go.kr:8081/ep/preparation/order/orderPlanList.do?taskClCd=5&orderOrganNm=${agencyQuery}&orderNm=${titleQuery}`;
        } else {
            return `https://www.g2b.go.kr/search/search.do?category=total&kwd=${titleQuery || agencyQuery}`;
        }
    };

    // 조건 초기화
    const handleReset = () => {
        setSearchTitle('');
        setSearchAgency('');
        setSelectedType('ALL');
        setMinBudget('0');
    };

    // 엑셀 다운로드
    const handleExportExcel = () => {
        if (realApiData.length === 0) {
            alert('다운로드할 실제 나라장터 사전규격 데이터가 없습니다. 상단 [사전규격 실시간 API 조회 실행] 버튼을 눌러주세요.');
            return;
        }

        const excelRows = realApiData.map((item, idx) => ({
            '번호': idx + 1,
            '구분': item.type,
            '사전규격 사업명': item.title,
            '수요기관': item.agency,
            '분류/참조번호': item.category,
            '추정예산 (원)': item.budget,
            '추정예산 (억원)': (item.budget / 100000000).toFixed(1) + ' 억',
            '등록일': item.regDate,
            '의견마감일': item.dueDate,
            '진행상태': item.status,
            '상세설명': item.description,
            '나라장터 바로가기 URL': item.detailUrl
        }));

        const worksheet = XLSX.utils.json_to_sheet(excelRows);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, '조달청_사전규격_실시간목록');

        const todayStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        XLSX.writeFile(workbook, `조달청_나라장터_사전규격_실시간조회_${todayStr}.xlsx`);
    };

    // 필터링된 데이터 (최소 예산 등 적용)
    const displayList = useMemo(() => {
        return realApiData.filter(item => {
            if (minBudget !== '0') {
                const minVal = parseFloat(minBudget) * 100000000;
                if (item.budget > 0 && item.budget < minVal) return false;
            }
            return true;
        });
    }, [realApiData, minBudget]);

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
                                조달청 나라장터 사전규격 실시간 API 연동
                                <span style={{ fontSize: '12px', padding: '3px 8px', borderRadius: '12px', background: 'rgba(16, 185, 129, 0.15)', color: 'var(--success-color)', fontWeight: 600 }}>
                                    HrcspSgntrPrcureDetailInfoService02 승인 연동
                                </span>
                            </h2>
                            <p style={{ margin: '4px 0 0', fontSize: '13.5px', color: 'var(--text-secondary)' }}>
                                공공데이터포털에서 승인받은 <strong>조달청 사전규격정보서비스 API</strong>를 이용하여 나라장터의 실제 최신 공고 데이터를 실시간으로 수집·표출합니다.
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
                            <Key size={15} /> 공공데이터포털 API키 {dataGovApiKey ? '등록완료' : '설정'}
                        </button>
                        <button
                            onClick={handleExportExcel}
                            disabled={realApiData.length === 0}
                            className="interactive"
                            style={{
                                display: 'flex', alignItems: 'center', gap: '6px',
                                padding: '10px 18px', borderRadius: '8px',
                                background: realApiData.length > 0 ? 'linear-gradient(135deg, #10b981, #059669)' : 'rgba(255,255,255,0.05)',
                                border: 'none',
                                color: 'white', fontSize: '13.5px', fontWeight: 700, cursor: realApiData.length > 0 ? 'pointer' : 'not-allowed',
                                opacity: realApiData.length > 0 ? 1 : 0.5,
                                boxShadow: realApiData.length > 0 ? '0 4px 14px rgba(16, 185, 129, 0.25)' : 'none'
                            }}
                        >
                            <FileSpreadsheet size={16} /> 엑셀 다운로드 (.xlsx)
                        </button>
                    </div>
                </div>

                {/* API 키 설정 창 */}
                {showApiKeySetting && (
                    <div className="animate-fade-in" style={{
                        padding: '18px', background: 'rgba(0,0,0,0.3)',
                        border: '1px solid rgba(59, 130, 246, 0.3)', borderRadius: '12px',
                        display: 'flex', flexDirection: 'column', gap: '12px'
                    }}>
                        <div style={{ fontSize: '14px', fontWeight: 700, color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '6px' }}>
                            <Key size={16} color="var(--accent-blue)" /> 공공데이터포털 조달청_나라장터 사전규격정보서비스 API 인증키 등록
                        </div>
                        <div style={{ fontSize: '12.5px', color: 'var(--text-secondary)', lineHeight: '1.6' }}>
                            승인받으신 <strong>[조달청_나라장터 사전규격정보서비스 (HrcspSgntrPrcureDetailInfoService02)]</strong>의 일반 인증키(Encoding/Decoding 키)를 아래에 등록해 주세요.<br />
                            등록 후 검색어를 입력하고 <strong>[사전규격 실시간 API 조회 실행]</strong> 버튼을 누르면 나라장터의 실제 실시간 데이터가 화면에 출력됩니다.
                        </div>
                        <div style={{ display: 'flex', gap: '10px' }}>
                            <input
                                type="text"
                                value={dataGovApiKey}
                                onChange={(e) => handleSaveApiKey(e.target.value)}
                                placeholder="공공데이터포털 승인 서비스키(Encoding / Decoding Key) 입력"
                                style={{
                                    flex: 1, padding: '10px 14px', borderRadius: '8px',
                                    border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.3)',
                                    color: 'var(--text-primary)', fontSize: '13px'
                                }}
                            />
                            <button
                                onClick={() => { fetchG2bRealApiData(); setShowApiKeySetting(false); }}
                                style={{
                                    padding: '10px 20px', borderRadius: '8px', background: 'var(--accent-blue)',
                                    color: 'white', border: 'none', fontWeight: 700, fontSize: '13.5px', cursor: 'pointer'
                                }}
                            >
                                API 키 저장 & 실시간 조회
                            </button>
                        </div>
                    </div>
                )}

                {/* 검색 필터 및 API 조회 버튼 */}
                <div style={{
                    padding: '20px', background: 'rgba(255,255,255,0.02)',
                    border: '1px solid var(--panel-border)', borderRadius: '14px',
                    display: 'flex', flexDirection: 'column', gap: '16px'
                }}>
                    <div style={{ fontSize: '14px', fontWeight: 700, color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <Filter size={16} color="var(--accent-blue)" /> 조달청 사전규격 실시간 검색 조건
                    </div>

                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(260px, 1fr))', gap: '16px' }}>
                        {/* 1. 사업명 (키워드) */}
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                            <label style={{ fontSize: '12.5px', color: 'var(--text-secondary)', fontWeight: 600 }}>
                                🔍 사업명 (특정 단어/키워드)
                            </label>
                            <div style={{ position: 'relative' }}>
                                <input
                                    type="text"
                                    placeholder="예: ISP, AI, 정보시스템, 빅데이터"
                                    value={searchTitle}
                                    onChange={(e) => setSearchTitle(e.target.value)}
                                    onKeyDown={(e) => e.key === 'Enter' && fetchG2bRealApiData()}
                                    style={{
                                        width: '100%', padding: '10px 12px 10px 36px', borderRadius: '8px',
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
                                🏛️ 특정 수요기관 (발주처명 - 선택사항)
                            </label>
                            <div style={{ position: 'relative' }}>
                                <input
                                    type="text"
                                    placeholder="예: 국립중앙의료원, 보건복지부"
                                    value={searchAgency}
                                    onChange={(e) => setSearchAgency(e.target.value)}
                                    onKeyDown={(e) => e.key === 'Enter' && fetchG2bRealApiData()}
                                    style={{
                                        width: '100%', padding: '10px 12px 10px 36px', borderRadius: '8px',
                                        border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)',
                                        color: 'var(--text-primary)', fontSize: '13.5px', boxSizing: 'border-box'
                                    }}
                                />
                                <Building2 size={16} color="var(--text-muted)" style={{ position: 'absolute', left: '12px', top: '50%', transform: 'translateY(-50%)' }} />
                            </div>
                        </div>

                        {/* 3. 최소 예산 규모 */}
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                            <label style={{ fontSize: '12.5px', color: 'var(--text-secondary)', fontWeight: 600 }}>
                                💰 최소 예산 규모 (추정가격)
                            </label>
                            <select
                                value={minBudget}
                                onChange={(e) => setMinBudget(e.target.value)}
                                style={{
                                    width: '100%', padding: '10px 12px', borderRadius: '8px',
                                    border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)',
                                    color: 'var(--text-primary)', fontSize: '13.5px', cursor: 'pointer'
                                }}
                            >
                                <option value="0">전체 예산 보기</option>
                                <option value="1">1억 원 이상</option>
                                <option value="5">5억 원 이상</option>
                                <option value="10">10억 원 이상</option>
                            </select>
                        </div>
                    </div>

                    {/* 추천 사업명 태그 */}
                    <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap', paddingTop: '8px', borderTop: '1px dashed var(--panel-border)' }}>
                        <span style={{ fontSize: '12px', color: 'var(--text-muted)', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '4px' }}>
                            <Tag size={13} /> 추천 사업명 태그:
                        </span>
                        {recommendedTitles.map(t => (
                            <button
                                key={t}
                                onClick={() => { setSearchTitle(t); }}
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

                    {/* 🚀 사전규격 실시간 API 호출 버튼 */}
                    <div style={{ display: 'flex', gap: '12px', marginTop: '8px' }}>
                        <button
                            id="btn-fetch-g2b-spec"
                            onClick={fetchG2bRealApiData}
                            disabled={isLoading}
                            style={{
                                flex: 1, padding: '16px', borderRadius: '10px', border: 'none',
                                background: 'linear-gradient(135deg, #a855f7, #7c3aed)',
                                color: 'white', fontWeight: 800, fontSize: '15.5px', cursor: isLoading ? 'not-allowed' : 'pointer',
                                display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '10px',
                                boxShadow: '0 4px 20px rgba(168, 85, 247, 0.35)', transition: 'all 0.2s'
                            }}
                        >
                            {isLoading ? (
                                <><Loader2 size={20} className="animate-spin" /> 조달청 나라장터 실제 DB 실시간 호출 중...</>
                            ) : (
                                <><Globe size={20} /> 📋 나라장터 실시간 사전규격 API 데이터 조회하기 (실제 공고)</>
                            )}
                        </button>
                    </div>
                </div>

                {/* API 피드백 안내 */}
                {apiError && (
                    <div className="animate-fade-in" style={{ padding: '14px', background: 'rgba(239, 68, 68, 0.1)', border: '1px solid rgba(239, 68, 68, 0.3)', borderRadius: '10px', color: '#f87171', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <AlertTriangle size={18} /> {apiError}
                    </div>
                )}
                {apiSuccessCount !== null && (
                    <div className="animate-fade-in" style={{ padding: '14px', background: 'rgba(16, 185, 129, 0.1)', border: '1px solid rgba(16, 185, 129, 0.3)', borderRadius: '10px', color: 'var(--success-color)', fontSize: '13.5px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <CheckCircle size={18} /> 성공적으로 조달청 나라장터 실시간 사전규격 데이터 <strong>{apiSuccessCount}건</strong>을 수집하여 표시하였습니다.
                    </div>
                )}
            </div>

            {/* 실제 데이터 표출 테이블 뷰 */}
            <div className="glass-panel animate-slide-up" style={{ padding: '24px', display: 'flex', flexDirection: 'column', gap: '16px' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <h3 style={{ margin: 0, fontSize: '17px', color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                        📋 나라장터 실시간 사전규격 수집 목록
                        <span style={{ fontSize: '13px', color: 'var(--text-muted)', fontWeight: 400 }}>({displayList.length}건 수집됨)</span>
                    </h3>
                </div>

                {displayList.length === 0 ? (
                    <div style={{ padding: '60px 20px', textAlign: 'center', background: 'rgba(0,0,0,0.15)', borderRadius: '14px', border: '1px dashed var(--panel-border)', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '14px' }}>
                        <Info size={36} color="#a855f7" />
                        <div>
                            <div style={{ fontSize: '16px', fontWeight: 700, color: 'var(--text-primary)' }}>
                                공공데이터포털 API키를 확인하고 [조회하기] 버튼을 눌러주세요.
                            </div>
                            <div style={{ fontSize: '13px', color: 'var(--text-secondary)', marginTop: '6px', maxWidth: '520px', lineHeight: 1.6 }}>
                                상단의 <strong>[📋 나라장터 실시간 사전규격 API 데이터 조회하기]</strong> 버튼을 누르시면, 승인받으신 조달청 OpenAPI로부터 수집된 실시간 실제 사전규격 조달 리스트가 이곳에 출력됩니다.
                            </div>
                        </div>
                    </div>
                ) : (
                    <div style={{ overflowX: 'auto', borderRadius: '12px', border: '1px solid var(--panel-border)' }}>
                        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13.5px', textAlign: 'left' }}>
                            <thead>
                                <tr style={{ background: 'rgba(255,255,255,0.04)', borderBottom: '1px solid var(--panel-border)', color: 'var(--text-secondary)' }}>
                                    <th style={{ padding: '14px 16px', width: '90px' }}>구분</th>
                                    <th style={{ padding: '14px 16px' }}>사전규격 사업명</th>
                                    <th style={{ padding: '14px 16px', width: '180px' }}>수요기관 (발주처)</th>
                                    <th style={{ padding: '14px 16px', width: '130px' }}>추정예산 / 배정액</th>
                                    <th style={{ padding: '14px 16px', width: '110px' }}>등록일</th>
                                    <th style={{ padding: '14px 16px', width: '110px' }}>의견마감일</th>
                                    <th style={{ padding: '14px 16px', width: '90px', textAlign: 'center' }}>g2b 상세</th>
                                </tr>
                            </thead>
                            <tbody>
                                {displayList.map((item, idx) => (
                                    <tr key={item.id || idx} style={{ borderBottom: '1px solid var(--panel-border)', background: idx % 2 === 1 ? 'rgba(255,255,255,0.015)' : 'transparent' }}>
                                        <td style={{ padding: '14px 16px' }}>
                                            <span style={{
                                                display: 'inline-block', padding: '4px 9px', borderRadius: '6px', fontSize: '12px', fontWeight: 700,
                                                background: 'rgba(168, 85, 247, 0.15)', color: '#c084fc', border: '1px solid rgba(168, 85, 247, 0.3)'
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
                                            {item.budget > 0 ? `${(item.budget / 100000000).toFixed(1)}억 원` : '미공개/비밀'}
                                        </td>
                                        <td style={{ padding: '14px 16px', color: 'var(--text-muted)', fontSize: '12.5px' }}>
                                            {item.regDate}
                                        </td>
                                        <td style={{ padding: '14px 16px', color: '#fb923c', fontSize: '12.5px', fontWeight: 600 }}>
                                            {item.dueDate}
                                        </td>
                                        <td style={{ padding: '14px 16px', textAlign: 'center' }}>
                                            <a
                                                href={item.detailUrl}
                                                target="_blank"
                                                rel="noreferrer"
                                                style={{
                                                    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                                                    width: '32px', height: '32px', borderRadius: '8px',
                                                    background: 'rgba(168, 85, 247, 0.15)', color: '#c084fc',
                                                    border: '1px solid rgba(168, 85, 247, 0.3)', transition: 'all 0.2s'
                                                }}
                                                title="나라장터 사전규격 상세페이지 바로가기"
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
