import React, { useState, useEffect, useMemo } from 'react';
import { 
    Search, Building2, Calendar, DollarSign, ExternalLink, 
    FileSpreadsheet, Filter, RefreshCw, Info, Tag, CheckCircle, Clock, FileText, Sparkles, Key, Globe, Loader2, AlertTriangle, ShieldCheck
} from 'lucide-react';
import * as XLSX from 'xlsx';

export default function G2bSearch() {
    const [searchTitle, setSearchTitle] = useState('ISP');
    const [searchAgency, setSearchAgency] = useState('');
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
        const cleanKey = key.trim();
        setDataGovApiKey(cleanKey);
        try {
            localStorage.setItem('data_gov_api_key', cleanKey);
        } catch (e) {}
    };

    // YYYYMMDD0000 날짜 포맷 헬퍼 (최근 6개월)
    const getFormattedDates = () => {
        const now = new Date();
        const endYear = now.getFullYear();
        const endMonth = String(now.getMonth() + 1).padStart(2, '0');
        const endDate = String(now.getDate()).padStart(2, '0');
        const inqryEndDt = `${endYear}${endMonth}${endDate}2359`;

        const past = new Date();
        past.setMonth(past.getMonth() - 5);
        const bgnYear = past.getFullYear();
        const bgnMonth = String(past.getMonth() + 1).padStart(2, '0');
        const bgnDate = String(past.getDate()).padStart(2, '0');
        const inqryBgnDt = `${bgnYear}${bgnMonth}${bgnDate}0000`;

        return { inqryBgnDt, inqryEndDt };
    };

    // 🌐 스마트 XML/JSON 파서 헬퍼
    const parseApiResponse = (textData) => {
        let items = [];
        let resultCode = '00';
        let resultMsg = 'NORMAL SERVICE.';

        try {
            // 1. JSON 시도
            const jsonData = JSON.parse(textData);
            const bodyObj = jsonData?.response?.body;
            const headerObj = jsonData?.response?.header;

            if (headerObj) {
                resultCode = headerObj.resultCode || '00';
                resultMsg = headerObj.resultMsg || '';
            }

            let rawItems = bodyObj?.items;
            if (rawItems && !Array.isArray(rawItems) && typeof rawItems === 'object') {
                rawItems = rawItems.item ? (Array.isArray(rawItems.item) ? rawItems.item : [rawItems.item]) : [];
            }
            if (Array.isArray(rawItems)) {
                items = rawItems;
            }
        } catch (jsonErr) {
            // 2. XML 시도
            try {
                const parser = new DOMParser();
                const xmlDoc = parser.parseFromString(textData, "text/xml");
                
                const codeNode = xmlDoc.querySelector("resultCode") || xmlDoc.querySelector("returnReasonCode");
                const msgNode = xmlDoc.querySelector("resultMsg") || xmlDoc.querySelector("returnAuthMsg");
                if (codeNode) resultCode = codeNode.textContent;
                if (msgNode) resultMsg = msgNode.textContent;

                const itemNodes = xmlDoc.querySelectorAll("item");
                itemNodes.forEach(node => {
                    const itemObj = {};
                    node.childNodes.forEach(child => {
                        if (child.nodeType === 1) {
                            itemObj[child.tagName] = child.textContent;
                        }
                    });
                    items.push(itemObj);
                });
            } catch (xmlErr) {
                console.error("파싱 오류:", xmlErr);
            }
        }

        return { items, resultCode, resultMsg };
    };

    // 🌐 공공데이터포털 조달청 사전규격 OpenAPI 실시간 호출 (Multi-Proxy Chain & Smart Encoding)
    const fetchG2bRealApiData = async () => {
        const rawKey = dataGovApiKey.trim();
        if (!rawKey) {
            setApiError('공공데이터포털(data.go.kr) API 키를 먼저 등록해 주세요.');
            setShowApiKeySetting(true);
            return;
        }

        setIsLoading(true);
        setApiError(null);
        setApiSuccessCount(null);

        try {
            const { inqryBgnDt, inqryEndDt } = getFormattedDates();

            // 인코딩 키 및 디코딩 키 2가지 버전 준비
            let encodedKey = rawKey;
            let decodedKey = rawKey;
            try {
                if (rawKey.includes('%')) {
                    decodedKey = decodeURIComponent(rawKey);
                } else {
                    encodedKey = encodeURIComponent(rawKey);
                }
            } catch (e) {
                console.warn('Key transcode fail:', e);
            }

            const keyVariants = [encodedKey, decodedKey, rawKey];
            const uniqueKeys = [...new Set(keyVariants)];

            let fetchSuccess = false;
            let lastErrorText = '';
            let parsedResult = null;

            for (const keyToUse of uniqueKeys) {
                let baseUrl = `https://apis.data.go.kr/1230000/HrcspSgntrPrcureDetailInfoService02/getHrcspSgntrPrcureDetailInfoList02?serviceKey=${keyToUse}&type=json&numOfRows=100&pageNo=1&inqryBgnDt=${inqryBgnDt}&inqryEndDt=${inqryEndDt}`;

                if (searchTitle.trim()) {
                    baseUrl += `&prcurRqstPrdNm=${encodeURIComponent(searchTitle.trim())}`;
                }
                if (searchAgency.trim()) {
                    baseUrl += `&rlDmdOrganNm=${encodeURIComponent(searchAgency.trim())}`;
                }

                // 프록시 체인 시도 목록
                const proxyAttempts = [
                    { name: '직접 호출', url: baseUrl },
                    { name: 'CORS Proxy 1', url: `https://corsproxy.io/?${encodeURIComponent(baseUrl)}` },
                    { name: 'CORS Proxy 2', url: `https://api.allorigins.win/raw?url=${encodeURIComponent(baseUrl)}` },
                    { name: 'CORS Proxy 3', url: `https://thingproxy.freeboard.io/fetch/${baseUrl}` }
                ];

                for (const attempt of proxyAttempts) {
                    try {
                        console.log(`[나라장터 API] ${attempt.name} (Key: ${keyToUse.substring(0, 10)}...) 시도 중...`);
                        const res = await fetch(attempt.url, { method: 'GET' });
                        if (res.ok) {
                            const text = await res.text();
                            parsedResult = parseApiResponse(text);
                            if (parsedResult.resultCode === '00' || parsedResult.items.length > 0) {
                                fetchSuccess = true;
                                break;
                            } else {
                                lastErrorText = `API 코드 [${parsedResult.resultCode}]: ${parsedResult.resultMsg}`;
                            }
                        }
                    } catch (e) {
                        console.warn(`${attempt.name} 실패:`, e);
                    }
                }

                if (fetchSuccess) break;
            }

            if (fetchSuccess && parsedResult) {
                const formattedList = parsedResult.items.map((it, idx) => {
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
                if (lastErrorText.includes('SERVICE_KEY_IS_NOT_REGISTERED')) {
                    setApiError('서비스키 미등록: 공공데이터포털(data.go.kr)에서 조달청 사전규격/입찰공고 서비스 활용신청 승인 상태(약 1~2시간 승인 지연 가능)를 확인해 주십시오.');
                } else if (lastErrorText) {
                    setApiError(`조달청 응답: ${lastErrorText}`);
                } else {
                    setApiError('브라우저 보안(CORS) 제한으로 조달청 API 직접 수신이 차단되었습니다. 하단의 [나라장터 공식 검색] 버튼을 누르시면 해당 검색어로 즉시 조회하실 수 있습니다.');
                }
            }
        } catch (err) {
            console.error('나라장터 OpenAPI 연동 오류:', err);
            setApiError(`API 시스템 오류: ${err.message || '인증키 확인 및 브라우저 통신 상태 확인'}`);
        } finally {
            setIsLoading(false);
        }
    };

    // 초기 API키가 존재할 시 최초 1회 자동으로 실시간 조달청 DB 조회
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
        setMinBudget('0');
    };

    // 엑셀 다운로드
    const handleExportExcel = () => {
        if (realApiData.length === 0) {
            alert('다운로드할 실제 나라장터 사전규격 데이터가 없습니다. 상단 [사전규격 실시간 API 조회 실행] 버튼을 누르고 데이터를 먼저 조회해 주세요.');
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
                                    HrcspSgntrPrcureDetailInfoService02 연동
                                </span>
                            </h2>
                            <p style={{ margin: '4px 0 0', fontSize: '13.5px', color: 'var(--text-secondary)' }}>
                                승인받으신 <strong>조달청 사전규격정보서비스 API</strong>를 통해 나라장터의 실제 사전규격 공고 데이터를 실시간으로 파싱하여 수집합니다.
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
                            <Key size={16} color="var(--accent-blue)" /> 공공데이터포털 조달청_나라장터 사전규격정보서비스 API 인증키 입력
                        </div>
                        <div style={{ fontSize: '12.5px', color: 'var(--text-secondary)', lineHeight: '1.6' }}>
                            승인받으신 <strong>[조달청_나라장터 사전규격정보서비스 (HrcspSgntrPrcureDetailInfoService02)]</strong>의 서비스키를 입력해 주세요.<br />
                            * 💡 <strong>Tip:</strong> 만약 [Encoding 키]로 실패하는 경우 공공데이터포털 마이페이지의 <strong>[Decoding 키]</strong>를 복사하여 입력하시면 100% 정상 작동합니다.
                        </div>
                        <div style={{ display: 'flex', gap: '10px' }}>
                            <input
                                type="text"
                                value={dataGovApiKey}
                                onChange={(e) => handleSaveApiKey(e.target.value)}
                                placeholder="공공데이터포털 승인 서비스키 (Decoding 키 또는 Encoding 키) 입력"
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
                                <><Loader2 size={20} className="animate-spin" /> 조달청 나라장터 DB 실시간 연동 수집 중...</>
                            ) : (
                                <><Globe size={20} /> 📋 나라장터 실시간 사전규격 API 데이터 조회하기 (실제 공고)</>
                            )}
                        </button>
                    </div>
                </div>

                {/* API 피드백 및 도움말 안내 */}
                {apiError && (
                    <div className="animate-fade-in" style={{ padding: '16px', background: 'rgba(239, 68, 68, 0.1)', border: '1px solid rgba(239, 68, 68, 0.3)', borderRadius: '12px', color: '#f87171', fontSize: '13px', display: 'flex', flexDirection: 'column', gap: '8px' }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', fontWeight: 700, fontSize: '13.5px' }}>
                            <AlertTriangle size={18} /> API 호출 안내
                        </div>
                        <div>{apiError}</div>
                        <div style={{ fontSize: '12px', color: 'var(--text-secondary)', marginTop: '4px', lineHeight: 1.5, background: 'rgba(0,0,0,0.2)', padding: '10px', borderRadius: '8px' }}>
                            💡 <strong>해결 팁:</strong><br />
                            1. 공공데이터포털(data.go.kr) 마이페이지에서 <strong>[Decoding 키]</strong>를 입력해 보세요. (Encoding 키 사용 시 이중 인코딩 오류가 발생할 수 있습니다.)<br />
                            2. 활용신청 승인 후 조달청 시스템 서버 반영에 약 1시간 정도 지연이 발생할 수 있습니다.<br />
                            3. 또는 하단의 <strong>[g2b 공식 웹사이트 실시간 조달 검색]</strong> 버튼을 누르시면 키 입력 없이도 100% 동일한 실제 조달 데이터를 확인하실 수 있습니다.
                        </div>
                    </div>
                )}

                {apiSuccessCount !== null && (
                    <div className="animate-fade-in" style={{ padding: '14px', background: 'rgba(16, 185, 129, 0.1)', border: '1px solid rgba(16, 185, 129, 0.3)', borderRadius: '10px', color: 'var(--success-color)', fontSize: '13.5px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <ShieldCheck size={18} /> 성공적으로 조달청 나라장터 실시간 사전규격 데이터 <strong>{apiSuccessCount}건</strong>을 수집하여 표시하였습니다.
                    </div>
                )}
            </div>

            {/* 실제 데이터 표출 테이블 뷰 */}
            <div className="glass-panel animate-slide-up" style={{ padding: '24px', display: 'flex', flexDirection: 'column', gap: '16px' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px' }}>
                    <h3 style={{ margin: 0, fontSize: '17px', color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                        📋 조달청 나라장터 실제 사전규격 수집 목록
                        <span style={{ fontSize: '13px', color: 'var(--text-muted)', fontWeight: 400 }}>({displayList.length}건 표시중)</span>
                    </h3>

                    {/* g2b 공식 웹 검색 대체 연동 버튼 */}
                    <div style={{ display: 'flex', gap: '8px' }}>
                        <a
                            href={getG2bOfficialSearchUrl('preCom')}
                            target="_blank"
                            rel="noreferrer"
                            style={{
                                display: 'inline-flex', alignItems: 'center', gap: '6px',
                                padding: '8px 14px', borderRadius: '7px',
                                background: 'rgba(168, 85, 247, 0.12)', border: '1px solid rgba(168, 85, 247, 0.3)',
                                color: '#c084fc', fontWeight: 600, fontSize: '12.5px', textDecoration: 'none'
                            }}
                        >
                            🌐 g2b 공식 웹사이트 실시간 사전규격 검색 <ExternalLink size={13} />
                        </a>
                    </div>
                </div>

                {displayList.length === 0 ? (
                    <div style={{ padding: '60px 20px', textAlign: 'center', background: 'rgba(0,0,0,0.15)', borderRadius: '14px', border: '1px dashed var(--panel-border)', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '14px' }}>
                        <Info size={36} color="#a855f7" />
                        <div>
                            <div style={{ fontSize: '16px', fontWeight: 700, color: 'var(--text-primary)' }}>
                                공공데이터포털 API키를 확인하고 [조회하기] 버튼을 누르거나, g2b 실시간 웹검색을 클릭해 주세요.
                            </div>
                            <div style={{ fontSize: '13px', color: 'var(--text-secondary)', marginTop: '6px', maxWidth: '540px', lineHeight: 1.6 }}>
                                상단의 <strong>[📋 나라장터 실시간 사전규격 API 데이터 조회하기]</strong> 버튼을 누르시면, 승인받으신 조달청 OpenAPI로부터 수집된 실시간 사전규격 조달 리스트가 이곳에 출력을 시작합니다.
                            </div>
                        </div>

                        <a
                            href={getG2bOfficialSearchUrl('preCom')}
                            target="_blank"
                            rel="noreferrer"
                            style={{
                                display: 'inline-flex', alignItems: 'center', gap: '6px',
                                padding: '10px 20px', borderRadius: '9px',
                                background: 'linear-gradient(135deg, #a855f7, #9333ea)', color: 'white',
                                fontWeight: 700, fontSize: '14px', textDecoration: 'none',
                                boxShadow: '0 4px 14px rgba(168, 85, 247, 0.3)'
                            }}
                        >
                            🌐 g2b 나라장터 공식사이트 실시간 사전규격 보기 <ExternalLink size={15} />
                        </a>
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
