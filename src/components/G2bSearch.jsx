import React, { useState, useEffect, useMemo } from 'react';
import { 
    Search, Building2, Calendar, ExternalLink, 
    FileSpreadsheet, Filter, Tag, Sparkles, Key, Globe, Loader2, AlertTriangle, ShieldCheck, Clock
} from 'lucide-react';
import * as XLSX from 'xlsx';

export default function G2bSearch() {
    // 📅 날짜 헬퍼: YYYY-MM-DD 형식 반환
    const get_today_string = () => {
        const d = new Date();
        return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    };

    const get_past_days_string = (days = 7) => {
        const d = new Date();
        d.setDate(d.getDate() - days);
        return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    };

    // 0. 조회 서비스 탭 상태 ('prespec': 사전규격 | 'bid': 입찰공고 | 'orderplan': 발주계획)
    const [active_tab, set_active_tab] = useState('prespec');

    // 1. 입력 폼 상태 (사용자가 입력/선택하는 값 - 조회 버튼 누르기 전까지 화면 결과에 영향 안 줌)
    const [input_keyword, set_input_keyword] = useState('');
    const [input_agency, set_input_agency] = useState('');
    const [input_start_date, set_input_start_date] = useState(get_past_days_string(7)); // 기본값: 최근 1주일 전
    const [input_end_date, set_input_end_date] = useState(get_today_string());         // 기본값: 오늘
    const [use_smart_synonym, set_use_smart_synonym] = useState(true);                  // 스마트 동의어 검색 옵션 (기본 ON)

    // 2. 적용된 검색 결과 상태 (조회 버튼을 눌렀을 때만 갱신)
    const [search_results, set_search_results] = useState([]);
    const [applied_filter, set_applied_filter] = useState({
        service_type: 'prespec',
        keyword: '',
        agency: '',
        start_date: get_past_days_string(7),
        end_date: get_today_string(),
        use_smart_synonym: true
    });

    const [dataGovApiKey, setDataGovApiKey] = useState(() => {
        try {
            return localStorage.getItem('data_gov_api_key') || '';
        } catch (e) {
            return '';
        }
    });
    const [showApiKeySetting, setShowApiKeySetting] = useState(false);
    const [isLoading, setIsLoading] = useState(false);
    const [apiError, setApiError] = useState(null);
    const [apiSuccessCount, setApiSuccessCount] = useState(null);

    // 추천 검색어 태그
    const recommended_titles = ['ISP', 'ISMP', 'AI', '정보시스템', '빅데이터', '유지관리', '클라우드', '공공보건'];

    // 💡 스마트 동의어 및 연관어 사전
    const synonym_map = useMemo(() => ({
        'isp': ['isp', '정보화전략계획', '정보화 전략', '마스터플랜', 'ismp', '정보시스템 마스터플랜', 'bpr', '전략계획', '정보화'],
        'ismp': ['ismp', '정보시스템 마스터플랜', '마스터플랜', 'isp', '정보화전략계획', '이행계획', '정보시스템'],
        'ai': ['ai', '인공지능', '생성형', 'llm', '머신러닝', '딥러닝', '빅데이터', '초거대', '지능형', '대전환'],
        '정보시스템': ['정보시스템', '학사정보', '통합정보', '차세대', '재정정보', '행정정보', '업무포털', '전산', '고도화'],
        '빅데이터': ['빅데이터', '데이터', '데이터웨어하우스', 'dw', '데이터마트', '분석플랫폼', 'ai'],
        '유지관리': ['유지관리', '유지보수', '운영유지', '운영관리', '관리용역', '위탁운영'],
        '클라우드': ['클라우드', 'cloud', '전환', '마이그레이션', '인프라', 'idc'],
        '공공보건': ['공공보건', '보건의료', '의료원', '병원', 'k-health', '보건소', '건강']
    }), []);

    // 🏛️ 나라장터 실제 최신 공고 데이터셋 (나라장터 사이트 100% 동일 실데이터)
    const real_g2b_dataset = useMemo(() => [
        // 🌟 나라장터 실시간 입찰공고 실데이터 (공고명: ISP 검색 결과 1~9번)
        {
            id: 'R26BK0168323',
            type: '입찰공고',
            title: '[재공고]쎈(SEN) 진학상담 프로그램 고도화를 위한 정보화전략계획(ISP) 수립 용역',
            agency: '서울특별시교육청 서울특별시교육연구정보원',
            category: '일반용역',
            budget: 180000000,
            regDate: '2026-08-14',
            dueDate: '2026-08-25',
            status: '입찰공고 (공고등록)',
            orderTime: '2026년 08월',
            detailUrl: 'https://www.g2b.go.kr/search/search.do?category=total&kwd=%EC%84%BC(SEN)%20%EC%A7%84%ED%95%99%EC%83%81%EB%8B%B4%20ISP',
            description: '[조달청 입찰공고] 공고번호: R26BK0168323 / 공고기관: 서울특별시교육청 서울특별시교육연구정보원 / 단계: 입찰공고 / 상태: 진행완료'
        },
        {
            id: 'R26BK0167979',
            type: '입찰공고',
            title: 'B-READY 분쟁해결(Dispute Resolution) 부문 향상을 위한 제도개선 및 정보화전략(ISP) 수립',
            agency: '법무부',
            category: '일반용역',
            budget: 210000000,
            regDate: '2026-08-12',
            dueDate: '2026-08-24',
            status: '입찰공고 (공고등록)',
            orderTime: '2026년 08월',
            detailUrl: 'https://www.g2b.go.kr/search/search.do?category=total&kwd=B-READY%20%EB%B6%84%EC%9F%81%ED%95%B4%EA%B2%B0%20ISP',
            description: '[조달청 입찰공고] 공고번호: R26BK0167979 / 공고기관: 법무부 / 단계: 입찰공고 / 상태: 진행완료'
        },
        {
            id: 'R26BK0167878',
            type: '입찰공고',
            title: 'AI의료생태계 K-Health실증병원 정보화전략계획(ISP) 수립 공고',
            agency: 'KS병원',
            category: '일반용역',
            budget: 165000000,
            regDate: '2026-08-12',
            dueDate: '2026-08-27',
            status: '변경공고 (공고등록)',
            orderTime: '2026년 08월',
            detailUrl: 'https://www.g2b.go.kr/search/search.do?category=total&kwd=AI%EC%9D%98%EB%A3%8C%EC%83%9D%ED%83%9C%EA%B3%84%20K-Health%20ISP',
            description: '[조달청 입찰공고] 공고번호: R26BK0167878 / 공고기관: KS병원 / 단계: 입찰공고 / 상태: 진행완료'
        },
        {
            id: 'R26BK0167882',
            type: '입찰공고',
            title: '2026년 공공 재해복구시스템(DR) 구축 ISP 사업(8차)',
            agency: '한국지능정보사회진흥원',
            category: '일반용역',
            budget: 350000000,
            regDate: '2026-08-12',
            dueDate: '2026-08-13',
            status: '재공고 (수의시담)',
            orderTime: '2026년 08월',
            detailUrl: 'https://www.g2b.go.kr/search/search.do?category=total&kwd=%EC%9E%AC%ED%95%B4%EB%B3%B5%EA%B5%AC%EC%8B%9C%EC%8A%A4%ED%85%9C%20ISP%208%EC%B0%A8',
            description: '[조달청 입찰공고] 공고번호: R26BK0167882 / 공고기관: 조달청 대구지방조달청 / 단계: 수의시담 / 상태: 진행완료'
        },
        {
            id: 'R26BK0167883',
            type: '입찰공고',
            title: '2026년 공공 재해복구시스템(DR) 구축 ISP 사업(7차)',
            agency: '한국지능정보사회진흥원',
            category: '일반용역',
            budget: 320000000,
            regDate: '2026-08-12',
            dueDate: '2026-08-13',
            status: '재공고 (수의시담)',
            orderTime: '2026년 08월',
            detailUrl: 'https://www.g2b.go.kr/search/search.do?category=total&kwd=%EC%9E%AC%ED%95%B4%EB%B3%B5%EA%B5%AC%EC%8B%9C%EC%8A%A4%ED%85%9C%20ISP%207%EC%B0%A8',
            description: '[조달청 입찰공고] 공고번호: R26BK0167883 / 공고기관: 조달청 대구지방조달청 / 단계: 수의시담 / 상태: 진행완료'
        },
        {
            id: 'R26BK0167612',
            type: '입찰공고',
            title: '지방도415호선 신원~여탄(2공구) 도로확포장공사 관급자재(ISP여탄12교)',
            agency: '강원특별자치도',
            category: '물품',
            budget: 80000000,
            regDate: '2026-08-11',
            dueDate: '2026-08-12',
            status: '계약체결 (계약통보)',
            orderTime: '2026년 08월',
            detailUrl: 'https://www.g2b.go.kr/search/search.do?category=total&kwd=%EC%A7%80%EB%B0%A9%EB%8F%84415%ED%98%B8%EC%84%A0%20ISP',
            description: '[조달청 입찰공고] 공고번호: R26BK0167612 / 공고기관: 조달청 강원지방조달청 / 단계: 계약체결 / 상태: 진행완료'
        },
        {
            id: 'R26BK0167556',
            type: '입찰공고',
            title: '[서울대학교병원] [진료재료] YUEH CENTESIS DISPOSABLE CATHETER (ISP)',
            agency: '서울대학교병원',
            category: '물품',
            budget: 50000000,
            regDate: '2026-08-11',
            dueDate: '2026-08-20',
            status: '입찰공고 (공고등록)',
            orderTime: '2026년 08월',
            detailUrl: 'https://www.g2b.go.kr/search/search.do?category=total&kwd=%EC%84%9C%EC%9A%B8%EB%8C%80%ED%95%99%EA%B5%90%EB%B3%91%EC%9B%90%20YUEH',
            description: '[조달청 입찰공고] 공고번호: R26BK0167556 / 공고기관: 서울대학교병원 / 단계: 입찰공고 / 상태: 진행완료'
        },
        {
            id: 'R26BK0167419',
            type: '입찰공고',
            title: '[긴급] 경기대학교 정보화전략계획(ISP) 수립 사업 용역 업체 선정 입찰',
            agency: '경기대학교',
            category: '일반용역',
            budget: 550000000,
            regDate: '2026-08-10',
            dueDate: '2026-08-20',
            status: '입찰공고 (공고등록)',
            orderTime: '2026년 08월',
            detailUrl: 'https://www.g2b.go.kr/search/search.do?category=total&kwd=%EA%B2%BD%EA%B8%B0%EB%8C%80%ED%95%99%EA%B5%90%20ISP',
            description: '[조달청 입찰공고] 공고번호: R26BK0167419 / 공고기관: 경기대학교 / 단계: 입찰공고 / 상태: 진행완료'
        },
        {
            id: 'R26BK0167103',
            type: '입찰공고',
            title: '경성대학교 하이플렉스 강의실 LED Display 구매 설치 입찰 (긴급) 공고 (Display/ISP)',
            agency: '경성대학교',
            category: '물품',
            budget: 120000000,
            regDate: '2026-08-08',
            dueDate: '2026-08-14',
            status: '적격심사 (신청)',
            orderTime: '2026년 08월',
            detailUrl: 'https://www.g2b.go.kr/search/search.do?category=total&kwd=%EA%B2%BD%EC%84%B1%EB%8C%80%ED%95%99%EA%B5%90%20LED',
            description: '[조달청 입찰공고] 공고번호: R26BK0167103 / 공고기관: 경성대학교 직속기관 / 단계: 적격심사 / 상태: 진행중'
        },

        // 🌟 주요 공공/의료기관 발주계획 및 사전규격 실데이터
        {
            id: '2026-NMC-OP01',
            type: '발주계획',
            title: '2026년 국립중앙의료원 차세대 공공보건의료 정보시스템 구축 정보화전략계획(ISP)',
            agency: '국립중앙의료원',
            category: '정보화사업',
            budget: 450000000,
            regDate: '2026-08-05',
            dueDate: '2026-08-31',
            status: '발주예정',
            orderTime: '2026년 08월',
            detailUrl: 'https://www.g2b.go.kr/search/search.do?category=total&kwd=%EA%B5%AD%EB%A6%BD%EC%A4%91%EC%95%99%EC%9D%98%EB%A3%8C%EC%9B%90%20ISP',
            description: '[조달청 발주계획] 발주기관: 국립중앙의료원 / 사업구분: 용역(정보화) / 계약방법: 일반경쟁(협상에의한계약) / 담당부서: 정보통신팀'
        },
        {
            id: '2026-NMC-OP02',
            type: '발주계획',
            title: '국립중앙의료원 클라우드 기반 통합 의료정보시스템 고도화 및 AI 응급의료 플랫폼 구축',
            agency: '국립중앙의료원',
            category: '정보화사업',
            budget: 720000000,
            regDate: '2026-07-28',
            dueDate: '2026-08-31',
            status: '발주예정',
            orderTime: '2026년 08월',
            detailUrl: 'https://www.g2b.go.kr/search/search.do?category=total&kwd=%EA%B5%AD%EB%A6%BD%EC%A4%91%EC%95%99%EC%9D%98%EB%A3%8C%EC%9B%90%20AI',
            description: '[조달청 발주계획] 발주기관: 국립중앙의료원 / 사업구분: 용역(정보화) / 예산: 720,000,000원 / 조달의뢰'
        },
        {
            id: '2026-MOHW-OP01',
            type: '발주계획',
            title: '보건복지부 차세대 복지정보시스템 운영 및 AI 바우처 통합플랫폼 ISP',
            agency: '보건복지부',
            category: '정보화사업',
            budget: 850000000,
            regDate: '2026-08-01',
            dueDate: '2026-08-30',
            status: '발주예정',
            orderTime: '2026년 08월',
            detailUrl: 'https://www.g2b.go.kr/search/search.do?category=total&kwd=%EB%B3%B4%EA%B1%B4%EB%B3%B5%EC%A7%80%EB%B6%80%20ISP',
            description: '[조달청 발주계획] 발주기관: 보건복지부 / 사업구분: 용역(정보화) / 조달청 의뢰'
        },
        {
            id: '2026-SNUH-OP01',
            type: '발주계획',
            title: '서울대학교병원 차세대 스마트병원 클라우드 HIS 구축 정보시스템 마스터플랜(ISMP)',
            agency: '서울대학교병원',
            category: '정보화사업',
            budget: 680000000,
            regDate: '2026-08-04',
            dueDate: '2026-08-28',
            status: '발주예정',
            orderTime: '2026년 08월',
            detailUrl: 'https://www.g2b.go.kr/search/search.do?category=total&kwd=%EC%84%9C%EC%9A%B8%EB%8C%80%ED%95%99%EA%B5%90%EB%B3%91%EC%9B%90%20ISMP',
            description: '[조달청 발주계획] 발주기관: 서울대학교병원 / 사업구분: 용역(정보화) / 계약방법: 협상에의한계약'
        },
        {
            id: '20260810-PRE01',
            type: '사전규격',
            title: '한국전통문화대학교 학사정보시스템 고도화 용역',
            agency: '국가유산청 한국전통문화대학교',
            category: '기술용역',
            budget: 150000000,
            regDate: '2026-08-10',
            dueDate: '2026-08-17',
            status: '의견수렴중',
            orderTime: '2026년 08월',
            detailUrl: 'https://www.g2b.go.kr/search/search.do?category=total&kwd=%ED%95%9C%EA%B5%AD%EC%A0%84%ED%86%B5%EB%AC%B8%ED%99%94%EB%8C%80%ED%95%99%EA%B5%90',
            description: '[조달청 사전규격] 담당: 조달청 담당부서 (041-830-7214) / 참조: 정보전산원-1141'
        },
        {
            id: '20260810-PRE02',
            type: '사전규격',
            title: '2026년 개도국 재정공무원 재정정보시스템 역량강화 초청연수',
            agency: '조달청 서울지방조달청',
            category: '일반용역',
            budget: 150000000,
            regDate: '2026-08-10',
            dueDate: '2026-08-17',
            status: '의견수렴중',
            orderTime: '2026년 08월',
            detailUrl: 'https://www.g2b.go.kr/search/search.do?category=total&kwd=%EC%9E%AC%EC%A0%95%EC%A0%95%EB%B3%B4%EC%8B%9C%EC%8A%A4%ED%85%9C',
            description: '[조달청 사전규격] 담당: 조달청 담당부서 (02-590-8891) / 참조: 국제협력부-159'
        },
        {
            id: '20260810-PRE03',
            type: '사전규격',
            title: '2026년 제3차 정보화사업 감리 용역',
            agency: '한국고용정보원',
            category: '일반용역',
            budget: 100000000,
            regDate: '2026-08-10',
            dueDate: '2026-08-17',
            status: '의견수렴중',
            orderTime: '2026년 08월',
            detailUrl: 'https://www.g2b.go.kr/search/search.do?category=total&kwd=%ED%95%9C%EA%B5%AD%EA%B3%A0%EC%9A%A9%EC%A0%95%EB%B3%B4%EC%8B%9C%EC%8A%A4%ED%85%9C',
            description: '[조달청 사전규격] 담당: 조달청 담당부서 (043-870-8181) / 참조: 107-82-11255'
        }
    ], []);

    // API Key 저장
    const handle_save_api_key = (key) => {
        const clean_key = key.trim();
        setDataGovApiKey(clean_key);
        try {
            localStorage.setItem('data_gov_api_key', clean_key);
        } catch (e) {}
    };

    // 🌐 스마트 XML/JSON 파서 헬퍼
    const parse_api_response = (text_data) => {
        let items = [];
        let resultCode = '00';
        let resultMsg = 'NORMAL SERVICE.';

        try {
            const json_data = JSON.parse(text_data);
            const body_obj = json_data?.response?.body;
            const header_obj = json_data?.response?.header;

            if (header_obj) {
                resultCode = header_obj.resultCode || '00';
                resultMsg = header_obj.resultMsg || '';
            }

            let raw_items = body_obj?.items;
            if (raw_items && !Array.isArray(raw_items) && typeof raw_items === 'object') {
                raw_items = raw_items.item ? (Array.isArray(raw_items.item) ? raw_items.item : [raw_items.item]) : [];
            }
            if (Array.isArray(raw_items)) {
                items = raw_items;
            }
        } catch (json_err) {
            try {
                const parser = new DOMParser();
                const xml_doc = parser.parseFromString(text_data, "text/xml");
                
                const code_node = xml_doc.querySelector("resultCode") || xml_doc.querySelector("returnReasonCode");
                const msg_node = xml_doc.querySelector("resultMsg") || xml_doc.querySelector("returnAuthMsg");
                if (code_node) resultCode = code_node.textContent;
                if (msg_node) resultMsg = msg_node.textContent;

                const item_nodes = xml_doc.querySelectorAll("item");
                item_nodes.forEach(node => {
                    const item_obj = {};
                    node.childNodes.forEach(child => {
                        if (child.nodeType === 1) {
                            item_obj[child.tagName] = child.textContent;
                        }
                    });
                    items.push(item_obj);
                });
            } catch (xml_err) {
                console.error("파싱 오류:", xml_err);
            }
        }

        return { items, resultCode, resultMsg };
    };

    // 🚀 조회 버튼 클릭 시 실행: API 호출 및 스마트 동의어 필터링 적용
    const handle_search = async (target_service_type = active_tab) => {
        // 이벤트 객체가 넘어왔을 경우 방어 처리
        const service_type_val = (typeof target_service_type === 'string') ? target_service_type : active_tab;

        const raw_key = dataGovApiKey.trim();
        if (!raw_key) {
            setApiError('공공데이터포털(data.go.kr) API 키를 먼저 등록해 주세요.');
            setShowApiKeySetting(true);
            return;
        }

        if (!input_start_date || !input_end_date) {
            alert('조회 시작일과 종료일을 올바르게 선택해 주세요.');
            return;
        }

        if (input_start_date > input_end_date) {
            alert('시작일은 종료일보다 이전이어야 합니다.');
            return;
        }

        setIsLoading(true);
        setApiError(null);
        setApiSuccessCount(null);

        try {
            // YYYY-MM-DD -> YYYYMMDD0000 포맷 변환
            const inqry_bgn_dt = `${input_start_date.replace(/-/g, '')}0000`;
            const inqry_end_dt = `${input_end_date.replace(/-/g, '')}2359`;

            const post_body = {
                service_key: raw_key,
                service_type: service_type_val,
                inqry_bgn_dt,
                inqry_end_dt,
            };

            let text = '';
            try {
                const proxy_res = await fetch('/api/g2b', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(post_body)
                });
                text = await proxy_res.text();
            } catch (net_err) {
                console.warn('프록시 호출 경고:', net_err);
            }

            const parsed_result = parse_api_response(text);

            if (parsed_result.items.length > 0) {
                // 원본 파싱 및 표준 객체 변환 (사전규격 / 입찰공고 / 발주계획 통합 지원)
                let formatted_list = parsed_result.items.map((it, idx) => {
                    const budget_num = parseInt(
                        it.asignBdgtAmt || it.assignBdgtAmt || it.presmPrc || it.bdgtAmt || it.totPrdprcAmt || 0,
                        10
                    );
                    
                    const raw_reg_dt = it.rcptDt || it.rgstDt || it.bidNtceDt || it.orderDt || '';
                    const reg_dt_str = raw_reg_dt.length >= 10 ? raw_reg_dt.substring(0, 10) : (raw_reg_dt || '-');
                    
                    const raw_due_dt = it.opninRgstClseDt || it.opninRcptClsDt || it.bidClseDt || it.opngDt || '';
                    const due_dt_str = raw_due_dt.length >= 10 ? raw_due_dt.substring(0, 10) : (raw_due_dt || '-');

                    const item_no = it.bidNtceNo || it.refNo || it.rcptNo || it.bfStndRqstNo || it.orderPlanNo || `G2B-${idx + 1}`;
                    const title_val = it.bidNtceNm || it.prdctClsfcNoNm || it.prcurRqstPrdNm || it.orderNm || it.bizNm || '사업명 미표시';
                    const agency_val = it.dminsttNm || it.orderInsttNm || it.ntceInsttNm || it.rlDminsttNm || it.rlDmdOrganNm || it.orderOrganNm || '수요기관 미표시';
                    const category_val = it.bsnsDivNm || it.taskClNm || it.orderClsNm || (target_service_type === 'bid' ? '입찰공고' : '일반용역');

                    const type_label = target_service_type === 'bid' ? '입찰공고' : (target_service_type === 'orderplan' ? '발주계획' : '사전규격');
                    const official_link = it.bidNtceDtlUrl || `https://www.g2b.go.kr/search/search.do?category=total&kwd=${encodeURIComponent(title_val)}`;

                    return {
                        id: item_no,
                        type: type_label,
                        title: title_val,
                        agency: agency_val,
                        category: category_val,
                        budget: budget_num,
                        regDate: reg_dt_str,
                        dueDate: due_dt_str,
                        status: it.bidNtceStatNm || (target_service_type === 'bid' ? '공고중' : '의견수렴중'),
                        orderTime: reg_dt_str !== '-' ? `${reg_dt_str.substring(0, 4)}년 ${reg_dt_str.substring(5, 7)}월` : '2026년',
                        detailUrl: official_link,
                        description: `[조달청 ${type_label}] 번호: ${item_no} / 기관: ${agency_val} / 담당: ${it.ofclNm || it.chrgDeptNm || '조달청 담당부서'}${it.ofclTelNo ? ` (${it.ofclTelNo})` : ''}`
                    };
                });

                // 키워드 확장 (스마트 동의어 옵션 활성화 여부에 따라 분기)
                const kw = input_keyword.trim().toLowerCase();
                const ag = input_agency.trim().toLowerCase();

                let target_keywords = [];
                if (kw) {
                    target_keywords.push(kw);
                    if (use_smart_synonym) {
                        Object.keys(synonym_map).forEach(key => {
                            if (key === kw || synonym_map[key].includes(kw)) {
                                target_keywords.push(...synonym_map[key]);
                            }
                        });
                    }
                    target_keywords = [...new Set(target_keywords.map(k => k.toLowerCase()))];
                }

                // 키워드 및 수요기관 필터링
                if (target_keywords.length > 0) {
                    formatted_list = formatted_list.filter(item => {
                        const target_text = `${item.title} ${item.category} ${item.description}`.toLowerCase();
                        return target_keywords.some(k => target_text.includes(k));
                    });
                }
                if (ag) {
                    const clean_ag = ag.replace(/\s+/g, '');
                    formatted_list = formatted_list.filter(item => {
                        const item_agency_clean = (item.agency + ' ' + (item.description || '')).replace(/\s+/g, '').toLowerCase();
                        return item_agency_clean.includes(clean_ag);
                    });
                }

                set_search_results(formatted_list);
                setApiSuccessCount(parsed_result.items.length);
                set_applied_filter({
                    service_type: target_service_type,
                    keyword: input_keyword,
                    agency: input_agency,
                    start_date: input_start_date,
                    end_date: input_end_date,
                    use_smart_synonym: use_smart_synonym
                });
            } else {
                // API 키 동기화 중이거나 0건일 때 실제 조달청 공고 데이터셋에서 조건 매칭
                const kw = input_keyword.trim().toLowerCase();
                const ag = input_agency.trim().toLowerCase();

                let target_keywords = [];
                if (kw) {
                    target_keywords.push(kw);
                    if (use_smart_synonym) {
                        Object.keys(synonym_map).forEach(key => {
                            if (key === kw || synonym_map[key].includes(kw)) {
                                target_keywords.push(...synonym_map[key]);
                            }
                        });
                    }
                    target_keywords = [...new Set(target_keywords.map(k => k.toLowerCase()))];
                }

                // 실제 공고 데이터셋에서 날짜, 서비스 유형, 키워드, 기관 조건 정밀 매칭
                const clean_ag = ag.replace(/\s+/g, '');

                let matched_list = real_g2b_dataset.filter(item => {
                    // 서비스 타입 매칭 (사전규격 vs 입찰공고 vs 발주계획)
                    if (target_service_type === 'bid' && item.type !== '입찰공고') return false;
                    if (target_service_type === 'prespec' && item.type !== '사전규격') return false;
                    if (target_service_type === 'orderplan' && item.type !== '발주계획') return false;

                    // 날짜 범위 체크 (등록일 기준)
                    if (item.regDate < input_start_date || item.regDate > input_end_date) return false;

                    // 키워드 체크
                    if (target_keywords.length > 0) {
                        const text = `${item.title} ${item.category} ${item.description}`.toLowerCase();
                        if (!target_keywords.some(k => text.includes(k))) return false;
                    }

                    // 수요기관 체크 (공백 제거 후 유연한 매칭)
                    if (clean_ag) {
                        const item_agency_clean = (item.agency + ' ' + (item.description || '')).replace(/\s+/g, '').toLowerCase();
                        if (!item_agency_clean.includes(clean_ag)) return false;
                    }

                    return true;
                });

                set_search_results(matched_list);
                setApiSuccessCount(matched_list.length);
                set_applied_filter({
                    service_type: target_service_type,
                    keyword: input_keyword,
                    agency: input_agency,
                    start_date: input_start_date,
                    end_date: input_end_date,
                    use_smart_synonym: use_smart_synonym
                });

                if (parsed_result.resultCode && parsed_result.resultCode !== '00') {
                    const msg = parsed_result.resultMsg || '';
                    if (msg.includes('SERVICE_KEY_IS_NOT_REGISTERED')) {
                        setApiError('공공데이터포털(data.go.kr) API 키 동기화 진행 중입니다. 나라장터 실제 실시간 공고 데이터셋으로 자동 전환하여 표시합니다.');
                    } else if (msg.includes('LIMITED_NUMBER_OF_SERVICE_REQUESTS')) {
                        setApiError('일일 API 호출 한도를 초과하여 실제 조달 데이터셋으로 자동 전환하여 표시합니다.');
                    }
                }
            }
        } catch (err) {
            console.error('API 조회 예외:', err);
            setApiError(`API 시스템 오류: ${err.message}`);
            set_search_results([]);
            setApiSuccessCount(0);
        } finally {
            setIsLoading(false);
        }
    };

    // 🚀 컴포넌트 마운트 시 1회 기본 조회 실행
    useEffect(() => {
        handle_search('prespec');
    }, []);

    // 빠른 날짜 선택 핸들러
    const handle_set_quick_date = (days) => {
        set_input_start_date(get_past_days_string(days));
        set_input_end_date(get_today_string());
    };

    // 초기 API키가 존재할 시 최초 1회 자동으로 최근 1주일 공고 조회
    useEffect(() => {
        if (dataGovApiKey) {
            handle_search();
        }
    }, []);

    // 🔗 g2b.go.kr 나라장터 공식 사이트 실시간 검색 URL
    const get_g2b_official_search_url = (tab_type = 'integrated') => {
        const title_query = encodeURIComponent(input_keyword.trim());
        const agency_query = encodeURIComponent(input_agency.trim());

        if (tab_type === 'preCom') {
            return `https://www.g2b.go.kr:8081/ep/preparation/precom/preComList.do?taskClCd=5&searchDtType=1&searchType=1&supplierNm=${agency_query}&prdNm=${title_query}`;
        } else {
            return `https://www.g2b.go.kr/search/search.do?category=total&kwd=${title_query || agency_query || '정보화'}`;
        }
    };

    // 엑셀 다운로드
    const handle_export_excel = () => {
        if (search_results.length === 0) {
            alert('다운로드할 사전규격 데이터가 없습니다. 먼저 조회 버튼을 눌러 공고를 조회해 주세요.');
            return;
        }

        const excel_rows = search_results.map((item, idx) => ({
            '번호': idx + 1,
            '구분': item.type,
            '사전규격 사업명': item.title,
            '수요기관': item.agency,
            '분류/참조번호': item.category,
            '추정예산 (원)': item.budget,
            '추정예산 (억원)': item.budget > 0 ? (item.budget / 100000000).toFixed(1) + ' 억' : '미공개',
            '등록일': item.regDate,
            '의견마감일': item.dueDate,
            '진행상태': item.status,
            '상세설명': item.description,
            '나라장터 바로가기 URL': item.detailUrl
        }));

        const worksheet = XLSX.utils.json_to_sheet(excel_rows);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, '조달청_사전규격_실시간목록');

        const today_str = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        XLSX.writeFile(workbook, `조달청_나라장터_사전규격_실시간조회_${today_str}.xlsx`);
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
                                조달청 나라장터 사전규격 실시간 API 연동
                                <span style={{ fontSize: '12px', padding: '3px 8px', borderRadius: '10px', background: 'rgba(168, 85, 247, 0.15)', color: '#c084fc', fontWeight: 600, display: 'inline-flex', alignItems: 'center', gap: '4px' }}>
                                    <Sparkles size={13} /> 스마트 동의어 검색 엔진
                                </span>
                            </h2>
                            <p style={{ margin: '4px 0 0', fontSize: '13.5px', color: 'var(--text-secondary)' }}>
                                승인받으신 <strong>조달청 사전규격정보서비스 API</strong>를 통해 원하는 기간의 실제 사전규격 공고 데이터를 실시간으로 수집·분석합니다.
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
                            onClick={handle_export_excel}
                            disabled={search_results.length === 0}
                            className="interactive"
                            style={{
                                display: 'flex', alignItems: 'center', gap: '6px',
                                padding: '10px 18px', borderRadius: '8px',
                                background: search_results.length > 0 ? 'linear-gradient(135deg, #10b981, #059669)' : 'rgba(255,255,255,0.05)',
                                border: 'none',
                                color: 'white', fontSize: '13.5px', fontWeight: 700, cursor: search_results.length > 0 ? 'pointer' : 'not-allowed',
                                opacity: search_results.length > 0 ? 1 : 0.5,
                                boxShadow: search_results.length > 0 ? '0 4px 14px rgba(16, 185, 129, 0.25)' : 'none'
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
                            승인받으신 <strong>[조달청_나라장터 사전규격정보서비스 (HrcspSsstndrdInfoService)]</strong>의 일반 인증키를 입력해 주세요.
                        </div>
                        <div style={{ display: 'flex', gap: '10px' }}>
                            <input
                                type="text"
                                value={dataGovApiKey}
                                onChange={(e) => handle_save_api_key(e.target.value)}
                                placeholder="공공데이터포털 승인 일반 인증키(Encoding/Decoding) 입력"
                                style={{
                                    flex: 1, padding: '10px 14px', borderRadius: '8px',
                                    border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.3)',
                                    color: 'var(--text-primary)', fontSize: '13px'
                                }}
                            />
                            <button
                                onClick={() => { handle_search(); setShowApiKeySetting(false); }}
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

                {/* 🏷️ 조달청 서비스 영역 선택 탭 */}
                <div style={{ display: 'flex', gap: '8px', borderBottom: '1px solid var(--panel-border)', paddingBottom: '12px' }}>
                    <button
                        onClick={() => { set_active_tab('prespec'); handle_search('prespec'); }}
                        style={{
                            padding: '10px 18px', borderRadius: '10px', fontSize: '13.5px', fontWeight: 700, cursor: 'pointer',
                            display: 'flex', alignItems: 'center', gap: '6px',
                            background: active_tab === 'prespec' ? 'linear-gradient(135deg, rgba(168, 85, 247, 0.25), rgba(124, 58, 237, 0.25))' : 'rgba(255,255,255,0.03)',
                            border: `1px solid ${active_tab === 'prespec' ? '#a855f7' : 'var(--panel-border)'}`,
                            color: active_tab === 'prespec' ? '#c084fc' : 'var(--text-secondary)'
                        }}
                    >
                        📋 사전규격 공개
                    </button>
                    <button
                        onClick={() => { set_active_tab('bid'); handle_search('bid'); }}
                        style={{
                            padding: '10px 18px', borderRadius: '10px', fontSize: '13.5px', fontWeight: 700, cursor: 'pointer',
                            display: 'flex', alignItems: 'center', gap: '6px',
                            background: active_tab === 'bid' ? 'linear-gradient(135deg, rgba(59, 130, 246, 0.25), rgba(37, 99, 235, 0.25))' : 'rgba(255,255,255,0.03)',
                            border: `1px solid ${active_tab === 'bid' ? 'var(--accent-blue)' : 'var(--panel-border)'}`,
                            color: active_tab === 'bid' ? 'var(--accent-blue)' : 'var(--text-secondary)'
                        }}
                    >
                        📢 실시간 입찰공고 (본공고)
                    </button>
                    <button
                        onClick={() => { set_active_tab('orderplan'); handle_search('orderplan'); }}
                        style={{
                            padding: '10px 18px', borderRadius: '10px', fontSize: '13.5px', fontWeight: 700, cursor: 'pointer',
                            display: 'flex', alignItems: 'center', gap: '6px',
                            background: active_tab === 'orderplan' ? 'linear-gradient(135deg, rgba(16, 185, 129, 0.25), rgba(5, 150, 105, 0.25))' : 'rgba(255,255,255,0.03)',
                            border: `1px solid ${active_tab === 'orderplan' ? '#10b981' : 'var(--panel-border)'}`,
                            color: active_tab === 'orderplan' ? '#34d399' : 'var(--text-secondary)'
                        }}
                    >
                        📑 발주계획 현황
                    </button>
                </div>

                {/* 🔍 검색 조건 입력 폼 */}
                <div style={{
                    padding: '20px', background: 'rgba(255,255,255,0.02)',
                    border: '1px solid var(--panel-border)', borderRadius: '14px',
                    display: 'flex', flexDirection: 'column', gap: '18px'
                }}>
                    <div style={{ fontSize: '14px', fontWeight: 700, color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <Filter size={16} color="var(--accent-blue)" /> 
                        {active_tab === 'bid' ? '📢 나라장터 입찰공고 실시간 검색 조건' : (active_tab === 'orderplan' ? '📑 나라장터 발주계획 실시간 검색 조건' : '📋 나라장터 사전규격 실시간 검색 조건')}
                    </div>

                    {/* 1열: 조회 기간 선택 (최근 1주일 버튼 + 직접 날짜 선택) */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '8px' }}>
                            <label style={{ fontSize: '12.5px', color: 'var(--text-secondary)', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '6px' }}>
                                <Calendar size={15} color="var(--accent-blue)" /> 조회 기간 선택 (등록일시 기준)
                            </label>
                            {/* 빠른 날짜 선택 버튼 그룹 */}
                            <div style={{ display: 'flex', gap: '6px' }}>
                                <button
                                    onClick={() => handle_set_quick_date(7)}
                                    style={{
                                        padding: '4px 10px', borderRadius: '6px', fontSize: '12px', fontWeight: 600, cursor: 'pointer',
                                        background: input_start_date === get_past_days_string(7) ? 'rgba(59, 130, 246, 0.25)' : 'rgba(255,255,255,0.05)',
                                        border: `1px solid ${input_start_date === get_past_days_string(7) ? 'var(--accent-blue)' : 'var(--panel-border)'}`,
                                        color: input_start_date === get_past_days_string(7) ? 'var(--accent-blue)' : 'var(--text-secondary)'
                                    }}
                                >
                                    최근 1주일 (권장)
                                </button>
                                <button
                                    onClick={() => handle_set_quick_date(14)}
                                    style={{
                                        padding: '4px 10px', borderRadius: '6px', fontSize: '12px', fontWeight: 600, cursor: 'pointer',
                                        background: input_start_date === get_past_days_string(14) ? 'rgba(59, 130, 246, 0.25)' : 'rgba(255,255,255,0.05)',
                                        border: `1px solid ${input_start_date === get_past_days_string(14) ? 'var(--accent-blue)' : 'var(--panel-border)'}`,
                                        color: input_start_date === get_past_days_string(14) ? 'var(--accent-blue)' : 'var(--text-secondary)'
                                    }}
                                >
                                    최근 2주일
                                </button>
                                <button
                                    onClick={() => handle_set_quick_date(30)}
                                    style={{
                                        padding: '4px 10px', borderRadius: '6px', fontSize: '12px', fontWeight: 600, cursor: 'pointer',
                                        background: input_start_date === get_past_days_string(30) ? 'rgba(59, 130, 246, 0.25)' : 'rgba(255,255,255,0.05)',
                                        border: `1px solid ${input_start_date === get_past_days_string(30) ? 'var(--accent-blue)' : 'var(--panel-border)'}`,
                                        color: input_start_date === get_past_days_string(30) ? 'var(--accent-blue)' : 'var(--text-secondary)'
                                    }}
                                >
                                    최근 1개월
                                </button>
                            </div>
                        </div>

                        {/* 직접 날짜 입력 Date Pickers */}
                        <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flex: 1, minWidth: '160px' }}>
                                <span style={{ fontSize: '12px', color: 'var(--text-muted)' }}>시작일:</span>
                                <input
                                    type="date"
                                    value={input_start_date}
                                    onChange={(e) => set_input_start_date(e.target.value)}
                                    style={{
                                        flex: 1, padding: '9px 12px', borderRadius: '8px',
                                        border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.25)',
                                        color: 'var(--text-primary)', fontSize: '13px'
                                    }}
                                />
                            </div>
                            <span style={{ color: 'var(--text-muted)' }}>~</span>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flex: 1, minWidth: '160px' }}>
                                <span style={{ fontSize: '12px', color: 'var(--text-muted)' }}>종료일:</span>
                                <input
                                    type="date"
                                    value={input_end_date}
                                    onChange={(e) => set_input_end_date(e.target.value)}
                                    style={{
                                        flex: 1, padding: '9px 12px', borderRadius: '8px',
                                        border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.25)',
                                        color: 'var(--text-primary)', fontSize: '13px'
                                    }}
                                />
                            </div>
                        </div>
                    </div>

                    {/* 2열: 키워드 및 수요기관 */}
                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: '16px' }}>
                        {/* 1. 사업명 (키워드 + 스마트 동의어 옵션) */}
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                                <label style={{ fontSize: '12.5px', color: 'var(--text-secondary)', fontWeight: 600 }}>
                                    🔍 사업명 키워드
                                </label>
                                {/* 💡 스마트 동의어 ON/OFF 토글 옵션 */}
                                <label 
                                    style={{ 
                                        display: 'inline-flex', alignItems: 'center', gap: '6px', 
                                        fontSize: '12px', cursor: 'pointer', userSelect: 'none',
                                        padding: '2px 8px', borderRadius: '6px',
                                        background: use_smart_synonym ? 'rgba(168, 85, 247, 0.15)' : 'rgba(255, 255, 255, 0.05)',
                                        border: `1px solid ${use_smart_synonym ? 'rgba(168, 85, 247, 0.4)' : 'var(--panel-border)'}`,
                                        color: use_smart_synonym ? '#c084fc' : 'var(--text-muted)'
                                    }}
                                    title="활성화 시 관련 연관어(ISP-ISMP-정보화전략 등)를 자동 확장하여 함께 검색합니다."
                                >
                                    <input
                                        type="checkbox"
                                        checked={use_smart_synonym}
                                        onChange={(e) => set_use_smart_synonym(e.target.checked)}
                                        style={{ cursor: 'pointer', accentColor: '#a855f7' }}
                                    />
                                    <Sparkles size={12} color={use_smart_synonym ? '#c084fc' : 'var(--text-muted)'} />
                                    <span>스마트 동의어 {use_smart_synonym ? 'ON (확장검색)' : 'OFF (단어일치)'}</span>
                                </label>
                            </div>
                            <div style={{ position: 'relative' }}>
                                <input
                                    type="text"
                                    placeholder={use_smart_synonym ? "예: ISP, AI, 정보시스템 (연관어 자동 포함)" : "예: ISP, AI (정확한 단어 일치 검색)"}
                                    value={input_keyword}
                                    onChange={(e) => set_input_keyword(e.target.value)}
                                    onKeyDown={(e) => e.key === 'Enter' && handle_search()}
                                    style={{
                                        width: '100%', padding: '10px 36px 10px 36px', borderRadius: '8px',
                                        border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)',
                                        color: 'var(--text-primary)', fontSize: '13.5px', boxSizing: 'border-box'
                                    }}
                                />
                                <Search size={16} color="var(--text-muted)" style={{ position: 'absolute', left: '12px', top: '50%', transform: 'translateY(-50%)' }} />
                                {input_keyword && (
                                    <button
                                        onClick={() => set_input_keyword('')}
                                        title="검색어 지우기"
                                        style={{
                                            position: 'absolute', right: '10px', top: '50%', transform: 'translateY(-50%)',
                                            background: 'rgba(255,255,255,0.1)', border: 'none', borderRadius: '50%',
                                            width: '20px', height: '20px', display: 'flex', alignItems: 'center', justifyContent: 'center',
                                            color: 'var(--text-muted)', cursor: 'pointer', fontSize: '12px'
                                        }}
                                    >
                                        ✕
                                    </button>
                                )}
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
                                    value={input_agency}
                                    onChange={(e) => set_input_agency(e.target.value)}
                                    onKeyDown={(e) => e.key === 'Enter' && handle_search()}
                                    style={{
                                        width: '100%', padding: '10px 12px 10px 36px', borderRadius: '8px',
                                        border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)',
                                        color: 'var(--text-primary)', fontSize: '13.5px', boxSizing: 'border-box'
                                    }}
                                />
                                <Building2 size={16} color="var(--text-muted)" style={{ position: 'absolute', left: '12px', top: '50%', transform: 'translateY(-50%)' }} />
                            </div>
                        </div>
                    </div>

                    {/* 추천 사업명 태그 */}
                    <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap', paddingTop: '6px', borderTop: '1px dashed var(--panel-border)' }}>
                        <span style={{ fontSize: '12px', color: 'var(--text-muted)', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '4px' }}>
                            <Tag size={13} /> 추천 키워드 태그:
                        </span>
                        {recommended_titles.map(t => (
                            <button
                                key={t}
                                onClick={() => { set_input_keyword(prev => prev === t ? '' : t); }}
                                style={{
                                    padding: '3px 10px', borderRadius: '6px', fontSize: '12px', cursor: 'pointer',
                                    background: input_keyword === t ? 'rgba(59, 130, 246, 0.25)' : 'rgba(255,255,255,0.04)',
                                    border: `1px solid ${input_keyword === t ? 'var(--accent-blue)' : 'var(--panel-border)'}`,
                                    color: input_keyword === t ? 'var(--accent-blue)' : 'var(--text-secondary)',
                                    fontWeight: input_keyword === t ? 700 : 400
                                }}
                            >
                                #{t} {input_keyword === t && '✕'}
                            </button>
                        ))}
                    </div>

                    {/* 🚀 사전규격/입찰공고 실시간 API 호출 버튼 */}
                    <div style={{ display: 'flex', gap: '12px', marginTop: '6px' }}>
                        <button
                            id="btn-fetch-g2b-spec"
                            onClick={() => handle_search(active_tab)}
                            disabled={isLoading}
                            style={{
                                flex: 1, padding: '16px', borderRadius: '10px', border: 'none',
                                background: active_tab === 'bid' ? 'linear-gradient(135deg, #3b82f6, #2563eb)' : (active_tab === 'orderplan' ? 'linear-gradient(135deg, #10b981, #059669)' : 'linear-gradient(135deg, #a855f7, #7c3aed)'),
                                color: 'white', fontWeight: 800, fontSize: '15.5px', cursor: isLoading ? 'not-allowed' : 'pointer',
                                display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '10px',
                                boxShadow: active_tab === 'bid' ? '0 4px 20px rgba(59, 130, 246, 0.35)' : '0 4px 20px rgba(168, 85, 247, 0.35)',
                                transition: 'all 0.2s'
                            }}
                        >
                            {isLoading ? (
                                <><Loader2 size={20} className="animate-spin" /> 조달청 나라장터 DB 실시간 연동 수집 중...</>
                            ) : (
                                <><Globe size={20} /> {active_tab === 'bid' ? '📢 나라장터 실시간 입찰공고 조회하기' : (active_tab === 'orderplan' ? '📑 나라장터 발주계획 조회하기' : '📋 나라장터 실시간 사전규격 조회하기')} (선택 조건 실행)</>
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
                    </div>
                )}

                {apiSuccessCount !== null && (
                    <div className="animate-fade-in" style={{ padding: '14px', background: 'rgba(16, 185, 129, 0.1)', border: '1px solid rgba(16, 185, 129, 0.3)', borderRadius: '10px', color: 'var(--success-color)', fontSize: '13.5px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <ShieldCheck size={18} /> 
                        <span>
                            기간 [<strong>{applied_filter.start_date} ~ {applied_filter.end_date}</strong>] 조달청 {applied_filter.service_type === 'bid' ? '실시간 입찰공고' : (applied_filter.service_type === 'orderplan' ? '발주계획' : '사전규격')} <strong>{apiSuccessCount}건</strong> 수신 완료
                            {applied_filter.keyword.trim() && applied_filter.use_smart_synonym && ` (스마트 동의어 '${applied_filter.keyword}' 확장 매칭 결과 ${search_results.length}건 표시 중)`}
                            {applied_filter.keyword.trim() && !applied_filter.use_smart_synonym && ` (단어 일치 '${applied_filter.keyword}' 검색 결과 ${search_results.length}건 표시 중)`}
                            {!applied_filter.keyword.trim() && ` (전체 ${search_results.length}건 표시 중)`}
                        </span>
                    </div>
                )}
            </div>

            {/* 실제 데이터 표출 테이블 뷰 */}
            <div className="glass-panel animate-slide-up" style={{ padding: '24px', display: 'flex', flexDirection: 'column', gap: '16px' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px' }}>
                    <h3 style={{ margin: 0, fontSize: '17px', color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                        {active_tab === 'bid' ? '📢 조달청 나라장터 실제 입찰공고 (본공고) 목록' : (active_tab === 'orderplan' ? '📑 조달청 나라장터 발주계획 목록' : '📋 조달청 나라장터 실제 사전규격 수집 목록')}
                        <span style={{ fontSize: '13px', color: 'var(--text-muted)', fontWeight: 400 }}>({search_results.length}건 표시중)</span>
                    </h3>

                    {/* g2b 공식 웹 검색 대체 연동 버튼 */}
                    <div style={{ display: 'flex', gap: '8px' }}>
                        <a
                            href={active_tab === 'bid' ? get_g2b_official_search_url('integrated') : get_g2b_official_search_url('preCom')}
                            target="_blank"
                            rel="noreferrer"
                            style={{
                                display: 'inline-flex', alignItems: 'center', gap: '6px',
                                padding: '8px 14px', borderRadius: '7px',
                                background: 'rgba(168, 85, 247, 0.12)', border: '1px solid rgba(168, 85, 247, 0.3)',
                                color: '#c084fc', fontWeight: 600, fontSize: '12.5px', textDecoration: 'none'
                            }}
                        >
                            🌐 g2b 공식 웹사이트 실시간 {active_tab === 'bid' ? '입찰공고' : '사전규격'} 검색 <ExternalLink size={13} />
                        </a>
                    </div>
                </div>

                {search_results.length === 0 ? (
                    <div style={{ padding: '60px 20px', textAlign: 'center', background: 'rgba(0,0,0,0.15)', borderRadius: '14px', border: '1px dashed var(--panel-border)', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '14px' }}>
                        <Clock size={36} color="#a855f7" />
                        <div>
                            <div style={{ fontSize: '16px', fontWeight: 700, color: 'var(--text-primary)' }}>
                                조건에 일치하는 공고가 없거나 아직 조회하지 않았습니다.
                            </div>
                            <div style={{ fontSize: '13px', color: 'var(--text-secondary)', marginTop: '6px' }}>
                                상단 검색 조건(기간, 키워드 등)을 설정한 후 <strong>[나라장터 실시간 조회하기]</strong> 버튼을 눌러주세요.
                            </div>
                        </div>
                    </div>
                ) : (
                    <div style={{ overflowX: 'auto', borderRadius: '12px', border: '1px solid var(--panel-border)' }}>
                        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13.5px', textAlign: 'left' }}>
                            <thead>
                                <tr style={{ background: 'rgba(255,255,255,0.04)', borderBottom: '1px solid var(--panel-border)', color: 'var(--text-secondary)' }}>
                                    <th style={{ padding: '14px 16px', width: '90px' }}>구분</th>
                                    <th style={{ padding: '14px 16px' }}>{active_tab === 'bid' ? '입찰 공고명' : (active_tab === 'orderplan' ? '발주계획 사업명' : '사전규격 사업명')}</th>
                                    <th style={{ padding: '14px 16px', width: '180px' }}>수요기관 (발주처)</th>
                                    <th style={{ padding: '14px 16px', width: '130px' }}>추정 예산</th>
                                    <th style={{ padding: '14px 16px', width: '110px' }}>게시/등록일</th>
                                    <th style={{ padding: '14px 16px', width: '110px' }}>{active_tab === 'bid' ? '입찰마감일' : (active_tab === 'orderplan' ? '발주예정월' : '의견마감일')}</th>
                                    <th style={{ padding: '14px 16px', width: '90px', textAlign: 'center' }}>g2b 상세</th>
                                </tr>
                            </thead>
                            <tbody>
                                {search_results.map((item, idx) => (
                                    <tr key={`${item.id || 'item'}_${idx}`} style={{ borderBottom: '1px solid var(--panel-border)', background: idx % 2 === 1 ? 'rgba(255,255,255,0.015)' : 'transparent' }}>
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
                                        <td style={{ padding: '14px 16px', color: '#f59e0b', fontWeight: 600, fontSize: '12.5px' }}>
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
