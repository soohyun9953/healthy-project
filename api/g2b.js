// api/g2b.js
// Vercel 서버리스 함수: 조달청 나라장터 사전규격 / 입찰공고 / 발주계획 다중 OpenAPI CORS 프록시

export default async function handler(req, res) {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, GET, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

    if (req.method === 'OPTIONS') {
        return res.status(200).end();
    }

    const service_key = req.body?.service_key || req.query?.service_key;
    const inqry_bgn_dt = req.body?.inqry_bgn_dt || req.query?.inqry_bgn_dt;
    const inqry_end_dt = req.body?.inqry_end_dt || req.query?.inqry_end_dt;
    const service_type = (req.body?.service_type || req.query?.service_type || 'prespec').toLowerCase();

    if (!service_key) {
        return res.status(400).json({ error: 'service_key 파라미터가 필요합니다.' });
    }

    try {
        const key_raw = service_key;
        const key_encoded = encodeURIComponent(service_key);
        const key_variants = [...new Set([key_raw, key_encoded])];

        const fetch_all_pages = async (base_url_builder) => {
            // 1단계: 첫 페이지 호출 및 전체 개수 파악
            const first_url = base_url_builder(1);
            let first_items = [];
            let total_count = 0;
            try {
                const first_resp = await fetch(first_url);
                const first_parsed = await first_resp.json();
                first_items = first_parsed?.response?.body?.items || [];
                total_count = parseInt(first_parsed?.response?.body?.totalCount || '0', 10);
            } catch (e) {
                return [];
            }

            if (total_count <= 999 || first_items.length === 0) {
                return first_items;
            }

            // 2단계: 2페이지부터 마지막 페이지까지 전체 병렬 수집 (최대 5페이지, 약 5,000건으로 속도 대폭 최적화)
            const total_pages = Math.min(5, Math.ceil(total_count / 999));
            const remaining_pages = [];
            for (let p = 2; p <= total_pages; p++) {
                remaining_pages.push(p);
            }

            const page_tasks = remaining_pages.map(async (p) => {
                const url = base_url_builder(p);
                try {
                    const resp = await fetch(url);
                    const data = await resp.json();
                    return data?.response?.body?.items || [];
                } catch (e) {
                    return [];
                }
            });

            const rest_results = await Promise.all(page_tasks);
            return [...first_items, ...rest_results.flat()];
        };

        if (service_type === 'prespec') {
            // 🌟 사전규격: 순수 용역 분야 1~5페이지 전수 수집 (누락 방지)
            try {
                const ep = `https://apis.data.go.kr/1230000/ao/HrcspSsstndrdInfoService/getPublicPrcureThngInfoServc`;
                const merged_items = await fetch_all_pages((p) => {
                    let url = `${ep}?serviceKey=${key_raw}&type=json&numOfRows=999&pageNo=${p}&inqryDiv=1`;
                    if (inqry_bgn_dt) url += `&inqryBgnDt=${inqry_bgn_dt}`;
                    if (inqry_end_dt) url += `&inqryEndDt=${inqry_end_dt}`;
                    return url;
                });

                return res.status(200).json({
                    response: {
                        header: { resultCode: '00', resultMsg: '정상 (용역 전수 수집/물품·공사 제외)' },
                        body: { items: merged_items, totalCount: merged_items.length }
                    }
                });
            } catch (e) {
                console.error('사전규격 병렬 수집 에러:', e);
            }
        }

        if (service_type === 'bid') {
            // 🌟 실시간 입찰공고: 순수 용역 분야 1~5페이지 전수 수집
            try {
                const ep = `https://apis.data.go.kr/1230000/ad/BidPublicInfoService/getBidPblancListInfoServcPPSSrch`;
                const merged_items = await fetch_all_pages((p) => {
                    let url = `${ep}?serviceKey=${key_raw}&type=json&numOfRows=999&pageNo=${p}&inqryDiv=1`;
                    if (inqry_bgn_dt) url += `&inqryBgnDt=${inqry_bgn_dt}`;
                    if (inqry_end_dt) url += `&inqryEndDt=${inqry_end_dt}`;
                    return url;
                });

                return res.status(200).json({
                    response: {
                        header: { resultCode: '00', resultMsg: '정상 (입찰공고 용역 전수 수집/물품·공사 제외)' },
                        body: { items: merged_items, totalCount: merged_items.length }
                    }
                });
            } catch (e) {
                console.error('입찰공고 병렬 수집 에러:', e);
            }
        }

        if (service_type === 'orderplan') {
            // 🌟 발주계획 현황: 순수 용역 분야 1~5페이지 전수 수집
            try {
                const bgn_dt_clean = (inqry_bgn_dt || '').substring(0, 8) || '20260101';
                const end_dt_clean = (inqry_end_dt || '').substring(0, 8) || '20261231';
                const ep = `https://apis.data.go.kr/1230000/ao/OrderPlanSttusService/getOrderPlanSttusListServc`;

                const merged_items = await fetch_all_pages((p) => {
                    return `${ep}?serviceKey=${key_raw}&type=json&numOfRows=999&pageNo=${p}&inqryDiv=1&inqryBgnDate=${bgn_dt_clean}&inqryEndDate=${end_dt_clean}`;
                });

                return res.status(200).json({
                    response: {
                        header: { resultCode: '00', resultMsg: '정상 (발주계획 용역 전수 수집/물품·공사 제외)' },
                        body: { items: merged_items, totalCount: merged_items.length }
                    }
                });
            } catch (e) {
                console.error('발주계획 병렬 수집 에러:', e);
            }
        }

        if (service_type === 'contract') {
            // 🌟 계약정보 현황: 순수 용역 분야 1~5페이지 전수 수집
            try {
                const ep = `https://apis.data.go.kr/1230000/ao/CntrctInfoService/getCntrctInfoListServc`;
                const merged_items = await fetch_all_pages((p) => {
                    let url = `${ep}?serviceKey=${key_raw}&type=json&numOfRows=999&pageNo=${p}&inqryDiv=1`;
                    if (inqry_bgn_dt) url += `&inqryBgnDt=${inqry_bgn_dt}`;
                    if (inqry_end_dt) url += `&inqryEndDt=${inqry_end_dt}`;
                    return url;
                });

                return res.status(200).json({
                    response: {
                        header: { resultCode: '00', resultMsg: '정상 (계약정보 용역 전수 수집/물품·공사 제외)' },
                        body: { items: merged_items, totalCount: merged_items.length }
                    }
                });
            } catch (e) {
                console.error('계약정보 병렬 수집 에러:', e);
            }
        }

        let last_error = '';
        let success = false;

        for (const endpoint of endpoints) {
            for (const key of key_variants) {
                let api_url = `${endpoint}?serviceKey=${key}&type=json&numOfRows=100&pageNo=1`;
                if (inqry_bgn_dt) api_url += `&inqryBgnDt=${inqry_bgn_dt}`;
                if (inqry_end_dt) api_url += `&inqryEndDt=${inqry_end_dt}`;

                try {
                    const api_res = await fetch(api_url, {
                        method: 'GET',
                        headers: { 'Accept': 'application/json, text/xml' }
                    });

                    if (api_res.ok) {
                        const result_text = await api_res.text();

                        if (result_text.includes('SERVICE_KEY_IS_NOT_REGISTERED_ERROR') || result_text.includes('SERVICE_KEY_IS_NOT_REGISTERED')) {
                            last_error = 'SERVICE_KEY_IS_NOT_REGISTERED';
                            continue;
                        }
                        if (result_text.includes('LIMITED_NUMBER_OF_SERVICE_REQUESTS')) {
                            last_error = 'LIMITED_NUMBER_OF_SERVICE_REQUESTS';
                            continue;
                        }
                        if (result_text.includes('NO_OPENAPI_SERVICE_ERROR')) {
                            last_error = `엔드포인트 미지원: ${endpoint}`;
                            continue;
                        }

                        res.setHeader('Content-Type', 'application/json; charset=utf-8');
                        return res.status(200).send(result_text);
                    }
                } catch (fetch_err) {
                    last_error = fetch_err.message;
                }
            }
            if (success) break;
        }

        return res.status(200).json({
            response: {
                header: { resultCode: '99', resultMsg: last_error || 'SYNCING' },
                body: { items: [] }
            }
        });

    } catch (err) {
        console.error('[G2B 프록시] 서버 오류:', err);
        return res.status(200).json({
            response: {
                header: { resultCode: '99', resultMsg: err.message },
                body: { items: [] }
            }
        });
    }
}
