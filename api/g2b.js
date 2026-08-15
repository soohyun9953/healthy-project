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

        let endpoints = [];
        if (service_type === 'bid') {
            // 입찰공고
            endpoints = [
                `https://apis.data.go.kr/1230000/ao/PubDataOpnStdService/getDataSetOpnStdBidPblancInfo`,
                `https://apis.data.go.kr/1230000/BidPublicInfoService02/getBidPblancListInfoServc`,
                `https://apis.data.go.kr/1230000/BidPublicInfoService02/getBidPblancListInfoServcPPSSrch`
            ];
        } else if (service_type === 'orderplan') {
            // 발주계획
            endpoints = [
                `https://apis.data.go.kr/1230000/ao/OrderPlanSttusService/getOrderPlanSttusList`,
                `https://apis.data.go.kr/1230000/OrderPlanSttusService/getOrderPlanSttusList`
            ];
        } else {
            // 사전규격 (기본)
            endpoints = [
                `https://apis.data.go.kr/1230000/ao/HrcspSsstndrdInfoService/getPublicPrcureThngInfoServcPPSSrch`,
                `https://apis.data.go.kr/1230000/ao/HrcspSsstndrdInfoService/getPublicPrcureThngInfoServc`,
                `https://apis.data.go.kr/1230000/ao/HrcspSsstndrdInfoService/getPublicPrcureThngInfoThngPPSSrch`,
                `https://apis.data.go.kr/1230000/ao/HrcspSsstndrdInfoService/getPublicPrcureThngInfoThng`
            ];
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
