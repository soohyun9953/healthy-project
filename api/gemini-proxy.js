// api/gemini-proxy.js
// Vercel 서버리스 함수: Gemini generateContent 호출 프록시.
// 브라우저가 Gemini API를 직접 호출하면 API 키가 URL 쿼리스트링(?key=...)에 그대로 남아
// 브라우저 히스토리/리퍼러/네트워크 로그에 노출된다. 이 프록시는 클라이언트가 보낸 키를
// 요청 헤더(x-goog-api-key)로만 Gemini에 전달하고, 서버에 키를 저장하지 않는다(BYOK 유지).

export default async function handler(req, res) {
    if (req.method !== 'POST') {
        return res.status(405).json({ error: 'POST 요청만 허용됩니다.' });
    }

    const { apiKey, modelId, contents, generationConfig } = req.body || {};

    if (!apiKey || !/^(AIza|AQ\.)/.test(String(apiKey))) {
        return res.status(400).json({ error: '유효한 Gemini API 키 형식이 아닙니다.' });
    }
    if (!modelId || !Array.isArray(contents)) {
        return res.status(400).json({ error: 'modelId와 contents가 필요합니다.' });
    }

    const normalizedModel = String(modelId).startsWith('models/') ? modelId : `models/${modelId}`;
    const targetUrl = `https://generativelanguage.googleapis.com/v1beta/${normalizedModel}:generateContent`;

    try {
        const upstream = await fetch(targetUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'x-goog-api-key': apiKey
            },
            body: JSON.stringify({
                contents,
                generationConfig: generationConfig || { temperature: 0.1 }
            })
        });

        const bodyText = await upstream.text();
        res.setHeader('Content-Type', 'application/json; charset=utf-8');
        return res.status(upstream.status).send(bodyText);
    } catch (err) {
        console.error('[Gemini Proxy] 호출 오류:', err);
        return res.status(502).json({ error: `Gemini 호출 실패: ${err.message}` });
    }
}
