import { FALLBACK_MODELS } from './geminiModels.js';

function recordUsage(modelName) {
    try {
        const usage = JSON.parse(localStorage.getItem('gemini_model_usage') || '{}');
        usage[modelName] = (usage[modelName] || 0) + 1;
        localStorage.setItem('gemini_model_usage', JSON.stringify(usage));
        window.dispatchEvent(new CustomEvent('gemini_usage_updated'));
    } catch (e) {
        console.error("Usage recording failed:", e);
    }
}

const systemPrompt = `당신은 최고 수준의 프레젠테이션 기획자이자 비즈니스 컨설턴트입니다.
단순히 주어진 텍스트를 옮겨 적는 것이 아니라, 문서의 "진짜 의도와 핵심 가치"를 완벽히 이해하고 재해석하여 전문가 수준의 슬라이드로 구성해야 합니다.

[내용 이해 및 분석 지침]
1. 원본 문서가 여러 슬라이드로 나뉘어 제공된 경우, 각 슬라이드 간의 논리적 연결성(Context)과 문서 전체의 목적을 먼저 파악하세요.
2. 조각난 텍스트라도 문맥을 유추하여 완성된 문장과 의미 있는 핵심 키워드로 복원 및 요약하세요.
3. 문서의 핵심 주제(Core Message)가 무엇인지 파악하고, 불필요한 미사여구는 과감히 생략하되 중요한 기술적/비즈니스적 팩트는 반드시 시각화하여 살리세요.

[슬라이드 구성 원칙]
1. 단순 나열이 아닌, 기승전결이 있는 스토리텔링 구조로 구성하세요.
2. 텍스트 위주의 지루한 구성보다는, 도형과 다이어그램을 활용한 시각화를 최우선으로 고려하세요.
3. 아키텍처, 시스템 구성, 프로세스 흐름, 핵심 기술 요소가 포함된 내용은 각각 ARCHITECTURE_LAYER, PROCESS_FLOW, KEYWORD_HIGHLIGHT 타입을 반드시 사용하여 고도화된 시각화 슬라이드를 생성하세요.

[출력 형식 제한 및 주의사항]
1. 반드시 다음 JSON 형식으로만 응답해야 합니다. 다른 설명이나 텍스트를 절대 포함하지 마세요.
2. JSON 내부의 문자열 값(Value) 안에 큰따옴표(")를 사용할 경우 반드시 이스케이프(\") 처리하거나 홑따옴표(')로 대체하세요.
3. 문자열 내부의 줄바꿈은 실제 엔터가 아닌 \n 으로만 작성해야 합니다. (JSON 파싱 오류 방지)

{
  "theme": "<추천 테마 색상 헥스코드 (예: #003366)>",
  "_overallContextAnalysis": "문서 전체를 관통하는 진짜 핵심 메시지와 기승전결 흐름을 1~2문장으로 요약 (슬라이드를 나누기 전에 전체 맥락을 먼저 여기서 완벽히 이해하고 정의하세요)",
  "slides": [
    {
      "type": "TITLE",
      "title": "<메인 제목>",
      "subtitle": "<부제목 (선택사항)>",
      "author": "<작성자 (선택사항)>"
    },
    {
      "type": "SECTION",
      "title": "<섹션 제목>"
    },
    {
      "type": "BULLET",
      "title": "<슬라이드 제목>",
      "bullets": ["<항목1>", "<항목2>", "<항목3>..."]
    },
    {
      "type": "TWO_COLUMN",
      "title": "<비교/대조 슬라이드 제목>",
      "leftTitle": "<좌측 소제목>",
      "leftBullets": ["<항목1>", "<항목2>"],
      "rightTitle": "<우측 소제목>",
      "rightBullets": ["<항목1>", "<항목2>"]
    },
    {
      "type": "ARCHITECTURE_LAYER",
      "title": "<레이어드 아키텍처 제목>",
      "layers": [
        { "name": "<계층명 (예: Presentation Layer)>", "items": ["<기술/모듈1>", "<기술/모듈2>"] }
      ]
    },
    {
      "type": "PROCESS_FLOW",
      "title": "<업무/데이터 흐름 제목>",
      "steps": [
        { "label": "<단계명>", "desc": "<단계 요약 설명>" }
      ]
    },
    {
      "type": "KEYWORD_HIGHLIGHT",
      "title": "<기술 키워드 강조 제목>",
      "keywords": [
        { "word": "<핵심 키워드>", "desc": "<키워드 설명>" }
      ]
    }
  ]
}`;

export async function analyzePptContent(inputText, emphasisText, inputSlideCount, apiKey, onProgress) {
    const keys = String(apiKey).split(',').map(k => k.trim()).filter(k => k.match(/^(AIza|AQ\.)/));
    if (keys.length === 0) {
        throw new Error("유효한 API 키가 제공되지 않았습니다.");
    }

    let currentKeyIndex = 0;
    let currentModelIndex = 0;

    const emphasisSection = emphasisText ? `\n\n[사용자 특별 강조 요청사항]\n다음 내용을 레이아웃 구성과 색상 및 키워드 선정 시 최우선으로 반영하고 눈에 띄게 강조하세요:\n${emphasisText}` : '';
    const slideConstraint = inputSlideCount > 0 ? `\n\n[중요: 슬라이드 장수 제약]\n원본 문서의 슬라이드 개수가 ${inputSlideCount}장입니다. **반드시 결과물도 정확히 ${inputSlideCount}장**이 되도록 구조화하세요.\n특히 1장일 경우, 모든 텍스트를 나열하지 말고 단 1장의 대시보드/인포그래픽 형태로 압축하여, 흐름도(PROCESS_FLOW)나 구조도(ARCHITECTURE_LAYER) 중심의 초고밀도 시각화 슬라이드를 완성하세요.` : '';
    
    const userInput = `[시스템 지시사항]\n${systemPrompt}${emphasisSection}${slideConstraint}\n\n[분석할 원본 내용]\n${inputText.substring(0, 30000)}`;

    if (onProgress) onProgress("입력된 내용의 시각화 흐름 및 구조 분석 중...");

    const fetchWithRetry = async (maxModelRetries = FALLBACK_MODELS.length) => {
        let modelRetries = 0;
        while (modelRetries < maxModelRetries) {
            const activeKey = keys[currentKeyIndex];
            const modelId = FALLBACK_MODELS[currentModelIndex];
            const fetchUrl = `https://generativelanguage.googleapis.com/v1beta/${modelId}:generateContent?key=${activeKey}`;
            
            const fetchOptions = {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    contents: [{ role: "user", parts: [{ text: userInput }] }],
                    generationConfig: { 
                        temperature: 0.3, 
                        maxOutputTokens: 8192,
                        responseMimeType: "application/json"
                    }
                })
            };

            const response = await fetch(fetchUrl, fetchOptions);
            
            if (response.ok) {
                recordUsage(modelId);
                return response;
            }

            const errData = await response.json().catch(() => ({}));
            const errMsg = errData.error?.message || response.statusText || '';
            const isModelUnavailable = response.status === 404 || response.status === 400
                || errMsg.toLowerCase().includes('not found')
                || errMsg.toLowerCase().includes('not supported')
                || errMsg.toLowerCase().includes('deprecated');

            // 1. 에러 발생 시 항상 다음 API 키를 먼저 시도 (API 키 오류, 429, 500+ 등 대비)
            if (keys.length > 1 && (currentKeyIndex + 1) < keys.length) {
                currentKeyIndex++;
                if (onProgress) onProgress(`API 오류 또는 할당량 초과. 다음 키로 시도 중 (${currentKeyIndex + 1}/${keys.length})`);
                continue;
            }

            // 2. 모든 키를 소진한 경우 다음 모델로 교체
            if (response.status === 429 || response.status >= 500 || isModelUnavailable) {
                modelRetries++;
                if (modelRetries < maxModelRetries) {
                    currentKeyIndex = 0;
                    const nextModelIndex = (currentModelIndex + 1) % FALLBACK_MODELS.length;
                    const nextModelName = FALLBACK_MODELS[nextModelIndex].split('/').pop();
                    const reason = response.status === 429 ? '할당량 초과' : (response.status >= 500 ? '서버 오류' : '모델 미지원/오류');
                    const currentModelName = modelId.split('/').pop();
                    
                    if (onProgress) onProgress(`[${reason}] [${currentModelName}] 소진 → 5초 후 [${nextModelName}]으로 변경하여 재시도합니다.`);
                    
                    await new Promise(resolve => setTimeout(resolve, 5000));
                    currentModelIndex = nextModelIndex;
                    continue;
                }
                
                throw new Error("모든 API 키와 모델의 사용 한도가 소진되었습니다.");
            }

            throw new Error(errMsg || response.statusText);
        }
        throw new Error("모든 모델을 시도했으나 응답을 받지 못했습니다.");
    };

    try {
        const response = await fetchWithRetry();
        const data = await response.json();
        let content = data.candidates?.[0]?.content?.parts?.[0]?.text || "{}";
        
        let jsonStr = content.trim();
        if (jsonStr.includes("\`\`\`")) {
            jsonStr = jsonStr.replace(/^[\s\S]*?\`\`\`(?:json|JSON)?\s*/, '');
            jsonStr = jsonStr.replace(/\s*\`\`\`[\s\S]*$/, '');
        }

        return JSON.parse(jsonStr);
    } catch (e) {
        console.error("PPT Analysis Error:", e);
        if (e.message.includes("Unterminated string") || e.message.includes("Unexpected token") || e.message.includes("JSON")) {
            throw new Error(`AI가 생성한 데이터 형식이 올바르지 않습니다. (JSON 파싱 오류)\n텍스트가 너무 길어 중간에 끊겼거나 특수문자 충돌일 수 있습니다. 다시 한번 시도해주세요.\n상세오류: ${e.message}`);
        }
        throw new Error(`PPT 구조화 실패: ${e.message}`);
    }
}
