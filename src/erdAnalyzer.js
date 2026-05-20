import { FALLBACK_MODELS } from './utils/geminiModels.js';

export async function analyzeERDWithLLM(documentText, apiKey, onProgress, selectedModel = 'auto', previousResult = null, additionalFeedback = null) {
    const keys = String(apiKey).split(',').map(k => k.trim()).filter(k => k.match(/^(AIza|AQ\.)/));
    if (keys.length === 0) {
        throw new Error("유효한 API 키가 제공되지 않았습니다.");
    }

    let currentKeyIndex = 0;

    // 사용량 기록 유틸리티
    const recordUsage = (modelName) => {
        try {
            const usage = JSON.parse(localStorage.getItem('gemini_model_usage') || '{}');
            usage[modelName] = (usage[modelName] || 0) + 1;
            localStorage.setItem('gemini_model_usage', JSON.stringify(usage));
            window.dispatchEvent(new CustomEvent('gemini_usage_updated'));
        } catch (e) {
            console.error("Usage recording failed:", e);
        }
    };

    const systemPrompt = `당신은 최고 수준의 데이터베이스 설계 전문가이자 데이터 아키텍트입니다.
입력된 비즈니스 요구사항 또는 프로젝트 문서를 분석하여 최적화된 논리적 데이터 모델을 도출하고 Mermaid.js erDiagram 형식으로 시각화 코드를 생성하는 것이 당신의 사명입니다.

[설계 원칙]
1. **정규화 준수**: 기본적으로 제3정규형(3NF)을 목표로 설계하십시오.
2. **엔티티 명명**: 비즈니스 맥락을 명확히 반영하는 영문 엔티티명과 한글 설명을 병기하십시오.
3. **속성 정의**: 모든 엔티티에 적절한 속성(Attribute)을 나열하고 PK(Primary Key)와 FK(Foreign Key)를 명확히 구분하십시오.
4. **관계 정의**: 엔티티 간의 관계 차수(1:1, 1:N, M:N)와 참여 제약 조건(Mandatory/Optional)을 논리적으로 설정하십시오.
5. **반정규화**: 성능상 반드시 필요한 경우에만 제한적으로 적용하고 그 근거를 제시하십시오.

[출력 형식]
반드시 프론트엔드 렌더링을 위해 아래 JSON 구조로만 응답하십시오.
(주의: **프로젝트의 모든 엔티티는 "entities" 배열에 단 하나도 빠짐없이 100% 기재되어야 합니다. 내용이 길어지더라도 절대 중간에 생략하거나 축약하지 마십시오.**)

{
  "entities": [
    {
      "name": "<엔티티 영문명>",
      "description": "<엔티티 한글 설명>",
      "reason": "<해당 엔티티 선정 사유>",
      "attributes": [
        { "name": "<속성 한글명>", "type": "<데이터타입>", "key": "<PK/FK/null>", "desc": "<속성 상세설명>" }
      ]
    }
  ],
  "relationships": [
    {
      "from": "<엔티티A_영문명>",
      "to": "<엔티티B_영문명>",
      "type": "<1:1, 1:N, M:N>",
      "desc": "<관계 설명 및 제약 조건>"
    }
  ],
  "normalizationNotes": "<정규화 준수 여부 및 반정규화 논거에 대한 상세 설명>",
  "summary": "<전체 데이터 모델링 전략 및 설계 방향 요약>",
  "mermaidCode": "<erDiagram으로 시작하는 Mermaid.js 시각화 코드>"
}

[Mermaid 작성 가이드]
- **erDiagram**으로 시작하십시오. (**절대로 Table { ... } 같은 DBML 형식을 섞지 마십시오.**)
- 엔티티 간의 관계는 'EntityA ||--o{ EntityB : "관련성"' 형식을 사용하십시오.
- 엔티티명이나 관계 설명에 공백이 있다면 반드시 " " (큰따옴표)로 감싸십시오.
- 속성이 있는 경우 아래 형식을 따르십시오:
  'EntityName {
    type "한글속성명" PK "상세설명"
  }'
- **중요**: 다이어그램 가독성을 위해 테이블 내 속성명은 가급적 **한글(논리명)**로 작성하십시오.
- **코드 내에 어떤 경우에도 역슬래시(\\) 문자를 포함하지 마십시오.**
- 관계 차수 기호(|o, o|, ||, }o, o{ 등)를 정확히 사용하십시오.

[데이터 안정성 가이드]
- **절대 주의**: JSON 내부의 모든 문자열값(특히 mermaidCode, reason, summary 등)에서 실제 줄바꿈(raw newline)을 사용하지 마십시오. 줄바꿈이 필요한 경우 반드시 \\n (이스케이프 문자)을 사용하십시오.
- **특수 문자 금지**: 엔티티명, 속성명, 설명 내에 역슬래시(\\), 줄바꿈, 탭 문자가 포함되지 않도록 하십시오.
- 모든 필드는 핵심 위주로 명확하고 간결하게 작성하여 전체 응답 길이를 최적화하십시오.
- 응답이 잘리지 않도록 최대 길이에 도달하기 전에 논리적으로 JSON을 마감하십시오.

[주의사항]
- 결과는 반드시 순수 JSON 형태여야 합니다.
- **다시 한 번 강력하게 지시합니다. JSON 구조의 entities 배열에 기입된 항목들은 생략되거나 축약되어서는 안 됩니다.**`;

    const isRefinement = previousResult && additionalFeedback;
    const hasInitialFeedback = !previousResult && additionalFeedback;
    
    let userInput = `
[시스템 지시사항]
${systemPrompt}

[입력 데이터 - 분석 대상 요구사항 문서]
${(documentText || '').substring(0, 15000)}
`;

    if (isRefinement) {
        userInput += `
[이전 설계 결과 (JSON)]
${JSON.stringify(previousResult, null, 2)}

[사용자 추가 요청사항 - 최우선 반영]
${additionalFeedback}

위의 '이전 설계 결과'에 '사용자 추가 요청사항'을 반영하여 설계를 수정하십시오. 
기존 설계의 장점은 유지하되, 요청된 변경 사항을 정확하게 적용하고 관련 엔티티나 관계를 함께 조정하십시오.
결과는 반드시 처음에 정의된 JSON 형식을 완벽하게 유지해야 합니다.
`;
    } else if (hasInitialFeedback) {
        userInput += `
[사용자 추가 요청 및 강조 사항]
${additionalFeedback}

요구사항 문서를 분석할 때 위의 '사용자 추가 요청 및 강조 사항'을 최우선적으로 고려하여 설계를 진행하십시오.
`;
    }

    if (onProgress) {
        onProgress(isRefinement ? "추가 요청사항을 반영하여 설계를 수정 중..." : "데이터베이스 모델 분석 및 ERD 설계 중...");
    }

    try {

        let initialModel = selectedModel && selectedModel !== 'auto' ? selectedModel : FALLBACK_MODELS[0];
        if (!initialModel.startsWith('models/')) initialModel = `models/${initialModel}`;
        
        let currentModelIndex = FALLBACK_MODELS.indexOf(initialModel);
        if (currentModelIndex === -1) currentModelIndex = 0;

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
                            temperature: 0.1, 
                            maxOutputTokens: 16384,
                            responseMimeType: "application/json"
                        }
                    })
                };

                const response = await fetch(fetchUrl, fetchOptions);
                
                if (response.ok) {
                    recordUsage(modelId); // 사용량 기록
                    return response;
                }

                const errData = await response.json().catch(() => ({}));
                const errMsg = errData.error?.message || response.statusText || '';
                const isModelUnavailable = response.status === 404 
                    || response.status === 400
                    || errMsg.toLowerCase().includes('not found')
                    || errMsg.toLowerCase().includes('not supported')
                    || errMsg.toLowerCase().includes('deprecated');

                // 1. 에러 발생 시 항상 다음 API 키를 먼저 시도
                if (keys.length > 1 && (currentKeyIndex + 1) < keys.length) {
                    currentKeyIndex++;
                    const reasonStr = response.status === 429 ? '할당량 초과' : (response.status >= 500 ? '서버 지연' : 'API 오류');
                    if (onProgress) onProgress(`[${reasonStr}] 다음 키로 교체 시도 중 (${currentKeyIndex + 1}/${keys.length})`);
                    continue;
                }

                // 2. 모든 키를 다 썼다면 모델 교체 시도
                if (response.status === 429 || response.status >= 500 || isModelUnavailable) {
                    modelRetries++;
                    if (modelRetries < maxModelRetries) {
                        currentKeyIndex = 0;
                        const nextModelIndex = (currentModelIndex + 1) % FALLBACK_MODELS.length;
                        const nextModelName = FALLBACK_MODELS[nextModelIndex].split('/').pop();
                        const reason = isModelUnavailable && response.status !== 429 ? '모델 미지원' : (response.status === 429 ? '할당량 소진' : '서버 혼잡');
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

        const response = await fetchWithRetry();
        const data = await response.json();
        let content = data.candidates?.[0]?.content?.parts?.[0]?.text || "{}";

        let jsonStr = content.trim();
        // 마크다운 코드블록 제거
        if (jsonStr.includes("```")) {
            jsonStr = jsonStr.replace(/^[\s\S]*?```(?:json|JSON)?\s*/, '');
            jsonStr = jsonStr.replace(/\s*```[\s\S]*$/, '');
        }

        // JSON 정제: char-by-char 스캔으로 문자열 내 비이스케이프 제어문자 처리
        const cleanControlChars = (s) => {
            let result = '';
            let inString = false;
            let escaped = false;
            for (let i = 0; i < s.length; i++) {
                const ch = s[i];
                if (escaped) { result += ch; escaped = false; continue; }
                if (ch === '\\') { escaped = true; result += ch; continue; }
                if (ch === '"') { inString = !inString; result += ch; continue; }
                if (inString) {
                    if (ch === '\n') { result += '\\n'; continue; }
                    if (ch === '\r') { result += '\\r'; continue; }
                    if (ch === '\t') { result += '\\t'; continue; }
                    if (ch.charCodeAt(0) < 32) continue; // 기타 제어문자 제거
                }
                result += ch;
            }
            return result;
        };

        const sanitizeJson = (str) => {
            // 1차: 원본 그대로
            try { return JSON.parse(str); } catch (_) {}

            // 2차: 제어문자 정제 후 파싱
            const cleaned = cleanControlChars(str);
            try { return JSON.parse(cleaned); } catch (_) {}

            // 3차: 후행 쉼표 제거 후 파싱
            try {
                return JSON.parse(cleaned.replace(/,\s*([}\]])/g, '$1'));
            } catch (_) {}

            // 4차: {…} 블록만 추출 후 파싱
            const match = cleaned.match(/\{[\s\S]*\}/);
            if (match) {
                try { return JSON.parse(match[0]); } catch (_) {}
                try { return JSON.parse(match[0].replace(/,\s*([}\]])/g, '$1')); } catch (_) {}
            }

            throw new Error("AI 응답을 JSON으로 파싱할 수 없습니다. 다시 시도해 주세요.");
        };

        return sanitizeJson(jsonStr);
    } catch (e) {
        console.error("ERD Analysis Error:", e);
        throw new Error(`ERD 설계 실패: ${e.message}`);
    }
}
