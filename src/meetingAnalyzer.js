// ─────────────────────────────────────────────
//  meetingAnalyzer.js
//  Gemini Multimodal Audio → 회의록 생성 파이프라인
// ─────────────────────────────────────────────

import { FALLBACK_MODELS } from './utils/geminiModels.js';

const INLINE_SIZE_LIMIT = 20 * 1024 * 1024; // 20MB

// ── 사용량 기록 ──────────────────────────────
function recordUsage(modelName) {
    try {
        const usage = JSON.parse(localStorage.getItem('gemini_model_usage') || '{}');
        usage[modelName] = (usage[modelName] || 0) + 1;
        localStorage.setItem('gemini_model_usage', JSON.stringify(usage));
        window.dispatchEvent(new CustomEvent('gemini_usage_updated'));
    } catch (e) {
        console.error('Usage recording failed:', e);
    }
}

// ── File → Base64 ────────────────────────────
function fileToBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => {
            const result = reader.result; // data:audio/xxx;base64,XXXX
            const base64 = result.split(',')[1];
            resolve(base64);
        };
        reader.onerror = reject;
        reader.readAsDataURL(file);
    });
}

// ── MIME 타입 매핑 ───────────────────────────
function getAudioMimeType(file) {
    if (file.type) return file.type;
    const ext = file.name.split('.').pop().toLowerCase();
    const map = {
        mp3: 'audio/mpeg',
        wav: 'audio/wav',
        m4a: 'audio/mp4',
        ogg: 'audio/ogg',
        flac: 'audio/flac',
        aac: 'audio/aac',
        webm: 'audio/webm',
        mp4: 'audio/mp4',
    };
    return map[ext] || 'audio/mpeg';
}

// ── Files API 업로드 (대용량) ─────────────────
async function uploadViaFilesAPI(file, apiKey, onProgress) {
    if (onProgress) onProgress('대용량 파일 감지 → Files API로 업로드 중...');

    const mimeType = getAudioMimeType(file);

    // 1. 업로드 세션 초기화
    const initRes = await fetch(
        `https://generativelanguage.googleapis.com/upload/v1beta/files?key=${apiKey}`,
        {
            method: 'POST',
            headers: {
                'X-Goog-Upload-Protocol': 'resumable',
                'X-Goog-Upload-Command': 'start',
                'X-Goog-Upload-Header-Content-Length': file.size,
                'X-Goog-Upload-Header-Content-Type': mimeType,
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({ file: { display_name: file.name } }),
        }
    );

    if (!initRes.ok) {
        const err = await initRes.json().catch(() => ({}));
        throw new Error(`Files API 초기화 실패: ${err.error?.message || initRes.statusText}`);
    }

    const uploadUrl = initRes.headers.get('X-Goog-Upload-URL');
    if (!uploadUrl) throw new Error('업로드 URL을 받지 못했습니다.');

    // 2. 파일 업로드
    if (onProgress) onProgress(`파일 업로드 중... (${(file.size / 1024 / 1024).toFixed(1)}MB)`);

    const uploadRes = await fetch(uploadUrl, {
        method: 'POST',
        headers: {
            'X-Goog-Upload-Command': 'upload, finalize',
            'X-Goog-Upload-Offset': '0',
            'Content-Type': mimeType,
        },
        body: file,
    });

    if (!uploadRes.ok) {
        const err = await uploadRes.json().catch(() => ({}));
        throw new Error(`파일 업로드 실패: ${err.error?.message || uploadRes.statusText}`);
    }

    const fileData = await uploadRes.json();
    const fileUri = fileData.file?.uri;
    const uploadedMime = fileData.file?.mimeType || mimeType;

    if (!fileUri) throw new Error('파일 URI를 받지 못했습니다.');
    if (onProgress) onProgress('파일 업로드 완료. AI 분석 시작...');

    return { fileUri, mimeType: uploadedMime };
}

// ── 메인 프롬프트 생성 ────────────────────────
function buildPrompt(terminology) {
    const termSection = terminology && terminology.length > 0
        ? `\n\n[전문 용어 사전 - 반드시 아래 표기법을 사용하여 변환할 것]\n${terminology.map(t => `- ${t.word}: ${t.desc}`).join('\n')}`
        : '';

    return `당신은 회의 내용을 분석하여 전문적이고 충실한 회의록을 작성하는 AI 전문가입니다.
회의 내용을 단순히 나열하는 것이 아니라, 맥락을 파악하고 핵심 인사이트와 향후 계획까지 심층 도출해야 합니다.

[처리 파이프라인 - 반드시 순서대로 수행]
1단계 - 오타 및 비문 수정: 음성에서 생성된 텍스트의 모든 오타, 문법 오류, 어색한 표현을 자연스러운 문어체로 교정하라.
2단계 - 화자 분류: 음성 톤, 문맥, 발언 패턴을 분석하여 화자를 A, B, C... 순으로 분류하라. 화자가 동시에 말하거나 발언이 섞인 경우, 문맥상 가장 개연성 있는 화자에 귀속시키고 [중복발언] 태그를 붙여라.
3단계 - 맥락 파악: 회의의 목적, 배경, 현재 상황(이슈/문제점)을 파악하라.
4단계 - 주제별 심층 요약: 회의에서 논의된 주요 주제를 2~6개로 분류하고, 각 주제에 대해 논의 내용, 쟁점, 합의된 방향을 충분히 요약하라.
5단계 - 핵심 인사이트 도출: 회의에서 도출된 중요한 통찰, 리스크, 기회 요인을 추출하라.
6단계 - 향후 계획 정리: 단기 실행 계획(즉시~1개월 이내)과 중장기 계획(1개월 이상)을 구분하여 정리하라.
7단계 - 결정사항 및 액션아이템 추출: 회의 안건, 결정된 사항, 구체적인 액션 아이템(담당자·기한 포함)을 추출하라.${termSection}

[예외 처리 규칙]
- 전문 용어: 용어 사전에 등록된 단어는 반드시 사전의 표기법을 따른다.
- 불명확한 발언: 음질 문제 등으로 인식 불가한 부분은 [불명확] 태그를 붙인다.
- 화자 겹침: 문맥을 분석하여 가장 개연성 있는 발언자에 귀속, [중복발언] 태그 표시.
- 담당자 추론: 액션 아이템의 담당자가 명확히 언급되지 않은 경우, 문맥상 가장 관련 있는 화자를 추론하여 기재하라.
- 향후 계획 추론: 명시적으로 계획이 언급되지 않아도 회의 맥락에서 추론 가능한 다음 단계를 기재하라.

[출력 형식 - 반드시 아래 JSON 형식으로만 출력. 다른 텍스트 금지]
{
  "meetingTitle": "<회의 내용을 바탕으로 AI가 추론한 회의 제목>",
  "meetingContext": "<회의 목적, 배경, 현재 상황(문제점·이슈)을 3~5문장으로 서술. 왜 이 회의가 열렸는지 설명>",
  "agenda": "<회의 안건을 2~4문장으로 구체적으로 요약>",
  "speakerCount": <화자 수 (숫자)>,
  "summary": "<전체 회의 핵심 내용을 4~6문장으로 충분히 요약. 단순 나열이 아닌 논의의 흐름과 결론 중심으로 서술>",
  "topicSummaries": [
    {
      "topic": "<주제 제목>",
      "content": "<해당 주제에 대한 논의 내용, 쟁점, 합의 방향 등을 3~5문장으로 상세 요약>",
      "result": "<이 주제의 최종 결론 또는 결정 방향 (1~2문장)>"
    }
  ],
  "keyInsights": [
    { "type": "<인사이트 유형: 기회|리스크|통찰|우선순위 중 택1>", "content": "<핵심 인사이트 내용>" }
  ],
  "decisions": ["<결정 사항 1>", "<결정 사항 2>"],
  "actionItems": [
    { "task": "<구체적인 할 일>", "owner": "<담당자 화자 또는 이름>", "deadline": "<언급된 기한, 없으면 null>", "priority": "<높음|보통|낮음 중 택1>" }
  ],
  "futurePlans": {
    "shortTerm": [
      { "plan": "<단기 계획 (즉시~1개월 이내)", "owner": "<담당자 또는 팀, 없으면 null>", "targetDate": "<목표 시점, 없으면 null>" }
    ],
    "longTerm": [
      { "plan": "<중장기 계획 (1개월 이상)", "owner": "<담당자 또는 팀, 없으면 null>", "targetDate": "<목표 시점, 없으면 null>" }
    ]
  },
  "keywords": ["<핵심 키워드 5~10개>"],
  "transcript": [
    { "speaker": "<화자 레이블 예: A>", "text": "<교정된 발언 내용>", "tag": "<정상|중복발언|불명확 중 택1>" }
  ]
}`;
}

// ── 메인 분석 함수 ────────────────────────────
export async function analyzeMeeting(inputData, inputType = 'audio', apiKey, terminology = [], onProgress) {
    const keys = String(apiKey).split(',').map(k => k.trim()).filter(k => k.match(/^(AIza|AQ\.)/));
    if (keys.length === 0) throw new Error('유효한 API 키가 없습니다.');

    let currentKeyIndex = 0;
    let currentModelIndex = 0;
    const prompt = buildPrompt(terminology);

    // ── 입력 페이로드 타입 처리 ─────────────
    let contentParts = [];

    if (inputType === 'text') {
        contentParts = [
            { text: "[회의 원본 텍스트]\n" + inputData },
            { text: prompt }
        ];
    } else {
        const mimeType = getAudioMimeType(inputData);
        let audioPart;
        if (inputData.size <= INLINE_SIZE_LIMIT) {
            if (onProgress) onProgress('오디오 파일 인코딩 중...');
            const base64 = await fileToBase64(inputData);
            audioPart = { inlineData: { mimeType, data: base64 } };
        } else {
            // 대용량: Files API 사용 (첫 번째 키 사용)
            const { fileUri, mimeType: uploadedMime } = await uploadViaFilesAPI(inputData, keys[0], onProgress);
            audioPart = { fileData: { mimeType: uploadedMime, fileUri } };
        }
        contentParts = [ audioPart, { text: prompt } ];
    }

    // ── Gemini API 요청 (모델/키 폴백) ──────────
    const fetchWithRetry = async () => {
        let modelRetries = 0;
        const maxRetries = FALLBACK_MODELS.length;

        while (modelRetries < maxRetries) {
            const activeKey = keys[currentKeyIndex];
            const modelId = FALLBACK_MODELS[currentModelIndex];
            const url = `https://generativelanguage.googleapis.com/v1beta/${modelId}:generateContent?key=${activeKey}`;

            let fetchTimer = null;
            if (onProgress) {
                const keyInfo = keys.length > 1 ? ` (키 ${currentKeyIndex + 1}/${keys.length})` : '';
                let elapsed = 0;
                onProgress(`${modelId.split('/').pop()} 로 분석 중...${keyInfo}`);
                fetchTimer = setInterval(() => {
                    elapsed++;
                    onProgress(`${modelId.split('/').pop()} 로 분석 중...${keyInfo} (${elapsed}초 경과)`);
                }, 1000);
            }

            const body = {
                contents: [{
                    role: 'user',
                    parts: contentParts
                }],
                generationConfig: { temperature: 0.2 }
            };

            let res;
            try {
                res = await fetch(url, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(body),
                });
            } finally {
                if (fetchTimer) clearInterval(fetchTimer);
            }

            if (res.ok) {
                if (onProgress) onProgress('분석 완료, 결과 처리 중...');
                recordUsage(modelId);
                return res;
            }

            const errData = await res.json().catch(() => ({}));
            const errMsg = errData.error?.message || res.statusText || '';
            const isModelUnsupported = res.status === 404 || res.status === 400
                || errMsg.toLowerCase().includes('not found')
                || errMsg.toLowerCase().includes('not supported')
                || errMsg.toLowerCase().includes('deprecated');

            const isOverloaded = res.status === 429 || res.status >= 500
                || errMsg.toLowerCase().includes('high demand')
                || errMsg.toLowerCase().includes('overloaded');

            // 1. 에러 발생 시 항상 다음 API 키를 먼저 시도
            if (keys.length > 1 && (currentKeyIndex + 1) < keys.length) {
                currentKeyIndex++;
                const reasonStr = res.status === 429 ? '할당량 초과' : (res.status >= 500 ? '서버 지연' : 'API 오류');
                if (onProgress) onProgress(`[${reasonStr}] 다음 키로 전환 (${currentKeyIndex + 1}/${keys.length})`);
                continue;
            }
            
            // 2. 모든 키를 다 썼다면 모델 교체 시도
            if (isOverloaded || isModelUnsupported) {
                modelRetries++;
                if (modelRetries < maxRetries) {
                    currentKeyIndex = 0;
                    currentModelIndex = (currentModelIndex + 1) % FALLBACK_MODELS.length;
                    const nextModel = FALLBACK_MODELS[currentModelIndex].split('/').pop();
                    const reason = isModelUnsupported ? '미지원 모델/오류' : (res.status === 429 ? '할당량 소진' : '서버 혼잡');
                    const currentModelName = modelId.split('/').pop();
                    if (onProgress) onProgress(`[${reason}] [${currentModelName}] 소진 → 5초 후 [${nextModel}](으)로 재시도...`);
                    await new Promise(r => setTimeout(r, 5000));
                    continue;
                }
                throw new Error('모든 API 키와 모델을 시도했으나 실패했습니다.');
            }

            throw new Error(errMsg || res.statusText);
        }
        throw new Error('모든 모델 시도 후 응답 없음');
    };

    if (onProgress) onProgress('Gemini AI 회의록 분석 중...');
    const response = await fetchWithRetry();
    const data = await response.json();

    let content = data.candidates?.[0]?.content?.parts?.[0]?.text || '{}';

    // ```json ... ``` 블록 제거
    if (content.includes('```')) {
        const match = content.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
        if (match?.[1]) content = match[1];
        else content = content.replace(/```(json)?/g, '').trim(); // 잘려서 안 닫힌 경우
    }

    try {
        return JSON.parse(content);
    } catch {
        // JSON 파싱 실패 시, 문자열 끝이 잘린 경우를 대비한 점진적 복구
        const fixAttempts = [
            content + ']', content + '}', content + ']}',
            content + '"]}', content + '}]}', content + '"]}]}',
            content + '"}'
        ];
        for (const fixed of fixAttempts) {
            try { return JSON.parse(fixed); } catch {}
        }
        
        // 정규식으로 직접 추출 (Fallback)
        let result = {
            meetingTitle: '회의록',
            meetingContext: '',
            agenda: '분석 결과를 파싱하지 못했습니다.',
            speakerCount: 0,
            summary: 'JSON 결과가 잘려 부분적으로 추출되었습니다.\n' + content.substring(0, 500) + '...',
            topicSummaries: [],
            keyInsights: [],
            decisions: [],
            actionItems: [],
            futurePlans: { shortTerm: [], longTerm: [] },
            keywords: [],
            transcript: []
        };
        
        const mTitle = content.match(/"meetingTitle"\s*:\s*"((?:[^"\\]|\\.)*)"/i);
        if (mTitle) result.meetingTitle = mTitle[1].replace(/\\"/g, '"');
        
        const mAgenda = content.match(/"agenda"\s*:\s*"((?:[^"\\]|\\.)*)"/i);
        if (mAgenda) result.agenda = mAgenda[1].replace(/\\"/g, '"');
        
        const mSummary = content.match(/"summary"\s*:\s*"((?:[^"\\]|\\.)*)"/i);
        if (mSummary) result.summary = mSummary[1].replace(/\\"/g, '"');
        
        const mSpeaker = content.match(/"speakerCount"\s*:\s*(\d+)/i);
        if (mSpeaker) result.speakerCount = parseInt(mSpeaker[1], 10);
        
        const dMatch = content.match(/"decisions"\s*:\s*\[(.*?)\]/s);
        if (dMatch) {
            const items = dMatch[1].match(/"((?:[^"\\]|\\.)*)"/g);
            if (items) result.decisions = items.map(s => s.replace(/^"|"$/g, '').replace(/\\"/g, '"'));
        }
        
        const aMatches = content.matchAll(/"task"\s*:\s*"((?:[^"\\]|\\.)*)",\s*"owner"\s*:\s*"((?:[^"\\]|\\.)*)",\s*"deadline"\s*:\s*"((?:[^"\\]|\\.)*)"/g);
        for (const m of aMatches) {
            result.actionItems.push({ 
                task: m[1].replace(/\\"/g, '"'), 
                owner: m[2].replace(/\\"/g, '"'), 
                deadline: m[3].replace(/\\"/g, '"') 
            });
        }
        
        const tMatches = content.matchAll(/"speaker"\s*:\s*"([^"]+)",\s*"text"\s*:\s*"((?:[^"\\]|\\.)*)",\s*"tag"\s*:\s*"([^"]+)"/g);
        for (const m of tMatches) {
            result.transcript.push({ speaker: m[1], text: m[2].replace(/\\"/g, '"'), tag: m[3] });
        }
        
        return result;
    }
}

// ─────────────────────────────────────────────
//  회의 내용 대상 AI 질의응답 (Q&A) 파이프라인
// ─────────────────────────────────────────────
export async function askMeetingQuestion(question, contextData, apiKey, chatHistory = []) {
    const keys = String(apiKey).split(',').map(k => k.trim()).filter(k => k.match(/^(AIza|AQ\.)/));
    if (keys.length === 0) throw new Error('유효한 Gemini API 키가 없습니다. 설정에서 API 키를 입력해 주세요.');

    let contextText = '';
    if (typeof contextData === 'string') {
        contextText = contextData;
    } else if (contextData && typeof contextData === 'object') {
        const sections = [];
        if (contextData.meetingTitle) sections.push(`[회의명]\n${contextData.meetingTitle}`);
        if (contextData.meetingContext) sections.push(`[회의 목적 및 배경]\n${contextData.meetingContext}`);
        if (contextData.agenda) sections.push(`[회의 안건]\n${contextData.agenda}`);
        if (contextData.summary) sections.push(`[종합 요약]\n${contextData.summary}`);
        
        if ((contextData.topicSummaries || []).length > 0) {
            sections.push(`[주제별 상세 논의 내용]\n` + contextData.topicSummaries.map((t, i) => `주제 ${i + 1}. ${t.topic}\n- 내용: ${t.content}${t.result ? `\n- 결론: ${t.result}` : ''}`).join('\n\n'));
        }
        if ((contextData.keyInsights || []).length > 0) {
            sections.push(`[핵심 인사이트]\n` + contextData.keyInsights.map((ins, i) => `${i + 1}. [${ins.type}] ${ins.content}`).join('\n'));
        }
        if ((contextData.decisions || []).length > 0) {
            sections.push(`[결정된 사항]\n` + contextData.decisions.map((d, i) => `- ${d}`).join('\n'));
        }
        if ((contextData.actionItems || []).length > 0) {
            sections.push(`[할 일(Action Items) 및 담당자]\n` + contextData.actionItems.map((a, i) => `${i + 1}. 과제: ${a.task} | 담당자: ${a.owner || '미정'} | 기한: ${a.deadline || '미정'} | 우선순위: ${a.priority || '보통'}`).join('\n'));
        }
        if (contextData.futurePlans) {
            const fp = contextData.futurePlans;
            const fpLines = [];
            if ((fp.shortTerm || []).length > 0) fpLines.push('단기 계획: ' + fp.shortTerm.map(p => `${p.plan} (담당: ${p.owner || '-'}, 목표: ${p.targetDate || '-'})`).join(', '));
            if ((fp.longTerm || []).length > 0) fpLines.push('중장기 계획: ' + fp.longTerm.map(p => `${p.plan} (담당: ${p.owner || '-'}, 목표: ${p.targetDate || '-'})`).join(', '));
            if (fpLines.length > 0) sections.push(`[향후 계획]\n` + fpLines.join('\n'));
        }
        if ((contextData.transcript || []).length > 0) {
            sections.push(`[화자별 발언록 상세 (Transcript)]\n` + contextData.transcript.map(t => `[${t.speaker}] ${t.text}`).join('\n'));
        }
        contextText = sections.join('\n\n====================\n\n');
    }

    const systemInstruction = `당신은 회의록 분석 및 질의응답을 수행하는 전문 AI 어시스턴트입니다.
제공된 [회의 내용 및 상세 발언록]을 바탕으로 사용자의 질문에 정확하고 구체적이며 친절하게 답변하세요.

[답변 원칙]
1. 회의 내용에 근거하여 사실에 입각해 명확히 답변하세요.
2. 회의에서 언급된 발언자, 결정사항, 과제 담당자, 일정, 리스크, 핵심 쟁점 등을 충실히 인용하세요.
3. 회의 내용에 명시되지 않은 추측이나 외부 지식은 "회의 내용에는 언급되지 않았으나"와 같이 구분하여 안내하세요.
4. 가독성을 위해 개조식(•, 번호)과 볼드체(**)를 적절히 활용하여 깔끔하게 정리하세요.`;

    const contents = [];

    // 이전 대화 히스토리 구성
    if (chatHistory && chatHistory.length > 0) {
        chatHistory.forEach(msg => {
            contents.push({
                role: msg.role === 'user' ? 'user' : 'model',
                parts: [{ text: msg.text }]
            });
        });
    }

    // 현재 사용자 질문 (컨텍스트 포함)
    const currentPrompt = `[회의 내용 및 상세 발언록]
${contextText.substring(0, 100000)}

[질문]
${question}`;

    contents.push({
        role: 'user',
        parts: [{ text: currentPrompt }]
    });

    const MODELS = FALLBACK_MODELS;
    let currentKeyIndex = 0;
    let currentModelIndex = 0;
    let lastError = null;

    while (currentModelIndex < MODELS.length) {
        const modelName = MODELS[currentModelIndex];
        const activeKey = keys[currentKeyIndex];
        const url = `https://generativelanguage.googleapis.com/v1beta/${modelName}:generateContent?key=${activeKey}`;

        try {
            const res = await fetch(url, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    contents: contents,
                    systemInstruction: {
                        parts: [{ text: systemInstruction }]
                    },
                    generationConfig: {
                        temperature: 0.2,
                        maxOutputTokens: 2048,
                    }
                })
            });

            if (res.ok) {
                const data = await res.json();
                const answer = data.candidates?.[0]?.content?.parts?.[0]?.text;
                if (answer) {
                    recordUsage(modelName);
                    return {
                        answer: answer.trim(),
                        modelUsed: modelName
                    };
                }
            }

            // 에러 응답 처리
            const errBody = await res.json().catch(() => ({}));
            lastError = new Error(`[${modelName}] API 호출 실패 (${res.status}): ${errBody.error?.message || res.statusText}`);

            // 키 로테이션 먼저 시도
            if (keys.length > 1 && (currentKeyIndex + 1) < keys.length) {
                currentKeyIndex++;
                continue;
            }

            // 키를 다 썼으면 다음 모델 시도
            currentKeyIndex = 0;
            currentModelIndex++;
        } catch (netErr) {
            lastError = netErr;
            if (keys.length > 1 && (currentKeyIndex + 1) < keys.length) {
                currentKeyIndex++;
                continue;
            }
            currentKeyIndex = 0;
            currentModelIndex++;
        }
    }

    throw lastError || new Error('모든 Gemini 모델 및 API 키로의 질의응답 요청이 실패했습니다.');
}
