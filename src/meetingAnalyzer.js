// ─────────────────────────────────────────────
//  meetingAnalyzer.js
//  Gemini 텍스트 기반 회의록 생성 파이프라인
// ─────────────────────────────────────────────

import { FALLBACK_MODELS } from './utils/geminiModels.js';

// 청크 하나당 최대 입력 글자 수. 회의록 분석은 원문 전체를 교정된 발언록으로
// "재구성"해야 하므로(출력 길이 ≈ 입력 길이), 오탈자 검사 등 다른 기능보다
// 보수적으로 작게 잡아 모델의 출력 토큰 한도를 넘지 않도록 한다.
const MAX_CHARS_PER_CHUNK = 8000;

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

// ─────────────────────────────────────────────
//  Gemini 호출 공통 유틸 (키/모델 폴백 + 백오프)
//  extractTermsFromText / analyzeMeeting / askMeetingQuestion 이 각자
//  거의 동일한 재시도 로직을 중복 구현하고 있던 것을 하나로 통합.
// ─────────────────────────────────────────────
export async function callGeminiWithFallback(keys, contents, {
    systemInstruction = null,
    generationConfig = { temperature: 0.2 },
    onProgress = null,
    progressLabel = 'AI 분석',
    showElapsed = false,
} = {}) {
    const MODELS = FALLBACK_MODELS;
    let currentKeyIndex = 0;
    let currentModelIndex = 0;
    let modelRetries = 0;
    const maxModelRetries = MODELS.length;

    while (modelRetries < maxModelRetries) {
        const activeKey = keys[currentKeyIndex];
        const modelId = MODELS[currentModelIndex];
        const url = `https://generativelanguage.googleapis.com/v1beta/${modelId}:generateContent?key=${activeKey}`;

        let elapsedTimer = null;
        if (onProgress) {
            const keyInfo = keys.length > 1 ? ` (키 ${currentKeyIndex + 1}/${keys.length})` : '';
            let elapsed = 0;
            onProgress(`${progressLabel}: ${modelId.split('/').pop()} 로 처리 중...${keyInfo}`);
            if (showElapsed) {
                elapsedTimer = setInterval(() => {
                    elapsed++;
                    onProgress(`${progressLabel}: ${modelId.split('/').pop()} 로 처리 중...${keyInfo} (${elapsed}초 경과)`);
                }, 1000);
            }
        }

        const body = { contents, generationConfig };
        if (systemInstruction) body.systemInstruction = { parts: [{ text: systemInstruction }] };

        let res;
        try {
            res = await fetch(url, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(body),
            });
        } catch (netErr) {
            if (keys.length > 1 && (currentKeyIndex + 1) < keys.length) { currentKeyIndex++; continue; }
            currentKeyIndex = 0;
            modelRetries++;
            if (modelRetries >= maxModelRetries) throw netErr;
            currentModelIndex = (currentModelIndex + 1) % MODELS.length;
            continue;
        } finally {
            if (elapsedTimer) clearInterval(elapsedTimer);
        }

        if (res.ok) {
            if (onProgress) onProgress(`${progressLabel}: 응답 처리 중...`);
            recordUsage(modelId);
            const data = await res.json();
            return { text: data.candidates?.[0]?.content?.parts?.[0]?.text || '', modelUsed: modelId.split('/').pop() };
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

        // 2. 모든 키를 다 썼다면, 과부하/미지원 오류에 한해 모델 교체 후 재시도
        currentKeyIndex = 0;
        if (isOverloaded || isModelUnsupported) {
            modelRetries++;
            if (modelRetries < maxModelRetries) {
                currentModelIndex = (currentModelIndex + 1) % MODELS.length;
                const nextModel = MODELS[currentModelIndex].split('/').pop();
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
}

// ── JSON 코드펜스(```json ... ```) 제거 ──────────
export function stripJsonFence(content) {
    if (!content || !content.includes('```')) return content || '';
    const match = content.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
    if (match?.[1]) return match[1];
    return content.replace(/```(json)?/g, '').trim(); // 잘려서 안 닫힌 경우
}

// ── JSON 파싱 + 잘린 응답에 대한 점진적 복구 시도 ──
function tryParseJsonLoose(content) {
    try { return JSON.parse(content); } catch { /* fallthrough */ }
    const fixAttempts = [
        content + ']', content + '}', content + ']}',
        content + '"]}', content + '}]}', content + '"]}]}',
        content + '"}'
    ];
    for (const fixed of fixAttempts) {
        try { return JSON.parse(fixed); } catch { /* try next */ }
    }
    return null;
}

// ── JSON 파싱이 끝내 실패했을 때, 정규식으로 알아볼 수 있는 필드만 복구 ──
function regexRecoverFields(content) {
    const result = {};

    const mTitle = content.match(/"meetingTitle"\s*:\s*"((?:[^"\\]|\\.)*)"/i);
    if (mTitle) result.meetingTitle = mTitle[1].replace(/\\"/g, '"');

    const mAgenda = content.match(/"agenda"\s*:\s*"((?:[^"\\]|\\.)*)"/i);
    if (mAgenda) result.agenda = mAgenda[1].replace(/\\"/g, '"');

    const mSummary = content.match(/"summary"\s*:\s*"((?:[^"\\]|\\.)*)"/i);
    if (mSummary) result.summary = mSummary[1].replace(/\\"/g, '"');

    const mChunkSummary = content.match(/"chunkSummary"\s*:\s*"((?:[^"\\]|\\.)*)"/i);
    if (mChunkSummary) result.chunkSummary = mChunkSummary[1].replace(/\\"/g, '"');

    const mSpeaker = content.match(/"speakerCount"\s*:\s*(\d+)/i);
    if (mSpeaker) result.speakerCount = parseInt(mSpeaker[1], 10);

    const dMatch = content.match(/"decisions"\s*:\s*\[(.*?)\]/s);
    if (dMatch) {
        const items = dMatch[1].match(/"((?:[^"\\]|\\.)*)"/g);
        if (items) result.decisions = items.map(s => s.replace(/^"|"$/g, '').replace(/\\"/g, '"'));
    }

    const kMatch = content.match(/"keywords"\s*:\s*\[(.*?)\]/s);
    if (kMatch) {
        const items = kMatch[1].match(/"((?:[^"\\]|\\.)*)"/g);
        if (items) result.keywords = items.map(s => s.replace(/^"|"$/g, '').replace(/\\"/g, '"'));
    }

    const actionItems = [];
    const aMatches = content.matchAll(/"task"\s*:\s*"((?:[^"\\]|\\.)*)",\s*"owner"\s*:\s*"((?:[^"\\]|\\.)*)",\s*"deadline"\s*:\s*"((?:[^"\\]|\\.)*)"/g);
    for (const m of aMatches) {
        actionItems.push({ task: m[1].replace(/\\"/g, '"'), owner: m[2].replace(/\\"/g, '"'), deadline: m[3].replace(/\\"/g, '"') });
    }
    if (actionItems.length > 0) result.actionItems = actionItems;

    const transcript = [];
    const tMatches = content.matchAll(/"speaker"\s*:\s*"([^"]+)",\s*"text"\s*:\s*"((?:[^"\\]|\\.)*)",\s*"tag"\s*:\s*"([^"]+)"/g);
    for (const m of tMatches) {
        transcript.push({ speaker: m[1], text: m[2].replace(/\\"/g, '"'), tag: m[3] });
    }
    if (transcript.length > 0) result.transcript = transcript;

    return result;
}

// ── 긴 텍스트를 문단 경계 기준으로 청크 분할 ──────
function splitTextIntoChunks(text, maxChars = MAX_CHARS_PER_CHUNK) {
    if (text.length <= maxChars) return [text];

    const paragraphs = text.split(/\n{2,}/);
    const chunks = [];
    let current = '';

    for (let para of paragraphs) {
        // 문단 하나가 그 자체로 최대 길이를 넘으면 강제로 잘라서 처리
        while (para.length > maxChars) {
            if (current) { chunks.push(current); current = ''; }
            chunks.push(para.slice(0, maxChars));
            para = para.slice(maxChars);
        }
        if (!current) {
            current = para;
        } else if (current.length + para.length + 2 <= maxChars) {
            current += '\n\n' + para;
        } else {
            chunks.push(current);
            current = para;
        }
    }
    if (current) chunks.push(current);
    return chunks;
}

// ── 전문 용어 사전 문구 ───────────────────────
function buildTermSection(terminology) {
    return terminology && terminology.length > 0
        ? `\n\n[전문 용어 사전 - 반드시 아래 표기법을 사용하여 변환할 것]\n${terminology.map(t => `- ${t.word}: ${t.desc}`).join('\n')}`
        : '';
}

// ── 메인(단일 호출) 프롬프트 생성 ─────────────
function buildPrompt(terminology) {
    const termSection = buildTermSection(terminology);

    return `당신은 회의 내용을 분석하여 전문적이고 충실한 회의록을 작성하는 AI 전문가입니다.
회의 내용을 단순히 나열하는 것이 아니라, 맥락을 파악하고 핵심 인사이트와 향후 계획까지 심층 도출해야 합니다.

[처리 파이프라인 - 반드시 순서대로 수행]
1단계 - 오타 및 비문 수정: 회의 녹취록 텍스트의 모든 오타, 문법 오류, 어색한 표현을 자연스러운 문어체로 교정하라.
2단계 - 화자 분류: 문맥과 발언 패턴을 분석하여 화자를 A, B, C... 순으로 분류하라. 화자가 동시에 말하거나 발언이 섞인 경우, 문맥상 가장 개연성 있는 화자에 귀속시키고 [중복발언] 태그를 붙여라.
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

// ── 구간(청크) 단위 정제·추출 프롬프트 ─────────
function buildChunkPrompt(terminology, speakerHint, chunkIndex, totalChunks) {
    const termSection = buildTermSection(terminology);

    return `당신은 긴 회의 녹취록을 구간(청크) 단위로 정제·구조화하는 AI입니다.
지금 처리할 내용은 전체 회의 녹취록 중 ${chunkIndex}/${totalChunks}번째 구간입니다. 이 구간 내용만 근거로 아래 작업을 수행하세요.

[처리 작업]
1. 오타 및 비문 수정: 텍스트의 오타, 문법 오류, 어색한 표현을 자연스러운 문어체로 교정하라.
2. 화자 분류: 문맥과 발언 패턴을 근거로 화자를 A, B, C... 로 분류하라.${speakerHint}
3. 이 구간에서 언급된 결정 사항, 액션 아이템(담당자·기한 포함), 핵심 키워드, 핵심 인사이트(기회/리스크/통찰/우선순위)를 추출하라. 해당 사항이 없으면 빈 배열로 두라.
4. 이 구간의 핵심 내용을 2~4문장으로 요약하라.${termSection}

[출력 형식 - 반드시 아래 JSON 형식으로만 출력, 다른 텍스트 금지]
{
  "transcript": [ { "speaker": "<예: A>", "text": "<교정된 발언 내용>", "tag": "<정상|중복발언|불명확 중 택1>" } ],
  "decisions": ["<이 구간에서 결정된 사항>"],
  "actionItems": [ { "task": "<할 일>", "owner": "<담당자>", "deadline": "<기한, 없으면 null>", "priority": "<높음|보통|낮음>" } ],
  "keywords": ["<핵심 키워드>"],
  "keyInsights": [ { "type": "<기회|리스크|통찰|우선순위>", "content": "<내용>" } ],
  "chunkSummary": "<이 구간 핵심 내용 2~4문장 요약>"
}`;
}

// ── 구간별 결과를 종합하는 최종 프롬프트 ───────
function buildSynthesisPrompt(terminology, chunkSummaries, mergedDecisions, mergedActionItems, mergedKeywords, mergedKeyInsights, speakerCount) {
    const termSection = buildTermSection(terminology);

    return `당신은 여러 구간으로 나뉘어 부분 분석된 회의 내용을 종합하여 최종 회의록을 작성하는 AI 전문가입니다.
아래는 긴 회의를 구간별로 나누어 미리 정제·추출한 결과입니다. 이를 바탕으로 전체 회의를 관통하는 하나의 완결된 회의록을 작성하세요.

[구간별 요약 목록 (시간 순서)]
${chunkSummaries.map((s, i) => `구간 ${i + 1}: ${s}`).join('\n') || '(요약 없음)'}

[구간들에서 추출된 결정 사항 후보]
${mergedDecisions.map(d => '- ' + d).join('\n') || '(없음)'}

[구간들에서 추출된 액션 아이템 후보]
${mergedActionItems.map(a => `- ${a.task} | 담당: ${a.owner || '미정'} | 기한: ${a.deadline || '미정'}`).join('\n') || '(없음)'}

[구간들에서 추출된 키워드 후보]
${mergedKeywords.join(', ') || '(없음)'}

[구간들에서 추출된 핵심 인사이트 후보]
${mergedKeyInsights.map(k => `[${k.type}] ${k.content}`).join('\n') || '(없음)'}

감지된 화자 수(추정): ${speakerCount}${termSection}

[작업]
- 위 자료를 종합해 회의 제목, 목적/배경, 안건, 전체 요약, 주제별 상세 논의(2~6개), 핵심 인사이트, 결정사항, 액션아이템, 향후 계획(단기/중장기), 핵심 키워드를 작성하라.
- 후보로 제시된 항목은 구간별로 중복되거나 파편적일 수 있으니, 의미가 겹치는 항목은 통합·정리하고 중복은 제거하라.

[출력 형식 - 반드시 아래 JSON 형식으로만 출력. 다른 텍스트 금지 (transcript 필드는 포함하지 말 것)]
{
  "meetingTitle": "<회의 제목>",
  "meetingContext": "<회의 목적, 배경, 현재 상황을 3~5문장으로 서술>",
  "agenda": "<회의 안건을 2~4문장으로 요약>",
  "speakerCount": <화자 수 (숫자)>,
  "summary": "<전체 회의 핵심 내용을 4~6문장으로 요약>",
  "topicSummaries": [
    { "topic": "<주제 제목>", "content": "<논의 내용 3~5문장>", "result": "<결론 1~2문장>" }
  ],
  "keyInsights": [ { "type": "<기회|리스크|통찰|우선순위>", "content": "<내용>" } ],
  "decisions": ["<정리된 결정 사항>"],
  "actionItems": [ { "task": "<할 일>", "owner": "<담당자>", "deadline": "<기한, 없으면 null>", "priority": "<높음|보통|낮음>" } ],
  "futurePlans": {
    "shortTerm": [ { "plan": "<단기 계획>", "owner": "<담당자, 없으면 null>", "targetDate": "<목표 시점, 없으면 null>" } ],
    "longTerm": [ { "plan": "<중장기 계획>", "owner": "<담당자, 없으면 null>", "targetDate": "<목표 시점, 없으면 null>" } ]
  },
  "keywords": ["<핵심 키워드 5~10개>"]
}`;
}

// ── 단일 호출 경로 (텍스트가 청크 기준 이하일 때) ──
async function analyzeSingleChunk(text, keys, terminology, onProgress) {
    const prompt = buildPrompt(terminology);
    const contents = [{ role: 'user', parts: [{ text: '[회의 원본 텍스트]\n' + text }, { text: prompt }] }];

    if (onProgress) onProgress('Gemini AI 회의록 분석 중...');
    const { text: raw } = await callGeminiWithFallback(keys, contents, {
        generationConfig: { temperature: 0.2 },
        onProgress,
        progressLabel: '회의록 분석',
        showElapsed: true,
    });

    const stripped = stripJsonFence(raw);
    let result = tryParseJsonLoose(stripped);
    let partial = false;

    if (!result) {
        partial = true;
        result = {
            meetingTitle: '회의록',
            meetingContext: '',
            agenda: '분석 결과를 파싱하지 못했습니다.',
            speakerCount: 0,
            summary: 'JSON 결과가 잘려 부분적으로 추출되었습니다.\n' + stripped.substring(0, 500) + '...',
            topicSummaries: [],
            keyInsights: [],
            decisions: [],
            actionItems: [],
            futurePlans: { shortTerm: [], longTerm: [] },
            keywords: [],
            transcript: [],
            ...regexRecoverFields(stripped),
        };
    }

    result._meta = { partial, totalChunks: 1, failedChunks: partial ? [1] : [] };
    return result;
}

// ── 청크 분할 경로 (긴 텍스트: 구간별 정제 → 종합) ──
async function analyzeChunkedText(chunks, keys, terminology, onProgress) {
    if (onProgress) onProgress(`텍스트가 길어 ${chunks.length}개 구간으로 나누어 분석합니다...`);

    const transcript = [];
    const decisions = [];
    const actionItems = [];
    const keywords = [];
    const keyInsights = [];
    const chunkSummaries = [];
    const speakersSoFar = new Set();
    const failedChunks = [];
    let partial = false;

    for (let i = 0; i < chunks.length; i++) {
        if (onProgress) onProgress(`구간 ${i + 1}/${chunks.length} 분석 중...`);

        const lastUtterance = transcript.length > 0 ? transcript[transcript.length - 1] : null;
        const speakerHint = speakersSoFar.size > 0
            ? `\n\n[화자 연속성 참고] 지금까지 등장한 화자: ${Array.from(speakersSoFar).join(', ')}. 동일 인물이면 같은 레이블을 유지하고, 새로운 인물만 다음 알파벳을 사용하세요.${lastUtterance ? `\n직전 구간 마지막 발언: [${lastUtterance.speaker}] ${lastUtterance.text.slice(-200)}` : ''}`
            : '';

        const chunkPrompt = buildChunkPrompt(terminology, speakerHint, i + 1, chunks.length);
        const contents = [{
            role: 'user',
            parts: [{ text: `[회의 녹취록 구간 ${i + 1}/${chunks.length}]\n` + chunks[i] }, { text: chunkPrompt }]
        }];

        let chunkResult = null;
        try {
            const { text: raw } = await callGeminiWithFallback(keys, contents, {
                generationConfig: { temperature: 0.2 },
                onProgress,
                progressLabel: `구간 ${i + 1}/${chunks.length} 분석`,
            });
            const stripped = stripJsonFence(raw);
            chunkResult = tryParseJsonLoose(stripped);
            if (!chunkResult) {
                chunkResult = regexRecoverFields(stripped);
                partial = true;
                failedChunks.push(i + 1);
            }
        } catch {
            partial = true;
            failedChunks.push(i + 1);
            chunkResult = {};
        }

        (chunkResult.transcript || []).forEach(t => { transcript.push(t); speakersSoFar.add(t.speaker); });
        (chunkResult.decisions || []).forEach(d => decisions.push(d));
        (chunkResult.actionItems || []).forEach(a => actionItems.push(a));
        (chunkResult.keywords || []).forEach(k => keywords.push(k));
        (chunkResult.keyInsights || []).forEach(k => keyInsights.push(k));
        if (chunkResult.chunkSummary) chunkSummaries.push(chunkResult.chunkSummary);

        // 다음 구간 호출 전 짧은 대기 (rate limit 완화)
        if (i < chunks.length - 1) {
            await new Promise(r => setTimeout(r, 1500));
        }
    }

    if (onProgress) onProgress('구간별 분석 완료, 전체 회의록으로 종합하는 중...');

    const dedupedKeywords = Array.from(new Set(keywords)).slice(0, 15);
    const dedupedDecisions = Array.from(new Set(decisions));

    const synthesisPrompt = buildSynthesisPrompt(terminology, chunkSummaries, dedupedDecisions, actionItems, dedupedKeywords, keyInsights, speakersSoFar.size);
    const synthesisContents = [{ role: 'user', parts: [{ text: synthesisPrompt }] }];

    let finalResult;
    try {
        const { text: raw } = await callGeminiWithFallback(keys, synthesisContents, {
            generationConfig: { temperature: 0.2 },
            onProgress,
            progressLabel: '전체 회의록 종합',
            showElapsed: true,
        });
        const stripped = stripJsonFence(raw);
        finalResult = tryParseJsonLoose(stripped);
        if (!finalResult) {
            partial = true;
            finalResult = {
                meetingTitle: '회의록',
                meetingContext: '',
                agenda: '종합 결과를 파싱하지 못했습니다.',
                speakerCount: speakersSoFar.size,
                summary: 'AI 응답을 파싱하지 못해 구간별 추출 결과로 일부만 구성되었습니다.',
                topicSummaries: [],
                keyInsights,
                decisions: dedupedDecisions,
                actionItems,
                futurePlans: { shortTerm: [], longTerm: [] },
                keywords: dedupedKeywords,
                ...regexRecoverFields(stripped),
            };
        }
    } catch (e) {
        partial = true;
        finalResult = {
            meetingTitle: '회의록',
            meetingContext: '',
            agenda: '종합 단계에서 오류가 발생했습니다: ' + (e.message || ''),
            speakerCount: speakersSoFar.size,
            summary: '구간별 분석은 완료되었으나 최종 종합에 실패했습니다. 구간별로 추출된 결정 사항/액션 아이템은 아래에 표시됩니다.',
            topicSummaries: [],
            keyInsights,
            decisions: dedupedDecisions,
            actionItems,
            futurePlans: { shortTerm: [], longTerm: [] },
            keywords: dedupedKeywords,
        };
    }

    finalResult.transcript = transcript;
    finalResult._meta = { partial, totalChunks: chunks.length, failedChunks };
    return finalResult;
}

// ── 메인 분석 함수 ────────────────────────────
export async function analyzeMeeting(text, apiKey, terminology = [], onProgress) {
    const keys = String(apiKey).split(',').map(k => k.trim()).filter(k => k.match(/^(AIza|AQ\.)/));
    if (keys.length === 0) throw new Error('유효한 API 키가 없습니다.');

    const chunks = splitTextIntoChunks(text);

    if (chunks.length <= 1) {
        return analyzeSingleChunk(text, keys, terminology, onProgress);
    }
    return analyzeChunkedText(chunks, keys, terminology, onProgress);
}

// ─────────────────────────────────────────────
//  회의 내용 대상 AI 질의응답 (Q&A) 파이프라인
// ─────────────────────────────────────────────
// Q&A에 실어보낼 전체 컨텍스트 상한과, 그중 발언록(transcript)에 최소로 보장할 글자 수.
// 발언록은 [화자별 발언록 상세] 섹션으로 다른 요약 섹션들 뒤에 이어붙기 때문에,
// 예전에는 contextText.substring(0, 100000)로 "뒤에서부터" 잘려나가 회의 후반부
// 발언이 통째로 사라지는 문제가 있었다. 이제는 요약 섹션과 발언록의 예산을 분리하고,
// 발언록이 예산을 초과하면 질문과 관련도 높은 발언을 우선 선별해 담는다.
const MAX_QA_CONTEXT_CHARS = 120000;
const MIN_TRANSCRIPT_BUDGET_CHARS = 20000;

const QA_STOPWORDS = new Set([
    '그리고', '그러나', '그래서', '하지만', '또한', '대해', '대한', '에서', '에게', '으로', '로써',
    '에서는', '것을', '것은', '합니다', '했습니다', '입니다', '있습니다', '됩니다', '무엇', '어떤',
    '대해서', '누가', '언제', '어디', '어떻게', '왜', '알려줘', '알려주세요', '정리해줘', '해줘',
]);

function tokenizeQuestion(question) {
    return Array.from(new Set(
        (question.match(/[가-힣a-zA-Z0-9]{2,}/g) || [])
            .map(w => w.toLowerCase())
            .filter(w => !QA_STOPWORDS.has(w))
    ));
}

// 긴 텍스트(발언록 등)를 줄 단위로 받아, 예산 안에서 질문과 관련도 높은 줄을 우선 채택한다.
// 회의의 시작/끝 부분은 관련도와 무관하게 최소한씩 포함해 전체 맥락을 유지한다.
function selectRelevantLines(lines, question, maxChars) {
    const full = lines.join('\n');
    if (full.length <= maxChars || lines.length === 0) return full;

    const keywords = tokenizeQuestion(question);
    const edgeCount = 5;
    const alwaysInclude = new Set();
    for (let i = 0; i < Math.min(edgeCount, lines.length); i++) alwaysInclude.add(i);
    for (let i = Math.max(0, lines.length - edgeCount); i < lines.length; i++) alwaysInclude.add(i);

    const scored = lines.map((line, idx) => {
        if (alwaysInclude.has(idx) || keywords.length === 0) return { idx, score: 0 };
        const lower = line.toLowerCase();
        let score = 0;
        for (const kw of keywords) { if (lower.includes(kw)) score++; }
        return { idx, score };
    }).filter(s => s.score > 0).sort((a, b) => b.score - a.score || a.idx - b.idx);

    const selectedIdx = new Set(alwaysInclude);
    let currentLen = Array.from(alwaysInclude).reduce((sum, i) => sum + lines[i].length + 1, 0);

    for (const s of scored) {
        if (currentLen + lines[s.idx].length + 1 > maxChars) continue;
        selectedIdx.add(s.idx);
        currentLen += lines[s.idx].length + 1;
    }

    // 관련 발언이 부족해 예산이 많이 남으면, 앞에서부터 순서대로 채워 최대한 맥락을 보강한다.
    if (currentLen < maxChars * 0.5) {
        for (let i = 0; i < lines.length && currentLen < maxChars; i++) {
            if (selectedIdx.has(i)) continue;
            if (currentLen + lines[i].length + 1 > maxChars) continue;
            selectedIdx.add(i);
            currentLen += lines[i].length + 1;
        }
    }

    const sortedIdx = Array.from(selectedIdx).sort((a, b) => a - b);
    const resultLines = [];
    let lastIdx = -2;
    for (const i of sortedIdx) {
        if (i !== lastIdx + 1 && resultLines.length > 0) resultLines.push('...(중략)...');
        resultLines.push(lines[i]);
        lastIdx = i;
    }
    return resultLines.join('\n');
}

export async function askMeetingQuestion(question, contextData, apiKey, chatHistory = []) {
    const keys = String(apiKey).split(',').map(k => k.trim()).filter(k => k.match(/^(AIza|AQ\.)/));
    if (keys.length === 0) throw new Error('유효한 Gemini API 키가 없습니다. 설정에서 API 키를 입력해 주세요.');

    let contextText = '';
    if (typeof contextData === 'string') {
        const lines = contextData.split('\n');
        contextText = selectRelevantLines(lines, question, MAX_QA_CONTEXT_CHARS);
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

        const metaText = sections.join('\n\n====================\n\n');

        if ((contextData.transcript || []).length > 0) {
            const transcriptLines = contextData.transcript.map(t => `[${t.speaker}] ${t.text}`);
            const transcriptBudget = Math.max(MIN_TRANSCRIPT_BUDGET_CHARS, MAX_QA_CONTEXT_CHARS - metaText.length - 200);
            const transcriptSection = selectRelevantLines(transcriptLines, question, transcriptBudget);
            sections.push(`[화자별 발언록 상세 (Transcript) - 질문과 관련도가 높은 발언 위주로 발췌됨]\n${transcriptSection}`);
        }

        contextText = sections.join('\n\n====================\n\n');
    }

    const systemInstruction = `당신은 회의록 분석 및 질의응답을 수행하는 전문 AI 어시스턴트입니다.
제공된 [회의 내용 및 상세 발언록]을 바탕으로 사용자의 질문에 정확하고 구체적이며 친절하게 답변하세요.

[답변 원칙]
1. 회의 내용에 근거하여 사실에 입각해 명확히 답변하세요. 요약이나 짐작으로 얼버무리지 말고, 질문과 직접 관련된 세부 내용(수치, 일정, 담당자, 구체적 발언 등)을 최대한 구체적으로 제시하세요.
2. 답변의 핵심 근거가 되는 발언은 반드시 발언록 원문을 그대로 인용하세요. 인용 시 "[화자] "원문 문장 그대로""와 같이 화자 레이블과 큰따옴표를 함께 표기하고, 문장을 절대 바꿔 쓰지(의역하지) 마세요. 관련 발언이 여러 개면 2~4개까지 인용해도 됩니다.
3. 결정사항·과제·일정을 답할 때도 회의록에 적힌 표현을 최대한 그대로 살려서 답하세요.
4. 회의 내용에 명시되지 않은 추측이나 외부 지식은 "회의 내용에는 언급되지 않았으나"와 같이 명확히 구분하여 안내하세요.
5. 가독성을 위해 개조식(•, 번호)과 볼드체(**)를 적절히 활용하고, 인용 발언은 별도 줄로 구분해 정리하세요.`;

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
${contextText}

[질문]
${question}`;

    contents.push({
        role: 'user',
        parts: [{ text: currentPrompt }]
    });

    const { text: answer, modelUsed } = await callGeminiWithFallback(keys, contents, {
        systemInstruction,
        generationConfig: { temperature: 0.2, maxOutputTokens: 4096 },
    });

    if (!answer) throw new Error('AI로부터 빈 응답을 받았습니다.');

    return {
        answer: answer.trim(),
        modelUsed,
    };
}
