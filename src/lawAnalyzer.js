import { FALLBACK_MODELS } from './utils/geminiModels.js';

const GEMINI_BASE_URL = 'https://generativelanguage.googleapis.com/v1beta';
const MCP_SERVER_URL = 'https://lexguard-mcp.onrender.com/mcp';
const LOCAL_RAG_URL = 'http://localhost:8000/api/query';

async function callLexGuardMcp(query) {
  try {
    const payload = {
      jsonrpc: "2.0",
      id: "lexguard_req_" + Date.now(),
      method: "tools/call",
      params: {
        name: "legal_qa_tool",
        arguments: { query, max_results_per_type: 3 }
      }
    };

    const response = await fetch(MCP_SERVER_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'text/event-stream'
      },
      body: JSON.stringify(payload)
    });

    if (!response.ok) {
      console.warn("LexGuard MCP HTTP Error:", response.status);
      return { error: `MCP 서버 통신 실패 (${response.status})` };
    }

    const text = await response.text();
    const lines = text.split('\n');
    for (const line of lines) {
      if (line.startsWith('data: ')) {
        const jsonStr = line.replace('data: ', '').trim();
        if (jsonStr) {
          const parsed = JSON.parse(jsonStr);
          if (parsed.result) return parsed.result;
          if (parsed.error) return parsed.error;
        }
      }
    }
    return { error: "답변에서 유효한 데이터를 파싱하지 못했습니다." };
  } catch (err) {
    console.error("MCP Call Error:", err);
    return { error: err.message };
  }
}

export async function askLawAssistant(query, apiKey, history = [], onMcpCall = null) {
  if (!apiKey) {
    throw new Error("Gemini API Key가 필요합니다.");
  }

  const systemInstruction = `
당신은 대한민국 공공기관의 IT 및 공공사업 관련 규정, 법령, 가이드라인을 깊이 있게 이해하고 있는 'AI 법률 자문 에이전트(LexGuard MCP 기반)'입니다.

[핵심 지침]
1. 사용자의 질문에 법률/판례 검색이 필요하다고 판단되면 반드시 'legal_qa_tool' 도구를 사용하여 국가법령정보센터의 실시간 데이터를 조회하세요.
2. MCP 툴에서 반환된 진짜 데이터를 바탕으로 명확하고 전문적인 어조로 답변을 구성하세요.
3. 법령명 및 관련 근거를 제시할 때는 **반드시 법령명을 「」으로 감싸서** 정확히 표기하고, 조문 번호도 포함하세요. (예: 「소프트웨어 진흥법」 제43조 제1항, 「국가계약법」 제10조) 이 형식을 절대 생략하지 마세요.
4. 실무자(PM, 공무원 등)가 이해하기 쉽도록 핵심을 요약하고, 긴 텍스트는 불릿 기호(-, •)를 사용하여 가독성 있게 구조화하세요.
5. 답변의 서두나 말미에 '본 답변은 법적 판단을 대신하지 않으며 참고용입니다'라는 판단 유보 문구를 가볍게 포함해 주세요.
  `;

  // history를 Gemini API 포맷으로 변환 (user, model만 허용되며 function 호출 내역은 생략하거나 파싱해야 함, 여기서는 순수 텍스트 메시지만 필터링)
  const contents = [];
  for (const msg of history) {
      if (msg.role === 'user' || msg.role === 'model') {
          contents.push({
              role: msg.role,
              parts: [{ text: msg.text || "" }]
          });
      }
  }

  // 현재 질문 추가
  contents.push({
    role: 'user',
    parts: [{ text: query }]
  });

  const tools = [{
    functionDeclarations: [
      {
        name: "legal_qa_tool",
        description: "국가법령정보센터 최신 데이터를 기반으로 법령, 판례, 행정해석을 검색합니다. 사용자가 법률이나 규정에 대해 물어볼 때 무조건 이 툴을 호출하세요.",
        parameters: {
          type: "OBJECT",
          properties: {
            query: {
              type: "STRING",
              description: "검색에 사용할 핵심 키워드 조합 (예: '소프트웨어 하도급 제한 요건', '개인정보보호법 가명정보')"
            }
          },
          required: ["query"]
        }
      }
    ]
  }];

  const keys = String(apiKey).split(',').map(k => k.trim()).filter(k => k.match(/^(AIza|AQ\.)/));
  if (keys.length === 0) {
    throw new Error("유효한 Gemini API Key가 필요합니다.");
  }


  let currentKeyIndex = 0;
  let currentModelIndex = 0;

  const generateWithRetry = async (reqContents, useTools = true) => {
    let modelRetries = 0;
    const maxModelRetries = FALLBACK_MODELS.length;

    while (modelRetries < maxModelRetries) {
      const activeKey = keys[currentKeyIndex];
      const modelId = FALLBACK_MODELS[currentModelIndex];
      const fetchUrl = `${GEMINI_BASE_URL}/${modelId}:generateContent?key=${activeKey}`;

      const payload = {
        systemInstruction: { parts: [{ text: systemInstruction }] },
        contents: reqContents,
        tools: useTools ? tools : undefined,
        generationConfig: { temperature: 0.1, topK: 40, topP: 0.95 }
      };

      try {
        const response = await fetch(fetchUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
        });

        if (response.ok) {
          return await response.json();
        }

        const errorData = await response.json().catch(() => ({}));
        const errMsg = errorData.error?.message || `API 요청 실패 (${response.status})`;
        const isModelUnavailable = response.status === 404
          || response.status === 400
          || errMsg.toLowerCase().includes('not found')
          || errMsg.toLowerCase().includes('not supported')
          || errMsg.toLowerCase().includes('deprecated');

        // 1. 에러 발생 시 항상 다음 API 키를 먼저 시도
        if (keys.length > 1 && (currentKeyIndex + 1) < keys.length) {
          currentKeyIndex++;
          continue;
        }

        // 2. 모든 키를 다 썼다면 모델 교체 시도
        if (response.status === 429 || response.status >= 500 || isModelUnavailable) {
          modelRetries++;
          if (modelRetries < maxModelRetries) {
            currentKeyIndex = 0;
            currentModelIndex = (currentModelIndex + 1) % FALLBACK_MODELS.length;
            await new Promise(r => setTimeout(r, 5000)); // 5초 대기 후 모델 전환
            continue;
          }
        }

        throw new Error(errMsg);
      } catch (err) {
        if (modelRetries >= maxModelRetries - 1) throw err;
        modelRetries++;
        currentModelIndex = (currentModelIndex + 1) % FALLBACK_MODELS.length;
      }
    }
  };

  try {
    let data = await generateWithRetry(contents);
    let candidate = data.candidates && data.candidates[0];

    if (!candidate) throw new Error("API 응답에서 결과를 찾을 수 없습니다.");

    // Function Calling 부분이 있는지 모든 파트에서 확인
    let funcPart = candidate.content.parts.find(p => p.functionCall && p.functionCall.name === 'legal_qa_tool');
    
    if (funcPart) {
      const functionArgs = funcPart.functionCall.args;
      const mcpQuery = functionArgs.query || query;
      
      // UI에 MCP 호출 사실 알림
      if (onMcpCall) onMcpCall(mcpQuery);

      // LexGuard MCP 서버 실제 호출
      const mcpResult = await callLexGuardMcp(mcpQuery);

      // Gemini에게 함수 실행 결과 전달을 위해 히스토리에 추가
      contents.push(candidate.content); 
      contents.push({
        role: "function",
        parts: [{
          functionResponse: {
            name: "legal_qa_tool",
            response: mcpResult
          }
        }]
      });

      // 2차 생성 요청 (최종 답변 생성)
      data = await generateWithRetry(contents);
      candidate = data.candidates && data.candidates[0];
      if (!candidate) throw new Error("API 응답(2차)에서 결과를 찾을 수 없습니다.");
      
      // 2차 응답에서 텍스트 추출
      const finalParts = candidate.content.parts || [];
      const extractedText = finalParts.map(p => p.text || '').join('').trim();
      
      // 2차에서도 텍스트가 없으면(또다시 functionCall 시도 등) tools 없이 3차 요청으로 강제 텍스트 응답 유도
      if (!extractedText) {
        data = await generateWithRetry([...contents, candidate.content, {
          role: 'user',
          parts: [{ text: '위 조회 결과를 바탕으로 질문에 대한 답변을 한국어로 작성해 주세요.' }]
        }], false);
        candidate = data.candidates && data.candidates[0];
        if (candidate) {
          return candidate.content.parts.map(p => p.text || '').join('').trim() || '답변을 생성하지 못했습니다.';
        }
      }
      
      return extractedText;
    }

    // 함수 호출 없이 바로 텍스트를 리턴한 경우
    const extractedText = candidate.content.parts.map(p => p.text || '').join('').trim();
    return extractedText || "내용을 찾을 수 없습니다.";
  } catch (error) {
    console.error("[askLawAssistant] Error:", error);
    throw error;
  }
}

// MCP를 사용하지 않는 일반 AI 법률 자문 함수 (토큰 절약 및 빠른 응답용)
export async function askGeneralLawAssistant(query, apiKey, history = []) {
  if (!apiKey) throw new Error("Gemini API Key가 필요합니다.");

  const systemInstruction = `
당신은 대한민국 공공기관의 IT 및 공공사업 관련 규정에 조예가 깊은 'AI 법률 자문 어시스턴트'입니다.
현재 당신은 외부 실시간 법령 조회 도구(MCP)를 사용하지 않고 당신의 내부 지식만으로 답변하고 있습니다.

[지침]
1. 당신이 알고 있는 소프트웨어 진흥법, 국가계약법, 행정기관 정보시스템 구축운영 지침 등을 바탕으로 답변하세요.
2. 실시간 조회를 하지 않으므로 답변 끝에 "실시간 법령 조회를 하지 않은 기반 지식 답변이므로 최신 규정은 반드시 'AI 법률 자문 (MCP)' 메뉴나 법제처를 통해 재확인하시기 바랍니다."라는 안내를 포함하세요.
3. 법령명은 「」으로 감싸고 전문적인 어조를 유지하세요.
  `;

  const keys = String(apiKey).split(',').map(k => k.trim()).filter(k => k.match(/^(AIza|AQ\.)/));
  if (keys.length === 0) throw new Error("유효한 Gemini API Key가 필요합니다.");


  let currentKeyIndex = 0;
  let currentModelIndex = 0;

  const generateWithRetry = async (reqContents) => {
    let modelRetries = 0;
    while (modelRetries < FALLBACK_MODELS.length) {
      const activeKey = keys[currentKeyIndex];
      const modelId = FALLBACK_MODELS[currentModelIndex];
      
      try {
        const response = await fetch(`${GEMINI_BASE_URL}/${modelId}:generateContent?key=${activeKey}`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            systemInstruction: { parts: [{ text: systemInstruction }] },
            contents: reqContents,
            generationConfig: { temperature: 0.2 }
          }),
        });

        if (response.ok) return await response.json();

        const errData = await response.json().catch(() => ({}));
        const errMsg = errData.error?.message || `API 요청 실패 (${response.status})`;
        const isModelUnavailable = response.status === 404
          || response.status === 400
          || errMsg.toLowerCase().includes('not found')
          || errMsg.toLowerCase().includes('not supported')
          || errMsg.toLowerCase().includes('deprecated');

        // 1. 에러 발생 시 항상 다음 API 키를 먼저 시도
        if (keys.length > 1 && (currentKeyIndex + 1) < keys.length) {
          currentKeyIndex++;
          continue;
        }

        // 2. 모든 키를 다 썼다면 모델 교체 시도
        if (response.status === 429 || response.status >= 500 || isModelUnavailable) {
          modelRetries++;
          if (modelRetries < FALLBACK_MODELS.length) {
            currentKeyIndex = 0;
            currentModelIndex = (currentModelIndex + 1) % FALLBACK_MODELS.length;
            await new Promise(r => setTimeout(r, 5000));
            continue;
          }
        }
        throw new Error(errMsg);
      } catch (err) {
        if (modelRetries >= FALLBACK_MODELS.length - 1) throw err;
        modelRetries++;
        currentModelIndex++;
      }
    }
  };
  
  // Gemini API는 첫 번째 메시지가 'user'여야 하고, 역할이 번갈아 나와야 함.
  // 인사말(model)이 먼저 나오는 히스토리를 고려하여 필터링 함.
  const contents = [];
  const historyMessages = (history || []).filter(m => m.text && m.text.trim());
  
  for (const m of historyMessages) {
    if (m.role === 'user' || m.role === 'model') {
      // 첫 번째 메시지가 model이면 무시 (API 제약)
      if (contents.length === 0 && m.role === 'model') continue;
      
      // 연속된 동일 역할 방지 (단순화)
      if (contents.length > 0 && contents[contents.length - 1].role === m.role) continue;
      
      contents.push({ role: m.role, parts: [{ text: m.text }] });
    }
  }
  
  // 현재 질문 추가
  if (contents.length > 0 && contents[contents.length - 1].role === 'user') {
      // 마지막이 user면 텍스트를 합침 (또는 이전 user 메시지 무시)
      contents[contents.length - 1].parts[0].text += `\n\n${query}`;
  } else {
      contents.push({ role: 'user', parts: [{ text: query }] });
  }

  const data = await generateWithRetry(contents);
  return data.candidates?.[0]?.content?.parts?.[0]?.text || "답변을 생성할 수 없습니다.";
}

// Local RAG를 이용한 법률 자문 (RFP Validator 연동용)
export async function askLocalRagAssistant(query, apiKey, history = [], onRagCall = null) {
  if (!apiKey) throw new Error("Gemini API Key가 필요합니다.");

  const systemInstruction = `
당신은 대한민국 공공기관의 IT 및 공공사업 관련 규정, 법령, 가이드라인을 깊이 있게 이해하고 있는 'AI 법률 자문 에이전트(Local RAG 기반)'입니다.

[핵심 지침]
1. 사용자의 질문에 답변하기 위해 전달된 '검색된 법령(Local RAG 데이터)'을 최우선으로 참고하여 사실에만 근거하여 답변하세요.
2. 답변을 작성할 때 **반드시 근거 법령의 명칭(예: 「소프트웨어 진흥법」), 구체적인 조항 번호(예: 제43조), 그리고 해당 문장의 구체적인 항 번호(예: 제1항, 원문의 동그라미 숫자 ①은 제1항으로 표기) 및 호 번호(예: 제1호)를 누락 없이 구체적으로 명시**하십시오.
   - 잘못된 예: "규정에 따르면 하도급은 제한됩니다." (X)
   - 올바른 예: "「소프트웨어 진흥법」 제43조 제1항에 따르면, 소프트웨어사업자는 국가기관등과 체결한 소프트웨어사업 계약 금액의 100분의 50을 초과하여 하도급할 수 없습니다." (O)
3. 만약 제공된 검색 데이터 내 문장에 동그라미 숫자(①, ②, ③ 등)가 포함되어 있다면, 이를 '제1항', '제2항', '제3항'으로 매끄럽고 명확하게 풀어서 서술하십시오.
4. 실무자(PM, 공무원 등)가 이해하기 쉽도록 핵심을 요약하고, 긴 텍스트는 불릿 기호(-, •)를 사용하여 가독성 있게 구조화하세요.
5. 답변의 서두나 말미에 '본 답변은 법적 판단을 대신하지 않으며 참고용입니다 (Local RAG)'라는 판단 유보 문구를 가볍게 포함해 주세요.
  `;

  // 1. RAG 백엔드에 쿼리 요청
  if (onRagCall) onRagCall(query);
  let ragData = null;
  try {
    const ragResponse = await fetch(LOCAL_RAG_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ query: query })
    });
    
    if (ragResponse.ok) {
      ragData = await ragResponse.json();
    } else {
      console.warn(`RAG Server Error: ${ragResponse.status}`);
    }
  } catch(e) {
    console.error("Local RAG Query failed, falling back to general query:", e);
  }

  // RAG 검색 결과가 있는 경우 Gemini로 고품질 자문 내용 합성
  if (ragData && ragData.results && ragData.results.length > 0) {
    let context_str = "";
    ragData.results.forEach((doc, idx) => {
      context_str += `[참고 ${idx + 1}] ${doc.title} ${doc.article_header || ''}\n내용: ${doc.content}\n\n`;
    });

    const user_prompt = `[참고 자료]\n${context_str}\n\n[사용자 질문]\n${query}\n\n[답변]`;

    // Gemini 호출 설정
    const keys = String(apiKey).split(',').map(k => k.trim()).filter(k => k.match(/^(AIza|AQ\.)/));
    if (keys.length === 0) throw new Error("유효한 Gemini API Key가 필요합니다.");


    let currentKeyIndex = 0;
    let currentModelIndex = 0;

    const generateWithRetry = async (reqContents) => {
      let modelRetries = 0;
      while (modelRetries < FALLBACK_MODELS.length) {
        const activeKey = keys[currentKeyIndex];
        const modelId = FALLBACK_MODELS[currentModelIndex];
        
        try {
          const response = await fetch(`${GEMINI_BASE_URL}/${modelId}:generateContent?key=${activeKey}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              systemInstruction: { parts: [{ text: systemInstruction }] },
              contents: reqContents,
              generationConfig: { temperature: 0.1 } // 낮은 온도로 사실 기반 정확도 유지
            }),
          });

          if (response.ok) return await response.json();

          const errData = await response.json().catch(() => ({}));
          const errMsg = errData.error?.message || `API 요청 실패 (${response.status})`;
          const isModelUnavailable = response.status === 404
            || response.status === 400
            || errMsg.toLowerCase().includes('not found')
            || errMsg.toLowerCase().includes('not supported')
            || errMsg.toLowerCase().includes('deprecated');

          if (keys.length > 1 && (currentKeyIndex + 1) < keys.length) {
            currentKeyIndex++;
            continue;
          }

          if (response.status === 429 || response.status >= 500 || isModelUnavailable) {
            modelRetries++;
            if (modelRetries < FALLBACK_MODELS.length) {
              currentKeyIndex = 0;
              currentModelIndex = (currentModelIndex + 1) % FALLBACK_MODELS.length;
              await new Promise(r => setTimeout(r, 5000));
              continue;
            }
          }
          throw new Error(errMsg);
        } catch (err) {
          if (modelRetries >= FALLBACK_MODELS.length - 1) throw err;
          modelRetries++;
          currentModelIndex++;
        }
      }
    };

    const contents = [];
    const historyMessages = (history || []).filter(m => m.text && m.text.trim());
    
    for (const m of historyMessages) {
      if (m.role === 'user' || m.role === 'model') {
        if (contents.length === 0 && m.role === 'model') continue;
        if (contents.length > 0 && contents[contents.length - 1].role === m.role) continue;
        contents.push({ role: m.role, parts: [{ text: m.text }] });
      }
    }
    
    if (contents.length > 0 && contents[contents.length - 1].role === 'user') {
      contents[contents.length - 1].parts[0].text += `\n\n${user_prompt}`;
    } else {
      contents.push({ role: 'user', parts: [{ text: user_prompt }] });
    }

    try {
      const data = await generateWithRetry(contents);
      return data.candidates?.[0]?.content?.parts?.[0]?.text || (ragData ? ragData.answer : "답변을 생성할 수 없습니다.");
    } catch (geminiError) {
      console.error("Gemini RAG synthesis failed, falling back to local Qwen answer:", geminiError);
      return ragData ? ragData.answer : `답변 생성 실패: ${geminiError.message}`;
    }
  }

  // RAG 서버로부터 텍스트 답변(Qwen)만 돌아온 경우의 Fallback
  if (ragData && ragData.answer) {
    return ragData.answer;
  }

  return "로컬 RAG 서버에서 관련 법령 데이터를 찾지 못했거나 통신에 실패했습니다. (서버 작동 여부 및 ChromaDB 구축 상태를 확인해 주세요)";
}

