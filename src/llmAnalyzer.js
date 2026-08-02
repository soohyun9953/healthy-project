import { FALLBACK_MODELS } from './utils/geminiModels.js';

async function fetch_with_timeout(resource, options = {}) {
    const { timeout = 25000 } = options;
    const controller = new AbortController();
    const id = setTimeout(() => controller.abort(), timeout);
    
    try {
        const response = await fetch(resource, {
            ...options,
            signal: controller.signal
        });
        clearTimeout(id);
        return response;
    } catch (error) {
        clearTimeout(id);
        throw error;
    }
}

const sleep_delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

function split_text_into_chunks(text, max_chunk_size = 15000) {
    if (!text) return [];
    if (text.length <= max_chunk_size) return [text];
    
    const chunks = [];
    let current_index = 0;
    
    while (current_index < text.length) {
        let end_index = current_index + max_chunk_size;
        if (end_index >= text.length) {
            chunks.push(text.substring(current_index));
            break;
        }
        
        let last_newline = text.lastIndexOf('\n', end_index);
        if (last_newline > current_index + (max_chunk_size * 0.7)) {
            end_index = last_newline + 1;
        } else {
            let last_period = text.lastIndexOf('. ', end_index);
            if (last_period > current_index + (max_chunk_size * 0.7)) {
                end_index = last_period + 2;
            }
        }
        
        chunks.push(text.substring(current_index, end_index));
        current_index = end_index;
    }
    
    return chunks;
}



function splitTextAtNewline(text) {
    if (!text || text.length <= 1) return [text, ""];
    const mid = Math.floor(text.length / 2);
    
    let nextNL = text.indexOf('\n', mid);
    let prevNL = text.lastIndexOf('\n', mid);
    
    let splitIndex = -1;
    if (nextNL !== -1 && prevNL !== -1) {
        if ((nextNL - mid) < (mid - prevNL)) {
            splitIndex = nextNL;
        } else {
            splitIndex = prevNL;
        }
    } else if (nextNL !== -1) {
        splitIndex = nextNL;
    } else if (prevNL !== -1) {
        splitIndex = prevNL;
    } else {
        splitIndex = mid;
    }

    const part1 = String(text || "").substring(0, splitIndex).trim();
    const part2 = String(text || "").substring(splitIndex).trim();
    return [part1, part2];
}

function merge_multiple_results(results_array, is_typo_mode) {
    if (!results_array || results_array.length === 0) return null;
    const valid_results = results_array.filter(Boolean);
    if (valid_results.length === 0) return null;
    
    let merged = valid_results[0];
    for (let i = 1; i < valid_results.length; i++) {
        merged = mergeResults(merged, valid_results[i], is_typo_mode, i + 1);
    }
    
    return merged;
}

function mergeResults(res1, res2, isTypoMode, partNumber = 2) {
    const s1 = typeof res1?.score === 'number' && !isNaN(res1.score) ? res1.score : 80;
    const s2 = typeof res2?.score === 'number' && !isNaN(res2.score) ? res2.score : 80;
    
    const sum1 = res1?.summary && res1.summary !== 'undefined' ? res1.summary : '분석 완료';
    const sum2 = res2?.summary && res2.summary !== 'undefined' ? res2.summary : '분석 완료';

    const merged = {
        score: Math.round((s1 + s2) / 2),
        inspectionScope: res1?.inspectionScope || res2?.inspectionScope || null,
        summary: partNumber === 2 
            ? `[1부 분석 요약]\n${sum1}\n\n[2부 분석 요약]\n${sum2}`
            : `${sum1}\n\n[${partNumber}부 분석 요약]\n${sum2}`,
        requirementMapping: [],
        typos: []
    };

    if (isTypoMode) {
        const uniqueTypos = [];
        const seen = new Set();
        const addTypos = (typos) => {
            if (!typos || !Array.isArray(typos)) return;
            typos.forEach(typo => {
                const sig = `${typo.page || typo.location}_${typo.originalText || typo.errorText}_${typo.correction}`;
                if (!seen.has(sig)) {
                    seen.add(sig);
                    uniqueTypos.push(typo);
                }
            });
        };
        addTypos(res1?.typos);
        addTypos(res2?.typos);
        merged.typos = uniqueTypos;
    } else {
        let combinedMapping = [];
        if (res1?.requirementMapping && Array.isArray(res1.requirementMapping)) combinedMapping = combinedMapping.concat(res1.requirementMapping);
        if (res2?.requirementMapping && Array.isArray(res2.requirementMapping)) combinedMapping = combinedMapping.concat(res2.requirementMapping);
        
        combinedMapping.forEach((req, index) => {
            req.id = `REQ-${String(index + 1).padStart(3, '0')}`;
        });
        merged.requirementMapping = combinedMapping;

        merged.rtm = combinedMapping.map(req => ({
            type: req.type || '필수',
            requirement: req.requirement || '-',
            status: req.status || '미이행(X)',
            location: req.artifactSection || '해당 없음',
            category: req.category || '-',
            levelLabel: req.levelLabel || '개별문장'
        }));

        merged.omissions = combinedMapping
            .filter(req => req.status !== '이행(O)')
            .map(req => ({
                title: `[ID: ${req.id}] ${String(req.requirement || '').substring(0, 30)}...`,
                evidence: req.requirement || '-',
                reason: req.gap || '구체적인 수행/설계 방안이 누락되었습니다.',
                recommendation: '해당 요건을 만족하기 위한 구체적인 명세와 실행계획을 산출물에 추가해야 합니다.'
            }));

        const uniqueTypos = [];
        const seen = new Set();
        const addTypos = (typos) => {
            if (!typos || !Array.isArray(typos)) return;
            typos.forEach(typo => {
                const sig = `${typo.location || typo.page}_${typo.originalText}_${typo.correction}`;
                if (!seen.has(sig)) {
                    seen.add(sig);
                    uniqueTypos.push(typo);
                }
            });
        };
        addTypos(res1?.typos);
        addTypos(res2?.typos);
        merged.typos = uniqueTypos;
    }

    return merged;
}

async function analyze_with_ollama(prompt, model = 'qwen2.5:3b', onProgress) {
    if (onProgress) onProgress(`로컬 LLM (${model}) 분석 진행 중... (Ollama http://localhost:11434)`);
    try {
        const response = await fetch('http://localhost:11434/api/generate', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                model: model || 'qwen2.5:3b',
                prompt: prompt,
                format: 'json', // Ollama 강제 JSON 포맷 지정
                stream: false,
                options: { temperature: 0.1 }
            })
        });

        if (!response.ok) {
            if (response.status === 404) {
                throw new Error(`Ollama에 '${model}' 모델이 설치되어 있지 않습니다.\n터미널에서 [ollama run ${model}] 명령어로 모델을 먼저 받거나 설정에서 변경해 주세요.`);
            }
            throw new Error(`로컬 LLM (Ollama) 응답 오류: HTTP ${response.status}`);
        }

        const data = await response.json();
        const raw_text = data.response || "{}";

        // 강력한 JSON 추출, 복구 및 결과 정규화 파서
        const parse_and_normalize_ollama = (text) => {
            if (!text) return null;
            let clean = text.trim();
            let parsed = null;

            // 1. ```json ... ``` 패턴 추출
            if (clean.includes("```")) {
                const match = clean.match(/```(?:json)?\s*([\s\S]*?)\s*```/i);
                if (match && match[1]) {
                    clean = match[1].trim();
                }
            }

            // 2. JSON.parse 시도
            try {
                parsed = JSON.parse(clean);
            } catch (e1) {
                const start = clean.indexOf('{');
                const end = clean.lastIndexOf('}');
                if (start !== -1 && end !== -1 && end > start) {
                    const sub = clean.substring(start, end + 1);
                    try {
                        parsed = JSON.parse(sub);
                    } catch (e2) {
                        try {
                            const repaired = sub.replace(/,\s*([\}\]])/g, '$1');
                            parsed = JSON.parse(repaired);
                        } catch (e3) {}
                    }
                }
            }

            if (!parsed || typeof parsed !== 'object') {
                parsed = { summary: text };
            }

            // 3. 필드명 정규화
            const score = typeof parsed.score === 'number' && !isNaN(parsed.score) 
                ? parsed.score 
                : (typeof parsed.overallScore === 'number' ? parsed.overallScore : 85);

            // 재귀적 Deep Walker: 자율적 JSON 구조에서도 실제 의미있는 typos만 엄격히 필터링 수집
            const deep_extracted_typos = [];
            const summary_bullets = [];

            const walk_json_tree = (obj, path = "1페이지") => {
                if (!obj) return;
                if (typeof obj === 'string') {
                    const str = obj.trim();
                    if (str.length > 3) {
                        summary_bullets.push(`- [${path}] ${str.substring(0, 100)}`);
                    }
                    return;
                }
                if (Array.isArray(obj)) {
                    obj.forEach((item) => walk_json_tree(item, path));
                    return;
                }
                if (typeof obj === 'object') {
                    // 원문(originalText)과 수정안(correction)이 명시되어 있고 두 텍스트가 실제로 다를 때만 진짜 교정건으로 수집!
                    const orig = String(obj.originalText || obj.original || obj.errorText || obj.before || '').trim();
                    const corr = String(obj.correction || obj.correct || obj.suggestion || '').trim();
                    
                    if (orig && corr && orig !== corr && corr !== '문맥 검토 및 구체적 명세 보완 권고' && corr !== '보완 권고') {
                        deep_extracted_typos.push({
                            page: String(obj.page || obj.location || obj.section || path),
                            originalText: orig,
                            correction: corr,
                            errorType: String(obj.errorType || obj.type || obj.reason || obj.category || '[표현 품질] 띄어쓰기 및 맞춤법 교정')
                        });
                        return;
                    }

                    // 일반 키-값 객체 (예: {"슬라이드 26": { ... }})
                    for (const [key, value] of Object.entries(obj)) {
                        if (['score', 'summary', 'overview'].includes(key)) continue;
                        const next_path = (path === "1페이지" || !path) ? key : `${path} > ${key}`;
                        walk_json_tree(value, next_path);
                    }
                }
            };

            walk_json_tree(parsed);

            // summary 정교화: JSON 객체가 요약에 찍히지 않도록 가공
            let summary = "";
            if (typeof parsed.summary === 'string' && !parsed.summary.trim().startsWith('{')) {
                summary = parsed.summary;
            } else if (parsed.overview || parsed.analysis) {
                summary = String(parsed.overview || parsed.analysis);
            } else if (summary_bullets.length > 0) {
                summary = `[로컬 LLM 산출물 핵심 명세 요약]\n${summary_bullets.slice(0, 10).join('\n')}`;
            } else {
                summary = "로컬 LLM 점검 완료: 띄어쓰기 및 명백한 맞춤법 오류 위주로 정밀 검수를 완료했습니다.";
            }

            // typos 목록 정밀 병합 (원문 != 수정안 인 진짜 결함만 남김)
            let raw_typos = parsed.typos || parsed.typo_list || parsed.corrections || parsed.errors || parsed.issues || parsed.items || [];
            if (!Array.isArray(raw_typos) && typeof raw_typos === 'object') {
                raw_typos = Object.values(raw_typos);
            }
            
            const direct_typos = (Array.isArray(raw_typos) ? raw_typos : []).map(t => ({
                page: String(t.page || t.location || t.section || '1페이지'),
                originalText: String(t.originalText || t.original || t.errorText || t.before || '').trim(),
                correction: String(t.correction || t.correct || t.after || t.suggestion || '').trim(),
                errorType: String(t.errorType || t.type || t.reason || t.category || '[표현 품질] 띄어쓰기 및 맞춤법 교정')
            })).filter(t => t.originalText && t.correction && t.originalText !== t.correction && t.correction !== '문맥 검토 및 구체적 명세 보완 권고');

            // 직접 수집된 typos 중 유효한 건이 있으면 사용하고, 아니면 딥 워커 수집건 적용
            const final_typos = direct_typos.length > 0 ? direct_typos : deep_extracted_typos;

            // requirementMapping (요구사항 매핑) 유연한 추출
            let raw_reqs = parsed.requirementMapping || parsed.rtm || parsed.requirements || parsed.mapping || [];
            if (!Array.isArray(raw_reqs) && typeof raw_reqs === 'object') {
                raw_reqs = Object.values(raw_reqs);
            }
            
            const requirementMapping = (Array.isArray(raw_reqs) ? raw_reqs : []).map((r, idx) => ({
                id: r.id || `REQ-${String(idx + 1).padStart(3, '0')}`,
                category: String(r.category || '기능'),
                type: String(r.type || '필수'),
                levelLabel: String(r.levelLabel || '개별문장'),
                path: String(r.path || '본문'),
                requirement: String(r.requirement || r.req || ''),
                artifactSection: String(r.artifactSection || r.section || '해당 없음'),
                artifactContent: String(r.artifactContent || r.content || ''),
                status: String(r.status || '이행(O)'),
                gap: r.gap || null
            }));

            const rtm = requirementMapping.map(req => ({
                type: req.type,
                requirement: req.requirement,
                status: req.status,
                location: req.artifactSection,
                category: req.category,
                levelLabel: req.levelLabel
            }));

            const omissions = requirementMapping
                .filter(req => req.status !== '이행(O)')
                .map(req => ({
                    title: `[ID: ${req.id}] ${String(req.requirement || '').substring(0, 30)}...`,
                    evidence: req.requirement || '-',
                    reason: req.gap || '구체적인 수행 방안 보완이 필요합니다.',
                    recommendation: '실행 계획을 산출물에 추가하십시오.'
                }));

            return {
                score,
                summary,
                requirementMapping,
                rtm,
                omissions,
                typos: final_typos
            };
        };

        return parse_and_normalize_ollama(raw_text);

    } catch (err) {
        if (err.name === 'TypeError' && err.message.includes('fetch')) {
            throw new Error("로컬 LLM (Ollama) 서버가 실행되어 있지 않습니다.\nPC에서 Ollama 앱을 실행해 주세요. (http://localhost:11434)");
        }
        throw err;
    }
}

export async function analyzeDocumentsWithLLM(guidelineText, artifactText, inspectionScope, apiKey, glossaryText, onProgress, selectedModel = 'auto', isSubCall = false, ragContext = "", llmProvider = 'gemini', ollamaModel = 'qwen2.5:3b') {
    const keys = String(apiKey || '').split(',').map(k => k.trim()).filter(k => k.match(/^(AIza|AQ\.)/));
    if (llmProvider !== 'ollama' && keys.length === 0) {
        throw new Error("유효한 Gemini API 키가 제공되지 않았습니다. [설정] 메뉴에서 API 키를 등록하거나 '로컬 LLM (Ollama)'을 선택해 주세요.");
    }

    let currentKeyIndex = 0;
    const isOnlyTypoCheck = !guidelineText || guidelineText.trim() === '';

    if (!isSubCall) {
        if (isOnlyTypoCheck && artifactText && artifactText.length > 20000) {
            const chunks = split_text_into_chunks(artifactText, 18000);
            if (chunks.length > 1) {
                const results = [];
                for (let i = 0; i < chunks.length; i++) {
                    if (onProgress) onProgress(`산출물 용량이 커서 ${chunks.length}회로 나누어 분석을 진행합니다. (${i + 1}/${chunks.length}부 시작)`);
                    
                    if (i > 0) {
                        if (onProgress) onProgress(`Rate Limit 방지를 위해 3초 대기합니다...`);
                        await sleep_delay(3000);
                    }
                    
                    const res = await analyzeDocumentsWithLLM(guidelineText, chunks[i], inspectionScope, apiKey, glossaryText, onProgress, selectedModel, true, ragContext, llmProvider, ollamaModel);
                    results.push(res);
                }
                if (onProgress) onProgress("분석 결과 병합 중...");
                return merge_multiple_results(results, true);
            }
        }
        
        if (!isOnlyTypoCheck && guidelineText && guidelineText.length > 15000) {
            const chunks = split_text_into_chunks(guidelineText, 12000);
            if (chunks.length > 1) {
                const results = [];
                for (let i = 0; i < chunks.length; i++) {
                    if (onProgress) onProgress(`기준 문서 용량이 커서 ${chunks.length}회로 나누어 분석을 진행합니다. (${i + 1}/${chunks.length}부 시작)`);
                    
                    if (i > 0) {
                        if (onProgress) onProgress(`Rate Limit 방지를 위해 3초 대기합니다...`);
                        await sleep_delay(3000);
                    }
                    
                    const res = await analyzeDocumentsWithLLM(chunks[i], artifactText, inspectionScope, apiKey, glossaryText, onProgress, selectedModel, true, ragContext, llmProvider, ollamaModel);
                    results.push(res);
                }
                if (onProgress) onProgress("분석 결과 병합 중...");
                return merge_multiple_results(results, false);
            }
        }
    }
    
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

    let systemPrompt = '';
    if (onProgress) onProgress("분석 프롬프트 구성 중...");
    if (isOnlyTypoCheck) {
        systemPrompt = `[시스템 역할]
당신은 최고의 섬세함과 엄격함을 지닌 **'ISMP 산출물 하이브리드 품질 감사 에이전트'**입니다. 
당신의 임무는 단순한 오탈자 교정을 넘어, **[품질 5대 차원: 표현, 논리, 완결, 정합, 일관]** 관점에서 문서의 결함을 전수 조사하고 구체적인 교정안을 제시하는 것입니다.

[검토 기준 및 5대 품질 차원 핵심 규칙]
1. **표현 품질 (Expression)**: 
   - 오탈자, 띄어쓰기, 비표준 공백(\\xa0 등)을 전수 교정하되, **원문에 오류가 있을 때만** 지적하십시오.
   - [중요] 문서 내에 **이중 피동 표현**(\`~되어 집니다\`, \`~되어져야 함\`)이 **실제로 존재하는 경우에만** 지적하고 \`~됩니다\`, \`~해야 함\`으로 간결하게 교정하십시오. (원문에 없는 오류를 억지로 만들어내지 마십시오.)
   - 단위 대소문자 혼용(GB/gb, vCPU/Vcpu 등) 및 동일 개념의 다중 용어 사용이 **실제 발견될 경우에만** 지적하십시오.
   - '용어 사전' 제공 시 사전 정의된 표준 용어와의 일치 여부를 최우선 검증하십시오.
2. **논리 구조 (Logical Structure)**:
   - "Why → What → How → When" 흐름의 논리적 비약 여부, 현황/문제점과 개선 과제 간의 인과관계를 점검하십시오.
   - MECE(중복/누락 없음) 원칙 준수 여부를 확인하십시오.
   - **ID 정합성**: 기능 ID나 프로세스 ID의 일련번호 누락(Gap)이나 중복이 **명확히 확인되는 경우에만** 지적하십시오.
3. **내용 완결성 (Completeness)**:
   - 필수 섹션 누락, 이해관계자 관점 반영 부족을 도출하고, "다수", "상당수" 등 정량 데이터가 누락된 모호한 표현이 **원문에 쓰인 경우에만** 지적하십시오.
   - **I-P-O 정의 / 필수 속성**: 기능/프로세스 정의 시 '입력, 처리, 결과' 누락 또는 수행 주체, 선/후행 조건 등이 **실제 공란인 경우에만** 찾아내십시오.
4. **사업 정합성 (Strategic Alignment)**:
   - 기술된 제안 내용이 본 사업의 목적, RFP 핵심 요구사항, 최신 IT 트렌드에 비추어 구체적인 실행 방안을 담고 있는지 점검하여 '보완 권고'를 제시하십시오.
   - 예산 및 기간 측면의 현실성이 부족하거나 리스크 관리가 미흡한 경우 지적하십시오.
5. **일관성 (Consistency - 문서 내적 정합성)**:
   - 문서 내 서로 다른 페이지에서 동일 개체에 대해 명칭, 수치, 아키텍처 내역이 상충되거나 다르게 기술된 경우 '논리 상충'으로 지적하십시오.
   - AS-IS 문제점이 TO-BE에서 제대로 해소되도록 연결되어 있는지 점검하십시오.

[출력 가이드]
- 찾아낸 모든 결함을 하나도 빠짐없이 JSON 배열의 'typos' 항목에 담으십시오.
- errorType은 다음 5가지 중 하나를 선택하여 접두어로 명시하십시오: '[1. 표현 품질]', '[2. 논리 구조]', '[3. 내용 완결성]', '[4. 사업 정합성]', '[5. 일관성]'.
- **원문에 없는 오류를 스스로 지어내는 행위(Hallucination)를 엄격히 금지**하며, 확실한 결함만 도출하십시오. 중복 내역은 하나로 병합하십시오.

[중요 예외 규칙: 띄어쓰기 오류 지적 최소화 원칙]
다음의 경우는 **절대** 띄어쓰기 오류로 지적하지 마십시오:

① **IT·기술 복합 명사**: '데이터 전송', '데이터 수집', '데이터 처리', '정보 시스템', '업무 프로세스', '시스템 설계', '응용 프로그램', '네트워크 구성', '데이터 레이크', '데이터 파이프라인' 등 두 단어 이상이 결합된 IT 전문 복합 용어는 **띄어 써도 붙여 써도 모두 허용**되는 실무 관행입니다. 이를 오류로 지적하는 행위를 엄격히 금지합니다.

② **숫자+단위 붙여쓰기**: '6가지', '3개', '10명' 등 아라비아 숫자 뒤 단위/의존 명사 붙여쓰기는 절대 띄어쓰기 오류로 지적하지 마십시오.

③ **의심스러운 경우 지적 금지**: 해당 표현이 오류인지 올바른지 100% 확신할 수 없다면 지적하지 마십시오. **명백하고 확실한 오류만** 도출하십시오. (예: '데이터전 송'처럼 단어 중간에 공백이 삽입된 경우만 해당)

④ **원문 그대로 올바른 표현을 오류로 간주 금지**: AI가 스스로 "이렇게 쓰면 더 낫다"고 판단하여 올바른 원문을 오류로 지적하는 할루시네이션을 엄격히 금지합니다.

[출력 형식 및 필수 제약 사항]
[제약 1] 반드시 프론트엔드 표 렌더링을 위해 아래 JSON 데이터 배열로만 출력하라. (아래 필드명을 엄격히 유지할 것)
{
  "score": 100,
  "inspectionScope": "<점검범위 텍스트 또는 null>",
  "summary": "<전체 문서의 주요 내용 분석 및 5대 품질 차원에 기반한 종합 검토 의견 (매우 상세하게)>",
  "requirementMapping": [],
  "typos": [
    {
      "page": "<페이지 번호 또는 섹션/목차명>",
      "originalText": "<원문 문장 전체 또는 결함 내용 요약>",
      "correction": "<수정 제안 또는 구체적 보완 권고>",
      "errorType": "<'[1. 표현 품질] 오탈자', '[2. 논리 구조] 원인-결과 불일치' 등의 상세 사유>"
    }
  ]
}`;
    } else {
        systemPrompt = `당신은 최고 수준의 IT 감리 전문가이자 공공 프로젝트 산출물 검증 전문 에이전트입니다.
당신의 임무는 입력된 **'기준 문서(Base Document)'**와 **'산출물(Artifact)'**의 성격과 특성을 먼저 파악하고, 그 상관관계에 기반하여 이행 여부 및 내용적 충분성(Adequacy)을 지능적으로 검증하는 것입니다.

[검증 전 필수 분석: 문서의 특성 및 컨텍스트 파악]
- 분석 시작 전, 기준 문서와 산출물의 내용을 대조하여 각 문서가 프로젝트의 어느 단계(예: 요건 정의, 업무 프로세스 분석, 시스템 설계 등)에 해당하는지 파악하십시오.
- 입력된 문서의 특성을 고려하여 점검하십시오. (예: 기준 문서가 '프로세스 정의서'이고 산출물이 '응용아키텍처'라면, 업무 흐름이 아키텍처 컴포넌트나 인터페이스 설계에 어떻게 논리적으로 투영되었는지 도메인 지식을 활용하여 점검합니다.)

[핵심 검증 원칙 - 지능적 전수 조사]
1. **문장 단위 전수 추출 및 논리 대조**: 
   - 기준 문서의 모든 본문 문장을 독립된 요건으로 추출하고, 산출물에서 그 요건이 '문서의 목적과 성격에 맞게' 적절히 반영되었는지 확인하십시오.
2. **지능적 충분성(Adequacy) 판정 (오탈자 검사 제외)**:
   - 이 모드에서는 단순 맞춤법보다는 **내용의 실질적 완성도와 논리적 완결성**에 집중합니다. (오탈자 점검은 별도 모드이므로 여기서 수행하지 마십시오.)
   - **이행(O)**: 산출물의 특성에 맞게 기술 수준이 충분히 구체적이고 전문적으로 작성된 경우.
   - **부분 이행(△)**: 언급은 있으나 문서의 특성상 기대되는 상세도가 낮거나 실행 방안이 모호한 경우.
   - **미이행(X)**: 핵심 취지가 누락되었거나 문서 성격상 반드시 포함되어야 할 설계/수행 내용이 없는 경우.
3. **전문가적 Gap 분석**:
   - '부분 이행' 또는 '미이행' 시, 어떤 기술적/관리적 내용이 보완되어야 하는지 문서의 특성을 고려하여 구체적인 개선 방향을 'gap' 필드에 제시하십시오.
4. **구조적 결함 및 정합성 수색 (typos 배열 활용)**:
   - 오탈자가 아닌, **목차-본문 불일치, 수치 간의 모순, 존재하지 않는 기능 참조** 등 문서 전체의 구조적 결함을 발견 시 'typos' 배열에 전문적으로 기록하십시오.
   - **원문에 없는 결함을 스스로 지어내는 행위(Hallucination)를 엄격히 금지합니다.** 실제 존재하는 불일치나 모순만 지적하십시오.

[출력 형식 제한]
반드시 아래 JSON 형식으로만 출력하세요. 모든 항목은 JSON 배열 내의 개별 객체여야 합니다.
{
  "score": <총점(0~100 정수, 이행 비중 및 내용 충실도 기반)>,
  "inspectionScope": "<전달받은 점검범위 또는 null>",
  "summary": "<입력된 문서들의 특성(예: 프로세스 정의서 vs 설계서) 분석 결과와 이를 바탕으로 한 종합 검증 의견 (매우 상세하게)>",
  "requirementMapping": [
    {
      "id": "<REQ-001 부터 순차 부여>",
      "category": "<요구사항 카테고리>",
      "type": "<'필수' 또는 '선택'>",
      "levelLabel": "<'개별문장'>",
      "path": "<기준 문서 내 위치>",
      "requirement": "<기준 문서에서 추출된 개별 문장 원문 그대로>",
      "artifactSection": "<대응되는 산출물 위치 (없으면 '해당 없음')>",
      "artifactContent": "<산출물의 문서 특성에 맞춰 재구성된 설계/반영 내용 요약 (없으면 '관련 내용 없음')>",
      "status": "<'이행(O)', '부분 이행(△)', '미이행(X)' 중 택 1>",
      "gap": "<부족 사유 및 문서 특성을 고려한 구체적 보완 권고 (이행 시 null)>"
    }
  ],
  "typos": [
    {
      "location": "<위치>",
      "originalText": "<원문>",
      "correction": "<구조적 수정안>",
      "reason": "<[구조 결함], [논리 상충] 등 머리말을 포함한 분석 사유>"
    }
  ]
} `;
    }

    const userInput = isOnlyTypoCheck ? `
[시스템 지시사항]
${systemPrompt}

[입력 데이터]${glossaryText ? `\n--- 용어 사전 ---\n${String(glossaryText).substring(0, 50000)}` : ''}

--- 산출물 ---
${String(artifactText || '').substring(0, 2000000)}

--- 점검 범위 ---
${inspectionScope || '없음'}
` : `
[시스템 지시사항]
${systemPrompt}

[입력 데이터]${glossaryText ? `\n--- 용어 사전 ---\n${String(glossaryText).substring(0, 50000)}` : ''}

--- 기준 문서 ---
${String(guidelineText || '').substring(0, 500000)}

--- 산출물 ---
${String(artifactText || '').substring(0, 2000000)}

--- 점검 범위 (해당 내용이 있으면 위주로 더 엄격히 볼 것) ---
${inspectionScope || '없음'}
${ragContext ? `\n${ragContext}` : ''}
`;

    if (llmProvider === 'ollama') {
        const ollamaPrompt = `[시스템 지시사항]
당신은 ISMP 공공 산출물 검증 및 문서 교정 전문 AI 에이전트입니다.
입력된 산출물 텍스트를 전수 조사하여 반드시 아래 JSON 구조로만 결과를 생성하십시오. (추가 설명 금지)

{
  "score": 85,
  "summary": "입력된 산출물에 대한 5대 품질 관점에서의 종합 분석 소감 및 검토 의견 (최소 3문장 이상 구체적 작성)",
  "typos": [
    {
      "page": "위치/페이지",
      "originalText": "오류/보완 대상 원문 문장",
      "correction": "올바른 교정 및 보완 제안 문장",
      "errorType": "[1. 표현 품질] 띄어쓰기 및 맞춤법 교정"
    }
  ],
  "requirementMapping": [
    {
      "id": "REQ-001",
      "category": "기능",
      "type": "필수",
      "levelLabel": "개별문장",
      "path": "본문",
      "requirement": "주요 요구사항 또는 핵심 명세",
      "artifactSection": "산출물 관련 위치",
      "artifactContent": "반영 내역",
      "status": "이행(O)",
      "gap": null
    }
  ]
}

[분석 대상 데이터]
${userInput}
`;
        return await analyze_with_ollama(ollamaPrompt, ollamaModel, onProgress);
    }

    try {

        let initialModel = selectedModel && selectedModel !== 'auto' ? selectedModel : FALLBACK_MODELS[0];
        if (!initialModel.startsWith('models/')) initialModel = `models/${initialModel}`;
        
        let currentModelIndex = FALLBACK_MODELS.indexOf(initialModel);
        if (currentModelIndex === -1) currentModelIndex = 0;

        const fetchWithRetry = async (maxModelRetries = FALLBACK_MODELS.length) => {
            let modelRetries = 0;
            const error_log = []; // 각 시도별 실패 원인 수집
            
            while (modelRetries < maxModelRetries) {
                const activeKey = keys[currentKeyIndex];
                const keyLabel = `키${currentKeyIndex + 1}(${activeKey.substring(0, 8)}...)`;
                const modelId = FALLBACK_MODELS[currentModelIndex];
                const modelLabel = modelId.split('/').pop();
                const fetchUrl = `https://generativelanguage.googleapis.com/v1beta/${modelId}:generateContent?key=${activeKey}`;
                
                const fetchOptions = {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({
                        contents: [{ role: "user", parts: [{ text: userInput }] }],
                        generationConfig: { temperature: 0.1 }
                    })
                };

                if (onProgress) {
                    const keyInfo = keys.length > 1 ? ` (키 ${currentKeyIndex + 1}/${keys.length} 사용 중)` : '';
                    onProgress(`${modelLabel} 모델로 분석 요청 중...${keyInfo}`);
                }

                let response;
                try {
                    response = await fetch_with_timeout(fetchUrl, { ...fetchOptions, timeout: 25000 });
                } catch (fetchErr) {
                    const reason = fetchErr.name === 'AbortError' ? '25초 타임아웃 초과' : `네트워크 오류(${fetchErr.message})`;
                    error_log.push(`[${modelLabel} / ${keyLabel}] ${reason}`);
                    console.warn(`Fetch failed or timed out for ${modelId}:`, fetchErr);
                    if (keys.length > 1 && (currentKeyIndex + 1) < keys.length) {
                        currentKeyIndex++;
                        continue;
                    }
                    modelRetries++;
                    if (modelRetries < maxModelRetries) {
                        currentKeyIndex = 0;
                        currentModelIndex = (currentModelIndex + 1) % FALLBACK_MODELS.length;
                        await sleep_delay(5000);
                        continue;
                    }
                    throw new Error(`네트워크 타임아웃 또는 연결 지연이 반복되어 분석을 완료하지 못했습니다.\n\n[시도 기록]\n${error_log.join('\n')}`);
                }
                
                if (response.ok) {
                    recordUsage(modelId); // 사용량 기록
                    return response;
                }

                const errData = await response.json().catch(() => ({}));
                const errMsg = errData.error?.message || response.statusText || '';
                const httpStatus = response.status;
                const isModelUnavailable = httpStatus === 404
                    || httpStatus === 400
                    || errMsg.toLowerCase().includes('not found')
                    || errMsg.toLowerCase().includes('not supported')
                    || errMsg.toLowerCase().includes('deprecated');

                // 상세 실패 원인 분류 및 기록
                let fail_reason = `HTTP ${httpStatus}`;
                if (httpStatus === 429) fail_reason = `429 Rate Limit (할당량 초과)`;
                else if (httpStatus === 401) fail_reason = `401 인증 실패 (API 키 오류)`;
                else if (httpStatus === 403) fail_reason = `403 접근 거부 (키 권한 없음)`;
                else if (httpStatus === 404) fail_reason = `404 모델 없음 (지원 종료)`;
                else if (httpStatus >= 500) fail_reason = `${httpStatus} 서버 오류`;
                if (errMsg) fail_reason += ` - ${errMsg.substring(0, 100)}`;
                error_log.push(`[${modelLabel} / ${keyLabel}] ${fail_reason}`);

                if (keys.length > 1 && (currentKeyIndex + 1) < keys.length) {
                    currentKeyIndex++;
                    continue;
                }

                if (response.status === 429 || response.status >= 500 || isModelUnavailable) {
                    modelRetries++;
                    if (modelRetries < maxModelRetries) {
                        currentKeyIndex = 0;
                        currentModelIndex = (currentModelIndex + 1) % FALLBACK_MODELS.length;
                        await sleep_delay(5000);
                        continue;
                    }
                    
                    throw new Error(`모든 API 키와 모델의 사용 한도가 소진되었습니다.\n\n[시도별 실패 원인]\n${error_log.join('\n')}`);
                }
                
                throw new Error(`API 호출 실패: ${fail_reason}\n\n[시도 기록]\n${error_log.join('\n')}`);
            }
            throw new Error(`모든 모델을 시도했으나 응답을 받지 못했습니다.\n\n[시도 기록]\n${error_log.join('\n')}`);
        };


        const response = await fetchWithRetry();
        const data = await response.json();
        let content = data.candidates?.[0]?.content?.parts?.[0]?.text || "{}";

        if (content.includes("```")) {
            const match = content.match(/```(?:json)?\\s*([\\s\\S]*?)\\s*```/i);
            if (match && match[1]) {
                content = match[1];
            } else {
                content = content.replace(/```(?:json)?/gi, '').replace(/```/g, '');
            }
        }
        
        content = content.trim();
        const parsed = JSON.parse(content);

        if (parsed.requirementMapping && Array.isArray(parsed.requirementMapping)) {
            if (!parsed.rtm) {
                parsed.rtm = parsed.requirementMapping.map(req => ({
                    type: req.type || '필수',
                    requirement: req.requirement || '-',
                    status: req.status || '미이행(X)',
                    location: req.artifactSection || '해당 없음',
                    category: req.category || '-',
                    levelLabel: req.levelLabel || '개별문장'
                }));
            }
            if (!parsed.omissions) {
                parsed.omissions = parsed.requirementMapping
                    .filter(req => req.status !== '이행(O)')
                    .map(req => ({
                        title: `[ID: ${req.id || 'N/A'}] ${String(req.requirement || '').substring(0, 30)}...`,
                        evidence: req.requirement || '-',
                        reason: req.gap || '구체적인 수행/설계 방안이 누락되었습니다.',
                        recommendation: '해당 요건을 만족하기 위한 구체적인 명세와 실행계획을 산출물에 추가해야 합니다.'
                    }));
            }
        } else {
            parsed.requirementMapping = [];
            parsed.rtm = [];
            parsed.omissions = [];
        }
        
        if (!parsed.typos) {
            parsed.typos = [];
        } else {
            const uniqueTypos = [];
            const seen = new Set();
            parsed.typos.forEach(typo => {
                const signature = `${typo.page}_${typo.originalText}_${typo.correction}`;
                if (!seen.has(signature)) {
                    seen.add(signature);
                    uniqueTypos.push(typo);
                }
            });
            parsed.typos = uniqueTypos;
        }

        return parsed;
    } catch (e) {
        throw new Error(`Gemini 검증 실패: ${e.message}`);
    }
}

export async function askRagQuestion(docTitle, docContent, question, apiKey, onProgress) {
    const keys = String(apiKey).split(',').map(k => k.trim()).filter(k => k.match(/^(AIza|AQ\.)/));
    if (keys.length === 0) throw new Error("유효한 Gemini API Key가 없습니다.");

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

    const systemPrompt = `당신은 ISMP 산출물 전문 Q&A 어시스턴트입니다. 
제공된 문서 [${docTitle}]의 내용을 바탕으로 사용자의 질문에 전문적이고 친절하게 답변하십시오.`;

    const userInput = `
[문서 제목]: ${docTitle}
[사용자 질문]: ${question}
`;


    let currentKeyIndex = 0;
    let currentModelIndex = 0;

    const fetchWithRetry = async () => {
        const activeKey = keys[currentKeyIndex];
        const modelId = FALLBACK_MODELS[currentModelIndex];
        const fetchUrl = `https://generativelanguage.googleapis.com/v1beta/${modelId}:generateContent?key=${activeKey}`;
        
        const response = await fetch(fetchUrl, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                contents: [{ role: "user", parts: [{ text: userInput }] }]
            })
        });

        if (response.ok) {
            const data = await response.json();
            const answer = data.candidates?.[0]?.content?.parts?.[0]?.text;
            recordUsage(modelId);
            return answer;
        }
        throw new Error("Failed to fetch");
    };

    return await fetchWithRetry();
}

export async function askTotalRagQuestion(question, contextDocs, apiKey, onProgress) {
    const keys = String(apiKey).split(',').map(k => k.trim()).filter(k => k.match(/^(AIza|AQ\.)/));
    if (keys.length === 0) throw new Error("유효한 Gemini API Key가 없습니다.");

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

    let contextText = "";
    if (contextDocs && contextDocs.length > 0) {
        contextDocs.forEach((doc, idx) => {
            contextText += `\n[참고자료 ${idx + 1}: ${doc.title}]\n`;
            contextText += doc.content.substring(0, 1500) + (doc.content.length > 1500 ? "..." : "") + "\n";
        });
    } else {
        contextText = "관련된 참고 자료를 찾지 못했습니다.";
    }

    const systemPrompt = `당신은 공공 IT 사업 제안서·수행계획서·설계서 분야에 특화된 한국어 맞춤법 및 문서 교정 전문가이자 최고의 IT 전략 컨설턴트입니다.
제공되는 [참고 지식베이스 내용]을 기반으로 사용자의 질문에 정확하고 풍부한 내용으로 친절하게 답변해 주십시오. 
만약 참고 자료에 핵심 답변이 부재한 경우, 본인이 가지고 있는 IT 상식을 동원하여 구체적인 로드맵이나 대응 방안을 제시하고 출처를 밝혀주십시오.`;

    const userInput = `
--- [참고 지식베이스 내용] ---
${contextText}

--- [사용자 질문] ---
${question}
`;

    let currentKeyIndex = 0;
    let currentModelIndex = 0;

    const fetchWithRetry = async () => {
        const activeKey = keys[currentKeyIndex];
        const modelId = FALLBACK_MODELS[currentModelIndex];
        const fetchUrl = `https://generativelanguage.googleapis.com/v1beta/${modelId}:generateContent?key=${activeKey}`;
        
        const response = await fetch(fetchUrl, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                contents: [
                    { role: "user", parts: [{ text: systemPrompt + "\n" + userInput }] }
                ]
            })
        });

        if (response.ok) {
            const data = await response.json();
            const answer = data.candidates?.[0]?.content?.parts?.[0]?.text;
            recordUsage(modelId);
            return answer;
        }
        throw new Error("Failed to fetch");
    };

    return await fetchWithRetry();
}
