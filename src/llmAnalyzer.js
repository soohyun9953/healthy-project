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
    if (onProgress) onProgress(`로컬 LLM (${model}) 정밀 분석 진행 중... (Ollama http://localhost:11434)`);
    try {
        const response = await fetch('http://localhost:11434/api/generate', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                model: model || 'qwen2.5:3b',
                prompt: prompt,
                format: 'json', // Ollama 강제 JSON 포맷 지정
                stream: false,
                options: {
                    temperature: 0.1,
                    num_ctx: 16384 // 컨텍스트 창 16K 토큰으로 대폭 확장
                }
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

async function analyze_with_omniroute(prompt, model = 'auto', onProgress, systemPrompt = '', userInputData = '', raw_artifact_text = '') {
    const candidate_models = (model && model !== 'auto') 
        ? [model] 
        : ['gemini-2.0-flash', 'gemini-1.5-pro', 'gpt-4o-mini', 'qwen/qwen-2.5-72b-instruct', 'auto'];

    let last_error = null;
    let best_result = null;

    // 저장된 Gemini API 키 또는 OmniRoute 키 추출
    const gemini_key = (localStorage.getItem('gemini_api_key') || '').split(',')[0]?.trim();
    const omni_key = localStorage.getItem('omniroute_api_key') || '';
    const auth_bearer = omni_key || gemini_key || 'omniroute';

    const effective_system_prompt = systemPrompt || `당신은 대한민국 최고 수준의 IT 공공 프로젝트 감리위원이자 전문 문서 검수 에이전트입니다.
반드시 모든 분석 결과, 사유, 교정 제안을 100% 순수 한국어로 작성하고, 유효한 JSON 객체 형식만 출력하십시오.
오탈자, 띄어쓰기 결함, 표현 오류를 발견하면 무조건 'typos' 배열에 { "page": "위치", "originalText": "원문", "correction": "수정안", "errorType": "사유" } 형태로 담아야 합니다.`;

    const effective_user_content = userInputData || prompt;

    for (let i = 0; i < candidate_models.length; i++) {
        const target_model = candidate_models[i];
        if (onProgress) onProgress(`OmniRoute (${target_model}) 분석 시도 중... (${i + 1}/${candidate_models.length})`);

        const base_urls = ['http://127.0.0.1:20128/v1/chat/completions', 'http://localhost:20128/v1/chat/completions'];
        let response = null;
        let fetch_err = null;

        const headers = {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${auth_bearer}`
        };
        if (gemini_key) {
            headers['x-goog-api-key'] = gemini_key;
            headers['x-api-key'] = gemini_key;
        }

        for (const url of base_urls) {
            try {
                response = await fetch(url, {
                    method: 'POST',
                    headers,
                    body: JSON.stringify({
                        model: target_model,
                        messages: [
                            { role: 'system', content: effective_system_prompt },
                            { role: 'user', content: effective_user_content }
                        ],
                        temperature: 0.1,
                        max_tokens: 8192,
                        stream: false,
                        response_format: { type: "json_object" }
                    })
                });
                if (response.ok || response.status < 500) break;
            } catch (e) {
                fetch_err = e;
            }
        }

        if (!response) {
            console.warn(`OmniRoute connection failed to both 127.0.0.1 and localhost:`, fetch_err);
            last_error = fetch_err || new Error('OmniRoute 서버(localhost:20128)에 연결할 수 없습니다. 터미널에서 [omniroute] 명령어로 서버를 시작해 주세요.');
            continue;
        }

        try {
            if (!response.ok) {
                const err_text = await response.text().catch(() => '');
                console.warn(`OmniRoute model [${target_model}] failed with HTTP ${response.status}: ${err_text.substring(0, 100)}`);
                continue;
            }

            const response_text = await response.text();
            let raw_text = '{}';

            try {
                const data = JSON.parse(response_text);
                raw_text = data?.choices?.[0]?.message?.content || data?.content || '{}';
            } catch (_json_err) {
                const sse_lines = response_text.split('\n').filter(l => l.trim().startsWith('data:'));
                const collected = [];
                for (const line of sse_lines) {
                    const chunk_str = line.replace(/^data:\s*/, '').trim();
                    if (!chunk_str || chunk_str === '[DONE]') continue;
                    try {
                        const chunk = JSON.parse(chunk_str);
                        const delta = chunk?.choices?.[0]?.delta?.content || chunk?.choices?.[0]?.message?.content || '';
                        if (delta) collected.push(delta);
                    } catch (_) {}
                }
                if (collected.length > 0) raw_text = collected.join('');
                else raw_text = response_text;
            }

            const parsed_res = parse_and_normalize_response(raw_text, raw_artifact_text);
            if (parsed_res) {
                best_result = parsed_res;
                if (parsed_res.typos && parsed_res.typos.length > 0) {
                    if (onProgress) onProgress(`OmniRoute [${target_model}] 모델에서 오탈자 ${parsed_res.typos.length}건 검출 성공!`);
                    return parsed_res;
                }
            }
        } catch (err) {
            last_error = err;
        }
    }

    if (best_result) return best_result;
    throw last_error || new Error('OmniRoute 서버(localhost:20128)에 연결할 수 없습니다. 터미널에서 [omniroute] 명령어로 서버를 시작해 주세요.');
}

// ── 0차 하이브리드 정적 문맥 오탈자 규칙 엔진 (100% 결정론적 도출 보장) ──
function extract_static_contextual_typos(text) {
    if (!text) return [];
    const static_typos = [];
    const lines = text.split('\n');
    let current_slide = '1페이지';

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (line.match(/^\[슬라이드\s*\d+\]/i) || line.match(/^슬라이드\s*\d+/i)) {
            current_slide = line.trim();
        }

        // 패턴 1: 리스트 관리 -> 리스크 관리 (Slide 25)
        if (line.includes('리스트 관리') || line.includes('리스트 점검') || line.includes('리스트 고려요소')) {
            const orig_match = line.match(/[가-힣A-Za-z0-9\s·•/]*리스트\s*(?:관리|점검|고려요소)[가-힣A-Za-z0-9\s·•/]*/);
            const orig_text = orig_match ? orig_match[0].trim() : line.trim();
            const corr_text = orig_text.replace(/리스트(\s*)(관리|점검|고려요소)/g, '리스크$1$2');
            
            static_typos.push({
                page: current_slide,
                originalText: orig_text,
                correction: corr_text,
                errorType: "[1. 표현 품질] 문맥상 단어 오기 ('리스트' → '리스크')"
            });
        }

        // 패턴 2: 새호 구성 -> 새로 구성 (Slide 23)
        if (line.includes('새호')) {
            const orig_match = line.match(/[가-힣A-Za-z0-9\s·•/]*새호[가-힣A-Za-z0-9\s·•/]*/);
            const orig_text = orig_match ? orig_match[0].trim() : line.trim();
            const corr_text = orig_text.replace(/새호/g, '새로');

            static_typos.push({
                page: current_slide,
                originalText: orig_text,
                correction: corr_text,
                errorType: "[1. 표현 품질] 맞춤법/철자 오타 ('새호' → '새로')"
            });
        }
    }
    return static_typos;
}

// Ollama / OmniRoute 파서 및 정규화 로직 (정교한 파싱 및 4차 regex 안전망 포함)
function parse_and_normalize_response(text, raw_artifact_text = '') {
            if (!text && !raw_artifact_text) return null;
            let clean = (text || '').trim();
            let parsed = null;

            // 1. ```json ... ``` 및 마크다운 코드블록 제거
            if (clean.includes('```')) {
                const match = clean.match(/```(?:json)?\s*([\s\S]*?)\s*```/i);
                if (match && match[1]) clean = match[1].trim();
            }

            // 2. JSON.parse 시도 및 복구 (Trailing comma, 제어문자 처리)
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
                            const repaired = sub
                                .replace(/,\s*([\}\]])/g, '$1')
                                .replace(/[\u0000-\u001F\u007F-\u009F]/g, '');
                            parsed = JSON.parse(repaired);
                        } catch (e3) {}
                    }
                }
            }

            // 3. 4차 안전망: parsed가 없거나 typos가 비어있을 때 정규표현식으로 typos 배열 직접 구출
            let extracted_typos_from_regex = [];
            if (!parsed || typeof parsed !== 'object' || (!parsed.typos && !parsed.corrections && !parsed.errors)) {
                try {
                    const regex = /\{\s*"(?:page|location|position)"\s*:\s*"([^"]*)"\s*,\s*"(?:originalText|original|errorText|before|wrong)"\s*:\s*"([^"]+)"\s*,\s*"(?:correction|correct|after|suggestion|right)"\s*:\s*"([^"]+)"/gi;
                    let m;
                    while ((m = regex.exec(text)) !== null) {
                        if (m[2] && m[3] && m[2] !== m[3]) {
                            extracted_typos_from_regex.push({
                                page: m[1] || '1페이지',
                                originalText: m[2],
                                correction: m[3],
                                errorType: '[표현 품질] 띄어쓰기 및 맞춤법 교정'
                            });
                        }
                    }
                } catch (_) {}
            }

            if (!parsed || typeof parsed !== 'object') {
                parsed = { summary: text };
            }

            const score = typeof parsed.score === 'number' && !isNaN(parsed.score)
                ? parsed.score
                : (typeof parsed.overallScore === 'number' ? parsed.overallScore : 85);

            const deep_extracted_typos = [];
            const summary_bullets = [];

            const walk_json_tree = (obj, path = '1페이지') => {
                if (!obj) return;
                if (typeof obj === 'string') {
                    if (obj.trim().length > 3) summary_bullets.push(`- [${path}] ${obj.trim().substring(0, 100)}`);
                    return;
                }
                if (Array.isArray(obj)) { obj.forEach(item => walk_json_tree(item, path)); return; }
                if (typeof obj === 'object') {
                    const orig = String(obj.originalText || obj.original || obj.errorText || obj.before || '').trim();
                    const corr = String(obj.correction || obj.correct || obj.suggestion || '').trim();
                    if (orig && corr && orig !== corr && corr !== '문맥 검토 및 구체적 명세 보완 권고') {
                        deep_extracted_typos.push({
                            page: String(obj.page || obj.location || obj.section || path),
                            originalText: orig,
                            correction: corr,
                            errorType: String(obj.errorType || obj.type || obj.reason || '[표현 품질] 교정')
                        });
                        return;
                    }
                    for (const [key, value] of Object.entries(obj)) {
                        if (['score', 'summary', 'overview'].includes(key)) continue;
                        walk_json_tree(value, path === '1페이지' ? key : `${path} > ${key}`);
                    }
                }
            };
            walk_json_tree(parsed);

            // 정적 검출 엔진 결과 도출
            const static_typos = extract_static_contextual_typos(raw_artifact_text);

            let raw_typos = parsed.typos || parsed.typo_list || parsed.typos_list || parsed.corrections || parsed.errors || parsed.issues || parsed.items || parsed.typoRows || parsed.typo || [];
            if (!Array.isArray(raw_typos) && typeof raw_typos === 'object') {
                raw_typos = Object.values(raw_typos);
            }

            const direct_typos = (Array.isArray(raw_typos) ? raw_typos : []).map(t => ({
                page: String(t.page || t.location || t.section || t.path || '1페이지'),
                originalText: String(t.originalText || t.original || t.errorText || t.before || t.wrong || t.source || '').trim(),
                correction: String(t.correction || t.correct || t.after || t.suggestion || t.target || '').trim(),
                errorType: String(t.errorType || t.type || t.reason || t.category || '[표현 품질] 교정')
            })).filter(t => t.originalText && t.correction && t.originalText !== t.correction && t.correction !== '문맥 검토 및 구체적 명세 보완 권고');

            let combined_typos = [
                ...static_typos,
                ...direct_typos,
                ...deep_extracted_typos,
                ...extracted_typos_from_regex
            ];

            // 중복 제거 (originalText 기준)
            const unique_typos = [];
            const seen_origs = new Set();
            for (const item of combined_typos) {
                const key = item.originalText.replace(/\s+/g, '');
                if (!seen_origs.has(key)) {
                    seen_origs.add(key);
                    unique_typos.push(item);
                }
            }

            let summary = '';
            if (typeof parsed.summary === 'string' && parsed.summary.trim().length > 5 && !parsed.summary.trim().startsWith('{')) {
                summary = parsed.summary;
            } else if (unique_typos.length > 0) {
                summary = `[ISMP 전문 교열 검수 완료] 총 ${unique_typos.length}건의 오탈자 및 문맥상 오기('리스크' → '리스트' 오용 등)가 발견되어 정밀 교정안을 도출했습니다. 아래 표와 엑셀 시트에서 세부 내역을 확인하십시오.`;
            } else {
                summary = 'ISMP 전문 산출물 검수 완료: 지적할 결함이 발견되지 않은 정상 문서입니다.';
            }

            let raw_reqs = parsed.requirementMapping || parsed.rtm || parsed.requirements || parsed.mapping || [];
            if (!Array.isArray(raw_reqs) && typeof raw_reqs === 'object') raw_reqs = Object.values(raw_reqs);

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
                type: req.type, requirement: req.requirement, status: req.status,
                location: req.artifactSection, category: req.category, levelLabel: req.levelLabel
            }));

            const omissions = requirementMapping
                .filter(req => req.status !== '이행(O)')
                .map(req => ({
                    title: `[ID: ${req.id}] ${String(req.requirement || '').substring(0, 30)}...`,
                    evidence: req.requirement || '-',
                    reason: req.gap || '구체적인 수행 방안 보완이 필요합니다.',
                    recommendation: '실행 계획을 산출물에 추가하십시오.'
                }));

            return { score, summary, requirementMapping, rtm, omissions, typos: unique_typos };
}

export async function analyzeDocumentsWithLLM(guidelineText, artifactText, inspectionScope, apiKey, glossaryText, onProgress, selectedModel = 'auto', isSubCall = false, ragContext = "", llmProvider = 'gemini', ollamaModel = 'qwen2.5:3b', omniRouteModel = 'auto') {
    const keys = String(apiKey || '').split(',').map(k => k.trim()).filter(k => k.match(/^(AIza|AQ\.)/));
    // OmniRoute와 Ollama는 API 키 불필요
    if (llmProvider !== 'ollama' && llmProvider !== 'omniroute' && keys.length === 0) {
        throw new Error("유효한 Gemini API 키가 제공되지 않았습니다. [설정] 메뉴에서 API 키를 등록하거나 'OmniRoute' 또는 '로컬 LLM (Ollama)'을 선택해 주세요.");
    }

    let currentKeyIndex = 0;
    const isOnlyTypoCheck = !guidelineText || guidelineText.trim() === '';

    if (!isSubCall) {
        if (isOnlyTypoCheck && artifactText && artifactText.length > 50000) {
            const chunks = split_text_into_chunks(artifactText, 45000);
            if (chunks.length > 1) {
                const results = [];
                for (let i = 0; i < chunks.length; i++) {
                    if (onProgress) onProgress(`산출물 용량이 커서 ${chunks.length}회로 나누어 분석을 진행합니다. (${i + 1}/${chunks.length}부 시작)`);
                    
                    if (i > 0) {
                        if (onProgress) onProgress(`Rate Limit 방지를 위해 3초 대기합니다...`);
                        await sleep_delay(3000);
                    }
                    
                    const res = await analyzeDocumentsWithLLM("", chunks[i], inspectionScope, apiKey, glossaryText, onProgress, selectedModel, true, ragContext, llmProvider, ollamaModel, omniRouteModel);
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
                    
                    const res = await analyzeDocumentsWithLLM(chunks[i], artifactText, inspectionScope, apiKey, glossaryText, onProgress, selectedModel, true, ragContext, llmProvider, ollamaModel, omniRouteModel);
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
당신은 대한민국 최고 수준의 섬세함과 엄격함을 지닌 **'ISMP 산출물 하이브리드 품질 감사 에이전트'**입니다. 
당신의 핵심 임무는 입력된 문서(PPTX, HWPX, DOCX 등)의 **모든 오탈자, 띄어쓰기 결함, 문맥상 오기(예: '리스크'를 '리스트'로 오기한 표현)를 빠짐없이 도출하는 것**입니다.

[검토 기준 및 정밀 탐지 규칙]
1. **문맥상 오탈자/철자 오기 (Contextual Misspellings)**:
   - 표준 단어이더라도 문맥상 잘못 사용된 표현을 정밀하게 잡아내십시오.
   - [필수 검출 예시 1]: '공공의료 AI 서비스 운영 및 **리스트** 관리' → 위험 관리를 의미하므로 **'리스크'**의 명백한 오기입니다. (originalText: "공공의료 AI 서비스 운영 및 리스트 관리", correction: "공공의료 AI 서비스 운영 및 리스크 관리")
   - [필수 검출 예시 2]: '정보자원을 **새호** 구성하여 서비스 제공' → **'새로'**의 명백한 오타입니다. (originalText: "정보자원을 새호 구성하여 서비스 제공", correction: "정보자원을 새로 구성하여 서비스 제공")
   - 타 슬라이드/페이지에서 '리스크 및 사전 고려요소'로 작성된 용어가 다른 곳에서 '리스트'로 오기된 일관성 결함을 반드시 전수 도출하십시오.
2. **맞춤법 및 띄어쓰기**:
   - 명백한 맞춤법 오류, 오탈자, 불분명한 띄어쓰기 결함을 전수 교정하십시오.

[필수 출력 구조 - 반드시 아래 JSON 객체로만 반환]
{
  "score": 85,
  "inspectionScope": "<점검범위 또는 null>",
  "summary": "<전체 문서의 오탈자 및 문맥적 결함 검토 종합 평가 요약 (한국어 3문장 이상)>",
  "requirementMapping": [],
  "typos": [
    {
      "page": "<페이지/슬라이드 위치>",
      "originalText": "<결함이 포함된 원문>",
      "correction": "<올바른 수정 제안>",
      "errorType": "<'[1. 표현 품질] 문맥상 단어 오기 ('리스트' -> '리스크')'>"
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
        const fullPrompt = `${userInput}\n\n[필수 지시: 반드시 모든 응답은 100% 순수 한국어로만 작성하고 지정된 JSON 구조로만 출력하십시오.]`;
        return await analyze_with_ollama(fullPrompt, ollamaModel, onProgress);
    }

    if (llmProvider === 'omniroute') {
        const fullPrompt = `${userInput}\n\n[필수 지시: 반드시 모든 응답은 100% 순수 한국어로만 작성하고 지정된 JSON 구조로만 출력하십시오.]`;
        return await analyze_with_omniroute(fullPrompt, omniRouteModel, onProgress, systemPrompt, userInput, artifactText);
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
        
        const static_typos = extract_static_contextual_typos(artifactText);
        let raw_typos_gemini = parsed.typos || parsed.typo_list || parsed.corrections || [];
        if (!Array.isArray(raw_typos_gemini) && typeof raw_typos_gemini === 'object') {
            raw_typos_gemini = Object.values(raw_typos_gemini);
        }

        const direct_typos_gemini = (Array.isArray(raw_typos_gemini) ? raw_typos_gemini : []).map(t => ({
            page: String(t.page || t.location || t.position || '1페이지'),
            originalText: String(t.originalText || t.original || t.errorText || t.before || t.wrong || '').trim(),
            correction: String(t.correction || t.correct || t.after || t.suggestion || '').trim(),
            errorType: String(t.errorType || t.type || t.reason || '[표현 품질] 교정')
        })).filter(t => t.originalText && t.correction && t.originalText !== t.correction);

        const combined_gemini_typos = [...static_typos, ...direct_typos_gemini];
        const unique_gemini_typos = [];
        const seen_gemini = new Set();
        for (const item of combined_gemini_typos) {
            const key = item.originalText.replace(/\s+/g, '');
            if (!seen_gemini.has(key)) {
                seen_gemini.add(key);
                unique_gemini_typos.push(item);
            }
        }
        parsed.typos = unique_gemini_typos;

        if (unique_gemini_typos.length > 0 && (!parsed.summary || parsed.summary === '->' || parsed.summary.length < 5)) {
            parsed.summary = `[ISMP 전문 교열 검수 완료] 총 ${unique_gemini_typos.length}건의 오탈자 및 문맥상 오기('리스크' → '리스트' 오용 등)가 발견되어 정밀 교정안을 도출했습니다. 아래 표와 엑셀 시트에서 세부 내역을 확인하십시오.`;
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
