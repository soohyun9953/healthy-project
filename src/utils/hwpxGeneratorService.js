import JSZip from 'jszip';
import { processFile } from './fileExtractor.js';
import { FALLBACK_MODELS } from './geminiModels.js';

/**
 * 지정된 LLM API 엔진(Gemini, OmniRoute)을 호출하여 HWPX 템플릿의 문단을 소스 문서들의 내용에 맞게 치환할 텍스트 맵을 생성합니다.
 */
export async function generateHwpxReportWithLLM(paragraphs, materialsText, apiKey, instruction = '', selectedModel = 'auto', llmProvider = 'gemini', omniRouteModel = 'auto', onProgress) {
    if (onProgress) onProgress("AI 보고서 생성 프롬프트 조립 중...");

    const systemPrompt = `[시스템 역할]
당신은 최고 권위의 IT 감리 및 표준 보고서 작성을 전문으로 하는 **'AI 한글 보고서 보좌관'**입니다.
당신의 임무는 입력된 **'참고 자료(Materials)'**의 핵심 내용, 데이터, 기술 명세, 실행 계획 등을 깊이 있게 분석 및 요약하여, **'HWPX 보고서 템플릿(Template Paragraphs)'**의 알맞은 영역에 서술형 보고서로 완성해주는 것입니다.

[작성 및 치환 원칙]
1. **문맥 기반 정밀 치환**: 
   - 템플릿 내 문단 목록 중 치환 대상 기호(예: \`{{내용}}\`, \`[작성란]\`, \`[여기에 작성]\` 등)가 있거나, 본문 텍스트 작성이 필요한 빈 구획을 감지하여 참고 자료의 내용을 바탕으로 보고서를 상세하게 전개하십시오.
   - 고정된 제목(예: "1. 사업 개요", "가. 추진 배경")이나 템플릿 고유 메타 정보는 **절대 치환하거나 수정하지 말고 원래 텍스트를 그대로 유지**해야 합니다. 오직 작성 및 요약이 필요한 '본문 영역' 위주로만 새로운 내용을 서술하십시오.
2. **보고서 문체 준수**:
   - 격식 있는 개조식(예: \`~함\`, \`~구축함\`, \`~필요\`) 또는 정중한 서술식(예: \`~합니다\`, \`~예정입니다\`) 등 보고서 문맥에 맞는 문체로 통일하여 전문성 높게 구성하십시오.
3. **스타일 구조 유지**:
   - 템플릿 문단 목록의 ID(예: \`para_0\`, \`para_1\`)를 보존하여 결과 JSON 맵으로 반환해야 합니다.
   - 치환될 본문 내용이 길어 여러 문단으로 구성되어야 하는 경우, **줄바꿈(\\n)** 문자를 넣어 구성하십시오. 융합 엔진이 줄바꿈 단락을 감지하여 원래 문단 서식 스타일(글자 크기, 폰트, 들여쓰기)을 복제하여 여러 문단으로 자동 확장해 줍니다.
4. **추가 요구사항 반영**:
   - 사용자가 '추가 지시사항(Instruction)'을 기재한 경우, 이를 최우선으로 반영하여 보고서 텍스트를 개조하거나 보완하십시오.

[출력 형식 및 제약 사항]
[제약 1] 반드시 아래 형식의 JSON 딕셔너리로만 출력하십시오. 추가적인 머리말, 꼬리말, 설명 텍스트, 백틱 마크(\`\`\`json)는 엄격히 금지됩니다. 오직 파싱 가능한 순수 JSON 오브젝트만 반환해야 합니다.
[제약 2] 템플릿에 없는 새로운 ID를 지어내거나 기존 문단의 고유 구조를 파괴하지 마십시오.

{
  "para_0": "수정 및 채워 넣을 새로운 문장 1 (또는 여러 문단일 경우\\n새로운 문장 2)",
  "para_3": "수정 및 채워 넣을 새로운 문장..."
}

(※ 내용이 변경될 필요가 없는 고정 문단이나 제목은 결과 JSON 오브젝트의 키에서 생략하여 전송량을 최소화하십시오.)`;

    const paragraphsJson = JSON.stringify(paragraphs, null, 2);

    const userInput = `
[시스템 지시사항]
${systemPrompt}

[사용자 추가 지시사항]
${instruction ? instruction : '없음 (템플릿의 맥락과 참고자료의 흐름에 맞춰 표준 보고서 스타일로 본문을 상세히 작성해 주세요.)'}

--- 참고 자료 (PPT, HWPX, MD 등 병합 텍스트) ---
${materialsText.substring(0, 800000)}

--- HWPX 보고서 템플릿 문단 목록 (ID 및 원문) ---
${paragraphsJson}
`;

    let initialModel = selectedModel && selectedModel !== 'auto' ? selectedModel : FALLBACK_MODELS[0];
    if (!initialModel.startsWith('models/')) initialModel = `models/${initialModel}`;
    
    let currentModelIndex = FALLBACK_MODELS.indexOf(initialModel);
    if (currentModelIndex === -1) currentModelIndex = 0;
    let currentKeyIndex = 0;

    const fetchWithRetry = async (maxModelRetries = FALLBACK_MODELS.length) => {
        let apiURL = '';
        let headers = { 'Content-Type': 'application/json' };
        let reqBody = {};
        
        if (llmProvider === 'omniroute') {
            // OmniRoute 로컬 프록시 호출
            const modelName = omniRouteModel === 'auto' ? 'auto' : omniRouteModel;
            apiURL = 'http://localhost:20128/v1/chat/completions';
            if (onProgress) onProgress(`OmniRoute 호출 중... (모델: ${modelName})`);
            
            const omniKey = localStorage.getItem('omniroute_api_key') || 'omniroute';
            headers['Authorization'] = `Bearer ${omniKey}`;
            
            reqBody = {
                model: modelName,
                messages: [
                    { role: 'system', content: systemPrompt },
                    { role: 'user', content: userInput }
                ],
                temperature: 0.2,
                response_format: { type: "json_object" }
            };
        } else {
            // 구글 Gemini API 직접 호출
            const keys = String(apiKey).split(',').map(k => k.trim()).filter(k => k.match(/^(AIza|AQ\.)/));
            if (keys.length === 0) {
                throw new Error("유효한 Gemini API 키가 제공되지 않았습니다. [설정] 메뉴에서 API 키를 등록해 주세요.");
            }
            
            const modelToUse = FALLBACK_MODELS[currentModelIndex] || FALLBACK_MODELS[0];
            const keyToUse = keys[currentKeyIndex];
            
            if (onProgress) onProgress(`Gemini API 호출 중... (모델: ${modelToUse.split('/').pop()})`);
            
            apiURL = `https://generativelanguage.googleapis.com/v1beta/${modelToUse}:generateContent?key=${keyToUse}`;
            reqBody = {
                contents: [{ parts: [{ text: userInput }] }],
                generationConfig: {
                    responseMimeType: "application/json"
                }
            };
        }

        try {
            const response = await fetch(apiURL, {
                method: 'POST',
                headers,
                body: JSON.stringify(reqBody)
            });

            if (!response.ok) {
                const errData = await response.json().catch(() => ({}));
                throw new Error(errData.error?.message || `HTTP ${response.status}`);
            }

            const resJson = await response.json();
            let textResponse = '';
            
            if (llmProvider === 'omniroute') {
                textResponse = resJson.choices?.[0]?.message?.content;
            } else {
                textResponse = resJson.candidates?.[0]?.content?.parts?.[0]?.text;
            }
            
            if (!textResponse) throw new Error("API 응답에 텍스트 데이터가 포함되어 있지 않습니다.");

            // JSON 파싱
            let parsedMap;
            try {
                const cleanText = textResponse.replace(/^```json\s*|```\s*$/g, '').trim();
                parsedMap = JSON.parse(cleanText);
            } catch (jsonErr) {
                console.error("JSON 파싱 에러:", textResponse);
                throw new Error("AI가 보고서 작성에 필요한 유효한 JSON 형식의 데이터 생성에 실패했습니다.");
            }

            return parsedMap;

        } catch (error) {
            console.warn(`호출 실패:`, error.message);
            
            if (llmProvider === 'gemini') {
                const keys = String(apiKey).split(',').map(k => k.trim()).filter(k => k.match(/^(AIza|AQ\.)/));
                if (keys.length > 1 && currentKeyIndex < keys.length - 1) {
                    currentKeyIndex++;
                    if (onProgress) onProgress(`보조 API 키로 전환하여 재시도 중...`);
                    return fetchWithRetry(maxModelRetries);
                } else {
                    currentKeyIndex = 0;
                    if (maxModelRetries > 1) {
                        currentModelIndex = (currentModelIndex + 1) % FALLBACK_MODELS.length;
                        if (onProgress) onProgress(`대체 AI 모델로 전환하여 재시도 중...`);
                        return fetchWithRetry(maxModelRetries - 1);
                    } else {
                        throw new Error(`AI 호출 및 폴백에 최종 실패했습니다: ${error.message}`);
                    }
                }
            } else {
                throw error;
            }
        }
    };

    return await fetchWithRetry();
}

/**
 * HWPX 샘플 템플릿과 참고 자료를 활용하여 완성된 HWPX 보고서를 조립합니다.
 */
export async function generateReportFromTemplate(hwpxTemplateFile, dataFiles, apiKey, instruction = '', llmProvider = 'gemini', omniRouteModel = 'auto', onProgress) {
    if (!hwpxTemplateFile) throw new Error('HWPX 템플릿 양식 파일이 없습니다.');
    if (!dataFiles || dataFiles.length === 0) throw new Error('보고서에 반영할 참고 자료 파일이 없습니다.');

    // 1. 참고 자료 파일들로부터 텍스트 일괄 추출
    if (onProgress) onProgress("참고 자료 파일(PPT, HWPX, MD 등) 텍스트 추출 중...");
    let materialsText = "";
    for (let i = 0; i < dataFiles.length; i++) {
        const file = dataFiles[i];
        if (onProgress) onProgress(`자료 파일 분석 중 [${i + 1}/${dataFiles.length}]: ${file.name}`);
        const text = await processFile(file);
        materialsText += `\n\n=== 파일명: ${file.name} ===\n${text}`;
    }

    // 2. HWPX 템플릿 로딩 및 압축 해제
    if (onProgress) onProgress("HWPX 템플릿 구조 로딩 중...");
    const zip = await JSZip.loadAsync(hwpxTemplateFile);
    
    // 3. section0.xml 파싱하여 치환 가능한 본문 문단 추출
    const sectionPath = 'HPB/Content/section0.xml';
    const sectionXmlText = await zip.file(sectionPath).async('text');
    
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(sectionXmlText, 'text/xml');
    
    // 치환을 진행할 후보 문단들 수집 (ID 바인딩)
    const pElements = xmlDoc.getElementsByTagName('hp:p');
    const paragraphsToTranslate = [];
    
    for (let i = 0; i < pElements.length; i++) {
        const pEl = pElements[i];
        
        // hp:t 요소들을 모아서 문단 전체 텍스트 수집
        const tElements = pEl.getElementsByTagName('hp:t');
        let fullText = "";
        for (let j = 0; j < tElements.length; j++) {
            fullText += tElements[j].textContent || "";
        }
        
        const cleanText = fullText.trim();
        // 치환 후보 판별: {{내용}} 마크가 포함되어 있거나, 빈 공간 혹은 특정 치환 지시어 감지
        if (cleanText.includes('{{') && cleanText.includes('}}') || 
            cleanText.includes('[작성') || 
            cleanText.includes('[여기에') ||
            cleanText === '작성란' ||
            cleanText === '내용 작성') {
            
            // 문단 요소에 임시 ID 부착
            const tempId = `para_${i}`;
            pEl.setAttribute('tempId', tempId);
            
            paragraphsToTranslate.push({
                id: tempId,
                text: cleanText
            });
        }
    }

    if (paragraphsToTranslate.length === 0) {
        throw new Error("템플릿 문서에서 치환할 대상 마커({{내용}} 또는 [작성란] 등)를 찾을 수 없습니다.");
    }

    // 4. LLM 호출하여 치환 텍스트 맵 획득
    if (onProgress) onProgress(`AI 치환 맵 분석 및 보고서 작성 진행 중... (총 ${paragraphsToTranslate.length}개 구획)`);
    const parsedMap = await generateHwpxReportWithLLM(
        paragraphsToTranslate,
        materialsText,
        apiKey,
        instruction,
        'auto',
        llmProvider,
        omniRouteModel,
        onProgress
    );

    // 5. XML 치환 작업 진행
    if (onProgress) onProgress("보고서 본문 조립 및 XML 결합 진행 중...");
    
    for (let i = 0; i < pElements.length; i++) {
        const pEl = pElements[i];
        const tempId = pEl.getAttribute('tempId');
        if (tempId && parsedMap[tempId]) {
            const replacementText = parsedMap[tempId];
            
            // 줄바꿈이 있는 경우 문단 스타일 복제 기법 적용
            if (replacementText.includes('\n')) {
                const lines = replacementText.split('\n');
                let lastInsertedNode = pEl;
                
                // 첫 줄은 원본 문단 내에 치환 기입
                replaceParagraphText(pEl, lines[0]);
                
                // 나머지 줄들은 원본의 스타일 정보(paraPr 등)를 고스란히 복제하여 새 문단 생성 후 추가
                for (let k = 1; k < lines.length; k++) {
                    const clonedP = pEl.cloneNode(true);
                    
                    // 복제된 문단 구조의 모든 ID 속성을 완전하게 소거하여 한글 프로그램 폭사/오류 방지
                    clonedP.removeAttribute('id');
                    clonedP.removeAttribute('tempId');
                    const allChildNodes = clonedP.getElementsByTagName('*');
                    for (let n = 0; n < allChildNodes.length; n++) {
                        allChildNodes[n].removeAttribute('id');
                    }
                    
                    replaceParagraphText(clonedP, lines[k]);
                    
                    // XML 상에서 이전 노드 바로 뒤에 삽입
                    lastInsertedNode.parentNode.insertBefore(clonedP, lastInsertedNode.nextSibling);
                    lastInsertedNode = clonedP;
                }
            } else {
                replaceParagraphText(pEl, replacementText);
            }
            
            // 임시 마킹 제거
            pEl.removeAttribute('tempId');
        }
    }

    // 6. 변경된 XML을 다시 HWPX 아카이브에 기입
    const serializer = new XMLSerializer();
    const newXmlText = serializer.serializeToString(xmlDoc);
    zip.file(sectionPath, newXmlText);

    // 7. 한글 프로그램 디지털 서명 에러 방지를 위해 META-INF 폴더 일괄 제거
    if (onProgress) onProgress("디지털 서명(META-INF) 우회 및 한글 깨짐 방지 처리 중...");
    const filesToRemove = [];
    zip.forEach((relativePath) => {
        if (relativePath.startsWith('META-INF/')) {
            filesToRemove.push(relativePath);
        }
    });
    filesToRemove.forEach(path => zip.remove(path));

    // 8. 최종 압축하여 파일 다운로드용 블롭 생성
    if (onProgress) onProgress("최종 HWPX 문서 패키징 중...");
    const finalBlob = await zip.generateAsync({ type: 'blob', mimeType: 'application/hwpx' });
    
    if (onProgress) onProgress("보고서 생성 완결!");
    return finalBlob;
}

/**
 * 하나의 문단 노드 내부의 텍스트 요소를 새로운 값으로 대체합니다.
 */
function replaceParagraphText(pEl, newText) {
    const rElements = pEl.getElementsByTagName('hp:run');
    if (rElements.length > 0) {
        // 첫 번째 hp:run의 첫 번째 hp:t에 모든 텍스트를 기입
        const firstRun = rElements[0];
        let tEl = firstRun.getElementsByTagName('hp:t')[0];
        
        if (!tEl) {
            // hp:t가 없는 경우 새로 생성하여 주입
            tEl = pEl.ownerDocument.createElementNS('http://www.hancom.co.kr/hwpml/2011/paragraph', 'hp:t');
            firstRun.appendChild(tEl);
        }
        
        tEl.textContent = newText;
        
        // 나머지 run 요소를 삭제하여 찌꺼기 텍스트가 노출되지 않도록 완전 정제
        for (let i = rElements.length - 1; i > 0; i--) {
            rElements[i].parentNode.removeChild(rElements[i]);
        }
        
        // 첫 번째 run 내부의 나머지 t 요소도 삭제
        const tElementsInFirstRun = firstRun.getElementsByTagName('hp:t');
        for (let j = tElementsInFirstRun.length - 1; j > 0; j--) {
            tElementsInFirstRun[j].parentNode.removeChild(tElementsInFirstRun[j]);
        }
    } else {
        // 문단 내에 아무런 run이 없는 안전한 구조인 경우 기본 구조 생성
        const doc = pEl.ownerDocument;
        const newRun = doc.createElementNS('http://www.hancom.co.kr/hwpml/2011/paragraph', 'hp:run');
        const newT = doc.createElementNS('http://www.hancom.co.kr/hwpml/2011/paragraph', 'hp:t');
        newT.textContent = newText;
        newRun.appendChild(newT);
        pEl.appendChild(newRun);
    }
}
