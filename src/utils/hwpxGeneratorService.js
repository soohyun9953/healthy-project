import JSZip from 'jszip';
import { processFile } from './fileExtractor.js';
import { FALLBACK_MODELS } from './geminiModels.js';

/**
 * Gemini LLM API를 호출하여 HWPX 템플릿의 문단을 소스 문서들의 내용에 맞게 치환할 텍스트 맵을 생성합니다.
 */
export async function generateHwpxReportWithLLM(paragraphs, materialsText, apiKey, instruction = '', selectedModel = 'auto', onProgress) {
    const keys = String(apiKey).split(',').map(k => k.trim()).filter(k => k.match(/^(AIza|AQ\.)/));
    if (keys.length === 0) {
        throw new Error("유효한 API 키가 제공되지 않았습니다.");
    }

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
        const modelToUse = FALLBACK_MODELS[currentModelIndex] || FALLBACK_MODELS[0];
        const keyToUse = keys[currentKeyIndex];

        if (onProgress) onProgress(`AI 엔진 호출 중... (모델: ${modelToUse.split('/').pop()})`);

        try {
            const apiURL = `https://generativelanguage.googleapis.com/v1beta/${modelToUse}:generateContent?key=${keyToUse}`;
            const response = await fetch(apiURL, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    contents: [{ parts: [{ text: userInput }] }],
                    generationConfig: {
                        responseMimeType: "application/json"
                    }
                })
            });

            if (!response.ok) {
                const errData = await response.json().catch(() => ({}));
                throw new Error(errData.error?.message || `HTTP ${response.status}`);
            }

            const resJson = await response.json();
            const textResponse = resJson.candidates?.[0]?.content?.parts?.[0]?.text;
            if (!textResponse) throw new Error("API 응답 구조에 유효한 텍스트가 없습니다.");
            
            // 사용량 기록
            try {
                const modelShortName = modelToUse.split('/').pop();
                const usage = JSON.parse(localStorage.getItem('gemini_model_usage') || '{}');
                usage[modelShortName] = (usage[modelShortName] || 0) + 1;
                localStorage.setItem('gemini_model_usage', JSON.stringify(usage));
                window.dispatchEvent(new CustomEvent('gemini_usage_updated'));
            } catch (e) {
                console.error("Usage logging failed:", e);
            }

            // JSON 파싱
            let parsedMap;
            try {
                // 혹시 모를 마크다운 백틱 정제
                const cleanText = textResponse.replace(/^```json\s*|```\s*$/g, '').trim();
                parsedMap = JSON.parse(cleanText);
            } catch (jsonErr) {
                console.error("JSON parse error on LLM response:", textResponse);
                throw new Error("AI가 유효한 JSON 형식의 치환 맵을 생성하지 못했습니다. 다시 시도해 주세요.");
            }

            return parsedMap;

        } catch (error) {
            console.warn(`모델 ${modelToUse} 호출 실패:`, error.message);
            
            // 키 또는 모델 순환 교체 시도
            if (keys.length > 1 && currentKeyIndex < keys.length - 1) {
                currentKeyIndex++;
                if (onProgress) onProgress(`보조 API 키로 전환하여 재시도 중...`);
                return fetchWithRetry(maxModelRetries);
            } else {
                currentKeyIndex = 0; // 키 초기화
                if (maxModelRetries > 1) {
                    currentModelIndex = (currentModelIndex + 1) % FALLBACK_MODELS.length;
                    if (onProgress) onProgress(`대체 AI 모델로 전환하여 재시도 중...`);
                    return fetchWithRetry(maxModelRetries - 1);
                } else {
                    throw new Error(`AI 호출 및 폴백에 최종 실패했습니다: ${error.message}`);
                }
            }
        }
    };

    return await fetchWithRetry();
}

/**
 * HWPX 샘플 템플릿과 참고 자료를 활용하여 완성된 HWPX 보고서를 조립합니다.
 */
export async function generateReportFromTemplate(hwpxTemplateFile, dataFiles, apiKey, instruction = '', onProgress) {
    if (!hwpxTemplateFile) throw new Error('HWPX 템플릿 양식 파일이 없습니다.');
    if (!dataFiles || dataFiles.length === 0) throw new Error('보고서에 반영할 참고 자료 파일이 없습니다.');

    // 1. 참고 자료 파일들로부터 텍스트 일괄 추출
    if (onProgress) onProgress("참고 자료 파일(PPT, HWPX, MD 등) 텍스트 추출 중...");
    const extractedTexts = [];
    for (let i = 0; i < dataFiles.length; i++) {
        const file = dataFiles[i];
        if (onProgress) onProgress(`자료 파일 읽는 중 (${i+1}/${dataFiles.length}): ${file.name}`);
        try {
            const fileData = await processFile(file);
            extractedTexts.push(`[자료 파일명: ${file.name}]\n${fileData.text}`);
        } catch (fileErr) {
            console.warn(`${file.name} 파일 텍스트 추출 실패:`, fileErr);
            extractedTexts.push(`[자료 파일명: ${file.name} (텍스트 추출 에러: ${fileErr.message})]`);
        }
    }
    const mergedMaterialsText = extractedTexts.join('\n\n=========================================\n\n');

    // 2. HWPX 템플릿 압축 풀기 및 section0.xml 파싱
    if (onProgress) onProgress("HWPX 양식 템플릿 로딩 및 XML 구조 분석 중...");
    const templateBuffer = await hwpxTemplateFile.arrayBuffer();
    const templateZip = new JSZip();
    await templateZip.loadAsync(templateBuffer);

    // HWPX 본문 XML 파일 찾기 (기본 section0.xml)
    const section0Path = 'Contents/section0.xml';
    if (!templateZip.files[section0Path]) {
        throw new Error('HWPX 템플릿 내부에서 Contents/section0.xml을 찾을 수 없습니다.');
    }
    
    const sectionXmlStr = await templateZip.files[section0Path].async('text');
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(sectionXmlStr, 'application/xml');
    
    if (xmlDoc.getElementsByTagName('parsererror').length > 0) {
        throw new Error('HWPX section0.xml 구조 해석 오류가 발생했습니다.');
    }

    // 3. XML에서 모든 문단 (<hp:p>) 구조 추출 및 고유 ID 부여
    const pNodes = xmlDoc.getElementsByTagName('hp:p');
    if (pNodes.length === 0) {
        throw new Error('HWPX 템플릿 내부에 문단(<hp:p>)이 존재하지 않는 비정상적인 구조입니다.');
    }

    const templateParagraphs = [];
    const nodeMap = new Map(); // ID -> XML Node 매핑

    for (let i = 0; i < pNodes.length; i++) {
        const pNode = pNodes[i];
        
        // 문단 내 모든 텍스트(<hp:t>) 수집
        const tNodes = pNode.getElementsByTagName('hp:t');
        let textContent = '';
        for (let j = 0; j < tNodes.length; j++) {
            textContent += tNodes[j].textContent || '';
        }
        
        const trimmed = textContent.trim();
        // 텍스트가 존재하는 문단 위주로 AI 분석에 전송 (메타 정보 유실 차단)
        if (trimmed.length > 0) {
            const pId = `para_${i}`;
            templateParagraphs.push({
                id: pId,
                text: trimmed
            });
            nodeMap.set(pId, pNode);
        }
    }

    if (templateParagraphs.length === 0) {
        throw new Error('HWPX 템플릿에 치환할 본문 텍스트가 존재하지 않습니다.');
    }

    // 4. Gemini API를 통해 치환 맵 받아오기
    if (onProgress) onProgress("AI 보고서 내용 생성 요청 중...");
    const replacementMap = await generateHwpxReportWithLLM(
        templateParagraphs,
        mergedMaterialsText,
        apiKey,
        instruction,
        'auto',
        onProgress
    );

    // 5. 생성된 치환 맵을 기반으로 HWPX XML 수정
    if (onProgress) onProgress("생성된 내용을 한글 보고서 구조에 이식 중...");
    let replaceCount = 0;

    for (const [pId, replacementText] of Object.entries(replacementMap)) {
        const originalNode = nodeMap.get(pId);
        if (!originalNode || !replacementText) continue;

        // 개행(\n)이 포함되어 있다면 복수 문단으로 분할 복제 확장
        const textLines = replacementText.split('\n').map(l => l.trim()).filter(l => l.length > 0);
        if (textLines.length === 0) continue;

        let lastInsertedNode = originalNode;

        for (let lIdx = 0; lIdx < textLines.length; lIdx++) {
            const lineText = textLines[lIdx];
            let targetNode;

            if (lIdx === 0) {
                // 첫 단락은 기존 노드를 그대로 재활용하여 덮어쓰기
                targetNode = originalNode;
            } else {
                // 두 번째 단락부터는 기존 문단 노드를 복제(Clone)하여 스타일에 어긋나지 않게 추가
                targetNode = originalNode.cloneNode(true);
                // 복제된 노드를 이전 노드 바로 다음에 삽입
                originalNode.parentNode.insertBefore(targetNode, lastInsertedNode.nextSibling);
                lastInsertedNode = targetNode;
            }

            // 문단 내부의 모든 <hp:t>를 정리하고 새 본문 텍스트 적용
            const tNodes = targetNode.getElementsByTagName('hp:t');
            if (tNodes.length > 0) {
                // 첫 번째 <hp:t> 에 텍스트를 기입하고, 스타일 깨짐 방지를 위해 나머지 <hp:t> 들은 빈 값 처리
                tNodes[0].textContent = lineText;
                for (let tIdx = 1; tIdx < tNodes.length; tIdx++) {
                    tNodes[tIdx].textContent = '';
                }
            } else {
                // 만약 <hp:t> 노드가 없다면 강제로 생성 (안전장치)
                const runNode = targetNode.getElementsByTagName('hp:run')[0];
                if (runNode) {
                    const tElem = xmlDoc.createElement('hp:t');
                    tElem.textContent = lineText;
                    runNode.appendChild(tElem);
                }
            }
            replaceCount++;
        }
    }

    if (onProgress) onProgress(`총 ${replaceCount}개 문단 구조 생성 및 이식 완료. 파일 빌드 중...`);

    // 6. section0.xml을 직렬화하여 zip에 다시 담기
    const serializer = new XMLSerializer();
    const updatedXmlStr = serializer.serializeToString(xmlDoc);
    templateZip.file(section0Path, updatedXmlStr);

    // 디지털 서명 무결성 검증 우회: 본문 내용 치환으로 인한 '문서 손상/변조' 경고를 방지하기 위해 
    // META-INF 디렉토리 및 하위 서명 파일들을 완전히 제거합니다.
    Object.keys(templateZip.files).forEach(filePath => {
        if (filePath.startsWith('META-INF/') || filePath === 'META-INF') {
            templateZip.remove(filePath);
        }
    });

    // HWPX 파일 압축 빌드
    const finalBuffer = await templateZip.generateAsync({ type: 'blob' });
    if (onProgress) onProgress("보고서 HWPX 파일 다운로드 준비 완료!");
    
    // 다운로드 결과를 트래킹하기 위해 치환 건수 메타데이터 주입
    finalBuffer.fusedCount = replaceCount;
    return finalBuffer;
}
