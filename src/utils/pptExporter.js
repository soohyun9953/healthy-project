import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { saveAs } from 'file-saver';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';

/**
 * 사용자가 저장할 위치와 파일명을 선택할 수 있도록 다이얼로그를 띄워 저장합니다.
 * File System Access API를 지원하지 않는 브라우저에서는 기본 다운로드 방식으로 동작합니다.
 */
export async function saveFileWithLocationPicker(blob, defaultFileName) {
    if ('showSaveFilePicker' in window) {
        try {
            const handle = await window.showSaveFilePicker({
                suggestedName: defaultFileName,
                types: [{
                    description: 'PowerPoint Presentation',
                    accept: { 'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx'] },
                }],
            });
            const writable = await handle.createWritable();
            await writable.write(blob);
            await writable.close();
            return true;
        } catch (err) {
            if (err.name === 'AbortError') {
                return false;
            }
            console.error('File System Access API 에러, 기본 다운로드 방식으로 전환합니다.', err);
        }
    }
    // Fallback
    saveAs(blob, defaultFileName);
    return true;
}

/**
 * 엑셀 파일을 읽어서 JSON 배열로 변환합니다.
 */
export async function parseExcelData(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet);
                resolve(jsonData);
            } catch (error) {
                reject(error);
            }
        };
        reader.onerror = (error) => reject(error);
        reader.readAsArrayBuffer(file);
    });
}

/**
 * 템플릿의 slide1.xml을 분석하여 최대 인덱스(chunkSize)를 감지합니다.
 */
function detectTemplateChunkSize(zip) {
    const slide1Xml = zip.file('ppt/slides/slide1.xml')?.asText();
    if (!slide1Xml) return 1;
    
    let maxIdx = 1;
    const tagMatchRegex = /{([^{}]+?)_(\d+)}/g;
    let match;
    while ((match = tagMatchRegex.exec(slide1Xml)) !== null) {
        const idx = parseInt(match[2]);
        if (!isNaN(idx) && idx > maxIdx) maxIdx = idx;
    }
    return maxIdx;
}

/**
 * [V15] 정밀 문자열 패칭 방식 (Precision String Patching)
 * 브라우저 DOMParser의 네임스페이스 오염을 피하기 위해 원본 문자열을 직접 편집합니다.
 */
function duplicateSlides(zip, count, chunkSize) {
    if (count <= 1) {
        // 1번만 있더라도 1번 슬라이드 태그는 고쳐줘야 함
        let sld1 = zip.file('ppt/slides/slide1.xml').asText();
        sld1 = sld1.replace(/{([^{}]+?)(?:_(\d+))?}/g, (match, key, rowInSlide) => {
            const rowIdxInRange = rowInSlide ? parseInt(rowInSlide) : 1;
            return `{${key}_${rowIdxInRange}}`;
        });
        zip.file('ppt/slides/slide1.xml', sld1);
        return;
    }

    // 1. 원본 소스 획득
    const presXml = zip.file('ppt/presentation.xml').asText();
    const ctXml = zip.file('[Content_Types].xml').asText();
    const presRelsXml = zip.file('ppt/_rels/presentation.xml.rels').asText();
    const sld1Xml = zip.file('ppt/slides/slide1.xml').asText();
    const sld1RelsXml = zip.file('ppt/slides/_rels/slide1.xml.rels')?.asText();

    // 2. ID 분석 (문자열 기반)
    let maxRidNum = 0;
    const ridMatches = presRelsXml.matchAll(/Id="rId(\d+)"/g);
    for (const m of ridMatches) {
        const n = parseInt(m[1]);
        if (n > maxRidNum) maxRidNum = n;
    }

    let maxSldIdNum = 255;
    const sldIdMatches = presXml.matchAll(/id="(\d+)"/g);
    for (const m of sldIdMatches) {
        const n = parseInt(m[1]);
        if (n > maxSldIdNum && n < 1000000) maxSldIdNum = n;
    }

    // 3. 1번 슬라이드 태그 정규화
    const sld1Fixed = sld1Xml.replace(/{([^{}]+?)(?:_(\d+))?}/g, (match, key, rowInSlide) => {
        const rowIdxInRange = rowInSlide ? parseInt(rowInSlide) : 1;
        return `{${key}_${rowIdxInRange}}`;
    });
    zip.file('ppt/slides/slide1.xml', sld1Fixed);

    // 4. 새 슬라이드 데이터 조립
    let newSldIdEntries = "";
    let newContentTypeEntries = "";
    let newRelEntries = "";

    for (let i = 2; i <= count; i++) {
        const rId = `rId${maxRidNum + (i - 1)}`;
        const sldId = maxSldIdNum + (i - 1);
        const slideFileName = `slide${i}.xml`;
        const slidePath = `ppt/slides/${slideFileName}`;

        // Metadata entries
        newSldIdEntries += `<p:sldId id="${sldId}" r:id="${rId}"/>`;
        newContentTypeEntries += `<Override PartName="/${slidePath}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
        newRelEntries += `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/${slideFileName}"/>`;

        // Slide Content (변환)
        let slideNStr = sld1Xml.replace(/{([^{}]+?)(?:_(\d+))?}/g, (match, key, rowInSlide) => {
            const rowIdxInRange = rowInSlide ? parseInt(rowInSlide) : 1;
            const globalRowIdx = (i - 1) * chunkSize + rowIdxInRange;
            return `{${key}_${globalRowIdx}}`;
        });
        // Shape ID 충돌 방지
        slideNStr = slideNStr.replace(/ id="(\d+)"/g, (match, idStr) => ` id="${parseInt(idStr) + ((i-1) * 1000)}"`);
        
        zip.file(slidePath, slideNStr);
        if (sld1RelsXml) zip.file(`ppt/slides/_rels/${slideFileName}.rels`, sld1RelsXml);
    }

    // 5. 문자열 패치 적용 - 원본 구조의 네임스페이스를 해치지 않음
    
    // [Presentation.xml] - 기존 슬라이드 삭제(1번 제외) 후 새 슬라이드 삽입
    // 1번 슬라이드 정보(첫번째 p:sldId)만 남깁니다.
    const sldIdLstMatch = presXml.match(/<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/);
    if (sldIdLstMatch) {
        const firstSldId = sldIdLstMatch[1].match(/<p:sldId[^>]+>/)?.[0] || "";
        const updatedSldIdLst = `<p:sldIdLst>${firstSldId}${newSldIdEntries}</p:sldIdLst>`;
        const newPresXml = presXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, updatedSldIdLst);
        zip.file('ppt/presentation.xml', newPresXml);
    }

    // [[Content_Types].xml] - Override 추가
    const newCtXml = ctXml.replace('</Types>', `${newContentTypeEntries}</Types>`);
    zip.file('[Content_Types].xml', newCtXml);

    // [presentation.xml.rels] - Relationship 추가
    const newPresRelsXml = presRelsXml.replace('</Relationships>', `${newRelEntries}</Relationships>`);
    zip.file('ppt/_rels/presentation.xml.rels', newPresRelsXml);

    // [app.xml] - 슬라이드 개수 업데이트
    const appXml = zip.file('docProps/app.xml')?.asText();
    if (appXml) {
        let newAppXml = appXml.replace(/<Slides>\d+<\/Slides>/, `<Slides>${count}</Slides>`);
        // vt:vector size 업데이트 및 추가 lpstr 삽입
        const vtMatch = newAppXml.match(/<vt:vector[^>]+size="(\d+)"[^>]+baseType="lpstr">([\s\S]*?)<\/vt:vector>/);
        if (vtMatch) {
            const firstLpstr = vtMatch[2].match(/<vt:lpstr>[\s\S]*?<\/vt:lpstr>/)?.[0] || "<vt:lpstr>Slide 1</vt:lpstr>";
            let newLpstrs = firstLpstr;
            for (let i = 2; i <= count; i++) newLpstrs += `<vt:lpstr>Slide ${i}</vt:lpstr>`;
            
            const updatedVector = vtMatch[0]
                .replace(/size="\d+"/, `size="${count}"`)
                .replace(vtMatch[2], newLpstrs);
            newAppXml = newAppXml.replace(vtMatch[0], updatedVector);
        }
        zip.file('docProps/app.xml', newAppXml);
    }
}

/**
 * PPT 템플릿과 데이터를 머지하여 PPT를 생성합니다.
 */
export async function generatePptFromTemplate(pptTemplateFile, dataRows, generationMode = 'single', chunkSizeArg = 10) {
    const templateArrayBuffer = await pptTemplateFile.arrayBuffer();

    if (generationMode === 'single') {
        try {
            const zip = new PizZip(templateArrayBuffer);
            const templateChunkSize = detectTemplateChunkSize(zip);
            const slideCount = Math.ceil(dataRows.length / templateChunkSize);
            
            console.log(`PPT v15 Patching: Rows=${dataRows.length}, ChunkSize=${templateChunkSize}, Slides=${slideCount}`);
            
            duplicateSlides(zip, slideCount, templateChunkSize);

            const doc = new Docxtemplater(zip, {
                paragraphLoop: true,
                linebreaks: true,
                nullGetter() { return ''; }
            });

            const flatData = {};
            dataRows.forEach((rowObj, idx) => {
                const rowNum = idx + 1;
                for (const key in rowObj) {
                    let val = rowObj[key];
                    if (typeof val === 'string') {
                        // 엑셀에서 넘어온 \r\n을 \n으로 정규화하여 PPT 내 이중 공백 방지
                        val = val.replace(/\r\n/g, '\n');
                    }
                    flatData[`${key}_${rowNum}`] = val;
                }
            });

            doc.render(flatData);

            const blob = doc.getZip().generate({
                type: 'blob',
                mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
            });

            await saveFileWithLocationPicker(blob, '통합_데이터_리포트.pptx');
        } catch (error) {
            console.error('단일 PPT 생성 오류:', error);
            throw error;
        }
    } else {
        // 분할 모드 등 (필요시 duplicateSlidesV15 적용)
    }
}

async function createAndDownloadZip(files, zipFileName) {
    const zip = new JSZip();
    for (const file of files) zip.file(file.name, file.blob);
    const content = await zip.generateAsync({ type: 'blob' });
    saveAs(content, zipFileName);
}

/**
 * PPT 내 지정한 텍스트의 디자인을 일괄 변경합니다 (윤곽선: 흰색 실선, 투명도 100%).
 */
export async function applyTextDesignToPpt(pptFile, targetText) {
    if (!pptFile) {
        throw new Error('PPT 파일이 필요합니다.');
    }
    const trimmedTargetText = targetText ? targetText.trim() : '';

    const arrayBuffer = await pptFile.arrayBuffer();
    const zip = new PizZip(arrayBuffer);

    const allFiles = Object.keys(zip.files);
    const targetFilesSet = new Set();

    if (trimmedTargetText !== '') {
        const rawTarget = trimmedTargetText.replace(/\s+/g, '');
        const checkFiles = allFiles.filter(p => p.endsWith('.xml') && 
            (p.startsWith('ppt/slides/slide') || p.startsWith('ppt/slideLayouts/') || p.startsWith('ppt/slideMasters/'))
        );
        
        checkFiles.forEach(slidePath => {
            let slideText = '';
            const filesInSlide = [slidePath];

            const slideFileName = slidePath.split('/').pop();
            const relsPath = slidePath.replace(slideFileName, '_rels/' + slideFileName + '.rels');
            if (zip.files[relsPath]) {
                try {
                    const parser = new DOMParser();
                    const relsStr = zip.file(relsPath).asText();
                    const relsDoc = parser.parseFromString(relsStr, 'application/xml');
                    const rels = relsDoc.getElementsByTagName('Relationship');
                    for (let i = 0; i < rels.length; i++) {
                        const target = rels[i].getAttribute('Target');
                        if (target && target.endsWith('.xml')) {
                            const targetFileName = target.split('/').pop();
                            const actualPath = Object.keys(zip.files).find(p => p.endsWith(targetFileName) && p.startsWith('ppt/'));
                            if (actualPath && !filesInSlide.includes(actualPath)) {
                                filesInSlide.push(actualPath);
                            }
                        }
                    }
                } catch (e) {
                    console.error('Error parsing rels for slide:', slidePath, e);
                }
            }

            const parser = new DOMParser();
            filesInSlide.forEach(fp => {
                try {
                    const doc = parser.parseFromString(zip.file(fp).asText(), 'application/xml');
                    const allNodes = doc.getElementsByTagName('*');
                    for (let i = 0; i < allNodes.length; i++) {
                        const localName = allNodes[i].localName || allNodes[i].tagName.split(':').pop();
                        if (localName === 't') {
                            slideText += allNodes[i].textContent;
                        }
                    }
                } catch (e) {
                    // 무시
                }
            });

            const rawSlideText = slideText.replace(/\s+/g, '');
            if (rawSlideText.includes(rawTarget)) {
                filesInSlide.forEach(fp => targetFilesSet.add(fp));
            }
        });
    } else {
        allFiles.forEach(path => {
            if (path.endsWith('.xml') && path.startsWith('ppt/')) {
                if (!path.includes('presentation.xml') && 
                    !path.includes('presProps.xml') && 
                    !path.includes('viewProps.xml') && 
                    !path.includes('tableStyles.xml')) {
                    targetFilesSet.add(path);
                }
            }
        });
    }

    if (targetFilesSet.size === 0) {
        throw new Error('PPT 파일에서 대상 XML을 찾을 수 없거나 대상 텍스트가 포함된 슬라이드가 없습니다.');
    }

    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    const nsA = 'http://schemas.openxmlformats.org/drawingml/2006/main';

    targetFilesSet.forEach(slidePath => {
        let slideXmlStr = zip.file(slidePath).asText();
        const xmlDoc = parser.parseFromString(slideXmlStr, 'application/xml');

        const parserError = xmlDoc.getElementsByTagName('parsererror');
        if (parserError.length > 0) {
            console.error('XML Parsing Error in file:', slidePath);
            return;
        }

        function applyLnToRPr(rPr) {
            let existingLn = null;
            for (let j = 0; j < rPr.childNodes.length; j++) {
                const child = rPr.childNodes[j];
                if (child.nodeType === 1) {
                    const localName = child.localName || child.tagName.split(':').pop();
                    if (localName === 'ln') {
                        existingLn = child;
                        break;
                    }
                }
            }
            if (existingLn) {
                rPr.removeChild(existingLn);
            }

            const ln = xmlDoc.createElementNS(nsA, 'a:ln');
            ln.setAttribute('w', '9525');
            ln.setAttribute('cmpd', 'sng');

            const solidFill = xmlDoc.createElementNS(nsA, 'a:solidFill');
            const srgbClr = xmlDoc.createElementNS(nsA, 'a:srgbClr');
            srgbClr.setAttribute('val', 'FFFFFF');

            // 투명도 100% 복구
            const alpha = xmlDoc.createElementNS(nsA, 'a:alpha');
            alpha.setAttribute('val', '0');

            srgbClr.appendChild(alpha);
            solidFill.appendChild(srgbClr);
            ln.appendChild(solidFill);
            
            const prstDash = xmlDoc.createElementNS(nsA, 'a:prstDash');
            prstDash.setAttribute('val', 'solid');
            ln.appendChild(prstDash);
            
            // 파워포인트 스키마(CT_TextCharacterProperties)에서 
            // <a:ln>은 반드시 가장 첫 번째 자식 요소로 위치해야 합니다.
            // (도형의 경우 fill 다음에 ln이 오지만, 텍스트는 ln이 fill보다 먼저 와야 합니다)
            rPr.insertBefore(ln, rPr.firstChild);
        }

        const allElements = xmlDoc.getElementsByTagName('*');
        for (let i = 0; i < allElements.length; i++) {
            const el = allElements[i];
            if (el.nodeType !== 1) continue;
            const localName = el.localName || el.tagName.split(':').pop();
            
            if (localName === 'r' || localName === 'fld' || localName === 'br') {
                let rPr = null;
                for (let j = 0; j < el.childNodes.length; j++) {
                    const child = el.childNodes[j];
                    if (child.nodeType === 1) {
                        const childLocalName = child.localName || child.tagName.split(':').pop();
                        if (childLocalName === 'rPr') {
                            rPr = child;
                            break;
                        }
                    }
                }
                if (!rPr) {
                    rPr = xmlDoc.createElementNS(nsA, 'a:rPr');
                    el.insertBefore(rPr, el.firstChild);
                }
                applyLnToRPr(rPr);
            } else if (localName === 'endParaRPr' || localName === 'defRPr') {
                applyLnToRPr(el);
            }
        }

        const updatedXmlStr = serializer.serializeToString(xmlDoc);
        zip.file(slidePath, updatedXmlStr);
    });

    const blob = zip.generate({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    });

    await saveFileWithLocationPicker(blob, `수정_${pptFile.name}`);
}

/**
 * PPT 내 텍스트를 찾아 일괄 수정합니다. (형식: "기존단어(새단어), 기존단어2(새단어2)")
 */
export async function replaceWordsInPpt(pptFile, replaceRulesStr) {
    if (!pptFile) throw new Error('PPT 파일이 필요합니다.');
    if (!replaceRulesStr || !replaceRulesStr.trim()) throw new Error('수정할 단어 규칙을 입력해주세요.');

    const rules = [];
    const parts = replaceRulesStr.split(',');
    for (const part of parts) {
        const trimmed = part.trim();
        if (!trimmed) continue;
        const match = trimmed.match(/^(.+?)\((.+?)\)$/);
        if (match) {
            rules.push({ oldWord: match[1].trim(), newWord: match[2].trim() });
        } else {
            throw new Error(`규칙 형식이 올바르지 않습니다: "${trimmed}" (올바른 형식 예: 애플리케이션(어플리케이션))`);
        }
    }

    if (rules.length === 0) throw new Error('유효한 치환 규칙이 없습니다.');

    const arrayBuffer = await pptFile.arrayBuffer();
    const zip = new PizZip(arrayBuffer);
    const allFiles = Object.keys(zip.files);

    const targetFiles = allFiles.filter(p => p.endsWith('.xml') && p.startsWith('ppt/'));

    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    let hasChanges = false;

    targetFiles.forEach(slidePath => {
        let slideXmlStr = zip.file(slidePath).asText();
        
        // 최적화: 치환할 단어가 XML 원본 문자열에 하나라도 있는지 빠른 검사
        // (파워포인트가 단어를 쪼개서 저장한 경우는 이 단순 치환 방식으로는 잡기 어려우나, 대부분의 일반 텍스트에 적용 가능)
        let containsAny = false;
        for (const rule of rules) {
            if (slideXmlStr.includes(rule.oldWord)) {
                containsAny = true;
                break;
            }
        }
        
        if (!containsAny) return;

        const xmlDoc = parser.parseFromString(slideXmlStr, 'application/xml');
        const parserError = xmlDoc.getElementsByTagName('parsererror');
        if (parserError.length > 0) return;

        let fileChanged = false;
        const allElements = xmlDoc.getElementsByTagName('*');
        
        for (let i = 0; i < allElements.length; i++) {
            const el = allElements[i];
            if (el.nodeType !== 1) continue;
            
            const localName = el.localName || el.tagName.split(':').pop();
            if (localName === 't') {
                let text = el.textContent;
                let originalText = text;
                
                for (const rule of rules) {
                    text = text.split(rule.oldWord).join(rule.newWord);
                }
                
                if (text !== originalText) {
                    el.textContent = text;
                    fileChanged = true;
                    hasChanges = true;
                }
            }
        }

        if (fileChanged) {
            zip.file(slidePath, serializer.serializeToString(xmlDoc));
        }
    });

    if (!hasChanges) {
        throw new Error('PPT 파일 내에서 해당 단어를 찾을 수 없거나 이미 모두 수정되었습니다.');
    }

    const blob = zip.generate({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    });

    await saveFileWithLocationPicker(blob, `단어수정_${pptFile.name}`);
}

/**
 * pptxgenjs로 생성한 슬라이드(aiGenBlob)를 사용자가 업로드한 원본 마스터(masterFile)에 덮어씌웁니다.
 * 원본의 배경, 로고, 테마 색상(slideMaster, slideLayout)은 유지하면서,
 * 내용물은 AI가 새로 그린 슬라이드들로 완전히 교체하는 하이브리드 병합 엔진입니다.
 */
export async function injectSlidesIntoMaster(masterFile, aiGenBlob) {
    const masterBuffer = await masterFile.arrayBuffer();
    const aiGenBuffer = await aiGenBlob.arrayBuffer();

    const zipMaster = new PizZip(masterBuffer);
    const zipAi = new PizZip(aiGenBuffer);

    // 1. 마스터에서 사용할 기준 레이아웃 타겟 탐색 (기본적으로 첫 번째 슬라이드의 레이아웃 사용)
    let masterLayoutTarget = '../slideLayouts/slideLayout1.xml';
    const firstSlideRelsStr = zipMaster.file('ppt/slides/_rels/slide1.xml.rels')?.asText();
    if (firstSlideRelsStr) {
        const layoutMatch = firstSlideRelsStr.match(/Target="([^"]*slideLayout[^"]*)"/);
        if (layoutMatch) {
            masterLayoutTarget = layoutMatch[1];
        }
    }

    // 2. 마스터의 기존 슬라이드 모두 제거
    let ctXml = zipMaster.file('[Content_Types].xml').asText();
    ctXml = ctXml.replace(/<Override PartName="\/ppt\/slides\/slide\d+\.xml"[^>]*>\s*/g, '');

    let presRelsXml = zipMaster.file('ppt/_rels/presentation.xml.rels').asText();
    presRelsXml = presRelsXml.replace(/<Relationship Id="[^"]+" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/slide" Target="[^"]+"\s*\/>\s*/g, '');

    let presXml = zipMaster.file('ppt/presentation.xml').asText();
    presXml = presXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, '<p:sldIdLst></p:sldIdLst>');

    // 기존 슬라이드 파일 삭제
    for (const key of Object.keys(zipMaster.files)) {
        if (key.startsWith('ppt/slides/slide') || key.startsWith('ppt/slides/_rels/slide')) {
            zipMaster.remove(key);
        }
    }

    // 3. AI가 생성한 슬라이드들을 마스터에 주입
    const aiSlides = Object.keys(zipAi.files).filter(k => k.match(/^ppt\/slides\/slide\d+\.xml$/));
    
    let rIdCounter = 1000;
    let sldIdCounter = 2000;
    
    let newOverrides = '';
    let newPresRels = '';
    let newSldIds = '';

    for (let i = 1; i <= aiSlides.length; i++) {
        const slidePath = `ppt/slides/slide${i}.xml`;
        const relsPath = `ppt/slides/_rels/slide${i}.xml.rels`;
        
        const slideStr = zipAi.file(slidePath)?.asText();
        let relsStr = zipAi.file(relsPath)?.asText();
        
        if (!slideStr || !relsStr) continue;
        // 레이아웃 참조를 마스터의 레이아웃으로 변경
        relsStr = relsStr.replace(/Target="([^"]*slideLayout[^"]*)"/, `Target="${masterLayoutTarget}"`);
        
        zipMaster.file(slidePath, slideStr);
        zipMaster.file(relsPath, relsStr);

        const rId = `rId${rIdCounter++}`;
        const sldId = sldIdCounter++;

        newOverrides += `<Override PartName="/${slidePath}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
        newPresRels += `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i}.xml"/>`;
        newSldIds += `<p:sldId id="${sldId}" r:id="${rId}"/>`;
    }

    // 변경된 메타데이터 갱신
    ctXml = ctXml.replace('</Types>', `${newOverrides}</Types>`);
    zipMaster.file('[Content_Types].xml', ctXml);

    presRelsXml = presRelsXml.replace('</Relationships>', `${newPresRels}</Relationships>`);
    zipMaster.file('ppt/_rels/presentation.xml.rels', presRelsXml);

    presXml = presXml.replace('<p:sldIdLst></p:sldIdLst>', `<p:sldIdLst>${newSldIds}</p:sldIdLst>`);
    
    // 슬라이드 개수 업데이트
    const appXmlPath = 'docProps/app.xml';
    if (zipMaster.files[appXmlPath]) {
        let appXml = zipMaster.file(appXmlPath).asText();
        appXml = appXml.replace(/<Slides>\d+<\/Slides>/, `<Slides>${aiSlides.length}</Slides>`);
        zipMaster.file(appXmlPath, appXml);
    }
    
    zipMaster.file('ppt/presentation.xml', presXml);

    const mergedBlob = zipMaster.generate({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    });

    return mergedBlob;
}

/**
 * [신규] PPT 파일에 단어 일괄 수정과 텍스트 디자인 일괄 변경을 동시에 적용하여 Blob을 반환합니다.
 * @param {File} pptFile 처리할 PPT 파일
 * @param {Object} options { replaceRules: Array, fontRules: Array, applyDesign: boolean, targetText: string }
 * @returns {Promise<Blob>} 변환된 PPT 파일 Blob
 */
export async function processPptBatch(pptFile, options) {
    if (!pptFile) throw new Error('PPT 파일이 필요합니다.');
    
    const { 
        replaceRules = [], 
        fontRules = [], 
        fontSizeRules = [],
        applyDesign = false, 
        applyTableDesign = false, 
        targetText = '' 
    } = options;
    
    if (replaceRules.length === 0 && fontRules.length === 0 && !applyDesign && fontSizeRules.length === 0 && !applyTableDesign) {
        throw new Error('적용할 변경 사항이 없습니다.');
    }

    const arrayBuffer = await pptFile.arrayBuffer();
    const zip = new PizZip(arrayBuffer);
    const allFiles = Object.keys(zip.files);
    
    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    const nsA = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    
    let hasChanges = false;
    let totalTablesCount = 0; // PPT 내에 감지된 총 테이블(표) 요소 개수 누적기
    
    // 타겟 슬라이드 XML 파일 목록
    const targetFiles = allFiles.filter(p => p.endsWith('.xml') && 
        (p.startsWith('ppt/slides/slide') || p.startsWith('ppt/slideLayouts/') || p.startsWith('ppt/slideMasters/') || p.startsWith('ppt/theme/'))
    );

    // 텍스트 디자인 대상 탐색 로직 (적용 시에만)
    const designTargetFilesSet = new Set();
    const trimmedTargetText = targetText ? targetText.trim() : '';
    const rawTarget = trimmedTargetText.replace(/\s+/g, '');

    if (applyDesign) {
        if (trimmedTargetText !== '') {
            targetFiles.forEach(slidePath => {
                if (slidePath.startsWith('ppt/theme/')) return; // 테마는 디자인 변경 대상에서 제외
                
                let slideText = '';
                const filesInSlide = [slidePath];

                const slideFileName = slidePath.split('/').pop();
                const relsPath = slidePath.replace(slideFileName, '_rels/' + slideFileName + '.rels');
                if (zip.files[relsPath]) {
                    try {
                        const relsStr = zip.file(relsPath).asText();
                        const relsDoc = parser.parseFromString(relsStr, 'application/xml');
                        const rels = relsDoc.getElementsByTagName('Relationship');
                        for (let i = 0; i < rels.length; i++) {
                            const target = rels[i].getAttribute('Target');
                            if (target && target.endsWith('.xml')) {
                                const targetFileName = target.split('/').pop();
                                const actualPath = allFiles.find(p => p.endsWith(targetFileName) && p.startsWith('ppt/'));
                                if (actualPath && !filesInSlide.includes(actualPath)) {
                                    filesInSlide.push(actualPath);
                                }
                            }
                        }
                    } catch (e) {
                        console.error('Error parsing rels:', e);
                    }
                }

                filesInSlide.forEach(fp => {
                    try {
                        const doc = parser.parseFromString(zip.file(fp).asText(), 'application/xml');
                        const allNodes = doc.getElementsByTagName('*');
                        for (let i = 0; i < allNodes.length; i++) {
                            const localName = allNodes[i].localName || allNodes[i].tagName.split(':').pop();
                            if (localName === 't') slideText += allNodes[i].textContent;
                        }
                    } catch (e) {}
                });

                if (slideText.replace(/\s+/g, '').includes(rawTarget)) {
                    filesInSlide.forEach(fp => designTargetFilesSet.add(fp));
                }
            });
        } else {
            targetFiles.forEach(path => {
                if (!path.startsWith('ppt/theme/')) designTargetFilesSet.add(path);
            });
        }
    }

    targetFiles.forEach(slidePath => {
        let slideXmlStr = zip.file(slidePath).asText();
        let fileChanged = false;
        let designChanged = false;
        
        const xmlDoc = parser.parseFromString(slideXmlStr, 'application/xml');
        if (xmlDoc.getElementsByTagName('parsererror').length > 0) return;
        
        // [신규] 테이블(표) 표준 디자인 일괄 변경 로직
        if (applyTableDesign && !slidePath.startsWith('ppt/theme/')) {
            const nsA = 'http://schemas.openxmlformats.org/drawingml/2006/main';
            
            // 💡 브라우저 네임스페이스 감지 실패 현상을 원천 방지하기 위해 전체 노드 순회 수집 적용
            const tblElements = [];
            const allNodes = xmlDoc.getElementsByTagName('*');
            for (let i = 0; i < allNodes.length; i++) {
                const node = allNodes[i];
                if (node.nodeType === 1) {
                    const localName = node.localName || node.tagName.split(':').pop();
                    if (localName === 'tbl') {
                        tblElements.push(node);
                    }
                }
            }

            totalTablesCount += tblElements.length;
            let tableChanged = false;
            for (let tIdx = 0; tIdx < tblElements.length; tIdx++) {
                const tbl = tblElements[tIdx];
                
                // 테이블의 모든 행(a:tr) 수집
                const trs = [];
                for (let k = 0; k < tbl.childNodes.length; k++) {
                    const child = tbl.childNodes[k];
                    if (child.nodeType === 1) {
                        const localName = child.localName || child.tagName.split(':').pop();
                        if (localName === 'tr') {
                            trs.push(child);
                        }
                    }
                }
                
                // 각 행 및 셀 순회하며 포맷팅
                for (let rIdx = 0; rIdx < trs.length; rIdx++) {
                    const tr = trs[rIdx];
                    
                    // tr의 모든 셀(a:tc) 수집
                    const tcs = [];
                    for (let k = 0; k < tr.childNodes.length; k++) {
                        const child = tr.childNodes[k];
                        if (child.nodeType === 1) {
                            const localName = child.localName || child.tagName.split(':').pop();
                            if (localName === 'tc') {
                                tcs.push(child);
                            }
                        }
                    }
                    
                    for (let cIdx = 0; cIdx < tcs.length; cIdx++) {
                        const tc = tcs[cIdx];
                        
                        // 셀 속성 tcPr 찾기 및 생성
                        let tcPr = null;
                        for (let k = 0; k < tc.childNodes.length; k++) {
                            const child = tc.childNodes[k];
                            if (child.nodeType === 1) {
                                const localName = child.localName || child.tagName.split(':').pop();
                                if (localName === 'tcPr') {
                                    tcPr = child;
                                    break;
                                }
                            }
                        }
                        
                        if (!tcPr) {
                            tcPr = xmlDoc.createElementNS(nsA, 'a:tcPr');
                            tc.insertBefore(tcPr, tc.firstChild);
                        }
                        
                        // [규칙 1]: 모든 셀에 대해 테두리 선 굵기 0.5pt (w="6350"), 기본 색상 127,127,127 (#7F7F7F) 적용
                        // 단, 첫 번째 행(rIdx === 0)의 내부 실선(경계선)만 흰색(#FFFFFF)으로 적용하고 외곽선은 회색(#7F7F7F) 유지
                        const borderNames = ['lnL', 'lnR', 'lnT', 'lnB'];
                        
                        // 1. 기존에 존재할 수 있는 모든 테두리 노드들을 완벽하게 선제적으로 제거하여 중복 및 굵기 간섭 원천 방지
                        borderNames.forEach(bName => {
                            const existingLns = [];
                            for (let k = 0; k < tcPr.childNodes.length; k++) {
                                const child = tcPr.childNodes[k];
                                if (child.nodeType === 1 && (child.localName === bName || child.tagName.split(':').pop() === bName)) {
                                    existingLns.push(child);
                                }
                            }
                            existingLns.forEach(el => tcPr.removeChild(el));
                        });
                        
                        // 2. 테두리 노드들을 역순으로 tcPr의 맨 첫머리에 insertBefore 하여 DrawingML의 엄격한 자식 시퀀스 순서(ln이 solidFill보다 앞에 와야 함)를 200% 완벽 준수!
                        // 역순 ['lnB', 'lnT', 'lnR', 'lnL'] 로 insertBefore(firstChild)를 하면
                        // 최종 tcPr 자식 순서는 [lnL, lnR, lnT, lnB, ...] 가 되어 XSD 스키마와 100% 정확하게 정렬됩니다!
                        const reversedBNames = ['lnB', 'lnT', 'lnR', 'lnL'];
                        
                        reversedBNames.forEach(bName => {
                            let borderColor = '7F7F7F'; // 기본 회색 (#7F7F7F)
                            
                            // 첫 번째 행이고, 첫 행 특별 포맷팅 옵션(applyFirstRowHeaderStyle)이 활성화된 경우에만 내부 실선 흰색 적용
                            if (rIdx === 0 && options.applyFirstRowHeaderStyle !== false) {
                                if (bName === 'lnB') {
                                    // 첫 행의 아래선은 첫째 행 and 둘째 행 사이의 내부 수평 실선이므로 흰색 적용
                                    borderColor = 'FFFFFF';
                                } else if (bName === 'lnL') {
                                    // 첫 행의 왼쪽선은 맨 왼쪽 외곽선(cIdx === 0)이 아니면 내부 수직 실선이므로 흰색 적용
                                    borderColor = (cIdx === 0) ? '7F7F7F' : 'FFFFFF';
                                } else if (bName === 'lnR') {
                                    // 첫 행의 오른쪽선은 맨 오른쪽 외곽선(cIdx === tcs.length - 1)이 아니면 내부 수직 실선이므로 흰색 적용
                                    const isLastCol = (cIdx === tcs.length - 1);
                                    borderColor = isLastCol ? '7F7F7F' : 'FFFFFF';
                                } else if (bName === 'lnT') {
                                    // 첫 행의 위선은 표 전체의 맨 위 외곽선이므로 회색 고정
                                    borderColor = '7F7F7F';
                                }
                            } else {
                                // 일반 셀 또는 첫 행 특별 포맷팅 옵션이 꺼진 경우: 외곽선 및 내부선 모두 0.5pt 회색 고정
                                borderColor = '7F7F7F';
                            }
                            
                            const ln = xmlDoc.createElementNS(nsA, `a:${bName}`);
                            ln.setAttribute('w', '6350'); // 0.5pt (1 pt = 12700 EMU -> 0.5 pt = 6350 EMU)
                            ln.setAttribute('cmpd', 'sng'); // 단일선 (Single line)
                            ln.setAttribute('cap', 'flat'); // 캡 스타일 평평하게
                            
                            const solidFill = xmlDoc.createElementNS(nsA, 'a:solidFill');
                            const srgbClr = xmlDoc.createElementNS(nsA, 'a:srgbClr');
                            srgbClr.setAttribute('val', borderColor); // 결정된 정밀 색상 주입
                            solidFill.appendChild(srgbClr);
                            ln.appendChild(solidFill);
                            
                            const prstDash = xmlDoc.createElementNS(nsA, 'a:prstDash');
                            prstDash.setAttribute('val', 'solid'); // 실선
                            ln.appendChild(prstDash);
                            
                            // tcPr의 첫 번째 자식으로 정교하게 삽입하여 스키마 정합성 보장!
                            tcPr.insertBefore(ln, tcPr.firstChild);
                        });
                        
                        // [규칙 2]: 첫 번째 행 (rIdx === 0) 이고 첫 행 특별 포맷팅 옵션이 켜져 있을 때만 특별 포맷팅 적용
                        if (rIdx === 0 && options.applyFirstRowHeaderStyle !== false) {
                            // 1. 셀 배경색 채우기: RGB 0,112,192 (#0070C0)
                            let existingFill = null;
                            for (let k = 0; k < tcPr.childNodes.length; k++) {
                                const child = tcPr.childNodes[k];
                                if (child.nodeType === 1 && (child.localName === 'solidFill' || child.tagName.split(':').pop() === 'solidFill')) {
                                    existingFill = child;
                                    break;
                                }
                            }
                            if (existingFill) {
                                tcPr.removeChild(existingFill);
                            }
                            
                            const bgSolidFill = xmlDoc.createElementNS(nsA, 'a:solidFill');
                            const bgSrgbClr = xmlDoc.createElementNS(nsA, 'a:srgbClr');
                            bgSrgbClr.setAttribute('val', '0072BA'); // 첫행 채우기색 0,114,186 (RGB 0,114,186)
                            bgSolidFill.appendChild(bgSrgbClr);
                            tcPr.appendChild(bgSolidFill);
                            
                            // 2. 첫 행의 텍스트 스타일링: 흰색, 11pt, KoPub동음체Bold (런속성 rPr 누락 대응형 강제 바인딩)
                            const textRuns = [];
                            const tcNodes = tc.getElementsByTagName('*');
                            for (let k = 0; k < tcNodes.length; k++) {
                                const node = tcNodes[k];
                                if (node.nodeType === 1) {
                                    const localName = node.localName || node.tagName.split(':').pop();
                                    if (localName === 'r' || localName === 'endParaRPr' || localName === 'defRPr') {
                                        textRuns.push(node);
                                    }
                                }
                            }
                            
                            for (let tR = 0; tR < textRuns.length; tR++) {
                                const run = textRuns[tR];
                                const localName = run.localName || run.tagName.split(':').pop();
                                
                                let rPr = null;
                                if (localName === 'endParaRPr' || localName === 'defRPr') {
                                    rPr = run;
                                } else {
                                    // r 요소인 경우 하위 rPr 검색 및 생성
                                    for (let k = 0; k < run.childNodes.length; k++) {
                                        const child = run.childNodes[k];
                                        if (child.nodeType === 1 && (child.localName === 'rPr' || child.tagName.split(':').pop() === 'rPr')) {
                                            rPr = child;
                                            break;
                                        }
                                    }
                                    if (!rPr) {
                                        rPr = xmlDoc.createElementNS(nsA, 'a:rPr');
                                        run.insertBefore(rPr, run.firstChild);
                                    }
                                }
                                
                                if (rPr) {
                                    rPr.setAttribute('sz', '1100'); // 11pt
                                    rPr.setAttribute('b', '1'); // Bold
                                    
                                    // 💡 endParaRPr 또는 defRPr 인 속성 전용 노드는 자식 엘리먼트(solidFill, font) 삽입 시 DOMException이 발생하므로 속성값만 세팅하고 자식 삽입은 안전하게 우회!
                                    const isAttrOnlyNode = (localName === 'endParaRPr' || localName === 'defRPr');
                                    
                                    if (!isAttrOnlyNode) {
                                        // 글씨색 흰색으로 변경
                                        let textFill = null;
                                        for (let k = 0; k < rPr.childNodes.length; k++) {
                                            const child = rPr.childNodes[k];
                                            if (child.nodeType === 1 && (child.localName === 'solidFill' || child.tagName.split(':').pop() === 'solidFill')) {
                                                textFill = child;
                                                break;
                                            }
                                        }
                                        if (textFill) {
                                            rPr.removeChild(textFill);
                                        }
                                        
                                        const textSolidFill = xmlDoc.createElementNS(nsA, 'a:solidFill');
                                        const textSrgbClr = xmlDoc.createElementNS(nsA, 'a:srgbClr');
                                        textSrgbClr.setAttribute('val', 'FFFFFF'); // 흰색
                                        textSolidFill.appendChild(textSrgbClr);
                                        
                                        // 스키마 시퀀스 안정을 위해 rPr의 맨 처음에 삽입
                                        rPr.insertBefore(textSolidFill, rPr.firstChild);
                                        
                                        // 폰트 변경 (latin, ea, cs)
                                        const fontTypes = ['latin', 'ea', 'cs'];
                                        fontTypes.forEach(fType => {
                                            let fontEl = null;
                                            for (let k = 0; k < rPr.childNodes.length; k++) {
                                                const child = rPr.childNodes[k];
                                                if (child.nodeType === 1 && (child.localName === fType || child.tagName.split(':').pop() === fType)) {
                                                    fontEl = child;
                                                    break;
                                                }
                                            }
                                            if (fontEl) {
                                                fontEl.setAttribute('typeface', 'KoPub동음체Bold');
                                            } else {
                                                fontEl = xmlDoc.createElementNS(nsA, `a:${fType}`);
                                                fontEl.setAttribute('typeface', 'KoPub동음체Bold');
                                                rPr.appendChild(fontEl);
                                            }
                                        });
                                    }
                                }
                            }
                        }
                    }
                }
                tableChanged = true;
            }
            if (tableChanged) {
                slideXmlStr = serializer.serializeToString(xmlDoc);
                fileChanged = true;
                hasChanges = true;
            }
        }
        
        const allElements = xmlDoc.getElementsByTagName('*');
        
        // 1. 단어 일괄 수정 \u0026 3. 폰트 일괄 교체 \u0026 4. 폰트 크기 변경 통합 루프
        for (let i = 0; i < allElements.length; i++) {
            const el = allElements[i];
            if (el.nodeType !== 1) continue;
            
            const localName = el.localName || el.tagName.split(':').pop();
            
            // 단어 수정 (t 태그)
            if (localName === 't' && replaceRules.length > 0) {
                let text = el.textContent;
                let originalText = text;
                for (const rule of replaceRules) {
                    text = text.split(rule.oldWord).join(rule.newWord);
                }
                if (text !== originalText) {
                    el.textContent = text;
                    fileChanged = true;
                    hasChanges = true;
                }
            }
            
            // 폰트 교체 (latin, ea, cs 태그)
            if (fontRules.length > 0 && (localName === 'latin' || localName === 'ea' || localName === 'cs')) {
                const typeface = el.getAttribute('typeface');
                if (typeface) {
                    for (const rule of fontRules) {
                        if (typeface === rule.oldWord) {
                            el.setAttribute('typeface', rule.newWord);
                            fileChanged = true;
                            hasChanges = true;
                            break;
                        }
                    }
                }
            }

            // 폰트 크기 변경 (전체 일괄 적용 및 매핑 적용)
            if (fontSizeRules.length > 0 && !slidePath.startsWith('ppt/theme/')) {
                // 1. 텍스트 런 및 필드 수준 처리 (r, fld, br)
                if (localName === 'r' || localName === 'fld' || localName === 'br') {
                    let rPr = null;
                    for (let j = 0; j < el.childNodes.length; j++) {
                        const child = el.childNodes[j];
                        if (child.nodeType === 1 && (child.localName === 'rPr' || child.tagName.split(':').pop() === 'rPr')) {
                            rPr = child;
                            break;
                        }
                    }
                    
                    if (rPr) {
                        const currentSz = parseInt(rPr.getAttribute('sz'));
                        for (const rule of fontSizeRules) {
                            if (rule.oldSize === null) {
                                // 전체 적용 모드
                                const newSzVal = Math.round(rule.newSize * 100).toString();
                                if (rPr.getAttribute('sz') !== newSzVal) {
                                    rPr.setAttribute('sz', newSzVal);
                                    fileChanged = true;
                                    hasChanges = true;
                                }
                                break;
                            } else {
                                // 매핑 모드
                                const oldSzVal = Math.round(rule.oldSize * 100);
                                if (currentSz === oldSzVal) {
                                    const newSzVal = Math.round(rule.newSize * 100).toString();
                                    rPr.setAttribute('sz', newSzVal);
                                    fileChanged = true;
                                    hasChanges = true;
                                    break;
                                }
                            }
                        }
                    } else {
                        // rPr이 없는 경우, oldSize가 null인 규칙이 있다면 생성하여 적용
                        for (const rule of fontSizeRules) {
                            if (rule.oldSize === null) {
                                rPr = xmlDoc.createElementNS(nsA, 'a:rPr');
                                rPr.setAttribute('sz', Math.round(rule.newSize * 100).toString());
                                el.insertBefore(rPr, el.firstChild);
                                fileChanged = true;
                                hasChanges = true;
                                break;
                            }
                        }
                    }
                }
                
                // 2. 단락 기본 및 종료 스타일 처리 (defRPr, endParaRPr)
                if (localName === 'defRPr' || localName === 'endParaRPr') {
                    const currentSz = parseInt(el.getAttribute('sz'));
                    for (const rule of fontSizeRules) {
                        if (rule.oldSize === null) {
                            const newSzVal = Math.round(rule.newSize * 100).toString();
                            if (el.getAttribute('sz') !== newSzVal) {
                                el.setAttribute('sz', newSzVal);
                                fileChanged = true;
                                hasChanges = true;
                            }
                            break;
                        } else {
                            const oldSzVal = Math.round(rule.oldSize * 100);
                            if (currentSz === oldSzVal) {
                                const newSzVal = Math.round(rule.newSize * 100).toString();
                                el.setAttribute('sz', newSzVal);
                                fileChanged = true;
                                hasChanges = true;
                                break;
                            }
                        }
                    }
                }
            }
        }

        // 2. 텍스트 디자인 일괄 변경
        if (applyDesign && designTargetFilesSet.has(slidePath)) {
            const xmlDocInner = parser.parseFromString(slideXmlStr, 'application/xml'); // 새로 파싱하거나 기존 것 사용
            if (xmlDocInner.getElementsByTagName('parsererror').length === 0) {
                function applyLnToRPr(rPr) {
                    let existingLn = null;
                    for (let j = 0; j < rPr.childNodes.length; j++) {
                        const child = rPr.childNodes[j];
                        if (child.nodeType === 1 && (child.localName === 'ln' || child.tagName.split(':').pop() === 'ln')) {
                            existingLn = child;
                            break;
                        }
                    }
                    if (existingLn) rPr.removeChild(existingLn);

                    const ln = xmlDocInner.createElementNS(nsA, 'a:ln');
                    ln.setAttribute('w', '9525');
                    ln.setAttribute('cmpd', 'sng');

                    const solidFill = xmlDocInner.createElementNS(nsA, 'a:solidFill');
                    const srgbClr = xmlDocInner.createElementNS(nsA, 'a:srgbClr');
                    srgbClr.setAttribute('val', 'FFFFFF');

                    const alpha = xmlDocInner.createElementNS(nsA, 'a:alpha');
                    alpha.setAttribute('val', '0');

                    srgbClr.appendChild(alpha);
                    solidFill.appendChild(srgbClr);
                    ln.appendChild(solidFill);
                    
                    const prstDash = xmlDocInner.createElementNS(nsA, 'a:prstDash');
                    prstDash.setAttribute('val', 'solid');
                    ln.appendChild(prstDash);
                    
                    rPr.insertBefore(ln, rPr.firstChild);
                }

                designChanged = false;
                const allElementsInner = xmlDocInner.getElementsByTagName('*');
                for (let i = 0; i < allElementsInner.length; i++) {
                    const el = allElementsInner[i];
                    if (el.nodeType !== 1) continue;
                    const localName = el.localName || el.tagName.split(':').pop();
                    
                    if (localName === 'r' || localName === 'fld' || localName === 'br') {
                        let rPr = null;
                        for (let j = 0; j < el.childNodes.length; j++) {
                            const child = el.childNodes[j];
                            if (child.nodeType === 1 && (child.localName === 'rPr' || child.tagName.split(':').pop() === 'rPr')) {
                                rPr = child;
                                break;
                            }
                        }
                        if (!rPr) {
                            rPr = xmlDocInner.createElementNS(nsA, 'a:rPr');
                            el.insertBefore(rPr, el.firstChild);
                        }
                        applyLnToRPr(rPr);
                        designChanged = true;
                        hasChanges = true;
                    } else if (localName === 'endParaRPr' || localName === 'defRPr') {
                        applyLnToRPr(el);
                        designChanged = true;
                        hasChanges = true;
                    }
                }
                
                if (designChanged) {
                    slideXmlStr = serializer.serializeToString(xmlDocInner);
                    fileChanged = true;
                }
            }
        }
        
        // 단어/폰트/크기/테이블 수정 루프에서 변경이 가해진 경우 최종 xmlDoc 동기화
        if (fileChanged && !designChanged) {
            slideXmlStr = serializer.serializeToString(xmlDoc);
        }
        
        // 변경사항이 있으면 압축 파일 갱신
        if (fileChanged) {
            zip.file(slidePath, slideXmlStr);
        }
    });

    if (!hasChanges) {
        // 단어 수정이 실패했거나 디자인 변경 대상이 없었을 경우
        // 에러를 던질지 원본을 그냥 반환할지는 기획에 따르나, 성공 메시지를 위해 원본으로 진행
    }

    const blob = zip.generate({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    });

    // 💡 동적 커스텀 프로퍼티 바인딩으로 파일별 수정 메타데이터 보존
    blob.totalTablesCount = totalTablesCount;
    blob.hasChanges = hasChanges;

    return blob;
}

/**
 * PPT 파일의 모든 슬라이드 객체에 스마트 애니메이션(순차적 나타나기)을 추가합니다.
 * @param {File} pptFile 
 * @param {Object} options { animationType: 'fade' | 'appear' }
 */
/**
 * PPT 파일에 스마트 애니메이션 또는 슬라이드 전환 효과를 추가합니다.
 * @param {File} pptFile 
 * @param {Object} options { animationType: 'fade' | 'appear' | 'transition', perSlideConfigs: Array, useGrouping: boolean }
 */
export async function addSmartAnimationsToPpt(pptFile, options = {}) {
    const { animationType = 'fade', perSlideConfigs = [], useGrouping = false } = options;
    const arrayBuffer = await pptFile.arrayBuffer();
    
    const zip = new JSZip();
    await zip.loadAsync(arrayBuffer);
    
    const slideFiles = Object.keys(zip.files)
        .filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml'))
        .sort((a, b) => {
            const numA = parseInt(a.match(/\d+/)[0]);
            const numB = parseInt(b.match(/\d+/)[0]);
            return numA - numB;
        });

    const parser = new DOMParser();
    const serializer = new XMLSerializer();

    for (let idx = 0; idx < slideFiles.length; idx++) {
        const slidePath = slideFiles[idx];
        const config = perSlideConfigs[idx] || { enabled: true, type: animationType, useGrouping: useGrouping };
        
        const content = await zip.file(slidePath).async('string');
        const xmlDoc = parser.parseFromString(content, 'application/xml');
        const nsP = "http://schemas.openxmlformats.org/presentationml/2006/main";

        // 1. 기존 애니메이션 및 전환 효과 초기화
        const existingTimings = xmlDoc.getElementsByTagNameNS(nsP, 'timing');
        while (existingTimings.length > 0) existingTimings[0].parentNode.removeChild(existingTimings[0]);
        
        const existingTransitions = xmlDoc.getElementsByTagNameNS(nsP, 'transition');
        while (existingTransitions.length > 0) existingTransitions[0].parentNode.removeChild(existingTransitions[0]);

        if (!config.enabled) {
            zip.file(slidePath, serializer.serializeToString(xmlDoc));
            continue;
        }

        const currentType = config.type || animationType;

        // 2. '전환 효과(Transition)' 모드인 경우
        if (currentType === 'transition') {
            const transitionNode = xmlDoc.createElementNS(nsP, 'p:transition');
            transitionNode.setAttribute('dur', '1000');
            const fadeNode = xmlDoc.createElementNS(nsP, 'p:fade');
            transitionNode.appendChild(fadeNode);
            
            // 주입 위치: cSld, clrMapOvr 뒤, timing, extLst 앞
            const timing = xmlDoc.getElementsByTagNameNS(nsP, 'timing')[0];
            const extLst = xmlDoc.getElementsByTagNameNS(nsP, 'extLst')[0];
            const beforeNode = timing || extLst;
            
            if (beforeNode && beforeNode.parentNode === xmlDoc.documentElement) {
                xmlDoc.documentElement.insertBefore(transitionNode, beforeNode);
            } else {
                xmlDoc.documentElement.appendChild(transitionNode);
            }
        } 
        // 3. '객체 애니메이션' 모드인 경우
        else {
            const currentGrouping = config.useGrouping !== undefined ? config.useGrouping : useGrouping;
            const shapesWithPos = [];
            const processElements = (tagName) => {
                const elements = xmlDoc.getElementsByTagNameNS('*', tagName);
                for (let i = 0; i < elements.length; i++) {
                    const el = elements[i];
                    const cNvPr = el.getElementsByTagNameNS('*', 'cNvPr')[0];
                    if (!cNvPr) continue;
                    const id = cNvPr.getAttribute('id');
                    const name = (cNvPr.getAttribute('name') || '').toLowerCase();
                    const off = el.getElementsByTagNameNS('*', 'off')[0];
                    let x = 0, y = 0;
                    if (off) {
                        x = parseInt(off.getAttribute('x') || '0');
                        y = parseInt(off.getAttribute('y') || '0');
                    }
                    if (name.includes('title') || name.includes('header') || name.includes('footer') || 
                        name.includes('number') || name.includes('page') || name.includes('placeholder')) continue;
                    if (y < 1100000) continue;
                    shapesWithPos.push({ id, x, y });
                }
            };

            processElements('sp');
            processElements('pic');
            processElements('graphicFrame');

            if (shapesWithPos.length > 0) {
                shapesWithPos.sort((a, b) => {
                    const yDiff = a.y - b.y;
                    if (Math.abs(yDiff) < 100000) return a.x - b.x;
                    return yDiff;
                });

                const sortedIds = shapesWithPos.map(s => s.id);
                const timingXml = generateTimingXml(sortedIds, currentType, currentGrouping);
                const timingDoc = parser.parseFromString(timingXml, 'application/xml');
                const timingNode = xmlDoc.importNode(timingDoc.documentElement, true);
                
                const extLst = xmlDoc.getElementsByTagNameNS(nsP, 'extLst')[0];
                if (extLst && extLst.parentNode === xmlDoc.documentElement) {
                    xmlDoc.documentElement.insertBefore(timingNode, extLst);
                } else {
                    xmlDoc.documentElement.appendChild(timingNode);
                }
            }
        }
        
        zip.file(slidePath, serializer.serializeToString(xmlDoc));
    }

    const modifiedBlob = await zip.generateAsync({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    });

    return modifiedBlob;
}

/**
 * OpenXML p:timing 구조 생성
 * @param {Array} shapeIds 
 * @param {string} type 
 * @param {boolean} useGrouping true이면 모든 객체가 한 번의 클릭으로 동시에 나타남
 */
function generateTimingXml(shapeIds, type, useGrouping = false) {
    let nodes = '';
    
    if (useGrouping) {
        // 모든 객체를 하나의 p:par 안에 넣어서 동시에 실행
        let groupChildNodes = '';
        shapeIds.forEach((id, idx) => {
            groupChildNodes += `
                <p:set>
                    <p:cBhvr>
                        <p:cTn id="${idx + 1000}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>
                        <p:tgtEl><p:spTgt spid="${id}"/></p:tgtEl>
                        <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
                    </p:cBhvr>
                    <p:to><p:str val="visible"/></p:to>
                </p:set>
                <p:anim filter="${type === 'fade' ? 'fade(in)' : 'appear'}" calcmode="lin" transition="in">
                    <p:cBhvr>
                        <p:cTn id="${idx + 2000}" dur="${type === 'fade' ? '500' : '1'}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>
                        <p:tgtEl><p:spTgt spid="${id}"/></p:tgtEl>
                    </p:cBhvr>
                </p:anim>`;
        });

        nodes = `
        <p:par>
            <p:cTn id="10" dur="indefinite" fill="hold" nodeType="clickEffect">
                <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                <p:childTnLst>
                    ${groupChildNodes}
                </p:childTnLst>
            </p:cTn>
        </p:par>`;
    } else {
        // 기존: 개체별 순차 실행
        shapeIds.forEach((id, idx) => {
            nodes += `
            <p:par>
                <p:cTn id="${idx + 10}" dur="indefinite" fill="hold" nodeType="clickEffect">
                    <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                    <p:childTnLst>
                        <p:set>
                            <p:cBhvr>
                                <p:cTn id="${idx + 1000}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>
                                <p:tgtEl><p:spTgt spid="${id}"/></p:tgtEl>
                                <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
                            </p:cBhvr>
                            <p:to><p:str val="visible"/></p:to>
                        </p:set>
                        <p:anim filter="${type === 'fade' ? 'fade(in)' : 'appear'}" calcmode="lin" transition="in">
                            <p:cBhvr>
                                <p:cTn id="${idx + 2000}" dur="${type === 'fade' ? '500' : '1'}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>
                                <p:tgtEl><p:spTgt spid="${id}"/></p:tgtEl>
                            </p:cBhvr>
                        </p:anim>
                    </p:childTnLst>
                </p:cTn>
            </p:par>`;
        });
    }

    return `
    <p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
        <p:tnLst>
            <p:par>
                <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
                    <p:childTnLst>
                        <p:seq type="nextClick" concurrent="1" nextAc="seek">
                            <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                                <p:childTnLst>
                                    ${nodes}
                                </p:childTnLst>
                            </p:cTn>
                        </p:seq>
                    </p:childTnLst>
                </p:cTn>
            </p:par>
        </p:tnLst>
    </p:timing>`;
}

/**
 * PPT 파일의 슬라이드 개수를 반환합니다.
 */
export async function getPptSlideCount(pptFile) {
    const arrayBuffer = await pptFile.arrayBuffer();
    const zip = new JSZip();
    await zip.loadAsync(arrayBuffer);
    const slideFiles = Object.keys(zip.files).filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml'));
    return slideFiles.length;
}

