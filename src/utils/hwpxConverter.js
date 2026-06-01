import JSZip from 'jszip';
import { HWPXBuilder, write } from 'hwpx-js';

/**
 * PPTX 파일의 내용을 파싱하여 표준 한글 규격인 HWPX 파일 Blob으로 변환합니다.
 * @param {File} pptxFile 
 * @returns {Promise<Blob>} 변환된 HWPX Blob (통계 프로퍼티가 포함됨)
 */
export async function convertPptxToHwpx(pptxFile) {
    if (!pptxFile) throw new Error('변환할 파워포인트(PPTX) 파일이 존재하지 않습니다.');

    const arrayBuffer = await pptxFile.arrayBuffer();
    const zip = new JSZip();
    await zip.loadAsync(arrayBuffer);

    // 1. 슬라이드 XML 파일 탐색 및 숫자순 정렬
    const slideFiles = Object.keys(zip.files)
        .filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml'))
        .sort((a, b) => {
            const numA = parseInt(a.match(/\d+/)[0]);
            const numB = parseInt(b.match(/\d+/)[0]);
            return numA - numB;
        });

    if (slideFiles.length === 0) {
        throw new Error('파워포인트 파일 내에 슬라이드 XML 데이터가 존재하지 않습니다.');
    }

    const builder = new HWPXBuilder();
    builder.setPageSettings({
        pageWidth: 210, // A4 세로 규격 (mm)
        pageHeight: 297,
        pageMargin: { left: 20, right: 20, top: 20, bottom: 20, header: 15, footer: 15 }
    });

    const parser = new DOMParser();
    const nsA = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    const nsP = 'http://schemas.openxmlformats.org/presentationml/2006/main';
    const nsR = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

    let totalParagraphs = 0;
    let totalTables = 0;
    let totalImages = 0;

    // 슬라이드별 루프 가동
    for (let idx = 0; idx < slideFiles.length; idx++) {
        const slidePath = slideFiles[idx];
        const slideXmlStr = await zip.files[slidePath].async('text');
        const xmlDoc = parser.parseFromString(slideXmlStr, 'application/xml');

        if (xmlDoc.getElementsByTagName('parsererror').length > 0) {
            console.error(`Slide ${idx + 1} 파싱 에러 건너뜀`);
            continue;
        }

        // Relationships (.rels) 파일을 읽어 이미지 리소스 경로 해독 준비
        const slideNum = slidePath.match(/\d+/)[0];
        const relsPath = `ppt/slides/_rels/slide${slideNum}.xml.rels`;
        const relsMap = new Map();
        if (zip.files[relsPath]) {
            const relsStr = await zip.files[relsPath].async('text');
            const relsDoc = parser.parseFromString(relsStr, 'application/xml');
            const relationships = relsDoc.getElementsByTagName('Relationship');
            for (let r = 0; r < relationships.length; r++) {
                const rel = relationships[r];
                const rId = rel.getAttribute('Id');
                const target = rel.getAttribute('Target');
                relsMap.set(rId, target);
            }
        }

        const elements = [];

        // 도형 및 텍스트 상자 (<p:sp>) 추출
        const spList = xmlDoc.getElementsByTagNameNS(nsP, 'sp');
        for (let i = 0; i < spList.length; i++) {
            const sp = spList[i];
            
            let x = 0, y = 0;
            const off = sp.getElementsByTagNameNS(nsA, 'off')[0];
            if (off) {
                x = parseInt(off.getAttribute('x') || '0');
                y = parseInt(off.getAttribute('y') || '0');
            }

            const pList = sp.getElementsByTagNameNS(nsA, 'p');
            const paragraphs = [];
            for (let j = 0; j < pList.length; j++) {
                const p = pList[j];
                const tList = p.getElementsByTagNameNS(nsA, 't');
                let paraText = '';
                for (let k = 0; k < tList.length; k++) {
                    paraText += tList[k].textContent || '';
                }
                const trimmed = paraText.trim();
                if (trimmed) {
                    paragraphs.push(trimmed);
                }
            }

            if (paragraphs.length > 0) {
                elements.push({
                    type: 'text',
                    x,
                    y,
                    data: paragraphs
                });
            }
        }

        // 표 (<a:tbl>) 추출
        const tblList = xmlDoc.getElementsByTagNameNS(nsA, 'tbl');
        for (let i = 0; i < tblList.length; i++) {
            const tbl = tblList[i];
            
            let x = 0, y = 0;
            let parent = tbl.parentNode;
            while (parent) {
                const pName = parent.localName || parent.tagName.split(':').pop();
                if (pName === 'graphicFrame') {
                    const off = parent.getElementsByTagNameNS(nsA, 'off')[0];
                    if (off) {
                        x = parseInt(off.getAttribute('x') || '0');
                        y = parseInt(off.getAttribute('y') || '0');
                    }
                    break;
                }
                parent = parent.parentNode;
            }

            const trList = tbl.getElementsByTagNameNS(nsA, 'tr');
            const tableData = [];
            for (let r = 0; r < trList.length; r++) {
                const tr = trList[r];
                const tcList = tr.getElementsByTagNameNS(nsA, 'tc');
                const rowData = [];
                for (let c = 0; c < tcList.length; c++) {
                    const tc = tcList[c];
                    const tList = tc.getElementsByTagNameNS(nsA, 't');
                    let cellText = '';
                    for (let k = 0; k < tList.length; k++) {
                        cellText += tList[k].textContent || '';
                    }
                    rowData.push(cellText.trim());
                }
                tableData.push(rowData);
            }

            if (tableData.length > 0) {
                elements.push({
                    type: 'table',
                    x,
                    y,
                    data: tableData
                });
            }
        }

        // 이미지 (<p:pic>) 추출
        const picList = xmlDoc.getElementsByTagNameNS(nsP, 'pic');
        for (let i = 0; i < picList.length; i++) {
            const pic = picList[i];
            
            let x = 0, y = 0;
            const off = pic.getElementsByTagNameNS(nsA, 'off')[0];
            if (off) {
                x = parseInt(off.getAttribute('x') || '0');
                y = parseInt(off.getAttribute('y') || '0');
            }

            const blip = pic.getElementsByTagNameNS(nsA, 'blip')[0];
            if (blip) {
                const rId = blip.getAttributeNS(nsR, 'embed') || blip.getAttribute('r:embed');
                if (rId && relsMap.has(rId)) {
                    let targetPath = relsMap.get(rId);
                    if (targetPath.startsWith('../')) {
                        targetPath = 'ppt/' + targetPath.replace('../', '');
                    } else if (!targetPath.startsWith('ppt/')) {
                        targetPath = 'ppt/' + targetPath;
                    }

                    if (zip.files[targetPath]) {
                        const imgData = await zip.files[targetPath].async('uint8array');
                        const ext = targetPath.split('.').pop().toLowerCase();
                        elements.push({
                            type: 'image',
                            x,
                            y,
                            ext: ext === 'jpeg' ? 'jpg' : ext,
                            data: imgData
                        });
                    }
                }
            }
        }

        // 좌표 기준 순서 정렬
        elements.sort((a, b) => {
            const yDiff = a.y - b.y;
            if (Math.abs(yDiff) < 360000) {
                return a.x - b.x;
            }
            return yDiff;
        });

        builder.addParagraph(`■ 슬라이드 ${idx + 1}`, { fontSize: 13, bold: true, color: '#a855f7' });
        builder.addEmptyParagraph();

        for (const el of elements) {
            if (el.type === 'text') {
                el.data.forEach(text => {
                    builder.addParagraph(text, { fontSize: 11 });
                    totalParagraphs++;
                });
                builder.addEmptyParagraph();
            } else if (el.type === 'table') {
                builder.addTable(el.data, {
                    width: 170,
                    borderStyle: 'SOLID',
                    cellPadding: 1.2
                });
                totalTables++;
                builder.addEmptyParagraph();
            } else if (el.type === 'image') {
                try {
                    builder.addImage(el.data, el.ext, { width: 120, height: 90 });
                    totalImages++;
                    builder.addEmptyParagraph();
                } catch (imgErr) {
                    console.error("그림 주입 실패 건너뜀:", imgErr);
                }
            }
        }
    }

    const doc = builder.build();
    const uint8 = await write(doc);
    
    const blob = new Blob([uint8], { type: 'application/x-hwp-hwpx' });
    blob.totalSlidesCount = slideFiles.length;
    blob.totalParagraphsCount = totalParagraphs;
    blob.totalTablesCount = totalTables;
    blob.totalImagesCount = totalImages;

    return blob;
}

// ==================== PPTX 서식 이식형 융합 기능 섹션 ====================

/**
 * PPTX의 a:tc 노드에서 단락 및 런들을 순회하며 서식 정보(Bold, Color)와 글머리 기호(bullet)가 있는 텍스트 리스트를 수집합니다.
 */
function extractPptCellStyledText(tc, nsA) {
    const paragraphs = [];
    const pList = tc.getElementsByTagNameNS('*', 'p'); // a:p
    
    for (let i = 0; i < pList.length; i++) {
        const p = pList[i];
        const pRuns = [];
        
        // 1. 문단 글머리(불릿) 기호 감지 및 주입
        const pPr = p.getElementsByTagNameNS(nsA, 'pPr')[0] || p.getElementsByTagName('a:pPr')[0];
        if (pPr) {
            const buChar = pPr.getElementsByTagNameNS(nsA, 'buChar')[0] || pPr.getElementsByTagName('a:buChar')[0];
            if (buChar && buChar.hasAttribute('char')) {
                const bulletChar = buChar.getAttribute('char');
                pRuns.push({
                    text: bulletChar + ' ',
                    bold: false,
                    color: null
                });
            }
        }
        
        // 2. 개별 런 순회
        for (let j = 0; j < p.childNodes.length; j++) {
            const child = p.childNodes[j];
            const tagName = child.localName || child.tagName?.split(':').pop();
            
            if (tagName === 'r') {
                const textElem = child.getElementsByTagNameNS(nsA, 't')[0] || child.getElementsByTagName('a:t')[0];
                if (textElem && textElem.textContent) {
                    let bold = false;
                    let color = null;
                    
                    const rPr = child.getElementsByTagNameNS(nsA, 'rPr')[0] || child.getElementsByTagName('a:rPr')[0];
                    if (rPr) {
                        const bVal = rPr.getAttribute('b');
                        if (bVal === '1' || bVal === 'true') {
                            bold = true;
                        }
                        const solidFill = rPr.getElementsByTagNameNS(nsA, 'solidFill')[0] || rPr.getElementsByTagName('a:solidFill')[0];
                        if (solidFill) {
                            const srgbClr = solidFill.getElementsByTagNameNS(nsA, 'srgbClr')[0] || solidFill.getElementsByTagName('a:srgbClr')[0];
                            if (srgbClr && srgbClr.hasAttribute('val')) {
                                color = '#' + srgbClr.getAttribute('val');
                            }
                        }
                    }
                    pRuns.push({
                        text: textElem.textContent,
                        bold,
                        color
                    });
                }
            } else if (tagName === 'br') {
                pRuns.push({
                    text: '\n',
                    bold: false,
                    color: null
                });
            }
        }
        paragraphs.push(pRuns);
    }
    return paragraphs;
}

/**
 * 서식 정보가 있는 문단 리스트에서 순수 텍스트 줄글을 추출합니다.
 */
function getPptPlainStructuredText(paras) {
    return paras.map(p => p.map(r => r.text).join('')).join('\n').trim();
}

/**
 * 고유번호, 명칭, 정의 등 한 줄 평탄화 처리를 위한 자바스크립트 헬퍼 함수
 */
function flattenParagraphsToSingleLine(paragraphs) {
    if (!paragraphs || paragraphs.length === 0) return [];
    
    let singleRunText = "";
    let isBold = false;
    let color = null;
    
    for (let i = 0; i < paragraphs.length; i++) {
        const pRuns = paragraphs[i];
        let pText = "";
        for (let j = 0; j < pRuns.length; j++) {
            const run = pRuns[j];
            if (run.text === '• ') continue;
            
            const t = run.text.replace(/\n/g, ' ').replace(/\r/g, ' ').trim();
            if (t) {
                pText += t;
                if (run.bold) isBold = true;
                if (run.color) color = run.color;
            }
        }
        if (pText) {
            if (singleRunText) {
                singleRunText += " " + pText;
            } else {
                singleRunText = pText;
            }
        }
    }
    
    if (singleRunText) {
        return [[{ text: singleRunText, bold: isBold, color: color }]];
    }
    return [];
}

/**
 * HWPX header.xml 내 charProperties에 글자 모양을 동적 등록해주는 헬퍼 클래스
 */
class HwpxCharPrRegistry {
    constructor(headerDoc) {
        this.headerDoc = headerDoc;
        this.hpNS = 'http://www.hancom.co.kr/hwpml/2011/paragraph';
        this.hhNS = 'http://www.hancom.co.kr/hwpml/2011/head';
        
        this.charProperties = headerDoc.getElementsByTagNameNS(this.hhNS, 'charProperties')[0] 
                           || headerDoc.getElementsByTagName('hh:charProperties')[0];
                           
        this.registry = new Map(); // key: baseId-bold-color -> charPrId
        this.maxId = -1;
        
        if (this.charProperties) {
            const charPrList = this.charProperties.getElementsByTagNameNS(this.hhNS, 'charPr')
                            || this.charProperties.getElementsByTagName('hh:charPr');
            for (let i = 0; i < charPrList.length; i++) {
                const cId = parseInt(charPrList[i].getAttribute('id') || '-1');
                if (cId > this.maxId) {
                    this.maxId = cId;
                }
            }
        }
    }
    
    getOrCreateCharPr(baseCharPrId, bold, color) {
        const charPrList = this.charProperties.getElementsByTagNameNS(this.hhNS, 'charPr')
                        || this.charProperties.getElementsByTagName('hh:charPr');
                        
        let basePr = null;
        for (let i = 0; i < charPrList.length; i++) {
            if (charPrList[i].getAttribute('id') === String(baseCharPrId)) {
                basePr = charPrList[i];
                break;
            }
        }
        
        if (!basePr) {
            for (let i = 0; i < charPrList.length; i++) {
                if (charPrList[i].getAttribute('id') === '0') {
                    basePr = charPrList[i];
                    break;
                }
            }
        }
        
        if (!basePr && charPrList.length > 0) {
            basePr = charPrList[0];
        }
        
        const targetColor = color || (basePr ? basePr.getAttribute('textColor') : '#000000') || '#000000';
        
        const cacheKey = `${baseCharPrId}-${bold}-${targetColor}`;
        if (this.registry.has(cacheKey)) {
            return this.registry.get(cacheKey);
        }
        
        const parentHasBold = basePr ? (basePr.getElementsByTagNameNS(this.hhNS, 'bold')[0] || basePr.getElementsByTagName('hh:bold')[0]) !== undefined : false;
        const parentColor = basePr ? basePr.getAttribute('textColor') : '#000000';
        
        if (parentHasBold === bold && parentColor === targetColor) {
            return baseCharPrId;
        }
        
        this.maxId++;
        const newId = String(this.maxId);
        
        const newPr = basePr.cloneNode(true);
        newPr.setAttribute('id', newId);
        newPr.setAttribute('textColor', targetColor);
        
        let boldElem = newPr.getElementsByTagNameNS(this.hhNS, 'bold')[0] 
                    || newPr.getElementsByTagName('hh:bold')[0];
                    
        if (bold) {
            if (!boldElem) {
                boldElem = this.headerDoc.createElementNS(this.hhNS, 'hh:bold');
                
                let inserted = false;
                const childNodes = Array.from(newPr.childNodes);
                for (let i = 0; i < childNodes.length; i++) {
                    const tag = childNodes[i].localName || childNodes[i].tagName?.split(':').pop();
                    if (['underline', 'strikeout', 'outline', 'shadow'].includes(tag)) {
                        newPr.insertBefore(boldElem, childNodes[i]);
                        inserted = true;
                        break;
                    }
                }
                if (!inserted) {
                    newPr.appendChild(boldElem);
                }
            }
        } else {
            if (boldElem) {
                newPr.removeChild(boldElem);
            }
        }
        
        this.charProperties.appendChild(newPr);
        this.charProperties.setAttribute('itemCnt', String(this.charProperties.getElementsByTagNameNS(this.hhNS, 'charPr').length || charPrList.length + 1));
        
        this.registry.set(cacheKey, newId);
        return newId;
    }
}

/**
 * PPTX의 표 데이터를 파싱하여 HWPX 양식 문서에 주입 및 병합/분할 생성합니다.
 * @param {File} pptxFile 
 * @param {File} hwpxTemplateFile 
 * @param {boolean} mergeMode 하나의 파일로 병합할지 여부 (true: 단일 HWPX 병합, false: 개별 HWPX들을 담은 ZIP 압축)
 * @returns {Promise<Blob>} 최종 생성된 HWPX Blob 혹은 ZIP Blob
 */
export async function fusePptToHwpxTemplate(pptxFile, hwpxTemplateFile, mergeMode = true) {
    if (!pptxFile || !hwpxTemplateFile) {
        throw new Error('PPTX 요건기술서 파일과 HWPX 양식 템플릿 파일을 모두 등록해주세요.');
    }

    const pptxBuffer = await pptxFile.arrayBuffer();
    const hwpxBuffer = await hwpxTemplateFile.arrayBuffer();

    const pptxZip = new JSZip();
    await pptxZip.loadAsync(pptxBuffer);

    // 1. PPTX 슬라이드별 표 데이터 파싱
    const slideFiles = Object.keys(pptxZip.files)
        .filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml'))
        .sort((a, b) => {
            const numA = parseInt(a.match(/\d+/)[0]);
            const numB = parseInt(b.match(/\d+/)[0]);
            return numA - numB;
        });

    const parser = new DOMParser();
    const nsA = 'http://schemas.openxmlformats.org/drawingml/2006/main';

    const requirementsMap = new Map();
    const requirementsOrder = [];
    const idPattern = /^[A-Za-z0-9]+[\-_][A-Za-z0-9]+$/;
    const cleanPattern = /^(.*?)\s*\(\s*\d+\s*\/\s*\d+\s*\)\s*$/; // (1/2) 괄호 검출 정규식

    for (const slidePath of slideFiles) {
        const slideXmlStr = await pptxZip.files[slidePath].async('text');
        const xmlDoc = parser.parseFromString(slideXmlStr, 'application/xml');
        
        const tblList = xmlDoc.getElementsByTagNameNS(nsA, 'tbl');
        for (let i = 0; i < tblList.length; i++) {
            const tbl = tblList[i];
            const trList = tbl.getElementsByTagNameNS(nsA, 'tr');
            if (trList.length === 0) continue;
            
            // 💡 [비즈니스 룰 2] 2번째 행이 1개의 열로 가로 병합된 표 스킵
            if (trList.length > 1) {
                const secondRowCells = trList[1].getElementsByTagNameNS(nsA, 'tc');
                let isMerged = false;
                
                if (secondRowCells.length === 1) {
                    isMerged = true;
                } else if (secondRowCells.length > 1) {
                    const firstCell = secondRowCells[0];
                    const gridSpan = firstCell.getAttribute('gridSpan');
                    if (gridSpan && parseInt(gridSpan) >= 4) {
                        isMerged = true;
                    } else {
                        const cellTexts = [];
                        for (let k = 0; k < secondRowCells.length; k++) {
                            const tc = secondRowCells[k];
                            const tList = tc.getElementsByTagNameNS(nsA, 't');
                            let tTxt = '';
                            for (let tIdx = 0; tIdx < tList.length; tIdx++) {
                                tTxt += tList[tIdx].textContent || '';
                            }
                            cellTexts.push(tTxt.trim());
                        }
                        const nonEmptyTexts = cellTexts.filter(t => t);
                        if (nonEmptyTexts.length <= 1 && cellTexts[0].length > 0) {
                            const firstTxt = cellTexts[0];
                            if (firstTxt.length > 40 || firstTxt.includes('※') || firstTxt.includes('■') || firstTxt.includes('요건')) {
                                isMerged = true;
                            }
                        }
                    }
                }
                
                if (isMerged) {
                    console.log(`[스킵] ${slidePath}의 표는 2번째 행이 1개 열로 병합되어 있어 작성을 건너뜁니다.`);
                    continue;
                }
            }
            
            const firstRowCells = trList[0].getElementsByTagNameNS(nsA, 'tc');
            const headerTexts = [];
            for (let k = 0; k < firstRowCells.length; k++) {
                const tc = firstRowCells[k];
                const tList = tc.getElementsByTagNameNS(nsA, 't');
                let hText = '';
                for (let tIdx = 0; tIdx < tList.length; tIdx++) {
                    hText += tList[tIdx].textContent || '';
                }
                headerTexts.push(hText.trim());
            }
            
            const firstColText = headerTexts[0] || '';
            const isFirstRowData = idPattern.test(firstColText) || (
                firstColText.length >= 3 && firstColText.length <= 15 &&
                !["고유번호", "요구사항", "분류", "No", "ID"].some(kw => firstColText.includes(kw))
            );
            
            let idIdx = 0, nameIdx = 1, descIdx = 2, detailIdx = 3;
            
            if (!isFirstRowData) {
                for (let cIdx = 0; cIdx < headerTexts.length; cIdx++) {
                    const hText = headerTexts[cIdx];
                    if (hText.length > 30) continue;
                    
                    if (hText.includes('고유번호') || hText.includes('ID')) {
                        idIdx = cIdx;
                    } else if (hText.includes('명칭') || hText.includes('요구사항명')) {
                        nameIdx = cIdx;
                    } else if (hText.includes('정의') || hText.includes('개요')) {
                        descIdx = cIdx;
                    } else if (hText.includes('세부내용') || hText.includes('요건') || hText.includes('상세설명')) {
                        detailIdx = cIdx;
                    }
                }
            }
            
            for (let j = 0; j < trList.length; j++) {
                const tr = trList[j];
                const tcList = tr.getElementsByTagNameNS(nsA, 'tc');
                
                if (tcList.length >= 4) {
                    const safeIdIdx = idIdx < tcList.length ? idIdx : 0;
                    const safeNameIdx = nameIdx < tcList.length ? nameIdx : 1;
                    const safeDescIdx = descIdx < tcList.length ? descIdx : 2;
                    const safeDetailIdx = detailIdx < tcList.length ? detailIdx : (tcList.length > 3 ? 3 : tcList.length - 1);
                    
                    const idParas = extractPptCellStyledText(tcList[safeIdIdx], nsA);
                    const nameParas = extractPptCellStyledText(tcList[safeNameIdx], nsA);
                    const descParas = extractPptCellStyledText(tcList[safeDescIdx], nsA);
                    const detailParas = extractPptCellStyledText(tcList[safeDetailIdx], nsA);
                    
                    const col0 = getPptPlainStructuredText(idParas).replace(/\n/g, ' ').replace(/\r/g, ' ').trim();
                    
                    const isHeader = col0.includes('고유번호') || col0.includes('요구사항') || col0.includes('분류') || col0.includes('No');
                    const hasValidId = idPattern.test(col0) || (col0.length >= 3 && col0.length <= 15 && !isHeader) || cleanPattern.test(col0);
                    
                    if (hasValidId && !isHeader) {
                        // 💡 [비즈니스 룰 1] (1/2) 괄호 패턴 동일 고유번호 병합 처리
                        let col0Clean = col0;
                        const match = cleanPattern.exec(col0);
                        if (match) {
                            col0Clean = match[1].trim();
                        }
                        
                        if (requirementsMap.has(col0Clean)) {
                            const existing = requirementsMap.get(col0Clean);
                            existing.detail_styled.push(...detailParas);
                            
                            const existingName = getPptPlainStructuredText(existing.name_styled).replace(/\n/g, ' ').replace(/\r/g, ' ').trim();
                            const newName = getPptPlainStructuredText(nameParas).replace(/\n/g, ' ').replace(/\r/g, ' ').trim();
                            if (existingName !== newName) {
                                existing.name_styled.push(...nameParas);
                            }
                            
                            const existingDesc = getPptPlainStructuredText(existing.desc_styled).replace(/\n/g, ' ').replace(/\r/g, ' ').trim();
                            const newDesc = getPptPlainStructuredText(descParas).replace(/\n/g, ' ').replace(/\r/g, ' ').trim();
                            if (existingDesc !== newDesc) {
                                existing.desc_styled.push(...descParas);
                            }
                        } else {
                            const reqItem = {
                                id: col0Clean,
                                // 고유번호용 id_styled 는 중복 치환 버그 예방 위해 깨끗한 단일 런 구조로 재정의
                                id_styled: [[{ text: col0Clean, bold: false, color: null }]],
                                name_styled: nameParas,
                                desc_styled: descParas,
                                detail_styled: detailParas
                            };
                            
                            requirementsMap.set(col0Clean, reqItem);
                            requirementsOrder.push(col0Clean);
                        }
                    }
                }
            }
        }
    }

    const requirements = requirementsOrder.map(k => requirementsMap.get(k));
    if (requirements.length === 0) {
        throw new Error('PPTX 파일 내에서 규격에 맞는 요구사항 기술서 표 데이터를 추출할 수 없습니다.');
    }

    const hpNS = 'http://www.hancom.co.kr/hwpml/2011/paragraph';

    // 셀 내용 치환 헬퍼 함수
    function fillHwpxCell(tc, paragraphs, registry, getNextId) {
        const subList = tc.getElementsByTagNameNS(hpNS, 'subList')[0] || tc.getElementsByTagName('hp:subList')[0];
        if (!subList) return;
        
        const existingPs = tc.getElementsByTagNameNS(hpNS, 'p') || tc.getElementsByTagName('hp:p');
        const firstP = existingPs.length > 0 ? existingPs[0] : null;
        
        const paraPrIDRef = firstP ? firstP.getAttribute('paraPrIDRef') : null;
        const styleIDRef = firstP ? firstP.getAttribute('styleIDRef') : null;
        
        let baseCharPrIDRef = '0';
        if (firstP) {
            const firstRun = firstP.getElementsByTagNameNS(hpNS, 'run')[0] || firstP.getElementsByTagName('hp:run')[0];
            if (firstRun) {
                baseCharPrIDRef = firstRun.getAttribute('charPrIDRef') || '0';
            }
        }
        
        while (subList.firstChild) {
            subList.removeChild(subList.firstChild);
        }
        
        paragraphs.forEach(pRuns => {
            const newP = document.createElementNS(hpNS, 'hp:p');
            
            if (firstP) {
                // 기존 단락의 모든 속성(id, paraPrIDRef, styleIDRef, pageBreak, columnBreak, merged 등) 완벽 복사 상속
                for (let i = 0; i < firstP.attributes.length; i++) {
                    const attr = firstP.attributes[i];
                    newP.setAttribute(attr.name, attr.value);
                }
            } else {
                if (paraPrIDRef) newP.setAttribute('paraPrIDRef', paraPrIDRef);
                if (styleIDRef) newP.setAttribute('styleIDRef', styleIDRef);
            }
            
            // 💡 [크래시 해결의 핵심] 줄 나눔 기준이 "어절"일 때 문단의 id 속성이 없거나 중복되면 한글이 폭사하므로 무조건 고유 정수 ID로 덮어씀
            if (getNextId) {
                newP.setAttribute('id', getNextId());
            }
            
            if (pRuns.length === 0) {
                const run = document.createElementNS(hpNS, 'hp:run');
                run.setAttribute('charPrIDRef', baseCharPrIDRef);
                const t = document.createElementNS(hpNS, 'hp:t');
                t.textContent = '';
                run.appendChild(t);
                newP.appendChild(run);
            } else {
                pRuns.forEach(runData => {
                    if (runData.text === '\n') {
                        const run = document.createElementNS(hpNS, 'hp:run');
                        run.setAttribute('charPrIDRef', baseCharPrIDRef);
                        const br = document.createElementNS(hpNS, 'hp:br');
                        run.appendChild(br);
                        newP.appendChild(run);
                    } else {
                        const run = document.createElementNS(hpNS, 'hp:run');
                        run.setAttribute('charPrIDRef', baseCharPrIDRef);
                        const t = document.createElementNS(hpNS, 'hp:t');
                        t.textContent = runData.text;
                        run.appendChild(t);
                        newP.appendChild(run);
                    }
                });
            }
            subList.appendChild(newP);
        });
    }

    function applyRequirementToP(pNode, req, registry, getNextId) {
        const tbl = pNode.getElementsByTagNameNS(hpNS, 'tbl')[0] || pNode.getElementsByTagName('hp:tbl')[0];
        if (!tbl) return;
        
        const tcList = tbl.getElementsByTagNameNS(hpNS, 'tc') || tbl.getElementsByTagName('hp:tc');
        for (let i = 0; i < tcList.length; i++) {
            const tc = tcList[i];
            const tNodes = tc.getElementsByTagNameNS(hpNS, 't') || tc.getElementsByTagName('hp:t');
            let cellText = '';
            for (let j = 0; j < tNodes.length; j++) {
                cellText += tNodes[j].textContent || '';
            }
            cellText = cellText.trim();
            
            if (cellText.includes('{고유번호}') || cellText.includes('{요구사항 고유번호}')) {
                let cleanIdStyled = flattenParagraphsToSingleLine(req.id_styled);
                if (cleanIdStyled.length === 0) {
                    cleanIdStyled = [[{ text: req.id, bold: false, color: null }]];
                }
                fillHwpxCell(tc, cleanIdStyled, registry, getNextId);
            } else if (cellText.includes('{요구사항 명칭}')) {
                const cleanNameStyled = flattenParagraphsToSingleLine(req.name_styled);
                fillHwpxCell(tc, cleanNameStyled, registry, getNextId);
            } else if (cellText.includes('{정의}')) {
                const cleanDescStyled = flattenParagraphsToSingleLine(req.desc_styled);
                fillHwpxCell(tc, cleanDescStyled, registry, getNextId);
            } else if (cellText.includes('{세부내용}')) {
                fillHwpxCell(tc, req.detail_styled, registry, getNextId);
            }
        }
    }

    const serializer = new XMLSerializer();

    if (mergeMode) {
        const hwpxZip = new JSZip();
        await hwpxZip.loadAsync(hwpxBuffer);
        
        const headerXmlStr = await hwpxZip.files['Contents/header.xml'].async('text');
        const headerDoc = parser.parseFromString(headerXmlStr, 'application/xml');
        const registry = new HwpxCharPrRegistry(headerDoc);
        
        const sectionXmlStr = await hwpxZip.files['Contents/section0.xml'].async('text');
        const sectionDoc = parser.parseFromString(sectionXmlStr, 'application/xml');
        
        // 💡 기존 문서에 존재하는 모든 hp:p 단락의 ID 중에서 최댓값을 수집하여 고유 ID 카운터 초기화
        const existingPs = sectionDoc.getElementsByTagNameNS(hpNS, 'p');
        let maxId = 0;
        for (let i = 0; i < existingPs.length; i++) {
            const idVal = parseInt(existingPs[i].getAttribute('id') || '0', 10);
            if (!isNaN(idVal) && idVal > maxId) {
                maxId = idVal;
            }
        }
        let currentId = Math.max(maxId + 1, 3000000000);
        const getNextId = () => String(currentId++);

        const allPs = sectionDoc.getElementsByTagNameNS(hpNS, 'p');
        let templateP = null;
        for (let i = 0; i < allPs.length; i++) {
            const p = allPs[i];
            const tbls = p.getElementsByTagNameNS(hpNS, 'tbl');
            if (tbls.length > 0) {
                templateP = p;
                break;
            }
        }
        
        if (!templateP) {
            throw new Error('HWPX 양식 템플릿 파일에서 요구사항 기술서 표 양식을 찾을 수 없습니다.');
        }
        
        const parentNode = templateP.parentNode;
        const fragment = sectionDoc.createDocumentFragment();
        
        requirements.forEach((req, idx) => {
            const clonedP = templateP.cloneNode(true);
            // 💡 표를 담고 있는 외곽 hp:p 단락에도 고유 ID를 부여하여 유실 방지
            clonedP.setAttribute('id', getNextId());
            
            if (idx > 0) {
                clonedP.setAttribute('pageBreak', '1');
            } else {
                clonedP.setAttribute('pageBreak', '0');
            }
            
            applyRequirementToP(clonedP, req, registry, getNextId);
            fragment.appendChild(clonedP);
            
            if (idx < requirements.length - 1) {
                const spacerP = sectionDoc.createElementNS(hpNS, 'hp:p');
                // 💡 문단 구분용 스페이서 문단에도 고유 ID를 부여
                spacerP.setAttribute('id', getNextId());
                spacerP.setAttribute('paraPrIDRef', '62');
                spacerP.setAttribute('styleIDRef', '0');
                spacerP.setAttribute('pageBreak', '0');
                spacerP.setAttribute('columnBreak', '0');
                spacerP.setAttribute('merged', '0');
                
                const run = sectionDoc.createElementNS(hpNS, 'hp:run');
                run.setAttribute('charPrIDRef', '0');
                const t = sectionDoc.createElementNS(hpNS, 'hp:t');
                t.textContent = '';
                run.appendChild(t);
                spacerP.appendChild(run);
                fragment.appendChild(spacerP);
            }
        });
        
        parentNode.replaceChild(fragment, templateP);
        
        const finalSectionXmlStr = serializer.serializeToString(sectionDoc);
        const finalHeaderXmlStr = serializer.serializeToString(headerDoc);
        
        hwpxZip.file('Contents/section0.xml', finalSectionXmlStr);
        hwpxZip.file('Contents/header.xml', finalHeaderXmlStr);
        
        const finalBuffer = await hwpxZip.generateAsync({ type: 'uint8array' });
        const finalBlob = new Blob([finalBuffer], { type: 'application/x-hwp-hwpx' });
        
        finalBlob.fusedCount = requirements.length;
        finalBlob.requirementsList = requirements;
        finalBlob.mode = 'merge';
        return finalBlob;
    } else {
        const outZip = new JSZip();
        
        for (const req of requirements) {
            const hwpxZip = new JSZip();
            await hwpxZip.loadAsync(hwpxBuffer);
            
            const headerXmlStr = await hwpxZip.files['Contents/header.xml'].async('text');
            const headerDoc = parser.parseFromString(headerXmlStr, 'application/xml');
            const registry = new HwpxCharPrRegistry(headerDoc);
            
            const sectionXmlStr = await hwpxZip.files['Contents/section0.xml'].async('text');
            const sectionDoc = parser.parseFromString(sectionXmlStr, 'application/xml');
            
            // 💡 기존 문서에 존재하는 모든 hp:p 단락의 ID 중에서 최댓값을 수집하여 고유 ID 카운터 초기화
            const existingPs = sectionDoc.getElementsByTagNameNS(hpNS, 'p');
            let maxId = 0;
            for (let i = 0; i < existingPs.length; i++) {
                const idVal = parseInt(existingPs[i].getAttribute('id') || '0', 10);
                if (!isNaN(idVal) && idVal > maxId) {
                    maxId = idVal;
                }
            }
            let currentId = Math.max(maxId + 1, 3000000000);
            const getNextId = () => String(currentId++);

            const allPs = sectionDoc.getElementsByTagNameNS(hpNS, 'p');
            let templateP = null;
            for (let i = 0; i < allPs.length; i++) {
                const p = allPs[i];
                const tbls = p.getElementsByTagNameNS(hpNS, 'tbl');
                if (tbls.length > 0) {
                    templateP = p;
                    break;
                }
            }
            
            if (templateP) {
                const clonedP = templateP.cloneNode(true);
                // 💡 표를 담고 있는 외곽 hp:p 단락에도 고유 ID를 부여하여 유실 방지
                clonedP.setAttribute('id', getNextId());
                clonedP.setAttribute('pageBreak', '0');
                
                applyRequirementToP(clonedP, req, registry, getNextId);
                templateP.parentNode.replaceChild(clonedP, templateP);
            }
            
            const finalSectionXmlStr = serializer.serializeToString(sectionDoc);
            const finalHeaderXmlStr = serializer.serializeToString(headerDoc);
            
            hwpxZip.file('Contents/section0.xml', finalSectionXmlStr);
            hwpxZip.file('Contents/header.xml', finalHeaderXmlStr);
            
            const fileBuffer = await hwpxZip.generateAsync({ type: 'uint8array' });
            
            const safeId = req.id.replace(/[\\/:*?"<>|]/g, '_');
            const safeName = req.name.replace(/[\\/:*?"<>|]/g, '_').substring(0, 20);
            const fileName = `요구사항_${safeId}_${safeName}.hwpx`;
            
            outZip.file(fileName, fileBuffer);
        }
        
        const zipBuffer = await outZip.generateAsync({ type: 'uint8array' });
        const finalBlob = new Blob([zipBuffer], { type: 'application/zip' });
        
        finalBlob.fusedCount = requirements.length;
        finalBlob.requirementsList = requirements;
        finalBlob.mode = 'zip';
        return finalBlob;
    }
}
