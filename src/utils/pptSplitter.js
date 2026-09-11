import JSZip from 'jszip';
import { saveAs } from 'file-saver';

/**
 * PPTX 파일의 슬라이드 구조 및 목차(대주제, 장, 간지) 자동 분석 함수
 * @param {File|ArrayBuffer} fileInput 
 * @returns {Promise<{ totalSlides: number, sections: Array<{ id: number, title: string, startSlide: number, endSlide: number, slideCount: number, previewText: string }>, slidesInfo: Array<{ slideNum: number, title: string, fullText: string, isMajor: boolean }> }>}
 */
export async function analyzePptxSections(fileInput) {
    const arrayBuffer = fileInput instanceof ArrayBuffer ? fileInput : await fileInput.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);
    const parser = new DOMParser();

    // 1. 슬라이드 파일 목록 추출 및 순서 정렬
    const slideEntries = Object.keys(zip.files).filter(path => 
        path.startsWith('ppt/slides/slide') && path.endsWith('.xml') && !path.includes('_rels')
    );

    slideEntries.sort((a, b) => {
        const numA = parseInt(a.replace(/[^0-9]/g, ''), 10) || 0;
        const numB = parseInt(b.replace(/[^0-9]/g, ''), 10) || 0;
        return numA - numB;
    });

    const totalSlides = slideEntries.length;
    if (totalSlides === 0) {
        throw new Error('PPTX 파일 내에 슬라이드가 존재하지 않습니다.');
    }

    const slidesInfo = [];

    // 대주제/장/간지 패턴 정규식
    const romanMajorRegex = /^\s*(?:제\s*([0-9IVXLCDM]+)\s*장|([IVXLCDM]+)\s*[\.\:\-\)]|\b([0-9]+)\s*장\b)/i;
    const numMajorRegex = /^\s*([0-9]+)\s*[\.\:\-\)]\s*([^\n]+)/;

    for (let i = 0; i < slideEntries.length; i++) {
        const slidePath = slideEntries[i];
        const slideNum = i + 1;
        const slideXmlStr = await zip.file(slidePath).async('text');
        const xmlDoc = parser.parseFromString(slideXmlStr, 'application/xml');

        // 텍스트 추출 (p 단위)
        const pNodes = xmlDoc.getElementsByTagName('a:p');
        const paragraphs = [];
        let slideTitle = '';
        let isTitleFound = false;

        // Title Placeholder 찾기
        const spNodes = xmlDoc.getElementsByTagName('p:sp');
        for (let j = 0; j < spNodes.length; j++) {
            const sp = spNodes[j];
            const phNode = sp.getElementsByTagName('p:ph')[0];
            const phType = phNode ? phNode.getAttribute('type') : '';
            
            const tList = [];
            const tNodes = sp.getElementsByTagName('a:t');
            for (let k = 0; k < tNodes.length; k++) {
                tList.push(tNodes[k].textContent || '');
            }
            const spText = tList.join(' ').trim();

            if (spText) {
                if ((phType === 'title' || phType === 'ctrTitle') && !isTitleFound) {
                    slideTitle = spText;
                    isTitleFound = true;
                }
            }
        }

        // 전체 슬라이드 텍스트 수집
        for (let pIdx = 0; pIdx < pNodes.length; pIdx++) {
            const pEl = pNodes[pIdx];
            const tEls = pEl.getElementsByTagName('a:t');
            let pStr = '';
            for (let tIdx = 0; tIdx < tEls.length; tIdx++) {
                pStr += (tEls[tIdx].textContent || '');
            }
            if (pStr.trim()) {
                paragraphs.push(pStr.trim());
                if (!slideTitle && pStr.trim().length > 1) {
                    slideTitle = pStr.trim();
                }
            }
        }

        const fullText = paragraphs.join('\n');
        const cleanedTitle = (slideTitle || `슬라이드 ${slideNum}`).replace(/[\\/:*?"<>|]/g, '').trim();

        // 대주제 여부 판별 (로마자/장 번호/간지 형태)
        const isMajor = romanMajorRegex.test(cleanedTitle) || 
                       (slideNum === 1) || 
                       (paragraphs.length <= 3 && cleanedTitle.length < 30 && (numMajorRegex.test(cleanedTitle) || cleanedTitle.includes('개요') || cleanedTitle.includes('목표') || cleanedTitle.includes('설계') || cleanedTitle.includes('분석') || cleanedTitle.includes('전략')));

        slidesInfo.push({
            slideNum,
            title: cleanedTitle || `슬라이드 ${slideNum}`,
            fullText,
            previewText: paragraphs.slice(0, 3).join(' / ') || '텍스트 없음',
            isMajor
        });
    }

    // 2. 목차/섹션 군집화 (자동 분할 구역 생성)
    const rawSections = [];
    let currentSection = null;

    for (let i = 0; i < slidesInfo.length; i++) {
        const slide = slidesInfo[i];
        
        // 새로운 섹션 시작 조건: 첫 슬라이드 또는 대주제로 감지된 슬라이드
        if (i === 0 || (slide.isMajor && i > 0)) {
            if (currentSection) {
                currentSection.endSlide = slide.slideNum - 1;
                currentSection.slideCount = currentSection.endSlide - currentSection.startSlide + 1;
                rawSections.push(currentSection);
            }
            currentSection = {
                id: rawSections.length + 1,
                title: slide.title,
                startSlide: slide.slideNum,
                endSlide: totalSlides,
                slideCount: 1,
                previewText: slide.previewText,
                selected: true
            };
        }
    }

    if (currentSection) {
        currentSection.endSlide = totalSlides;
        currentSection.slideCount = currentSection.endSlide - currentSection.startSlide + 1;
        rawSections.push(currentSection);
    }

    // 섹션이 1개만 도출되었거나 감지가 부족한 경우: 10장 단위 또는 기본 분할 보강
    let finalSections = rawSections;
    if (finalSections.length <= 1 && totalSlides > 15) {
        finalSections = [];
        const chunkSize = Math.ceil(totalSlides / 4);
        for (let s = 1; s <= totalSlides; s += chunkSize) {
            const e = Math.min(s + chunkSize - 1, totalSlides);
            const leaderSlide = slidesInfo[s - 1];
            finalSections.push({
                id: finalSections.length + 1,
                title: leaderSlide.title || `${s}장~${e}장`,
                startSlide: s,
                endSlide: e,
                slideCount: e - s + 1,
                previewText: leaderSlide.previewText,
                selected: true
            });
        }
    }

    return {
        totalSlides,
        sections: finalSections,
        slidesInfo
    };
}

/**
 * 원본 PPTX에서 특정 슬라이드 범위만 추출하여 독립된 PPTX 파일(Blob)을 생성하는 함수
 * @param {ArrayBuffer} originalBuffer 
 * @param {number} startSlide 1-indexed
 * @param {number} endSlide 1-indexed
 * @returns {Promise<Blob>}
 */
export async function createSubPptx(originalBuffer, startSlide, endSlide) {
    const zip = await JSZip.loadAsync(originalBuffer);
    const parser = new DOMParser();
    const serializer = new XMLSerializer();

    // 1. 원본 슬라이드 목록 파악
    const slideEntries = Object.keys(zip.files).filter(path => 
        path.startsWith('ppt/slides/slide') && path.endsWith('.xml') && !path.includes('_rels')
    );
    slideEntries.sort((a, b) => {
        const numA = parseInt(a.replace(/[^0-9]/g, ''), 10) || 0;
        const numB = parseInt(b.replace(/[^0-9]/g, ''), 10) || 0;
        return numA - numB;
    });

    const keepSlideIndices = [];
    for (let i = startSlide; i <= endSlide; i++) {
        if (i >= 1 && i <= slideEntries.length) {
            keepSlideIndices.push(i);
        }
    }

    if (keepSlideIndices.length === 0) {
        throw new Error(`유효하지 않은 슬라이드 범위입니다: ${startSlide} ~ ${endSlide}`);
    }

    // 2. presentation.xml 및 presentation.xml.rels 읽기
    const presXmlStr = await zip.file('ppt/presentation.xml').async('text');
    const presDoc = parser.parseFromString(presXmlStr, 'application/xml');

    const presRelsStr = await zip.file('ppt/_rels/presentation.xml.rels').async('text');
    const presRelsDoc = parser.parseFromString(presRelsStr, 'application/xml');

    // 슬라이드 ID 리스트 엘리먼트 (p:sldIdLst)
    const sldIdLst = presDoc.getElementsByTagName('p:sldIdLst')[0] || presDoc.getElementsByTagName('sldIdLst')[0];
    const sldIdNodes = sldIdLst ? Array.from(sldIdLst.getElementsByTagName('p:sldId') || sldIdLst.getElementsByTagName('sldId')) : [];

    // presentation.xml.rels의 모든 Relationship 파악
    const relsElements = Array.from(presRelsDoc.getElementsByTagName('Relationship'));

    // 원본 슬라이드 번호 -> relId 매핑
    const slideNumToRelMap = {};
    sldIdNodes.forEach((node, idx) => {
        const slideNum = idx + 1;
        const rId = node.getAttribute('r:id') || node.getAttribute('id');
        slideNumToRelMap[slideNum] = { rId, sldIdNode: node };
    });

    // 3. 유지할 슬라이드 파일들만 새로운 순번(1..K)으로 재구성
    const newZip = new JSZip();

    // 3-1. 원본 파일 중 슬라이드 관련 파일을 제외하고 기본 뼈대 복사
    const originalFileKeys = Object.keys(zip.files);
    for (const key of originalFileKeys) {
        if (key.startsWith('ppt/slides/slide') || key.startsWith('ppt/slides/_rels/slide')) {
            continue; // 슬라이드 본문 및 관계는 아래에서 별도 추가
        }
        if (!zip.files[key].dir) {
            const fileData = await zip.file(key).async('uint8array');
            newZip.file(key, fileData);
        }
    }

    // 3-2. 선택된 슬라이드들을 1부터 K까지 번호로 리매핑하여 새 ZIP에 저장
    const newSldIdLstNodes = [];
    const newPresRelsElements = [];

    // 기본 presentation.xml.rels에서 슬라이드가 아닌 관계들(마스터, 테마 등) 복사
    relsElements.forEach(rel => {
        const target = rel.getAttribute('Target') || '';
        const type = rel.getAttribute('Type') || '';
        if (!type.includes('/slide') || target.includes('slideMaster') || target.includes('slideLayout')) {
            newPresRelsElements.push(rel);
        }
    });

    for (let newIdx = 0; newIdx < keepSlideIndices.length; newIdx++) {
        const origSlideNum = keepSlideIndices[newIdx];
        const newSlideNum = newIdx + 1;

        const origSlidePath = `ppt/slides/slide${origSlideNum}.xml`;
        const origRelsPath = `ppt/slides/_rels/slide${origSlideNum}.xml.rels`;
        const newSlidePath = `ppt/slides/slide${newSlideNum}.xml`;
        const newRelsPath = `ppt/slides/_rels/slide${newSlideNum}.xml.rels`;

        // 슬라이드 본문 XML 복사
        if (zip.files[origSlidePath]) {
            const slideContent = await zip.file(origSlidePath).async('uint8array');
            newZip.file(newSlidePath, slideContent);
        }

        // 슬라이드 관계 XML 복사
        if (zip.files[origRelsPath]) {
            const relsContent = await zip.file(origRelsPath).async('uint8array');
            newZip.file(newRelsPath, relsContent);
        }

        // 새로운 presentation.xml.rels 항목 구성
        const newRId = `rId_split_sld_${newSlideNum}`;
        const newRelEl = presRelsDoc.createElement('Relationship');
        newRelEl.setAttribute('Id', newRId);
        newRelEl.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide');
        newRelEl.setAttribute('Target', `slides/slide${newSlideNum}.xml`);
        newPresRelsElements.push(newRelEl);

        // 새로운 presentation.xml sldId 노드 구성
        const origInfo = slideNumToRelMap[origSlideNum];
        const origIdVal = origInfo?.sldIdNode?.getAttribute('id') || `${255 + newSlideNum}`;
        const newSldIdEl = presDoc.createElement('p:sldId');
        newSldIdEl.setAttribute('id', origIdVal);
        newSldIdEl.setAttribute('r:id', newRId);
        newSldIdLstNodes.push(newSldIdEl);
    }

    // 4. presentation.xml 갱신
    if (sldIdLst) {
        while (sldIdLst.firstChild) {
            sldIdLst.removeChild(sldIdLst.firstChild);
        }
        newSldIdLstNodes.forEach(node => sldIdLst.appendChild(node));
    }
    const updatedPresXmlStr = serializer.serializeToString(presDoc);
    newZip.file('ppt/presentation.xml', updatedPresXmlStr);

    // 5. presentation.xml.rels 갱신
    const relationshipsRoot = presRelsDoc.getElementsByTagName('Relationships')[0];
    if (relationshipsRoot) {
        while (relationshipsRoot.firstChild) {
            relationshipsRoot.removeChild(relationshipsRoot.firstChild);
        }
        newPresRelsElements.forEach(node => relationshipsRoot.appendChild(node));
    }
    const updatedPresRelsStr = serializer.serializeToString(presRelsDoc);
    newZip.file('ppt/_rels/presentation.xml.rels', updatedPresRelsStr);

    // 6. [Content_Types].xml 갱신 (슬라이드 목록 조정)
    if (newZip.files['[Content_Types].xml']) {
        const ctXmlStr = await newZip.file('[Content_Types].xml').async('text');
        const ctDoc = parser.parseFromString(ctXmlStr, 'application/xml');
        const typesRoot = ctDoc.getElementsByTagName('Types')[0];
        if (typesRoot) {
            const overrideNodes = Array.from(typesRoot.getElementsByTagName('Override'));
            // 기존 slide{N}.xml Override 제거
            overrideNodes.forEach(node => {
                const partName = node.getAttribute('PartName') || '';
                if (partName.startsWith('/ppt/slides/slide') && partName.endsWith('.xml')) {
                    typesRoot.removeChild(node);
                }
            });

            // 새로운 slide{1..K}.xml Override 추가
            for (let k = 1; k <= keepSlideIndices.length; k++) {
                const newOverride = ctDoc.createElement('Override');
                newOverride.setAttribute('PartName', `/ppt/slides/slide${k}.xml`);
                newOverride.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
                typesRoot.appendChild(newOverride);
            }

            const updatedCtXmlStr = serializer.serializeToString(ctDoc);
            newZip.file('[Content_Types].xml', updatedCtXmlStr);
        }
    }

    // 7. 새로운 독립 PPTX 파일 Blob 생성
    const outputBlob = await newZip.generateAsync({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    });

    return outputBlob;
}

/**
 * 지정된 섹션들을 각각 분할하여 ZIP 파일로 압축 다운로드
 * @param {ArrayBuffer} originalBuffer 
 * @param {string} originalFileName 
 * @param {Array<{ id: number, title: string, startSlide: number, endSlide: number }>} sections 
 * @param {function(string, number): void} onProgress 
 * @param {string} fileNamePrefix
 */
export async function downloadAllSectionsAsZip(originalBuffer, originalFileName, sections, onProgress, fileNamePrefix = '[분할]') {
    const zipArchive = new JSZip();
    const baseName = originalFileName.replace(/\.[^.]+$/, '');
    const prefix = fileNamePrefix ? `${fileNamePrefix.trim()}_` : '';

    for (let i = 0; i < sections.length; i++) {
        const sec = sections[i];
        if (onProgress) {
            onProgress(`[${i + 1}/${sections.length}] "${sec.title}" 섹션 PPTX 생성 중...`, Math.round(((i + 1) / sections.length) * 100));
        }

        const subBlob = await createSubPptx(originalBuffer, sec.startSlide, sec.endSlide);
        const subArrayBuffer = await subBlob.arrayBuffer();

        // 파일명 포맷: [접두사]_[01]_섹션명_원본파일명.pptx
        const cleanTitle = (sec.title || `섹션_${sec.id}`).replace(/[\\/:*?"<>|]/g, '_').trim();
        const subFileName = `${prefix}[${String(i + 1).padStart(2, '0')}]_${cleanTitle}_${baseName}.pptx`;
        zipArchive.file(subFileName, subArrayBuffer);
    }

    if (onProgress) onProgress('ZIP 압축 파일 생성 중...', 99);
    const zipBlob = await zipArchive.generateAsync({ type: 'blob' });
    const zipFileName = `${prefix}[목차분할]_${baseName}.zip`;
    saveAs(zipBlob, zipFileName);
    if (onProgress) onProgress('다운로드 완료!', 100);
}
