import JSZip from 'jszip';
import { saveAs } from 'file-saver';

// 로마자 숫자 매핑 테이블 (I..X, Ⅰ..Ⅹ)
const ROMAN_MAP = {
    'I': 1, 'II': 2, 'III': 3, 'IV': 4, 'V': 5,
    'VI': 6, 'VII': 7, 'VIII': 8, 'IX': 9, 'X': 10,
    'XI': 11, 'XII': 12, 'XIII': 13, 'XIV': 14, 'XV': 15,
    'Ⅰ': 1, 'Ⅱ': 2, 'Ⅲ': 3, 'Ⅳ': 4, 'Ⅴ': 5,
    'Ⅵ': 6, 'Ⅶ': 7, 'Ⅷ': 8, 'Ⅸ': 9, 'Ⅹ': 10,
    'Ⅺ': 11, 'Ⅻ': 12
};

/**
 * 텍스트에서 대목차(1, 2, 3 / I, II, III / 제1장 등) 및 중목차(1.1, 1.2 등) 정보를 정밀 추출
 * @param {string} text 
 * @returns {object|null}
 */
export function parseChapterInfo(text) {
    if (!text || typeof text !== 'string') return null;
    const clean = text.trim();
    if (!clean) return null;

    // 1. 로마자 대목차 패턴 (예: I. 사업 개요, Ⅰ. 현황 분석, 제I장, I-1. 추진방향)
    const romanMatch = clean.match(/^\s*(?:제\s*)?([IVXLCDMⅠ-Ⅹ]+)(?:\s*장|\s*부|\s*편|[\.\:\-\]\)\s])\s*(.*)$/i);
    if (romanMatch) {
        const romStr = romanMatch[1].toUpperCase();
        if (ROMAN_MAP[romStr]) {
            const num = ROMAN_MAP[romStr];
            const rest = (romanMatch[2] || '').trim();
            // 로마자 하위 (예: I-1. 세부사항)
            const subMatch = rest.match(/^[0-9]+[\.\:\-\)]\s*(.*)$/);
            const subRest = subMatch ? subMatch[1] : rest;
            return {
                type: 'roman',
                majorNum: num,
                majorKey: `R_${num}`,
                majorLabel: `${romStr}. ${rest || '섹션'}`,
                midNum: subMatch ? parseInt(rest, 10) : null,
                midKey: subMatch ? `R_${num}_${parseInt(rest, 10)}` : `R_${num}`,
                midLabel: rest ? `${romStr}. ${rest}` : `${romStr}`,
                rawTitle: clean
            };
        }
    }

    // 2. 한국어 장/부/Part 패턴 (예: 제1장 사업 개요, 제 2 부 현황, PART 1. 착수, 1장)
    const koreanMatch = clean.match(/^\s*(?:제\s*0*([1-9][0-9]?)\s*[장부편]|(?:PART|CHAPTER|SECTION)\s*0*([1-9][0-9]?)|0*([1-9][0-9]?)\s*[장부편])[\.\:\-\]\)\s]*(.*)$/i);
    if (koreanMatch) {
        const num = parseInt(koreanMatch[1] || koreanMatch[2] || koreanMatch[3], 10);
        const rest = (koreanMatch[4] || '').trim();
        return {
            type: 'korean_chapter',
            majorNum: num,
            majorKey: `K_${num}`,
            majorLabel: `제${num}장 ${rest}`.trim(),
            midNum: null,
            midKey: `K_${num}`,
            midLabel: `제${num}장 ${rest}`.trim(),
            rawTitle: clean
        };
    }

    // 3. 숫자 계층 목차 패턴 (예: 1.1.2 세부기능, 1.1 추진배경, 1. 사업개요, [1] 개요, 01. 개요, 1-1.)
    const numMatch = clean.match(/^\s*\[?\s*0*([1-9][0-9]?)(?:[\.\-]([0-9]+)(?:[\.\-]([0-9]+))?)?\s*[\.\:\-\]\)\s]\s*(.*)$/);
    if (numMatch) {
        const major = parseInt(numMatch[1], 10);
        const mid = numMatch[2] ? parseInt(numMatch[2], 10) : null;
        const sub = numMatch[3] ? parseInt(numMatch[3], 10) : null;
        const rest = (numMatch[4] || '').trim();

        // 1.1.x 형태는 대목차 1, 중목차 1.1로 맵핑
        const majorLabel = mid !== null 
            ? `${major}. 목차 ${major}` 
            : `${major}. ${rest || `대목차 ${major}`}`;

        const midLabel = mid !== null
            ? `${major}.${mid} ${rest || `중목차 ${major}.${mid}`}`
            : majorLabel;

        return {
            type: 'numeric',
            majorNum: major,
            majorKey: `N_${major}`,
            majorLabel: majorLabel,
            midNum: mid,
            midKey: mid !== null ? `N_${major}_${mid}` : `N_${major}`,
            midLabel: midLabel,
            subNum: sub,
            rawTitle: clean
        };
    }

    // 4. 단독 목차 키워드 (예: "1. 사업개요", "2. 현황분석" 등의 단순 넘버링)
    const simpleNumMatch = clean.match(/^0*([1-9][0-9]?)\s+([^\n]+)$/);
    if (simpleNumMatch && simpleNumMatch[2].length <= 30) {
        const major = parseInt(simpleNumMatch[1], 10);
        const rest = simpleNumMatch[2].trim();
        return {
            type: 'numeric_simple',
            majorNum: major,
            majorKey: `N_${major}`,
            majorLabel: `${major}. ${rest}`,
            midNum: null,
            midKey: `N_${major}`,
            midLabel: `${major}. ${rest}`,
            rawTitle: clean
        };
    }

    return null;
}

/**
 * PPTX 파일의 슬라이드 구조 및 목차(대주제, 장, 간지) 자동 분석 함수
 * @param {File|ArrayBuffer} fileInput 
 * @param {string} splitLevel 'major' (대목차: 1, 2, 3) | 'mid' (중목차: 1.1, 1.2)
 * @returns {Promise<{ totalSlides: number, sections: Array<{ id: number, title: string, startSlide: number, endSlide: number, slideCount: number, previewText: string, selected: boolean }>, slidesInfo: Array<object>, tocItems: Array<object> }>}
 */
export async function analyzePptxSections(fileInput, splitLevel = 'major') {
    const arrayBuffer = fileInput instanceof ArrayBuffer ? fileInput : await fileInput.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);
    const parser = new DOMParser();

    // 1. 슬라이드 파일 목록 추출 및 번호 순서 정렬
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
    const tocItems = []; // 목차 슬라이드에서 발견된 전체 항목들

    // 2. 전체 슬라이드 텍스트 및 제목 파싱
    for (let i = 0; i < slideEntries.length; i++) {
        const slidePath = slideEntries[i];
        const slideNum = i + 1;
        const slideXmlStr = await zip.file(slidePath).async('text');
        const xmlDoc = parser.parseFromString(slideXmlStr, 'application/xml');

        // 슬라이드 내 모든 p 노드 텍스트 수집
        const pNodes = xmlDoc.getElementsByTagName('a:p');
        const paragraphs = [];
        let explicitTitle = '';
        let isTitleFound = false;

        // Title Placeholder 찾기 (p:ph type="title" or "ctrTitle")
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
                    explicitTitle = spText;
                    isTitleFound = true;
                }
            }
        }

        // 전체 단락 텍스트 조합
        for (let pIdx = 0; pIdx < pNodes.length; pIdx++) {
            const pEl = pNodes[pIdx];
            const tEls = pEl.getElementsByTagName('a:t');
            let pStr = '';
            for (let tIdx = 0; tIdx < tEls.length; tIdx++) {
                pStr += (tEls[tIdx].textContent || '');
            }
            if (pStr.trim()) {
                paragraphs.push(pStr.trim());
                if (!explicitTitle && pStr.trim().length > 1 && pIdx === 0) {
                    explicitTitle = pStr.trim();
                }
            }
        }

        const fullText = paragraphs.join('\n');
        const slideTitle = (explicitTitle || `슬라이드 ${slideNum}`).replace(/[\\/:*?"<>|]/g, '').trim();

        // 목차 정보 분석 시도 (타이틀 우선 -> 첫 3개 단락 순회)
        let chapterInfo = parseChapterInfo(slideTitle);
        if (!chapterInfo) {
            for (const p of paragraphs.slice(0, 4)) {
                chapterInfo = parseChapterInfo(p);
                if (chapterInfo) break;
            }
        }

        // 목차(TOC) 슬라이드 여부 감지
        const isTocSlide = (slideTitle.includes('목차') || slideTitle.includes('CONTENTS') || slideTitle.includes('INDEX') || slideTitle.includes('차례')) && slideNum <= 5;
        if (isTocSlide) {
            // 목차 슬라이드의 각 단락에서 목차 항목 추출
            paragraphs.forEach(p => {
                const info = parseChapterInfo(p);
                if (info) tocItems.push(info);
            });
        }

        slidesInfo.push({
            slideNum,
            title: slideTitle,
            fullText,
            previewText: paragraphs.slice(0, 3).join(' / ') || '텍스트 없음',
            chapterInfo,
            isTocSlide
        });
    }

    // 3. 목차(대목차/중목차) 기준 클러스터링 및 분할 구역 생성
    const sections = [];
    let currentSection = null;
    let activeMajorKey = null;
    let activeMidKey = null;

    // 첫 슬라이드가 표지/목차인 경우 첫 번째 대목차가 나타나기 전까지의 범위 식별
    const firstChapterSlide = slidesInfo.find(s => s.chapterInfo && (splitLevel === 'mid' ? s.chapterInfo.midKey : s.chapterInfo.majorKey));
    const firstChapterSlideNum = firstChapterSlide ? firstChapterSlide.slideNum : 1;

    // 1번 슬라이드부터 첫 대목차 전까지 표지/목차 그룹 생성 (첫 대목차가 2페이지 이상 뒤에 있을 때)
    if (firstChapterSlideNum > 1) {
        sections.push({
            id: 1,
            title: '표지 및 목차',
            startSlide: 1,
            endSlide: firstChapterSlideNum - 1,
            slideCount: firstChapterSlideNum - 1,
            previewText: slidesInfo[0]?.previewText || '표지 및 목차 슬라이드',
            selected: true,
            isIntro: true
        });
    }

    for (let i = 0; i < slidesInfo.length; i++) {
        const slide = slidesInfo[i];
        if (slide.slideNum < firstChapterSlideNum) continue;

        const info = slide.chapterInfo;
        const targetKey = splitLevel === 'mid' 
            ? (info?.midKey || activeMidKey) 
            : (info?.majorKey || activeMajorKey);

        const targetLabel = splitLevel === 'mid'
            ? (info?.midLabel || slide.title)
            : (info?.majorLabel || slide.title);

        const isNewMajor = info && (splitLevel === 'mid' ? (info.midKey !== activeMidKey) : (info.majorKey !== activeMajorKey));

        if (isNewMajor || (!currentSection && slide.slideNum >= firstChapterSlideNum)) {
            // 기존 열려있는 섹션 마감
            if (currentSection) {
                currentSection.endSlide = slide.slideNum - 1;
                currentSection.slideCount = currentSection.endSlide - currentSection.startSlide + 1;
                sections.push(currentSection);
            }

            activeMajorKey = info?.majorKey || `UNKNOWN_${slide.slideNum}`;
            activeMidKey = info?.midKey || `UNKNOWN_${slide.slideNum}`;

            // 새로운 섹션 생성
            currentSection = {
                id: sections.length + 1,
                title: targetLabel || `목차 ${sections.length + 1}`,
                startSlide: slide.slideNum,
                endSlide: totalSlides,
                slideCount: 1,
                previewText: slide.previewText,
                selected: true,
                chapterInfo: info
            };
        }
    }

    if (currentSection) {
        currentSection.endSlide = totalSlides;
        currentSection.slideCount = currentSection.endSlide - currentSection.startSlide + 1;
        sections.push(currentSection);
    }

    // 만약 문서 전체에서 명시적 목차 번호(1., 1.1, I. 등)가 하나도 발견되지 않은 경우에만
    // 슬라이드 제목 변화 기반 또는 간지 슬라이드로 분할 (균등 분할 대신 제목 기반)
    if (sections.length <= 1) {
        sections.length = 0; // 초기화
        let fallbackSec = null;
        for (let i = 0; i < slidesInfo.length; i++) {
            const slide = slidesInfo[i];
            const isTitleSlide = (i === 0) || (slide.title && slide.fullText.length < 80);
            if (isTitleSlide || !fallbackSec) {
                if (fallbackSec) {
                    fallbackSec.endSlide = slide.slideNum - 1;
                    fallbackSec.slideCount = fallbackSec.endSlide - fallbackSec.startSlide + 1;
                    sections.push(fallbackSec);
                }
                fallbackSec = {
                    id: sections.length + 1,
                    title: slide.title || `섹션 ${sections.length + 1}`,
                    startSlide: slide.slideNum,
                    endSlide: totalSlides,
                    slideCount: 1,
                    previewText: slide.previewText,
                    selected: true
                };
            }
        }
        if (fallbackSec) {
            fallbackSec.endSlide = totalSlides;
            fallbackSec.slideCount = fallbackSec.endSlide - fallbackSec.startSlide + 1;
            sections.push(fallbackSec);
        }
    }

    // 섹션 ID 재부여
    sections.forEach((sec, idx) => {
        sec.id = idx + 1;
    });

    return {
        totalSlides,
        sections,
        slidesInfo,
        tocItems
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
            continue;
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
            overrideNodes.forEach(node => {
                const partName = node.getAttribute('PartName') || '';
                if (partName.startsWith('/ppt/slides/slide') && partName.endsWith('.xml')) {
                    typesRoot.removeChild(node);
                }
            });

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
