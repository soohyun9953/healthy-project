import JSZip from 'jszip';
import { saveAs } from 'file-saver';

// 로마자 숫자 매핑 테이블 (I..XV, Ⅰ..Ⅻ)
const ROMAN_MAP = {
    'I': 1, 'II': 2, 'III': 3, 'IV': 4, 'V': 5,
    'VI': 6, 'VII': 7, 'VIII': 8, 'IX': 9, 'X': 10,
    'XI': 11, 'XII': 12, 'XIII': 13, 'XIV': 14, 'XV': 15,
    'Ⅰ': 1, 'Ⅱ': 2, 'Ⅲ': 3, 'Ⅳ': 4, 'Ⅴ': 5,
    'Ⅵ': 6, 'Ⅶ': 7, 'Ⅷ': 8, 'Ⅸ': 9, 'Ⅹ': 10,
    'Ⅺ': 11, 'Ⅻ': 12
};

/**
 * 슬라이드 목록의 앞부분(또는 전체)에서 표준 목차(Canonical Chapters)를 추출
 * @param {Array<{ slideNum: number, paragraphs: string[], isKanji: boolean }>} slidesData 
 * @returns {Array<{ num: number, title: string, keywords: string[], type: string }>}
 */
function extractCanonicalChapters(slidesData) {
    const chapters = [];
    const maxScan = Math.min(12, slidesData.length);

    // 1. 앞쪽 1~12 슬라이드에서 목차/CHAPTER 간지 슬라이드 분석
    for (let sIdx = 0; sIdx < maxScan; sIdx++) {
        const ps = slidesData[sIdx].paragraphs;

        // 형태 A: "1. 목표모델 수립 개요", "2. 정보화 비전 및 전략 수립", "3. 개선과제 상세화", "4. 목표모델 설계"
        for (const p of ps) {
            const m = p.match(/^\s*0*([1-9][0-9]?)\.\s+([^\d\.\-\s].*)$/);
            if (m) {
                const cNum = parseInt(m[1], 10);
                const rawTitle = m[2].trim();
                if (rawTitle.length >= 2 && rawTitle.length <= 40 && !rawTitle.includes('http')) {
                    if (!chapters.some(c => c.num === cNum)) {
                        chapters.push({
                            num: cNum,
                            title: `${cNum}. ${rawTitle}`,
                            keywords: [rawTitle, rawTitle.replace(/\s+/g, '')],
                            type: 'numeric'
                        });
                    }
                }
            }

            // 형태 B: 로마자 "I. 사업 개요", "Ⅱ. 현황 분석", "Ⅲ. 목표 모델 수립"
            const romM = p.match(/^\s*(?:제\s*)?([IVXLCDMⅠ-Ⅹ]+)[\.\:\-\s]\s*(.+)$/i);
            if (romM) {
                const romStr = romM[1].toUpperCase();
                if (ROMAN_MAP[romStr]) {
                    const cNum = ROMAN_MAP[romStr];
                    const rawTitle = romM[2].trim();
                    if (rawTitle.length >= 2 && rawTitle.length <= 40) {
                        if (!chapters.some(c => c.num === cNum)) {
                            chapters.push({
                                num: cNum,
                                title: `${romStr}. ${rawTitle}`,
                                keywords: [rawTitle, rawTitle.replace(/\s+/g, '')],
                                type: 'roman'
                            });
                        }
                    }
                }
            }

            // 형태 C: "제1장 ...", "제2장 ..."
            const korM = p.match(/^\s*제\s*0*([1-9][0-9]?)\s*장\s*[\.\:\-\s]*(.*)$/i);
            if (korM) {
                const cNum = parseInt(korM[1], 10);
                const rawTitle = korM[2].trim();
                if (rawTitle.length <= 40) {
                    if (!chapters.some(c => c.num === cNum)) {
                        chapters.push({
                            num: cNum,
                            title: `제${cNum}장 ${rawTitle}`.trim(),
                            keywords: [rawTitle, rawTitle.replace(/\s+/g, '')],
                            type: 'korean_chapter'
                        });
                    }
                }
            }
        }

        // 형태 D: 번호 토큰('1', '2', '3', '4')과 제목이 연속된 p 노드로 나열된 경우 (예: CHAPTER 간지)
        for (let i = 0; i < ps.length; i++) {
            if (/^[1-9]$/.test(ps[i])) {
                const numVal = parseInt(ps[i], 10);
                for (let j = i + 1; j < Math.min(i + 12, ps.length); j++) {
                    const candidate = ps[j].trim();
                    if (candidate && !/^\d+$/.test(candidate) && candidate.length <= 35 && candidate.length >= 2) {
                        if (!['CHAPTER', 'C·H·A·P·T·E·R', 'ISP', 'CONTENTS', 'INDEX', '차례', '목차'].some(k => candidate.includes(k))) {
                            if (!chapters.some(c => c.num === numVal)) {
                                chapters.push({
                                    num: numVal,
                                    title: `${numVal}. ${candidate}`,
                                    keywords: [candidate, candidate.replace(/\s+/g, '')],
                                    type: 'numeric'
                                });
                            }
                            break;
                        }
                    }
                }
            }
        }
    }

    // 2. 만약 앞쪽에서 목차가 도출되지 않았을 경우, 전체 슬라이드의 브레드크럼/제목에서 대목차 스캔
    if (chapters.length <= 1) {
        for (let sIdx = 0; sIdx < slidesData.length; sIdx++) {
            const ps = slidesData[sIdx].paragraphs;
            for (const p of ps.slice(0, 5)) {
                const m = p.match(/^\s*0*([1-9][0-9]?)\.\s+([^\d\.\-\s].*)$/);
                if (m) {
                    const cNum = parseInt(m[1], 10);
                    const rawTitle = m[2].trim();
                    if (rawTitle.length >= 2 && rawTitle.length <= 35) {
                        if (!chapters.some(c => c.num === cNum)) {
                            chapters.push({
                                num: cNum,
                                title: `${cNum}. ${rawTitle}`,
                                keywords: [rawTitle, rawTitle.replace(/\s+/g, '')],
                                type: 'numeric'
                            });
                        }
                    }
                }
            }
        }
    }

    chapters.sort((a, b) => a.num - b.num);
    return chapters;
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

    // 1. 슬라이드 파일 목록 추출 및 정렬
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

    const slidesData = [];

    // 2. 전체 슬라이드 텍스트 수집 및 간지 슬라이드 여부 파악
    for (let i = 0; i < slideEntries.length; i++) {
        const slidePath = slideEntries[i];
        const slideNum = i + 1;
        const slideXmlStr = await zip.file(slidePath).async('text');
        const xmlDoc = parser.parseFromString(slideXmlStr, 'application/xml');

        const pNodes = xmlDoc.getElementsByTagName('a:p');
        const paragraphs = [];

        for (let pIdx = 0; pIdx < pNodes.length; pIdx++) {
            const pEl = pNodes[pIdx];
            const tEls = pEl.getElementsByTagName('a:t');
            let pStr = '';
            for (let tIdx = 0; tIdx < tEls.length; tIdx++) {
                pStr += (tEls[tIdx].textContent || '');
            }
            if (pStr.trim()) {
                paragraphs.push(pStr.trim());
            }
        }

        const isKanji = paragraphs.slice(0, 4).some(p => 
            p.includes('CHAPTER') || p.includes('C·H·A·P·T·E·R') || p.includes('목차') || p.includes('CONTENTS')
        );

        slidesData.push({
            slideNum,
            paragraphs,
            previewText: paragraphs.slice(0, 3).join(' / ') || '텍스트 없음',
            isKanji
        });
    }

    // 3. 목차(Canonical Chapters) 도출
    const canonicalChapters = extractCanonicalChapters(slidesData);

    // 4. 슬라이드별 대목차/중목차 판정 (단조 증가 규칙 적용: 본문 내 과거 목차 인용에 의한 역주행 방지)
    const classifiedSlides = [];
    let activeMajorNum = null;
    let activeMajorTitle = null;
    let activeMidNum = null;
    let activeMidTitle = null;

    const getSlideChapter = (slide, curMajor) => {
        const ps = slide.paragraphs;

        // 1순위: 슬라이드 상단 번호 체계 (예: "4.3.6.4.", "3.2.1.", "2.1.", "1.1.")
        for (const p of ps.slice(0, 4)) {
            const numM = p.match(/^\s*0*([1-9][0-9]?)\.([0-9]+)/);
            if (numM) {
                const major = parseInt(numM[1], 10);
                const mid = parseInt(numM[2], 10);
                // 단조 증가 허용: 현재 챕터 이상이거나 최초 진입일 때
                if (curMajor === null || major >= curMajor) {
                    const matchedC = canonicalChapters.find(c => c.num === major);
                    const title = matchedC ? matchedC.title : `${major}. 목차 ${major}`;
                    const midTitle = `${major}.${mid} 목차`;
                    return { major, title, mid, midTitle };
                }
            }
        }

        // 2순위: 명시적 대목차 브레드크럼 (예: "4. 목표모델 설계", "3. 개선과제 상세화")
        for (const p of ps) {
            for (const c of canonicalChapters) {
                const hasKeyword = c.keywords.some(k => p.includes(k));
                const isHeaderFormat = p.startsWith(`${c.num}.`) || p.startsWith(`제${c.num}장`) || p.startsWith(`[${c.num}]`) || p.length <= 35;
                if (hasKeyword && isHeaderFormat) {
                    if (curMajor === null || c.num >= curMajor) {
                        return { major: c.num, title: c.title, mid: null, midTitle: c.title };
                    }
                }
            }
        }

        // 3순위: 한국어/로마자 단독 넘버링
        for (const p of ps.slice(0, 3)) {
            const korM = p.match(/^\s*제\s*0*([1-9][0-9]?)\s*장\s*[\.\:\-\s]*(.*)$/);
            if (korM) {
                const num = parseInt(korM[1], 10);
                if (curMajor === null || num >= curMajor) {
                    const title = `제${num}장 ${korM[2] || ''}`.trim();
                    return { major: num, title, mid: null, midTitle: title };
                }
            }
        }

        return null;
    };

    for (let i = 0; i < slidesData.length; i++) {
        const s = slidesData[i];
        const res = getSlideChapter(s, activeMajorNum);
        if (res) {
            activeMajorNum = res.major;
            activeMajorTitle = res.title;
            if (res.mid !== null) {
                activeMidNum = res.mid;
                activeMidTitle = res.midTitle;
            }
        }
        classifiedSlides.push({
            slideNum: s.slideNum,
            majorNum: activeMajorNum,
            majorTitle: activeMajorTitle,
            midNum: activeMidNum,
            midTitle: activeMidTitle,
            previewText: s.previewText,
            isKanji: s.isKanji
        });
    }

    // 5. 간지(Kanji/Chapter Cover) 슬라이드 보정: 간지 슬라이드는 바로 뒤따라오는 챕터에 편입
    for (let i = 0; i < classifiedSlides.length - 1; i++) {
        const s = classifiedSlides[i];
        const nextS = classifiedSlides[i + 1];
        if (s.isKanji && nextS.majorNum !== null) {
            s.majorNum = nextS.majorNum;
            s.majorTitle = nextS.majorTitle;
            s.midNum = nextS.midNum;
            s.midTitle = nextS.midTitle;
        }
    }

    // 5-1. 표지와 목차는 무조건 첫 번째 목차(1장)에 같이 편입 (1번 슬라이드부터 시작)
    const firstChSlide = classifiedSlides.find(s => s.majorNum !== null);
    const defaultMajorNum = firstChSlide?.majorNum || canonicalChapters[0]?.num || 1;
    const defaultMajorTitle = firstChSlide?.majorTitle || canonicalChapters[0]?.title || `1. 목표모델 수립 개요`;
    const defaultMidNum = firstChSlide?.midNum || 1;
    const defaultMidTitle = firstChSlide?.midTitle || defaultMajorTitle;

    for (let i = 0; i < classifiedSlides.length; i++) {
        const s = classifiedSlides[i];
        if (s.majorNum === null) {
            s.majorNum = defaultMajorNum;
            s.majorTitle = defaultMajorTitle;
            s.midNum = defaultMidNum;
            s.midTitle = defaultMidTitle;
        } else {
            break; // 첫 챕터 도달 시 루프 종료
        }
    }

    // 6. 목차별 구간(Section) 군집화 (1번 슬라이드부터 시작)
    const finalSections = [];
    let curSec = null;

    for (let i = 0; i < classifiedSlides.length; i++) {
        const s = classifiedSlides[i];
        const currentKey = splitLevel === 'mid' 
            ? `${s.majorNum}_${s.midNum || 0}`
            : `${s.majorNum || 'UNKNOWN'}`;

        const currentTitle = splitLevel === 'mid'
            ? (s.midTitle || s.majorTitle || `슬라이드 ${s.slideNum}`)
            : (s.majorTitle || `목차 ${s.majorNum || finalSections.length + 1}`);

        if (!curSec || curSec.key !== currentKey) {
            if (curSec) {
                curSec.endSlide = s.slideNum - 1;
                curSec.slideCount = curSec.endSlide - curSec.startSlide + 1;
                finalSections.push(curSec);
            }

            curSec = {
                id: finalSections.length + 1,
                key: currentKey,
                title: currentTitle,
                startSlide: s.slideNum,
                endSlide: totalSlides,
                slideCount: 1,
                previewText: s.previewText,
                selected: true
            };
        } else {
            curSec.endSlide = s.slideNum;
            curSec.slideCount = curSec.endSlide - curSec.startSlide + 1;
        }
    }

    if (curSec) {
        curSec.endSlide = totalSlides;
        curSec.slideCount = curSec.endSlide - curSec.startSlide + 1;
        finalSections.push(curSec);
    }

    // fallback: 섹션이 1개 이하인 경우 제목 변화 기반 안전 분할
    if (finalSections.length <= 1) {
        finalSections.length = 0;
        let fbSec = null;
        for (let i = 0; i < slidesData.length; i++) {
            const s = slidesData[i];
            const isDivider = (i === 0) || s.isKanji;
            if (isDivider || !fbSec) {
                if (fbSec) {
                    fbSec.endSlide = s.slideNum - 1;
                    fbSec.slideCount = fbSec.endSlide - fbSec.startSlide + 1;
                    finalSections.push(fbSec);
                }
                fbSec = {
                    id: finalSections.length + 1,
                    title: s.paragraphs[0] || `섹션 ${finalSections.length + 1}`,
                    startSlide: s.slideNum,
                    endSlide: totalSlides,
                    slideCount: 1,
                    previewText: s.previewText,
                    selected: true
                };
            }
        }
        if (fbSec) {
            fbSec.endSlide = totalSlides;
            fbSec.slideCount = fbSec.endSlide - fbSec.startSlide + 1;
            finalSections.push(fbSec);
        }
    }

    // 순번 재색인
    finalSections.forEach((sec, idx) => {
        sec.id = idx + 1;
    });

    return {
        totalSlides,
        sections: finalSections,
        slidesInfo: classifiedSlides,
        tocItems: canonicalChapters
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
