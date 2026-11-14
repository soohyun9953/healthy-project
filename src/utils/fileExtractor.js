import * as pdfjsLib from 'pdfjs-dist';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';

// Vite 환경을 고려한 PDF 워커 파일 내부 임포트 방식
import pdfWorker from 'pdfjs-dist/build/pdf.worker.mjs?url';
pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

/** 파일 확장자 가져오기 */
export function getFileExtension(filename) {
    return (filename || '').split('.').pop().toLowerCase();
}

/** 지원 파일 확장자 매핑 */
export const SUPPORTED_EXTENSIONS = {
    text: ['txt', 'md', 'markdown', 'csv', 'json', 'html', 'xml'],
    pdf: ['pdf'],
    excel: ['xlsx', 'xls'],
    docx: ['docx'],
    pptx: ['pptx'],
    hwpx: ['hwpx'],
    unsupported: ['hwp', 'ppt', 'doc'],
};

export const ALL_ACCEPT = '.txt,.md,.markdown,.csv,.json,.html,.xml,.pdf,.xlsx,.xls,.doc,.docx,.hwp,.hwpx,.ppt,.pptx';

// PPT 슬라이드 하단에 표시되는 "장번호-슬라이드번호"(예: Ⅱ-15, 3-15) 페이지 라벨 패턴.
// 공공 ISP/ISMP 표준 PPT 양식의 하단 페이지 번호 영역은 좁고 짧은 텍스트이므로,
// 문단 전체가 이 형식과 정확히 일치할 때만 라벨로 인식해 오탐을 방지한다.
const CHAPTER_PAGE_LABEL_REGEX = /^([ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]{1,3}|[0-9]{1,2})\s*[-–—]\s*(\d{1,4})$/;

function extractChapterPageLabel(paragraphs) {
    for (const p of paragraphs) {
        if (p.length <= 10 && CHAPTER_PAGE_LABEL_REGEX.test(p)) {
            return p.replace(/\s+/g, '').replace(/[–—]/g, '-');
        }
    }
    return null;
}

/** 확장자 → 파일 타입 분류 */
export function classifyFile(ext) {
    for (const [type, exts] of Object.entries(SUPPORTED_EXTENSIONS)) {
        if (exts.includes(ext)) return type;
    }
    return 'text'; // 알 수 없으면 텍스트로 시도
}

// ── PDF 텍스트 추출 ──────────────────────
export async function extractTextFromPDF(file) {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    const allPageLines = [];
    const pageMap = {};

    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
        const page = await pdf.getPage(pageNum);
        const textContent = await page.getTextContent();
        const items = textContent.items;
        if (!items || items.length === 0) continue;

        const validItems = items.filter(item =>
            item.str && item.str.trim().length > 0 && item.transform
        );
        if (validItems.length === 0) continue;

        const LINE_THRESHOLD = 3;
        const lines = [];
        let currentLine = [];
        let currentY = null;

        const sorted = [...validItems].sort((a, b) => {
            const yDiff = b.transform[5] - a.transform[5];
            if (Math.abs(yDiff) > LINE_THRESHOLD) return yDiff;
            return a.transform[4] - b.transform[4];
        });

        for (const item of sorted) {
            const y = item.transform[5];
            const x = item.transform[4];
            if (currentY === null || Math.abs(y - currentY) > LINE_THRESHOLD) {
                if (currentLine.length > 0) lines.push(currentLine);
                currentLine = [{ text: item.str, x, width: item.width || 0 }];
                currentY = y;
            } else {
                currentLine.push({ text: item.str, x, width: item.width || 0 });
            }
        }
        if (currentLine.length > 0) lines.push(currentLine);

        const pageLines = [];
        for (const lineItems of lines) {
            lineItems.sort((a, b) => a.x - b.x);
            let lineText = '';
            for (let i = 0; i < lineItems.length; i++) {
                const item = lineItems[i];
                if (i > 0) {
                    const prev = lineItems[i - 1];
                    const gap = item.x - (prev.x + prev.width);
                    if (gap > 2) lineText += ' ';
                }
                lineText += item.text;
            }
            const trimmed = lineText.trim();
            if (trimmed.length > 0) {
                allPageLines.push(trimmed);
                pageLines.push(trimmed);
            }
        }

        if (pageLines.length > 0) {
            pageMap[pageNum] = pageLines;
        }

        if (pageNum < pdf.numPages) allPageLines.push('');
    }

    return {
        text: allPageLines
            .reduce((acc, line) => {
                if (line === '' && acc.length > 0 && acc[acc.length - 1] === '') return acc;
                acc.push(line);
                return acc;
            }, [])
            .join('\n')
            .trim(),
        pages: pdf.numPages,
        pageMap
    };
}

// ── 엑셀 텍스트 추출 ─────────────────────────────────────────
export async function extractTextFromExcel(file) {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const allText = [];
    const pageMap = {};

    let sheetIndex = 1;
    for (const sheetName of workbook.SheetNames) {
        const sheet = workbook.Sheets[sheetName];
        if (!sheet) continue;

        const sheetLines = [];
        if (workbook.SheetNames.length > 1) {
            allText.push(`[시트: ${sheetName}]`);
            sheetLines.push(`[시트: ${sheetName}]`);
        }

        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

        for (const row of rows) {
            const cells = row.map(cell => String(cell ?? '').trim()).filter(c => c.length > 0);
            if (cells.length === 0) continue;

            const line = cells.length === 1 ? cells[0] : cells.join(' | ');
            allText.push(line);
            sheetLines.push(line);
        }

        if (sheetLines.length > 0) {
            pageMap[sheetIndex] = sheetLines;
            sheetIndex++;
        }

        allText.push('');
    }

    return {
        text: allText.join('\n').trim(),
        pages: Math.max(1, sheetIndex - 1),
        pageMap
    };
}

// ── PPTX 보조 유틸: 관계(rels)·슬라이드 번호 필드·장 구분 헤더 ─────────────────
function parsePptxXml(str) {
    return new DOMParser().parseFromString(str, 'text/xml');
}

function descendantsByLocalName(el, name) {
    return Array.from(el.getElementsByTagNameNS('*', name));
}

// "ppt/slides/slide3.xml" + "../slideLayouts/slideLayout2.xml" → "ppt/slideLayouts/slideLayout2.xml"
function resolvePartPath(basePath, target) {
    if (target.startsWith('/')) return target.slice(1);
    const parts = basePath.split('/').slice(0, -1);
    target.split('/').forEach(seg => {
        if (seg === '..') parts.pop();
        else if (seg && seg !== '.') parts.push(seg);
    });
    return parts.join('/');
}

async function readPartRels(zip, partPath) {
    const segs = partPath.split('/');
    const file = segs.pop();
    const relsPath = `${segs.join('/')}/_rels/${file}.rels`;
    const rels = [];
    if (!zip.files[relsPath]) return rels;
    const doc = parsePptxXml(await zip.files[relsPath].async('string'));
    descendantsByLocalName(doc, 'Relationship').forEach(r => {
        if (r.getAttribute('TargetMode') === 'External') return;
        rels.push({
            id: r.getAttribute('Id'),
            type: r.getAttribute('Type') || '',
            target: resolvePartPath(partPath, r.getAttribute('Target') || '')
        });
    });
    return rels;
}

// 프레젠테이션의 실제 슬라이드 순서(presentation.xml sldIdLst). 병합된 파일은 slideN.xml 번호가 표시 순서와 다르다.
async function readPresentationOrder(zip, slideFiles) {
    try {
        if (!zip.files['ppt/presentation.xml']) return { order: slideFiles, firstSlideNum: 1 };
        const doc = parsePptxXml(await zip.files['ppt/presentation.xml'].async('string'));
        const pres = doc.documentElement;
        const firstSlideNum = parseInt(pres.getAttribute('firstSlideNum') || '1', 10) || 1;
        const rels = await readPartRels(zip, 'ppt/presentation.xml');
        const relMap = new Map(rels.map(r => [r.id, r.target]));
        const order = [];
        descendantsByLocalName(doc, 'sldId').forEach(el => {
            const rid = el.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id')
                || el.getAttribute('r:id');
            const target = relMap.get(rid);
            if (target && zip.files[target] && !order.includes(target)) order.push(target);
        });
        // sldIdLst에 없는 고아 슬라이드는 뒤에 붙인다
        slideFiles.forEach(f => { if (!order.includes(f)) order.push(f); });
        return { order, firstSlideNum };
    } catch (e) {
        console.warn('PPTX 슬라이드 순서 분석 실패, 파일명 순서 사용:', e);
        return { order: slideFiles, firstSlideNum: 1 };
    }
}

function shapePlaceholder(sp) {
    const ph = descendantsByLocalName(sp, 'ph')[0];
    if (!ph) return null;
    return { type: ph.getAttribute('type') || 'body', idx: ph.getAttribute('idx') || '' };
}

// 도형 안의 슬라이드 번호 필드(<a:fld type="slidenum">)와 그 앞뒤 고정 텍스트(예: "Ⅲ- ")를 찾는다
function findSlideNumberTemplates(doc) {
    const templates = [];
    descendantsByLocalName(doc, 'sp').forEach(sp => {
        const ph = shapePlaceholder(sp);
        descendantsByLocalName(sp, 'p').forEach(p => {
            if (p.namespaceURI && !p.namespaceURI.includes('drawingml')) return;
            let prefix = '';
            let suffix = '';
            let found = false;
            Array.from(p.childNodes || []).forEach(node => {
                if (node.nodeType !== 1) return;
                if (node.localName === 'fld' && (node.getAttribute('type') || '').toLowerCase() === 'slidenum' && !found) {
                    found = true;
                    return;
                }
                if (node.localName === 'r' || node.localName === 'fld') {
                    const t = descendantsByLocalName(node, 't').map(x => x.textContent || '').join('');
                    if (found) suffix += t; else prefix += t;
                }
            });
            if (found) templates.push({ prefix, suffix, ph });
        });
    });
    return templates;
}

// 장표 구조 정보: 제목 개체 틀, 거버닝 메시지(제목 외 개체 틀 중 가장 긴 문장), 개체 틀 문구(상단 경로 등), 표의 행·열
function extractSlideMeta(doc) {
    const meta = { title: '', governing: '', headings: [], tables: [] };
    const textOf = (el) => descendantsByLocalName(el, 'p')
        .map(p => descendantsByLocalName(p, 't').map(t => t.textContent || '').join('').trim())
        .filter(Boolean)
        .join(' ');
    descendantsByLocalName(doc, 'sp').forEach(sp => {
        const ph = shapePlaceholder(sp);
        if (!ph) return;
        if (['sldNum', 'dt', 'ftr', 'hdr'].includes(ph.type)) return;
        const text = textOf(sp).replace(/\s+/g, ' ').trim();
        if (!text) return;
        if (ph.type === 'title' || ph.type === 'ctrTitle') {
            meta.title = meta.title ? `${meta.title} ${text}` : text;
            return;
        }
        meta.headings.push(text);
        if (text.length >= 15 && text.length > meta.governing.length) meta.governing = text;
    });
    descendantsByLocalName(doc, 'tbl').forEach(tbl => {
        descendantsByLocalName(tbl, 'tr').forEach(tr => {
            const cells = descendantsByLocalName(tr, 'tc').map(tc => textOf(tc).replace(/\s+/g, ' ').trim());
            if (cells.some(Boolean)) meta.tables.push(cells);
        });
    });
    return meta;
}

// 장 구분 헤더(예: 우상단 "Ⅱ. 환경분석")
const CHAPTER_HEADER_REGEX = /^([ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ])\s*[.．]\s*([가-힣A-Za-z][^.。]{1,24})$/;

function nonPlaceholderParagraphs(doc) {
    const out = [];
    descendantsByLocalName(doc, 'sp').forEach(sp => {
        if (shapePlaceholder(sp)) return;
        descendantsByLocalName(sp, 'p').forEach(p => {
            const t = descendantsByLocalName(p, 't').map(x => x.textContent || '').join('').trim();
            if (t) out.push(t);
        });
    });
    return out;
}

function findChapterHeader(paragraphs) {
    const hits = [];
    paragraphs.forEach(p => {
        const m = p.trim().match(CHAPTER_HEADER_REGEX);
        if (m) hits.push({ roman: m[1], title: m[2].trim() });
    });
    // 표지·목차처럼 여러 장이 나열된 슬라이드는 장 헤더로 보지 않는다
    const distinct = new Set(hits.map(h => h.roman));
    return distinct.size === 1 ? hits[0] : null;
}

function buildLabelFromTemplate(template, displayNum) {
    const raw = `${template.prefix}${displayNum}${template.suffix}`.replace(/\s+/g, '').replace(/[–—]/g, '-');
    return CHAPTER_PAGE_LABEL_REGEX.test(raw) ? raw : null;
}

// ── PPTX 텍스트 추출 (DOMParser & 엔티티 디코딩 100% 전수 수집) ─────────────────
export async function extractTextFromPPTX(file) {
    const arrayBuffer = await file.arrayBuffer();
    let zip;
    try {
        zip = await JSZip.loadAsync(arrayBuffer);
    } catch (zipErr) {
        console.error('PPTX JSZip 로드 에러:', zipErr);
        if (zipErr.message.includes('signature') || zipErr.message.includes('corrupted') || zipErr.message.includes('zip') || zipErr.message.includes('expected')) {
            throw new Error('보안 프로그램(DRM)으로 암호화되었거나 손상된 PPTX 파일입니다. 사내 DRM 보안을 해제(사외반출용 등)하여 원본 평문 문서로 저장한 뒤 업로드해 주세요.');
        }
        throw zipErr;
    }
    const textBlocks = [];
    
    const slideRegex = /^ppt\/slides\/slide\d+\.xml$/;
    const slideFiles = Object.keys(zip.files).filter(name => slideRegex.test(name));
    
    slideFiles.sort((a, b) => {
        const numA = parseInt(a.match(/\d+/)[0], 10);
        const numB = parseInt(b.match(/\d+/)[0], 10);
        return numA - numB;
    });

    // 페이지 번호 = 프레젠테이션 표시 순서(파워포인트 화면의 슬라이드 번호)
    const { order: orderedSlides, firstSlideNum } = await readPresentationOrder(zip, slideFiles);

    const decodeXml = (str) => {
        if (!str) return '';
        return str
            .replace(/&amp;/g, '&')
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&quot;/g, '"')
            .replace(/&#39;/g, "'")
            .replace(/&apos;/g, "'");
    };

    const pageMap = {};
    const pageLabelMap = {}; // { [slideNum]: "Ⅱ-15" 같은 슬라이드 하단 표시 페이지 라벨 | null }
    const slideChapterMap = {}; // { [slideNum]: { roman: 'Ⅱ', title: '환경분석' } } 우상단 장 구분 헤더
    const slideMetaMap = {}; // { [slideNum]: { title, governing, headings, tables } } 장표 제목·거버닝·표 구조

    // 레이아웃/마스터는 수십~수백 장이 공유하므로 1회만 해석
    const partCache = new Map();
    const loadLayoutLike = async (partPath) => {
        if (partCache.has(partPath)) return partCache.get(partPath);
        let info = null;
        try {
            if (zip.files[partPath]) {
                const doc = parsePptxXml(await zip.files[partPath].async('string'));
                const rels = await readPartRels(zip, partPath);
                const showMasterSp = doc.documentElement.getAttribute('showMasterSp') !== '0';
                info = {
                    templates: findSlideNumberTemplates(doc),
                    chapter: findChapterHeader(nonPlaceholderParagraphs(doc)),
                    parent: (rels.find(r => /slideMaster$/.test(r.type)) || {}).target || null,
                    showMasterSp
                };
            }
        } catch (e) {
            console.warn(`PPTX '${partPath}' 해석 실패:`, e);
        }
        partCache.set(partPath, info);
        return info;
    };

    for (let pos = 0; pos < orderedSlides.length; pos++) {
        const fileName = orderedSlides[pos];
        const slideNum = pos + 1;
        const displayNum = firstSlideNum + pos;
        const content = await zip.files[fileName].async('string');
        const slideParagraphs = [];
        let slideDoc = null;

        try {
            slideDoc = parsePptxXml(content);
            const paragraphs = slideDoc.getElementsByTagNameNS('*', 'p');

            for (let i = 0; i < paragraphs.length; i++) {
                const p = paragraphs[i];
                const tElements = p.getElementsByTagNameNS('*', 't');
                let pText = '';
                for (let j = 0; j < tElements.length; j++) {
                    pText += tElements[j].textContent || '';
                }
                const trimmed = decodeXml(pText).trim();
                if (trimmed.length > 0) {
                    slideParagraphs.push(trimmed);
                }
            }
        } catch (_) {
            slideDoc = null;
            // DOMParser 실패 시 Regex 폴백
            const paragraphRegex = /<a:p[^>]*>([\s\S]*?)<\/a:p>/g;
            let pMatch;
            while ((pMatch = paragraphRegex.exec(content)) !== null) {
                const pContent = pMatch[1];
                const textRegex = /<a:t[^>]*>([\s\S]*?)<\/a:t>/g;
                let tMatch;
                const textRuns = [];
                while ((tMatch = textRegex.exec(pContent)) !== null) {
                    textRuns.push(decodeXml(tMatch[1].replace(/<[^>]+>/g, '')));
                }
                const joined = textRuns.join('').trim();
                if (joined.length > 0) slideParagraphs.push(joined);
            }
        }

        // 하단 페이지 라벨: ① 슬라이드의 번호 필드 → ② 레이아웃 → ③ 마스터("Ⅲ- <#>") → ④ 슬라이드에 직접 입력된 "Ⅱ- 4"
        let label = null;
        let chapter = null;
        if (slideDoc) {
            try {
                const slideTemplates = findSlideNumberTemplates(slideDoc);
                const slideHasSldNumPh = descendantsByLocalName(slideDoc, 'ph').some(ph => ph.getAttribute('type') === 'sldNum');
                const slideRels = await readPartRels(zip, fileName);
                const layoutPath = (slideRels.find(r => /slideLayout$/.test(r.type)) || {}).target;
                const layout = layoutPath ? await loadLayoutLike(layoutPath) : null;
                const master = layout && layout.parent ? await loadLayoutLike(layout.parent) : null;

                const pick = (templates, allowPh) => {
                    for (const t of templates) {
                        if (t.ph && !allowPh) continue;
                        const l = buildLabelFromTemplate(t, displayNum);
                        if (l) return l;
                    }
                    return null;
                };
                label = pick(slideTemplates, true);
                if (!label && layout) label = pick(layout.templates, false) || (slideHasSldNumPh ? pick(layout.templates.filter(t => t.ph?.type === 'sldNum'), true) : null);
                if (!label && master && layout.showMasterSp) label = pick(master.templates, false);
                if (!label && master && slideHasSldNumPh) label = pick(master.templates.filter(t => t.ph?.type === 'sldNum'), true);

                chapter = findChapterHeader(slideParagraphs)
                    || (layout && layout.chapter)
                    || (master && layout.showMasterSp ? master.chapter : null);
            } catch (e) {
                console.warn(`PPTX 슬라이드 ${slideNum} 페이지 라벨 분석 실패:`, e);
            }
        }
        if (!label) {
            const staticLabel = extractChapterPageLabel(slideParagraphs);
            // 숫자 접두 라벨("3-1")은 표 안 과제번호와 구별되지 않으므로 실제 슬라이드 번호와 일치할 때만 인정
            if (staticLabel && (/^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]/.test(staticLabel) || staticLabel.endsWith(`-${displayNum}`))) {
                label = staticLabel;
            }
        }
        if (!chapter && !slideDoc) chapter = findChapterHeader(slideParagraphs);

        if (chapter) slideChapterMap[slideNum] = chapter;
        if (slideDoc) {
            try {
                slideMetaMap[slideNum] = extractSlideMeta(slideDoc);
            } catch (e) {
                console.warn(`PPTX 슬라이드 ${slideNum} 구조 분석 실패:`, e);
            }
        }
        if (slideParagraphs.length > 0) {
            pageMap[slideNum] = slideParagraphs;
            textBlocks.push(`[슬라이드 ${slideNum}]\n` + slideParagraphs.join('\n'));
        }
        if (label) pageLabelMap[slideNum] = label;
    }

    // 하단 번호가 없는 일부 슬라이드(간지·부록 레이아웃 등): 파일 대부분이 "Ⅳ-슬라이드번호"로 일관되면 같은 장 번호로 보완
    const labeledEntries = Object.entries(pageLabelMap);
    if (labeledEntries.length > 0 && labeledEntries.length < orderedSlides.length) {
        const prefixCounts = new Map();
        let consistent = 0;
        labeledEntries.forEach(([n, label]) => {
            const [prefix, num] = label.split('-');
            prefixCounts.set(prefix, (prefixCounts.get(prefix) || 0) + 1);
            if (parseInt(num, 10) === firstSlideNum + parseInt(n, 10) - 1) consistent++;
        });
        const [topPrefix, topCount] = Array.from(prefixCounts.entries()).sort((a, b) => b[1] - a[1])[0];
        if (prefixCounts.size === 1 && topCount >= 2 && consistent === labeledEntries.length) {
            for (let pos = 1; pos <= orderedSlides.length; pos++) {
                if (!pageLabelMap[pos]) pageLabelMap[pos] = `${topPrefix}-${firstSlideNum + pos - 1}`;
            }
        }
    }

    // 하단 번호 필드가 전혀 없는 파일: 장 구분 헤더(Ⅱ. 환경분석)가 일관되면 "Ⅱ-슬라이드번호"로 보완
    if (Object.keys(pageLabelMap).length === 0) {
        const counts = new Map();
        Object.values(slideChapterMap).forEach(c => counts.set(c.roman, (counts.get(c.roman) || 0) + 1));
        const total = orderedSlides.length;
        counts.forEach((cnt, roman) => {
            if (cnt >= Math.max(2, total * 0.5)) {
                Object.entries(slideChapterMap).forEach(([n, c]) => {
                    if (c.roman === roman) pageLabelMap[n] = `${roman}-${firstSlideNum + parseInt(n, 10) - 1}`;
                });
            }
        });
    }

    return {
        text: textBlocks.join('\n\n').trim(),
        pages: orderedSlides.length,
        pageMap,
        pageLabelMap,
        slideChapterMap,
        slideMetaMap
    };
}

// ── HWPX 텍스트 추출 ─────────────────────────────────────────
export async function extractTextFromHWPX(file) {
    const arrayBuffer = await file.arrayBuffer();
    let zip;
    try {
        zip = await JSZip.loadAsync(arrayBuffer);
    } catch (zipErr) {
        console.error('HWPX JSZip 로드 에러:', zipErr);
        if (zipErr.message.includes('signature') || zipErr.message.includes('corrupted') || zipErr.message.includes('zip') || zipErr.message.includes('expected')) {
            throw new Error('보안 프로그램(DRM)으로 암호화되었거나 손상된 HWPX 파일입니다. 사내 DRM 보안을 해제(사외반출용 등)하여 원본 평문 문서로 저장한 뒤 업로드해 주세요.');
        }
        throw zipErr;
    }
    const textBlocks = [];
    const pageMap = {};
    
    const sectionRegex = /^Contents\/section\d+\.xml$/;
    const sectionFiles = Object.keys(zip.files).filter(name => sectionRegex.test(name));
    
    sectionFiles.sort((a, b) => {
        const numA = parseInt(a.match(/\d+/)[0], 10);
        const numB = parseInt(b.match(/\d+/)[0], 10);
        return numA - numB;
    });

    let currentPage = 1;
    let prevVertpos = -1;

    for (let sIdx = 0; sIdx < sectionFiles.length; sIdx++) {
        const fileName = sectionFiles[sIdx];
        const content = await zip.files[fileName].async('string');
        const paragraphRegex = /<hp:p[^>]*>([\s\S]*?)<\/hp:p>/g;
        let pMatch;
        const sectionParagraphs = [];

        while ((pMatch = paragraphRegex.exec(content)) !== null) {
            const pContent = pMatch[1];
            if (pContent.includes('<hp:pageBreak') || pContent.includes('pageBreak')) {
                currentPage++;
            }
            const textRegex = /<hp:t.*?>(.*?)<\/hp:t>/g;
            let tMatch;
            const textRuns = [];
            while ((tMatch = textRegex.exec(pContent)) !== null) {
                let text = tMatch[1].replace(/<[^>]+>/g, '');
                text = text.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
                textRuns.push(text);
            }
            const joined = textRuns.join('').trim();
            if (joined.length > 0) {
                sectionParagraphs.push(joined);
                if (!pageMap[currentPage]) pageMap[currentPage] = [];
                pageMap[currentPage].push(joined);
            }
        }

        if (sectionParagraphs.length > 0) {
            textBlocks.push(sectionParagraphs.join('\n'));
        }
    }
    
    if (Object.keys(pageMap).length === 0 && textBlocks.length > 0) {
        pageMap[1] = textBlocks;
    }

    return {
        text: textBlocks.join('\n\n').trim(),
        pages: Math.max(1, currentPage),
        pageMap
    };
}

// ── Word (DOCX) 텍스트 추출 ──────────────────────────────────
export async function extractTextFromDOCX(file) {
    const arrayBuffer = await file.arrayBuffer();
    let zip;
    try {
        zip = await JSZip.loadAsync(arrayBuffer);
    } catch (zipErr) {
        console.error('DOCX JSZip 로드 에러:', zipErr);
        if (zipErr.message.includes('signature') || zipErr.message.includes('corrupted') || zipErr.message.includes('zip') || zipErr.message.includes('expected')) {
            throw new Error('보안 프로그램(DRM)으로 암호화되었거나 손상된 Word(.docx) 파일입니다.');
        }
        throw zipErr;
    }

    const docFile = zip.file('word/document.xml');
    if (!docFile) {
        throw new Error('Word(.docx) 문서의 본문 구조(document.xml)를 찾을 수 없습니다.');
    }

    const content = await docFile.async('string');
    const paragraphs = [];
    const pageMap = {};
    let currentPage = 1;

    try {
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(content, 'text/xml');
        const pElements = xmlDoc.getElementsByTagNameNS('*', 'p');

        for (let i = 0; i < pElements.length; i++) {
            const p = pElements[i];
            
            // 페이지 나눔 태그 감지
            const brs = p.getElementsByTagNameNS('*', 'br');
            for (let b = 0; b < brs.length; b++) {
                const type = brs[b].getAttribute('w:type') || brs[b].getAttribute('type');
                if (type === 'page') currentPage++;
            }
            if (p.getElementsByTagNameNS('*', 'lastRenderedPageBreak').length > 0) {
                currentPage++;
            }

            const tElements = p.getElementsByTagNameNS('*', 't');
            let pText = '';
            for (let j = 0; j < tElements.length; j++) {
                pText += tElements[j].textContent || '';
            }
            const trimmed = pText.trim();
            if (trimmed.length > 0) {
                paragraphs.push(trimmed);
                if (!pageMap[currentPage]) pageMap[currentPage] = [];
                pageMap[currentPage].push(trimmed);
            }
        }
    } catch (_) {
        // DOMParser 폴백
        const pRegex = /<w:p[^>]*>([\s\S]*?)<\/w:p>/g;
        let pMatch;
        while ((pMatch = pRegex.exec(content)) !== null) {
            const pContent = pMatch[1];
            if (pContent.includes('type="page"') || pContent.includes('lastRenderedPageBreak')) {
                currentPage++;
            }
            const tRegex = /<w:t[^>]*>([\s\S]*?)<\/w:t>/g;
            let tMatch;
            const tRuns = [];
            while ((tMatch = tRegex.exec(pContent)) !== null) {
                tRuns.push(tMatch[1].replace(/<[^>]+>/g, ''));
            }
            const joined = tRuns.join('').trim();
            if (joined.length > 0) {
                paragraphs.push(joined);
                if (!pageMap[currentPage]) pageMap[currentPage] = [];
                pageMap[currentPage].push(joined);
            }
        }
    }

    if (Object.keys(pageMap).length === 0 && paragraphs.length > 0) {
        pageMap[1] = paragraphs;
    }

    return {
        text: paragraphs.join('\n\n').trim(),
        pages: Math.max(1, currentPage),
        pageMap
    };
}

// ── Markdown / Text 추출 (H1/H2 또는 구분선 기준 페이지 분할) ─────────
export async function extractTextFromMD(file) {
    const text = await file.text();
    const lines = text.split(/\r?\n/);
    const pageMap = {};
    let currentPage = 1;
    let currentParagraphs = [];

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        if (line.length === 0) continue;

        // 대제목(#) 또는 구분선(---)을 만나면 다음 논리적 페이지/섹션으로 간주
        if ((line.startsWith('# ') || line === '---' || line === '***') && currentParagraphs.length > 0) {
            pageMap[currentPage] = currentParagraphs;
            currentPage++;
            currentParagraphs = [];
            if (line !== '---' && line !== '***') {
                currentParagraphs.push(line);
            }
            continue;
        }

        currentParagraphs.push(line);
    }

    if (currentParagraphs.length > 0) {
        pageMap[currentPage] = currentParagraphs;
    }

    if (Object.keys(pageMap).length === 0 && text.trim().length > 0) {
        pageMap[1] = [text.trim()];
    }

    return {
        text: text.trim(),
        pages: Math.max(1, Object.keys(pageMap).length),
        pageMap
    };
}

/**
 * 모든 지원 파일 형식에서 구조화된 { text, pages, pageMap }을 추출하는 범용 함수
 */
export async function extractStructuredFileContent(file) {
    const ext = getFileExtension(file.name);
    let res;

    if (ext === 'pptx') {
        res = await extractTextFromPPTX(file);
    } else if (ext === 'hwpx') {
        res = await extractTextFromHWPX(file);
    } else if (ext === 'docx') {
        res = await extractTextFromDOCX(file);
    } else if (ext === 'pdf') {
        res = await extractTextFromPDF(file);
    } else if (ext === 'xlsx' || ext === 'xls') {
        res = await extractTextFromExcel(file);
    } else if (ext === 'md' || ext === 'markdown' || ext === 'txt' || ext === 'csv' || ext === 'json' || ext === 'html' || ext === 'xml') {
        res = await extractTextFromMD(file);
    } else {
        const text = await file.text();
        res = {
            text: text.trim(),
            pages: 1,
            pageMap: { 1: text.split('\n').map(l => l.trim()).filter(l => l.length > 0) }
        };
    }

    const pageMap = res.pageMap || {};
    if (Object.keys(pageMap).length === 0 && res.text) {
        pageMap[1] = res.text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
    }

    return {
        text: res.text || '',
        pages: res.pages || Object.keys(pageMap).length || 1,
        pageMap,
        pageLabelMap: res.pageLabelMap || {},
        slideChapterMap: res.slideChapterMap || {},
        slideMetaMap: res.slideMetaMap || {},
        fileName: file.name,
        fileSize: file.size,
        ext
    };
}

/** 통합 파일 처리 함수 */
export async function processFile(file) {
    const ext = getFileExtension(file.name);
    const type = classifyFile(ext);

    switch (type) {
        case 'pdf': {
            const result = await extractTextFromPDF(file);
            if (!result.text || result.text.trim().length === 0) {
                throw new Error('PDF에서 텍스트를 추출하지 못했습니다. 이미지 기반 PDF일 수 있습니다.');
            }
            return result;
        }
        case 'docx': {
            const result = await extractTextFromDOCX(file);
            if (!result.text || result.text.trim().length === 0) {
                throw new Error('Word(.docx) 파일에서 텍스트를 추출하지 못했습니다.');
            }
            return result;
        }
        case 'excel': {
            const result = await extractTextFromExcel(file);
            if (!result.text || result.text.trim().length === 0) {
                throw new Error('엑셀 파일에서 데이터를 추출하지 못했습니다. 파일이 비어있을 수 있습니다.');
            }
            return result;
        }
        case 'pptx': {
            const result = await extractTextFromPPTX(file);
            if (!result.text || result.text.trim().length === 0) {
                throw new Error('PPTX 파일에서 텍스트를 추출하지 못했습니다.');
            }
            return result;
        }
        case 'hwpx': {
            const result = await extractTextFromHWPX(file);
            if (!result.text || result.text.trim().length === 0) {
                throw new Error('HWPX 파일에서 텍스트를 추출하지 못했습니다.');
            }
            return result;
        }
        case 'unsupported':
            throw new Error(
                `${ext.toUpperCase()} 구형 바이너리 파일은 브라우저 공간에서 직접 읽을 수 없습니다. 가능한 최신 포맷인 HWPX(.hwpx), PPTX(.pptx), DOCX(.docx), XLSX(.xlsx) 또는 PDF로 변환 후 업로드해 주세요.`
            );
        case 'text':
        default: {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = (ev) => resolve({ text: ev.target.result, pages: 1 });
                reader.onerror = () => reject(new Error('파일 읽기에 실패했습니다.'));
                reader.readAsText(file, 'UTF-8');
            });
        }
    }
}
