/**
 * PPT 목차 생성기
 *
 * 1) 최종 산출물(PPTX)에서 장·절 목차와 시작 페이지(하단 표기 Ⅱ-61 등)를 추출하고
 * 2) 사용자가 올린 목차 샘플 PPTX의 목차 슬라이드 디자인(행 반복 구조)을 그대로 복제해 채운다.
 *
 * 샘플 인식 방식
 * - 토큰 방식: 샘플 문구에 {{장번호}} {{장제목}} {{번호}} {{제목}} {{페이지}} 가 있으면 그 위치를 사용
 * - 자동 방식: 같은 x 위치·크기로 반복되는 제목 칸을 '행'으로, 같은 줄의 숫자 칸을 번호, 오른쪽 페이지 칸,
 *   행 밖의 로마숫자(Ⅱ)·장 제목 칸을 장 슬롯으로 인식 ('내용' 같은 빈 칸도 행으로 인식)
 * - 항목이 행보다 많으면 마지막 행을 복제(간격 유지, 넘치면 균등 압축), 적으면 남는 행 삭제
 */
import JSZip from 'jszip';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const EMU_PER_INCH = 914400;

const ROMAN_TO_NUM = { 'Ⅰ': 1, 'Ⅱ': 2, 'Ⅲ': 3, 'Ⅳ': 4, 'Ⅴ': 5, 'Ⅵ': 6, 'Ⅶ': 7, 'Ⅷ': 8, 'Ⅸ': 9, 'Ⅹ': 10 };
export const NUM_TO_ROMAN = Object.fromEntries(Object.entries(ROMAN_TO_NUM).map(([k, v]) => [v, k]));

const NUM_RE = /^\s*(0?\d{1,2})\s*[.)]?\s*$/;
const ROMAN_RE = /^\s*([ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]|I{1,3}|IV|VI{0,3}|IX|X)\s*[.．]?\s*$/;
const PAGE_RE = /^\s*([ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]\s*[-–]\s*\d{1,4}(\s*[~∼]\s*\d{1,4})?|p\.?\s*\d{1,4}|\d{1,4}\s*(쪽|p)|·{2,}\s*\d{1,4})\s*$/i;
const TOC_MARKER_RE = /^\s*(목\s*차|CONTENTS?|INDEX|TABLE OF CONTENTS)\s*$/i;
const CHAPTER_WORD_RE = /^\s*CHAPTER\s*0?(\d{1,2})?\s*$/i;
const TOKEN_RE = /\{\{\s*(장번호|장제목|번호|제목|페이지)\s*\}\}/;
const COMPOUND_RE = /^\s*\d{1,2}(\.\d{1,2})+\s*\.?\s*$/;
const COUNT_RE = /^\s*\d+\s*개\s*절(\s*[·∙・,]\s*\d+\s*개\s*항)?\s*$/;
const BIG_ROMAN_RE = /^\s*[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]\s*$/;
const CONT_SUFFIX = ' (계속)';

// ─────────────────────────────────────────────────────────────────────────────
// XML 유틸
// ─────────────────────────────────────────────────────────────────────────────
const parseXml = (str) => new DOMParser().parseFromString(str, 'text/xml');
const serializeXml = (doc) => new XMLSerializer().serializeToString(doc);
const elementChildren = (el) => Array.from(el.childNodes || []).filter(n => n.nodeType === 1);
const firstChild = (el, name) => elementChildren(el).find(n => n.localName === name) || null;
const descendants = (el, name) => Array.from(el.getElementsByTagNameNS('*', name));

function partDir(path) { const i = path.lastIndexOf('/'); return i >= 0 ? path.slice(0, i) : ''; }
function resolvePath(fromDir, target) {
    if (target.startsWith('/')) return target.slice(1);
    const parts = fromDir ? fromDir.split('/') : [];
    target.split('/').forEach(seg => { if (seg === '..') parts.pop(); else if (seg && seg !== '.') parts.push(seg); });
    return parts.join('/');
}
const relsPathOf = (partPath) => `${partDir(partPath)}/_rels/${partPath.split('/').pop()}.rels`;

async function readRels(zip, partPath) {
    const f = zip.file(relsPathOf(partPath));
    if (!f) return [];
    const doc = parseXml(await f.async('string'));
    return descendants(doc, 'Relationship').map(r => ({
        id: r.getAttribute('Id'),
        type: r.getAttribute('Type') || '',
        target: r.getAttribute('Target') || '',
        external: r.getAttribute('TargetMode') === 'External',
        full: r.getAttribute('TargetMode') === 'External' ? null : resolvePath(partDir(partPath), r.getAttribute('Target') || '')
    }));
}

async function presentationSlides(zip) {
    const pres = parseXml(await zip.file('ppt/presentation.xml').async('string'));
    const rels = await readRels(zip, 'ppt/presentation.xml');
    const byId = new Map(rels.map(r => [r.id, r.full]));
    const out = [];
    descendants(pres, 'sldId').forEach(el => {
        const rid = el.getAttributeNS(R_NS, 'id') || el.getAttribute('r:id');
        const path = byId.get(rid);
        if (path && zip.file(path)) out.push(path);
    });
    return out;
}

// ─────────────────────────────────────────────────────────────────────────────
// 1) 최종 산출물 → 목차 모델
// ─────────────────────────────────────────────────────────────────────────────
/**
 * @param {Array<File>} files 최종 산출물 PPTX
 * @param {Function} buildDeliverableTocIndex ssotValidator의 절 색인 함수 (순환 의존 방지를 위해 주입)
 * @returns {Promise<Array<{roman,num,title,fileName,startLabel,include,sections:[{num,title,startLabel,pageRange,include}]}>>}
 */
const normTitle = (t) => String(t || '').replace(/[\s·ㆍ・.,()-]/g, '').toLowerCase();

export async function extractTocFromDeliverables(files, buildDeliverableTocIndex, formatPageLabelRanges) {
    const index = await buildDeliverableTocIndex(files.map(f => ({ file: f, fileName: f.name })));
    const chapters = index.map((doc, i) => {
        const labels = doc.content?.pageLabelMap || {};
        const firstPos = Object.keys(doc.content?.pageMap || {}).map(Number).sort((a, b) => a - b)[0];
        const chapterNum = doc.chapter || null;
        const sections = (doc.sections || [])
            .filter(s => s.title)
            .map(s => ({
                num: s.num,
                title: s.title,
                startLabel: s.labels?.[0] || (s.pages?.[0] != null ? String(s.pages[0]) : ''),
                pageRange: s.pages?.length ? formatPageLabelRanges(s.labels, s.pages) : '',
                include: true,
                // 항 (1.1, 1.2 …) - 샘플 목차가 그 깊이까지 있을 때 사용
                subsections: (s.subsections || []).filter(sub => sub.title).map(sub => ({
                    num: sub.num,
                    title: sub.title,
                    startLabel: sub.labels?.[0] || (sub.pages?.[0] != null ? String(sub.pages[0]) : ''),
                    include: true
                }))
            }));
        // 목차 슬라이드 번호 목차로 보완
        // - 본문에서 절을 하나도 못 찾았으면 목차 슬라이드 항목 전체
        // - 찾았으면 인식된 번호 범위 안의 빈 번호만 (예: 1장짜리 '1. 개요'). 범위 밖 번호는 본문에 없는 목차로 보고 제외
        const recognized = sections.map(s => s.num);
        const maxNum = recognized.length ? Math.max(...recognized) : Infinity;
        const firstPageOf = (sec) => (doc.sections || []).find(x => x.num === sec.num)?.pages?.[0];
        (doc.numericToc || []).forEach(t => {
            if (sections.some(s => s.num === t.num) || t.num > maxNum) return;
            // 시작 쪽: 다음 절 시작 전 슬라이드 중 이 제목이 나오는 첫 슬라이드 (여러 절이 나열된 목차 슬라이드는 제외)
            const next = sections.filter(s => s.num > t.num).sort((x, y) => x.num - y.num)[0];
            const nextPage = next ? firstPageOf(next) : null;
            const pageMap = doc.content?.pageMap || {};
            const want = normTitle(t.title);
            const hit = Object.keys(pageMap).map(Number).sort((x, y) => x - y).find(n => {
                if (nextPage != null && n >= nextPage) return false;
                const ps = (pageMap[n] || []).map(normTitle);
                const listed = sections.filter(s => ps.some(q => q.includes(normTitle(s.title)))).length;
                return listed < 2 && ps.some(q => q.includes(want));
            });
            sections.push({ num: t.num, title: t.title, startLabel: hit != null ? (labels[hit] || String(hit)) : '', pageRange: '', include: true, subsections: [] });
        });
        sections.sort((a, b) => a.num - b.num);
        return {
            key: `${doc.fileName}#${i}`,
            num: chapterNum,
            roman: chapterNum ? NUM_TO_ROMAN[chapterNum] : '',
            title: doc.chapterTitle || doc.fileName.replace(/\.pptx$/i, ''),
            fileName: doc.fileName,
            startLabel: firstPos != null ? (labels[firstPos] || String(firstPos)) : '',
            include: true,
            sections
        };
    });
    chapters.sort((a, b) => (a.num ?? 99) - (b.num ?? 99));
    return chapters;
}

// ─────────────────────────────────────────────────────────────────────────────
// 2) 목차 샘플 슬라이드 분석
// ─────────────────────────────────────────────────────────────────────────────
const SHAPE_TAGS = new Set(['sp', 'grpSp', 'cxnSp', 'pic', 'graphicFrame']);

function readXfrm(el) {
    const xfrm = descendants(el, 'xfrm').find(x => {
        // 이 요소 자신의 spPr/grpSpPr/xfrm(그래픽프레임은 p:xfrm)만
        const p = x.parentNode;
        return p && (p.parentNode === el || p === el);
    });
    if (!xfrm) return null;
    const off = firstChild(xfrm, 'off');
    const ext = firstChild(xfrm, 'ext');
    if (!off || !ext) return null;
    const r = {
        x: parseInt(off.getAttribute('x') || '0', 10), y: parseInt(off.getAttribute('y') || '0', 10),
        cx: parseInt(ext.getAttribute('cx') || '0', 10), cy: parseInt(ext.getAttribute('cy') || '0', 10),
        offEl: off,
        extEl: ext
    };
    const chOff = firstChild(xfrm, 'chOff');
    const chExt = firstChild(xfrm, 'chExt');
    if (chOff && chExt) {
        r.chx = parseInt(chOff.getAttribute('x') || '0', 10);
        r.chy = parseInt(chOff.getAttribute('y') || '0', 10);
        r.chcx = parseInt(chExt.getAttribute('cx') || '1', 10) || 1;
        r.chcy = parseInt(chExt.getAttribute('cy') || '1', 10) || 1;
    }
    return r;
}

function placeholderKey(sp) {
    const ph = descendants(sp, 'ph')[0];
    if (!ph) return null;
    return { type: ph.getAttribute('type') || 'body', idx: ph.getAttribute('idx') || '' };
}

function leafText(el) {
    const paras = descendants(el, 'p').filter(p => p.namespaceURI === A_NS);
    return paras.map(p => descendants(p, 't').map(t => t.textContent || '').join('')).join('\n').trim();
}

function leafFontSize(el) {
    const r = descendants(el, 'rPr').find(x => x.getAttribute('sz')) || descendants(el, 'endParaRPr').find(x => x.getAttribute('sz'))
        || descendants(el, 'defRPr').find(x => x.getAttribute('sz'));
    return r ? parseInt(r.getAttribute('sz'), 10) / 100 : null;
}

// 슬라이드의 모든 도형(잎)을 절대 좌표로 펼친다
function collectLeaves(spTree, layoutXfrms) {
    const leaves = [];
    const walk = (container, tf) => {
        elementChildren(container).forEach(el => {
            if (!SHAPE_TAGS.has(el.localName)) return;
            let box = readXfrm(el);
            if (!box && el.localName === 'sp') {
                const key = placeholderKey(el);
                if (key) box = layoutXfrms.get(`${key.type}|${key.idx}`) || layoutXfrms.get(`${key.type}|`) || null;
            }
            const abs = box ? {
                x: tf.ox + (box.x - tf.chx) * tf.sx, y: tf.oy + (box.y - tf.chy) * tf.sy,
                w: box.cx * tf.sx, h: box.cy * tf.sy
            } : null;
            if (el.localName === 'grpSp') {
                if (box && box.chcx != null) {
                    const sx = (box.cx / box.chcx) * tf.sx;
                    const sy = (box.cy / box.chcy) * tf.sy;
                    walk(el, { ox: abs.x, oy: abs.y, chx: box.chx, chy: box.chy, sx, sy });
                } else {
                    walk(el, tf);
                }
                return;
            }
            const text = el.localName === 'sp' ? leafText(el) : '';
            leaves.push({ el, box: abs, text, fontSize: el.localName === 'sp' ? leafFontSize(el) : null, scaleY: tf.sy, hasText: !!text, isPlaceholder: !!placeholderKey(el) });
        });
    };
    walk(spTree, { ox: 0, oy: 0, chx: 0, chy: 0, sx: 1, sy: 1 });
    return leaves.filter(l => l.box);
}

async function layoutPlaceholderXfrms(zip, slidePath) {
    const map = new Map();
    const rels = await readRels(zip, slidePath);
    const layout = rels.find(r => /slideLayout$/.test(r.type));
    const chain = [];
    if (layout?.full && zip.file(layout.full)) {
        chain.push(layout.full);
        const lrels = await readRels(zip, layout.full);
        const master = lrels.find(r => /slideMaster$/.test(r.type));
        if (master?.full && zip.file(master.full)) chain.push(master.full);
    }
    for (const part of chain) {
        const doc = parseXml(await zip.file(part).async('string'));
        descendants(doc, 'sp').forEach(sp => {
            const key = placeholderKey(sp);
            const box = readXfrm(sp);
            if (!key || !box) return;
            const k = `${key.type}|${key.idx}`;
            if (!map.has(k)) map.set(k, box);
            if (!map.has(`${key.type}|`)) map.set(`${key.type}|`, box);
        });
    }
    return map;
}

const center = (b) => ({ cx: b.x + b.w / 2, cy: b.y + b.h / 2 });
const isLine = (l) => !l.hasText && (l.el.localName === 'cxnSp' || l.box.h < 0.02 * EMU_PER_INCH);

// 글자 폭 추정(인치): 한글·전각 ≈ 0.95em, 영문 대문자 0.62em, 소문자·숫자 0.52em, 공백 0.28em
function textWidthIn(text, pt, bold = false) {
    let em = 0;
    for (const ch of String(text || '')) {
        if (/[\u1100-\u11ff\u3000-\u9fff\uac00-\ud7af\uff00-\uffef]/.test(ch)) em += 0.95;
        else if (/[A-Z]/.test(ch)) em += 0.62;
        else if (/[a-z0-9]/.test(ch)) em += 0.52;
        else if (ch === ' ') em += 0.28;
        else em += 0.34;
    }
    return em * (pt || 10) / 72 * (bold ? 1.04 : 1);
}
function isBoldLeaf(el) {
    const r = descendants(el, 'rPr').find(x => x.getAttribute('b') != null);
    return r ? r.getAttribute('b') === '1' : false;
}
// 텍스트 상자 좌우 여백(EMU), 기본 0.1in
function leafInsets(el) {
    const bp = descendants(el, 'bodyPr')[0];
    const v = (k) => (bp && bp.getAttribute(k) != null ? parseInt(bp.getAttribute(k), 10) : 91440);
    return { l: v('lIns'), r: v('rIns') };
}
// 부모 그룹 축척
function groupScaleOf(unit) {
    let sx = 1;
    let sy = 1;
    let p = unit.parentNode;
    while (p && p.localName === 'grpSp') {
        const box = readXfrm(p);
        if (box && box.chcx) { sx *= box.cx / box.chcx; sy *= box.cy / box.chcy; }
        p = p.parentNode;
    }
    return { sx: sx || 1, sy: sy || 1 };
}
// 크기 변경 (슬라이드 좌표 EMU 증감)
function resizeUnit(unit, dwEmu, dhEmu) {
    const box = readXfrm(unit);
    if (!box) return;
    const { sx, sy } = groupScaleOf(unit);
    if (dwEmu) box.extEl.setAttribute('cx', String(Math.max(0, Math.round(box.cx + dwEmu / sx))));
    if (dhEmu) box.extEl.setAttribute('cy', String(Math.max(0, Math.round(box.cy + dhEmu / sy))));
}
// 글자 크기 배율
function scaleFonts(el, factor) {
    if (!factor || Math.abs(factor - 1) < 0.01) return;
    ['rPr', 'endParaRPr', 'defRPr'].forEach(tag => descendants(el, tag).forEach(r => {
        const sz = r.getAttribute('sz');
        if (sz) r.setAttribute('sz', String(Math.max(600, Math.round((parseInt(sz, 10) * factor) / 50) * 50)));
    }));
}
// spTree 끝(extLst 앞)에 도형 추가
function appendShape(parent, node) {
    const last = elementChildren(parent).slice(-1)[0];
    if (last && last.localName === 'extLst') parent.insertBefore(node, last); else parent.appendChild(node);
}

// 번호 표기 형식: base(roman|compound|pad2|plain) + suffix('.', ')', '')
function numStyleOf(text) {
    const t = String(text || '').trim();
    if (!t) return null;
    let m = t.match(/^([ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ])\s*([.．)]?)$/);
    if (m) return { base: 'roman', suffix: m[2] ? '.' : '' };
    m = t.match(/^(\d{1,2}(?:\.\d{1,2})+)\s*([.．)]?)$/);
    if (m) return { base: 'compound', suffix: m[2] ? '.' : '' };
    m = t.match(/^(0\d)\s*([.．)]?)$/);
    if (m) return { base: 'pad2', suffix: m[2] === ')' ? ')' : (m[2] ? '.' : '') };
    m = t.match(/^(\d{1,2})\s*([.．)]?)$/);
    if (m) return { base: 'plain', suffix: m[2] === ')' ? ')' : (m[2] ? '.' : '') };
    return null;
}

// 제목 앞 번호(예: '1. 환경분석 개요', 'Ⅰ. 사업 개요', '1.1 분석목적') 형식
function titlePrefixOf(text) {
    const m = String(text || '').match(/^\s*([ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]|\d{1,2}(?:\.\d{1,2}){0,2})(\s*[.．]\s*|\s+)(?=\S)/);
    if (!m) return null;
    const style = numStyleOf(m[1]);
    if (!style) return null;
    return { base: style.base, sep: m[2].replace(/．/g, '.') };
}

const IN = EMU_PER_INCH;
const rangeY = (cl) => ({ top: Math.min(...cl.items.map(i => i.box.y)), bottom: Math.max(...cl.items.map(i => i.box.y + i.box.h)) });

// 두 제목 묶음이 같은 목록의 상·하위 단계인지 (들여쓰기 가까이 + 줄이 서로 끼어 있거나 바로 붙어 있음)
function sameList(a, b) {
    if (Math.abs(a.x - b.x) > 2.5 * IN) return false;
    const ra = rangeY(a);
    const rb = rangeY(b);
    const ys = (cl) => cl.items.map(i => i.box.y).sort((p, q) => p - q);
    const ya = ys(a);
    const yb = ys(b);
    const between = (outer, inner) => inner.some(y => outer.some((y0, k) => k < outer.length - 1 && y > y0 && y < outer[k + 1]));
    if (between(ya, yb) || between(yb, ya)) return true;
    const rowH = Math.max(...a.items.map(i => i.box.h), ...b.items.map(i => i.box.h));
    const gap = Math.max(ra.top, rb.top) - Math.min(ra.bottom, rb.bottom);
    return gap <= rowH * 1.5;
}

/**
 * 목차 템플릿 인식 (단계·열 구조 포함)
 * @returns {{mode, groups, rows, levels, templates, pitch, numStyles, titlePrefixes, topIsChapter,
 *            chapterRoman, chapterTitle, chapterWord, lastRowBottom}}
 */
function detectTemplate(slideDoc, layoutXfrms, slideSize) {
    const spTree = descendants(slideDoc, 'spTree')[0];
    const leaves = collectLeaves(spTree, layoutXfrms);
    const textLeaves = leaves.filter(l => l.hasText);
    const tokenMode = textLeaves.some(l => TOKEN_RE.test(l.text));

    let candidates;
    if (tokenMode) {
        candidates = textLeaves.filter(l => /\{\{\s*제목\s*\}\}/.test(l.text));
    } else {
        candidates = textLeaves.filter(l => {
            const t = l.text;
            if (NUM_RE.test(t) || COMPOUND_RE.test(t) || ROMAN_RE.test(t) || TOC_MARKER_RE.test(t) || CHAPTER_WORD_RE.test(t) || PAGE_RE.test(t) || COUNT_RE.test(t)) return false;
            if (t.length > 80 || t.split('\n').length > 2) return false;
            // 상단 머리글(사업명·장 구분)과 하단 바닥글 띠는 목차 항목에서 제외
            const cy = l.box.y + l.box.h / 2;
            if (cy < slideSize.cy * 0.1 || cy > slideSize.cy * 0.94) return false;
            return l.box.h < slideSize.cy * 0.2;
        });
    }

    // 같은 왼쪽 x·비슷한 높이·글자 크기 = 한 단계의 제목 칸 묶음
    const tolX = 0.2 * IN;
    const clusters = [];
    candidates.forEach(l => {
        const c = clusters.find(cl => Math.abs(cl.x - l.box.x) <= tolX
            && Math.abs(cl.h - l.box.h) <= Math.max(cl.h, l.box.h) * 0.6
            && (!cl.font || !l.fontSize || Math.abs(cl.font - l.fontSize) <= Math.max(cl.font, l.fontSize) * 0.15));
        if (c) c.items.push(l); else clusters.push({ x: l.box.x, h: l.box.h, font: l.fontSize, items: [l] });
    });
    // 들여쓰기가 작아(예: 0.16in) 한 묶음에 섞인 상·하위 단계를 x 간격으로 다시 나눔
    for (let ci = 0; ci < clusters.length; ci++) {
        const cl = clusters[ci];
        const xs = cl.items.map(i => i.box.x).sort((p, q) => p - q);
        let cut = -1;
        for (let k = 1; k < xs.length; k++) {
            if (xs[k] - xs[k - 1] >= 0.08 * IN && k >= 2 && xs.length - k >= 2) { cut = (xs[k] + xs[k - 1]) / 2; break; }
        }
        if (cut < 0) continue;
        const right = cl.items.filter(i => i.box.x > cut);
        cl.items = cl.items.filter(i => i.box.x <= cut);
        cl.x = Math.min(...cl.items.map(i => i.box.x));
        clusters.push({ x: Math.min(...right.map(i => i.box.x)), h: right[0].box.h, font: right[0].fontSize, items: right });
    }
    const distinctRows = (cl) => new Set(cl.items.map(i => Math.round(i.box.y / (0.05 * IN)))).size;
    const valid = tokenMode ? clusters : clusters.filter(cl => distinctRows(cl) >= 2);
    if (valid.length === 0) return null;
    const maxSize = Math.max(...valid.map(cl => cl.items.length));
    const big = tokenMode ? valid : valid.filter(cl => cl.items.length >= Math.max(2, Math.ceil(maxSize * 0.5)));
    const selected = [...big];
    // 개수는 적어도 큰 묶음과 같은 목록의 상·하위 단계면 포함 (예: 장 2개 + 절 10개)
    valid.forEach(cl => { if (!selected.includes(cl) && big.some(b => sameList(cl, b))) selected.push(cl); });
    // 열마다 하나뿐인 상위 제목(예: 2단 목차의 'Ⅰ. 사업 개요' / 'Ⅱ. 환경분석')
    // 번호 머리말이 있거나, 같은 모양이 여러 열에 반복될 때만 상위 단계로 인정 (번호 없는 장 제목 1개는 장 슬롯)
    if (!tokenMode) {
        const singles = clusters.filter(cl => !selected.includes(cl) && cl.items.length === 1
            && big.some(b => sameList(cl, b) && cl.x <= b.x + 0.05 * IN && (cl.font || 0) >= (b.font || 0)));
        singles.forEach(cl => {
            const twins = singles.filter(o => o !== cl && Math.abs(o.h - cl.h) <= cl.h * 0.2 && (!o.font || !cl.font || Math.abs(o.font - cl.font) <= cl.font * 0.15));
            if (titlePrefixOf(cl.items[0].text) || twins.length > 0) selected.push(cl);
        });
    }

    // 열(group) 구성: 같은 목록인 묶음끼리
    const groups = [];
    selected.sort((a, b) => a.x - b.x).forEach(cl => {
        const g = groups.find(gr => gr.clusters.some(m => sameList(cl, m)));
        if (g) g.clusters.push(cl); else groups.push({ clusters: [cl] });
    });
    groups.forEach(g => {
        g.x = Math.min(...g.clusters.map(c => c.x));
        g.right = Math.max(...g.clusters.flatMap(c => c.items.map(i => i.box.x + i.box.w)));
        // 단계: 왼쪽일수록, 같은 위치면 글자가 클수록 상위
        const order = [...g.clusters].sort((a, b) => (Math.round((a.x - b.x) / (0.1 * IN))) || ((b.font || 0) - (a.font || 0)) || (b.h - a.h));
        order.forEach((c, i) => { c.level = Math.min(i, 2); });
    });
    groups.sort((a, b) => a.x - b.x);
    groups.forEach((g, gi) => {
        const prevRight = gi === 0 ? 0 : groups[gi - 1].right;
        g.left = gi === 0 ? 0 : (prevRight + g.x) / 2;
        g.boundRight = gi === groups.length - 1 ? slideSize.cx : (g.right + groups[gi + 1].x) / 2;
        g.leftMost = Math.max(g.left, g.x - 1.5 * IN);
    });

    const rows = [];
    groups.forEach((g, gi) => {
        const items = g.clusters.flatMap(c => c.items.map(t => ({ t, level: c.level }))).sort((a, b) => a.t.box.y - b.t.box.y);
        items.forEach(({ t, level }, i) => {
            const c = center(t.box).cy;
            const prevGap = i > 0 ? c - center(items[i - 1].t.box).cy : Infinity;
            const nextGap = i < items.length - 1 ? center(items[i + 1].t.box).cy - c : Infinity;
            let half = Math.min(prevGap, nextGap) / 2;
            if (!isFinite(half)) half = t.box.h * 0.75;
            rows.push({ group: gi, level, title: t, center: c, half, leaves: [t] });
        });
    });

    // 같은 줄의 번호·페이지·장식 도형 (목록 영역 가까이에 있는 것만)
    const titleSet = new Set(rows.map(r => r.title));
    // 목록 옆의 큰 패널(장 번호·제목 상자) 안 도형은 행이 아님
    const within = (b, l) => { const c = center(l.box); return c.cx >= b.x && c.cx <= b.x + b.w && c.cy >= b.y && c.cy <= b.y + b.h; };
    const panels = leaves.filter(l => !l.hasText && l.el.localName === 'sp' && l.box.h > slideSize.cy * 0.3
        && !rows.some(r => within(l.box, r.title)));
    const assigned = new Set();
    leaves.forEach(l => {
        if (titleSet.has(l)) return;
        if (panels.includes(l) || panels.some(pn => within(pn.box, l))) return;
        const { cx, cy } = center(l.box);
        if (tokenMode && /\{\{\s*(장번호|장제목)\s*\}\}/.test(l.text)) return;
        if (l.hasText && ((ROMAN_RE.test(l.text) && !tokenMode) || TOC_MARKER_RE.test(l.text) || /CHAPTER/i.test(l.text))) {
            // 로마숫자는 목록 안에 있으면(장 단계 번호) 행에 포함, 밖이면 장 슬롯
            const inList = rows.some(r => Math.abs(cy - r.center) <= r.half * 0.95 && cx >= groups[r.group].leftMost && cx < r.title.box.x);
            if (!(ROMAN_RE.test(l.text) && inList)) return;
        }
        const row = rows.find(r => Math.abs(cy - r.center) <= r.half * 0.95
            && cx >= groups[r.group].leftMost && cx <= groups[r.group].boundRight
            && l.box.x >= groups[r.group].leftMost - 0.05 * IN
            && l.box.h <= Math.max(r.title.box.h * 3, r.half * 2.2));
        if (row) { row.leaves.push(l); assigned.add(l); }
    });
    // 행 사이 구분선: 바로 위 행의 밑줄로 묶음
    leaves.forEach(l => {
        if (titleSet.has(l) || assigned.has(l) || !isLine(l)) return;
        if (panels.some(pn => within(pn.box, l))) return;
        const { cx, cy } = center(l.box);
        const above = rows
            .filter(r => cx >= groups[r.group].leftMost && cx <= groups[r.group].boundRight && l.box.x >= groups[r.group].leftMost - 0.05 * IN
                && r.center < cy && cy - r.center <= r.half * 4)
            .sort((a, b) => b.center - a.center)[0];
        if (!above) return;
        const next = rows.find(r => r.group === above.group && r.center > above.center && r.center < cy);
        if (!next) above.leaves.push(l);
    });
    rows.forEach(r => {
        const tc = center(r.title.box).cx;
        r.leaves.forEach(l => {
            if (l === r.title || !l.hasText) return;
            if (tokenMode) {
                if (/\{\{\s*번호\s*\}\}/.test(l.text)) r.num = l;
                else if (/\{\{\s*페이지\s*\}\}/.test(l.text)) r.page = l;
                return;
            }
            const lc = center(l.box).cx;
            if ((COMPOUND_RE.test(l.text) || NUM_RE.test(l.text) || ROMAN_RE.test(l.text) || /^\s*\d{1,2}(\.\d{1,2})+\s*$/.test(l.text)) && lc < tc && !r.num) r.num = l;
            else if ((PAGE_RE.test(l.text) || NUM_RE.test(l.text)) && lc > tc && !r.page) r.page = l;
        });
    });

    // 제목 뒤 점선(리더): 제목 줄 높이 안의 가로선이 제목 오른쪽에 있으면 제목 길이에 맞춰 시작점을 옮김
    rows.forEach(r => {
        const tb = r.title.box;
        r.leader = r.leaves.find(l => isLine(l) && center(l.box).cy >= tb.y && center(l.box).cy <= tb.y + tb.h && l.box.x > tb.x + 0.1 * IN) || null;
    });

    // 장 슬롯 (목록 밖)
    const rowLeafSet = new Set(rows.flatMap(r => r.leaves));
    const outside = textLeaves.filter(l => !rowLeafSet.has(l));
    let chapterRoman = null;
    let chapterTitle = null;
    let chapterWord = null;
    if (tokenMode) {
        chapterRoman = outside.find(l => /\{\{\s*장번호\s*\}\}/.test(l.text)) || null;
        chapterTitle = outside.find(l => /\{\{\s*장제목\s*\}\}/.test(l.text)) || null;
    } else {
        chapterRoman = outside.find(l => ROMAN_RE.test(l.text)) || null;
        chapterWord = outside.find(l => /CHAPTER\s*0?\d{1,2}/i.test(l.text)) || null;
        if (chapterRoman) {
            const rc = center(chapterRoman.box);
            chapterTitle = outside
                .filter(l => l !== chapterRoman && l !== chapterWord && !TOC_MARKER_RE.test(l.text) && !CHAPTER_WORD_RE.test(l.text) && l.text.length <= 30)
                .map(l => ({ l, d: Math.hypot(center(l.box).cx - rc.cx, center(l.box).cy - rc.cy) }))
                .filter(x => x.d <= 3.5 * IN)
                .sort((a, b) => a.d - b.d)[0]?.l || null;
        }
    }

    // 장 부가 슬롯: 'n개 절' 문구, 머리글 등의 'Ⅲ. 현황분석' 같은 장 언급
    const chapterCount = tokenMode ? null : (outside.find(l => COUNT_RE.test(l.text)) || null);
    let mentionRe = null;
    let mentions = [];
    if (!tokenMode && chapterRoman && chapterTitle) {
        mentionRe = chapterMentionRe(chapterRoman.text.trim().replace(/[.．]$/, ''), chapterTitle.text.trim());
        mentions = outside.filter(l => l !== chapterTitle && l !== chapterRoman && mentionRe.test(l.text));
    }

    // 단계별 견본 행·간격·번호 형식
    const levels = Math.max(...rows.map(r => r.level)) + 1;
    const templates = {};
    const pitch = {};
    const pairPitch = {}; // '앞 단계>이 단계' → 간격
    const numStyles = {};
    const titlePrefixes = {};
    for (let lv = 0; lv < levels; lv++) {
        const lvRows = rows.filter(r => r.level === lv);
        templates[lv] = lvRows[0];
        const deltas = [];
        groups.forEach((g, gi) => {
            const gRows = rows.filter(r => r.group === gi).sort((a, b) => a.center - b.center);
            gRows.forEach((r, i) => { if (i > 0 && r.level === lv) deltas.push(r.center - gRows[i - 1].center); });
        });
        deltas.sort((a, b) => a - b);
        pitch[lv] = deltas.length ? deltas[Math.floor(deltas.length / 2)] : lvRows[0].title.box.h * 1.5;
        groups.forEach((g, gi) => {
            const gRows = rows.filter(r => r.group === gi).sort((a, b) => a.center - b.center);
            gRows.forEach((r, i) => {
                if (i === 0 || r.level !== lv) return;
                const key = `${gRows[i - 1].level}>${lv}`;
                (pairPitch[key] = pairPitch[key] || []).push(r.center - gRows[i - 1].center);
            });
        });
        const numRow = lvRows.find(r => r.num);
        numStyles[lv] = numRow ? (tokenMode ? { base: 'plain', suffix: '' } : numStyleOf(numRow.num.text)) : null;
        titlePrefixes[lv] = tokenMode ? null : (lvRows.map(r => titlePrefixOf(r.title.text)).find(Boolean) || null);
    }
    Object.keys(pairPitch).forEach(k => { const v = pairPitch[k].sort((a, b) => a - b); pairPitch[k] = v[Math.floor(v.length / 2)]; });
    const topStyle = numStyles[0] || titlePrefixes[0];
    const topIsChapter = !!topStyle && topStyle.base === 'roman';
    groups.forEach((g, gi) => {
        const gRows = rows.filter(r => r.group === gi);
        g.y0 = Math.min(...gRows.map(r => r.center));
    });
    const lastRowBottom = Math.max(...rows.map(r => r.title.box.y + r.title.box.h));

    rows.sort((a, b) => a.group - b.group || a.center - b.center);
    return {
        mode: tokenMode ? 'token' : 'auto',
        groups,
        rows,
        levels,
        templates,
        pitch,
        pairPitch,
        numStyles,
        titlePrefixes,
        topIsChapter,
        chapterRoman,
        chapterTitle,
        chapterWord,
        chapterCount,
        mentions,
        mentionRe,
        outsideLeaves: leaves.filter(l => !rowLeafSet.has(l) && !panels.includes(l)),
        lastRowBottom
    };
}

const escRe = (t) => String(t).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
// 'Ⅲ. 현황분석' / 'Ⅲ.현황분석' / 'Ⅲ 현황분석' 형태의 장 언급
function chapterMentionRe(roman, title) {
    return new RegExp(`${escRe(roman)}\\s*[.．]?\\s*${escRe(title)}`);
}
// 'n개 절 · m개 항' 문구
function countLabel(secCount, subCount, withSub) {
    return `${secCount}개 절${withSub && subCount ? ` · ${subCount}개 항` : ''}`;
}


async function slideSizeOf(zip) {
    const pres = parseXml(await zip.file('ppt/presentation.xml').async('string'));
    const sz = descendants(pres, 'sldSz')[0];
    return { cx: parseInt(sz?.getAttribute('cx') || '12192000', 10), cy: parseInt(sz?.getAttribute('cy') || '6858000', 10) };
}

// 슬라이드 하나의 목차 디자인 종류 판별: cards(카드형 전체 목차) / table(표형 상세 목차) / list(행 반복 목차)
function classifySlide(doc, xf, slideSize) {
    let cards = null;
    let tables = null;
    let det = null;
    try { cards = detectCards(doc, xf, slideSize); } catch { cards = null; }
    if (cards) {
        const rows = cards.cards.reduce((n, c) => n + c.rows.length, 0);
        return { kind: 'cards', rowCount: rows, columns: cards.cards.length, levels: 2, topIsChapter: true, mode: 'auto', hasNum: true, hasPage: cards.cards.some(c => c.rows.some(r => r.page)), hasChapterRoman: true, hasChapterTitle: cards.cards.some(c => c.title) };
    }
    try { tables = detectTables(doc, xf, slideSize); } catch { tables = null; }
    if (tables) {
        const t = tables.templates;
        const rows = tables.tbls.reduce((n, x) => n + x.trs.length, 0);
        return { kind: 'table', rowCount: rows, columns: tables.tbls.length, levels: (t.ch ? 1 : 0) + (t.sec ? 1 : 0) + (t.sub ? 1 : 0), topIsChapter: !!t.ch, mode: 'auto', hasNum: true, hasPage: tables.pageCol >= 0, hasChapterRoman: !!t.ch, hasChapterTitle: !!tables.mention };
    }
    try { det = detectTemplate(doc, xf, slideSize); } catch (e) { det = null; console.warn('목차 샘플 인식 실패', e); }
    if (!det) return { kind: null, rowCount: 0, columns: 0, levels: 0, topIsChapter: false, mode: null, hasNum: false, hasPage: false, hasChapterRoman: false, hasChapterTitle: false };
    return {
        kind: 'list',
        rowCount: det.rows.length,
        columns: det.groups.length,
        levels: det.levels,
        topIsChapter: !!det.topIsChapter,
        mode: det.mode,
        hasNum: det.rows.some(r => r.num) || Object.values(det.titlePrefixes).some(Boolean),
        hasPage: det.rows.some(r => r.page),
        hasChapterRoman: !!det.chapterRoman,
        hasChapterTitle: !!det.chapterTitle
    };
}

/**
 * 목차 샘플 파일 분석: 슬라이드별 디자인 종류와 쓰임(앞부분 전체·상세 목차 / 장별 목차) 추천
 */
export async function analyzeTocSample(file) {
    const zip = await JSZip.loadAsync(await file.arrayBuffer());
    const slides = await presentationSlides(zip);
    const slideSize = await slideSizeOf(zip);
    const results = [];
    for (let i = 0; i < slides.length; i++) {
        const path = slides[i];
        const doc = parseXml(await zip.file(path).async('string'));
        const xf = await layoutPlaceholderXfrms(zip, path);
        const info = classifySlide(doc, xf, slideSize);
        const allText = leafText(doc);
        const tocWord = /목\s*차|CONTENTS/i.test(allText);
        results.push({
            index: i,
            path,
            preview: allText.split('\n').map(x => x.trim()).filter(Boolean).slice(0, 4).join(' · ').slice(0, 80),
            ...info,
            score: (info.mode === 'token' ? 1000 : 0) + info.rowCount * 10 + (tocWord ? 25 : 0) + (info.hasChapterRoman ? 5 : 0)
        });
    }
    const usable = results.filter(r => r.kind && (r.rowCount >= 2 || (r.mode === 'token' && r.rowCount >= 1)));
    const byScore = (list) => [...list].sort((a, b) => b.score - a.score);
    const cards = usable.find(r => r.kind === 'cards');
    const tables = usable.filter(r => r.kind === 'table');
    const table = tables[0] || null;
    const lists = byScore(usable.filter(r => r.kind === 'list'));
    const overview = cards || lists.find(r => r.topIsChapter) || lists[0] || null;
    // 장별 세부목차: 테이블(표형)을 우선 사용 (단일 장 언급 슬라이드가 있으면 최우선, 없으면 다른 표 슬라이드, 없으면 첫 표 슬라이드, 그 다음 텍스트박스형 lists)
    const singleChapterTable = tables.find(r => r.preview && /(?:—|-)\s*[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ][.．]/.test(r.preview) && !/\//.test(r.preview));
    const chapterTable = singleChapterTable || tables.find(r => r !== table) || table || null;
    const chapter = chapterTable || lists.find(r => r !== overview) || (overview && overview.kind === 'list' ? overview : null);
    const defaults = {
        overviewIndex: overview ? overview.index : null,
        detailIndex: table ? table.index : null,
        chapterIndex: chapter ? chapter.index : null
    };
    const best = byScore(usable)[0] || null;
    return { slideCount: slides.length, slides: results, bestIndex: defaults.chapterIndex ?? (best ? best.index : null), defaults };
}

// ─────────────────────────────────────────────────────────────────────────────
// 3) 채우기
// ─────────────────────────────────────────────────────────────────────────────
function setLeafText(sp, text, tokenMap = null) {
    const txBody = descendants(sp, 'txBody')[0];
    if (!txBody) return;
    const paras = elementChildren(txBody).filter(n => n.localName === 'p');
    if (paras.length === 0) return;
    const doc = sp.ownerDocument;
    if (tokenMap) {
        // 토큰 방식: 각 문단의 텍스트를 합쳐 토큰만 치환
        paras.forEach(p => {
            const runs = elementChildren(p).filter(n => n.localName === 'r' || n.localName === 'fld');
            const full = runs.map(r => descendants(r, 't').map(t => t.textContent).join('')).join('');
            if (!TOKEN_RE.test(full)) return;
            let hadEmpty = false;
            let replaced = full.replace(/\{\{\s*(장번호|장제목|번호|제목|페이지)\s*\}\}/g, (m, k) => {
                const v = tokenMap[k] ?? '';
                if (!v) hadEmpty = true;
                return v;
            });
            // 빈 값으로 남은 구분 기호 정리 (예: '{{장번호}}. {{장제목}}' → 전체 목차에서 '. 목차' 방지)
            if (hadEmpty) replaced = replaced.replace(/^\s*[.．·\-–:]\s*/, '').replace(/\s*[.．·\-–:]\s*$/, '').replace(/\s*([.．·\-–])\s*\1/g, '$1');
            const first = runs[0];
            const t = descendants(first, 't')[0];
            if (t) t.textContent = replaced;
            runs.slice(1).forEach(r => p.removeChild(r));
        });
        return;
    }
    const p = paras[0];
    paras.slice(1).forEach(x => txBody.removeChild(x));
    const runs = elementChildren(p).filter(n => ['r', 'fld', 'br'].includes(n.localName));
    let run = runs.find(n => n.localName === 'r') || null;
    if (!run) {
        run = doc.createElementNS(A_NS, 'a:r');
        const end = firstChild(p, 'endParaRPr');
        if (end) {
            const rPr = doc.createElementNS(A_NS, 'a:rPr');
            Array.from(end.attributes).forEach(a => rPr.setAttribute(a.name, a.value));
            elementChildren(end).forEach(c => rPr.appendChild(c.cloneNode(true)));
            run.appendChild(rPr);
        }
        run.appendChild(doc.createElementNS(A_NS, 'a:t'));
        if (end) p.insertBefore(run, end); else p.appendChild(run);
    }
    runs.forEach(n => { if (n !== run) p.removeChild(n); });
    let t = firstChild(run, 't');
    if (!t) { t = doc.createElementNS(A_NS, 'a:t'); run.appendChild(t); }
    t.textContent = text;
    // 빈 글자면 문단 끝 서식(endParaRPr)이 줄 높이를 정하므로, 없으면 글자 서식을 복사 (기본 18pt로 커지는 것 방지)
    const rPr = firstChild(run, 'rPr');
    if (rPr && !firstChild(p, 'endParaRPr')) {
        const end = doc.createElementNS(A_NS, 'a:endParaRPr');
        Array.from(rPr.attributes).forEach(at => { if (at.name !== 'dirty' && at.name !== 'err') end.setAttribute(at.name, at.value); });
        elementChildren(rPr).forEach(c => end.appendChild(c.cloneNode(true)));
        p.appendChild(end);
    }
    // 빈 글자 run은 남기지 않음 (일부 뷰어가 빈 run을 기본 글자 크기로 계산)
    if (!text && firstChild(p, 'endParaRPr')) p.removeChild(run);
}

function numText(item, style) {
    if (!style) return '';
    let core;
    if (style.base === 'roman' || (item.isChapter && style.base !== 'compound')) core = item.roman || NUM_TO_ROMAN[item.num] || String(item.num ?? '');
    else if (style.base === 'compound') core = item.parentNum != null ? `${item.parentNum}.${item.num}` : String(item.num ?? '');
    else if (style.base === 'pad2') core = String(item.num ?? '').padStart(2, '0');
    else core = String(item.num ?? '');
    return core + (style.suffix || '');
}

// 행을 이루는 최상위 단위(그룹이면 그 행만 담은 가장 큰 그룹)
function rowUnits(row) {
    const leafEls = new Set(row.leaves.map(l => l.el));
    const units = new Set();
    row.leaves.forEach(l => {
        let unit = l.el;
        let parent = unit.parentNode;
        while (parent && parent.localName === 'grpSp') {
            const inner = Array.from(parent.getElementsByTagNameNS('*', '*')).filter(n => SHAPE_TAGS.has(n.localName) && n.localName !== 'grpSp');
            if (inner.length && inner.every(n => leafEls.has(n))) { unit = parent; parent = parent.parentNode; } else break;
        }
        units.add(unit);
    });
    return Array.from(units);
}

// 단위 요소를 슬라이드 좌표 (dx, dy)만큼 이동 (그룹 안이면 부모 그룹 축척 반영)
function shiftUnit(unit, dxEmu, dyEmu) {
    let sx = 1;
    let sy = 1;
    let p = unit.parentNode;
    while (p && p.localName === 'grpSp') {
        const box = readXfrm(p);
        if (box && box.chcx) { sx *= box.cx / box.chcx; sy *= box.cy / box.chcy; }
        p = p.parentNode;
    }
    const box = readXfrm(unit);
    if (!box) return;
    if (dxEmu) box.offEl.setAttribute('x', String(Math.round(box.x + dxEmu / (sx || 1))));
    if (dyEmu) box.offEl.setAttribute('y', String(Math.round(box.y + dyEmu / (sy || 1))));
}

function nextShapeId(doc) {
    let max = 0;
    descendants(doc, 'cNvPr').forEach(n => { max = Math.max(max, parseInt(n.getAttribute('id') || '0', 10)); });
    return max + 1;
}

// 요소 안 경로(자식 인덱스)로 복제본에서 같은 잎 찾기
function pathFrom(ancestor, node) {
    const path = [];
    let cur = node;
    while (cur && cur !== ancestor) {
        const parent = cur.parentNode;
        path.unshift(elementChildren(parent).indexOf(cur));
        cur = parent;
    }
    return path;
}
function followPath(root, path) {
    let cur = root;
    for (const i of path) cur = elementChildren(cur)[i];
    return cur;
}

/**
 * 템플릿 요약 (목차 깊이·장 단계 여부) - 항목 구성 결정용
 */
export function describeTemplate({ templateXml, layoutXfrms, slideSize }) {
    const det = detectTemplate(parseXml(templateXml), layoutXfrms, slideSize);
    if (!det) return null;
    return { levels: det.levels, topIsChapter: det.topIsChapter, columns: det.groups.length, rows: det.rows.length, mode: det.mode };
}

/**
 * 항목(단계 포함)을 템플릿 디자인으로 배치해 슬라이드 XML 목록을 만든다 (넘치면 여러 장)
 * @param {Array<{level,num,parentNum,roman,title,page}>} p.items
 * @param {{roman,title,num}} p.chapter 장 슬롯 값
 */
export function fillTemplateSlides({ templateXml, layoutXfrms, slideSize, items, chapter, showPage = true }) {
    const probe = detectTemplate(parseXml(templateXml), layoutXfrms, slideSize);
    if (!probe || probe.rows.length === 0) throw new Error('목차 샘플에서 반복되는 목차 항목 칸을 찾지 못했습니다.');
    const warnings = [];
    const list = items.filter(it => it.level < probe.levels);
    if (list.length === 0) return { slides: [], warnings };
    // 각 항목의 상위 항목 (열·장이 바뀌어 이어질 때 '(계속)' 행으로 다시 보여줌)
    const stack = [];
    list.forEach(it => {
        while (stack.length && stack[stack.length - 1].level >= it.level) stack.pop();
        it.ancestors = [...stack];
        stack.push(it);
    });

    // ── 배치 계산: 열마다 위에서부터 단계별 간격으로, 넘치면 다음 열·다음 장 ──
    const groups = probe.groups;
    const maxRowH = Math.max(...probe.rows.map(r => r.title.box.h));
    const bottomLimit = Math.max(probe.lastRowBottom, slideSize.cy * 0.9) - maxRowH / 2;
    // 단계 간 간격: 샘플에 없는 조합(예: 항 없는 절 다음 절)은 같은 앞 단계의 간격 중 큰 값
    const stepOf = (a, b) => {
        if (probe.pairPitch[`${a}>${b}`] != null) return probe.pairPitch[`${a}>${b}`];
        const same = Object.keys(probe.pairPitch).filter(k => k.startsWith(`${a}>`)).map(k => probe.pairPitch[k]);
        return same.length ? Math.max(...same) : probe.pitch[b];
    };
    const place = (seq, f, limit = bottomLimit) => {
        const out = [];
        let g = 0;
        let y = groups[0].y0;
        let prevLevel = null;
        const put = (item) => {
            if (prevLevel !== null) {
                const step = stepOf(prevLevel, item.level) * f;
                if (y + step > limit) return false;
                y += step;
            }
            out.push({ g, y, item });
            prevLevel = item.level;
            return true;
        };
        const cont = (a) => ({ ...a, title: `${a.title}${CONT_SUFFIX}`, cont: true });
        for (let i = 0; i < seq.length; i++) {
            const it = seq[i];
            const anc = it.ancestors || [];
            if (i === 0) {
                // 장이 바뀌어 하위 항목부터 시작하면 상위 항목을 '(계속)'으로
                anc.forEach(a => put(cont(a)));
                if (!put(it)) return null;
                continue;
            }
            if (put(it)) continue;
            // 다음 열: 열 끝에 상위 항목만 남지 않게 함께 옮기고, 나머지 상위는 '(계속)' 행으로
            const moved = [];
            while (out.length && out[out.length - 1].g === g && anc.includes(out[out.length - 1].item)
                && out.filter(o => o.g === g).length > 1) moved.unshift(out.pop().item);
            g += 1;
            if (g >= groups.length) return null;
            y = groups[g].y0;
            prevLevel = null;
            anc.filter(a => !moved.includes(a)).forEach(a => put(cont(a)));
            moved.forEach(m => put(m));
            if (!put(it)) return null;
        }
        return out;
    };
    // 줄이 서로 겹치지 않는 한도까지 간격을 좁힐 수 있음 (넘으면 다음 장)
    const MIN_F = Math.min(1, Math.max(0.55, ...Object.keys(probe.templates).map(lv => {
        const t = probe.templates[lv];
        return (t.title.box.h * 1.05) / (probe.pitch[lv] || t.title.box.h * 1.5);
    })));
    const paginate = (softMax) => {
        const out = [];
        let cur = [];
        list.forEach(it => {
            // 장 나눔 균형: 목표 개수에 이르면 상위 단계 항목에서 끊음
            if (softMax && cur.length >= softMax) {
                const carry = [];
                while (cur.length > 1 && cur[cur.length - 1].level < it.level) carry.unshift(cur.pop());
                out.push(cur);
                cur = carry;
            }
            cur.push(it);
            if (!place(cur, MIN_F) && cur.length > 1) {
                cur.pop();
                // 상위 항목만 덩그러니 남지 않게 다음 장으로 넘김
                const carry = [];
                while (cur.length > 1 && cur[cur.length - 1].level < it.level) carry.unshift(cur.pop());
                out.push(cur);
                cur = [...carry, it];
            }
        });
        if (cur.length) out.push(cur);
        return out;
    };
    let pages = paginate(0);
    if (pages.length > 1) {
        // 마지막 장에 몇 줄만 남지 않도록 장마다 비슷한 개수로 다시 나눔 (같은 장 수가 되는 가장 작은 목표치)
        const base = Math.ceil(list.length / pages.length);
        for (let extra = 0; extra <= 6; extra++) {
            const balanced = paginate(base + extra);
            if (balanced.length === pages.length) { pages = balanced; break; }
        }
    }
    if (pages.length > 1) warnings.push(`항목 ${list.length}개가 한 장에 다 들어가지 않아 ${pages.length}장으로 나눴습니다.`);

    const slides = pages.map((pageItems, pi) => {
        let f = 1;
        let pos = place(pageItems, 1);
        if (!pos) {
            let lo = MIN_F;
            let hi = 1;
            for (let k = 0; k < 12; k++) { const mid = (lo + hi) / 2; if (place(pageItems, mid)) lo = mid; else hi = mid; }
            f = lo;
            pos = place(pageItems, f) || place(pageItems, MIN_F);
            if (f < 0.95) warnings.push(`${pages.length > 1 ? `${pi + 1}번째 장: ` : ''}행 간격을 ${Math.round((1 - f) * 100)}% 좁혔습니다.`);
        }
        // 여러 열에 걸치면 열 높이를 고르게 (한 열에 몰리고 다음 열에 몇 줄만 남지 않도록)
        if (pos && new Set(pos.map(p => p.g)).size > 1) {
            let lo = Math.min(...groups.map(g => g.y0));
            let hi = bottomLimit;
            for (let k = 0; k < 14; k++) { const mid = (lo + hi) / 2; if (place(pageItems, f, mid)) hi = mid; else lo = mid; }
            pos = place(pageItems, f, hi) || pos;
        }

        const doc = parseXml(templateXml);
        const det = detectTemplate(doc, layoutXfrms, slideSize);
        const tokenMode = det.mode === 'token';

        // 장 슬롯 (머리글에 장 언급이 있으면 거기에 쪽 순서를 붙이고 장 제목 칸은 그대로)
        const hasMention = det.mentions.length > 0;
        const title = pi > 0 && chapter.title && !hasMention ? `${chapter.title}${CONT_SUFFIX}` : (chapter.title || '');
        const chapterTokens = { 장번호: chapter.roman || '', 장제목: title };
        if (hasMention) {
            const label = `${chapter.roman ? `${chapter.roman}. ` : ''}${chapter.title || ''}${pages.length > 1 ? ` (${pi + 1}/${pages.length})` : ''}`;
            det.mentions.forEach(m => setLeafText(m.el, m.text.replace(det.mentionRe, label)));
        }
        if (det.chapterCount && chapter.counts) {
            setLeafText(det.chapterCount.el, countLabel(chapter.counts.sections, chapter.counts.subsections, /항/.test(det.chapterCount.text)));
        }
        if (det.chapterRoman) setLeafText(det.chapterRoman.el, chapter.roman || '', tokenMode ? chapterTokens : null);
        if (det.chapterTitle) {
            setLeafText(det.chapterTitle.el, title, tokenMode ? chapterTokens : null);
            if (!tokenMode) wrapSlotTitle(det, title);
        }
        if (det.chapterWord) {
            setLeafText(det.chapterWord.el, chapter.num
                ? det.chapterWord.text.replace(/(CHAPTER\s*)0?\d{1,2}/i, (m, w) => `${w}${String(chapter.num).padStart(2, '0')}`)
                : 'CONTENTS');
        }

        // 단계별 견본(복제 원본)을 기억하고 샘플의 행은 모두 제거
        const tpl = {};
        for (let lv = 0; lv < det.levels; lv++) {
            const row = det.templates[lv];
            const units = rowUnits(row).map(u => ({ el: u, parent: u.parentNode }));
            const locate = (leaf) => {
                if (!leaf) return null;
                const ui = units.findIndex(u => u.el === leaf.el || isAncestor(u.el, leaf.el));
                return ui < 0 ? null : { ui, path: units[ui].el === leaf.el ? [] : pathFrom(units[ui].el, leaf.el) };
            };
            tpl[lv] = { row, units, title: locate(row.title), num: locate(row.num), page: locate(row.page), leader: locate(row.leader), groupX: det.groups[row.group].x };
        }
        det.rows.forEach(r => rowUnits(r).forEach(u => u.parentNode && u.parentNode.removeChild(u)));

        let idCounter = nextShapeId(doc);
        pos.forEach(({ g, y, item: it }) => {
            const t = tpl[it.level];
            const dx = det.groups[g].x - t.groupX;
            const dy = y - t.row.center;
            const clones = t.units.map(u => {
                const c = u.el.cloneNode(true);
                descendants(c, 'cNvPr').forEach(n => n.setAttribute('id', String(idCounter++)));
                u.parent.appendChild(c);
                shiftUnit(c, dx, dy);
                return c;
            });
            const leafOf = (loc) => (loc ? (loc.path.length ? followPath(clones[loc.ui], loc.path) : clones[loc.ui]) : null);
            const titleEl = leafOf(t.title);
            const numEl = leafOf(t.num);
            const pageEl = leafOf(t.page);
            const numStr = numText(it, det.numStyles[it.level]);
            const pageStr = showPage ? (it.page || '') : '';
            if (tokenMode) {
                const tokens = { 번호: numStr || String(it.num ?? ''), 제목: it.title, 페이지: pageStr };
                if (titleEl) setLeafText(titleEl, '', tokens);
                if (numEl) setLeafText(numEl, '', tokens);
                if (pageEl) setLeafText(pageEl, '', tokens);
            } else {
                const prefix = det.titlePrefixes[it.level];
                const titleText = prefix ? `${numText(it, { base: prefix.base, suffix: '' })}${prefix.sep}${it.title}` : it.title;
                if (titleEl) setLeafText(titleEl, titleText);
                if (numEl) setLeafText(numEl, numStr);
                if (pageEl) setLeafText(pageEl, pageStr);
                fitLeader(leafOf(t.leader), t.row, titleText, dx);
            }
        });
        return serializeXml(doc);
    });
    return { slides, warnings };
}

// 장 제목 칸이 좁아 두 줄이 되면 칸을 늘리고 그 아래 도형(구분선·'n개 절')을 내림, 두 줄로도 넘치면 글자 축소
function wrapSlotTitle(det, text) {
    const l = det.chapterTitle;
    const pt = l.fontSize;
    if (!pt) return;
    const ins = leafInsets(l.el);
    const avail = (l.box.w - ins.l - ins.r) / EMU_PER_INCH;
    const need = textWidthIn(text, pt, isBoldLeaf(l.el));
    if (need <= avail * 1.02) return;
    if (need > avail * 2) scaleFonts(l.el, Math.max(0.6, (avail * 2) / need));
    const lineH = Math.min(l.box.h, (pt * 1.25 / 72) * EMU_PER_INCH);
    resizeUnit(l.el, 0, lineH);
    const below = det.outsideLeaves.filter(o => o !== l && o.box.y >= l.box.y + l.box.h * 0.6 && o.box.y <= l.box.y + l.box.h + 1.2 * EMU_PER_INCH
        && o.box.x < l.box.x + l.box.w && o.box.x + o.box.w > l.box.x && o.box.w <= l.box.w * 1.5);
    below.forEach(o => shiftUnit(o.el, 0, lineH));
}

// 점선 리더를 제목 끝 바로 뒤에서 시작하도록 (끝점은 그대로, 제목이 너무 길면 숨김)
function fitLeader(leaderEl, row, titleText, dx) {
    if (!leaderEl || !row.leader) return;
    const algn = descendants(row.title.el, 'pPr')[0]?.getAttribute('algn');
    if (algn && algn !== 'l') return;
    const tb = row.title.box;
    const lb = row.leader.box;
    const ins = leafInsets(row.title.el);
    const start = tb.x + dx + ins.l + textWidthIn(titleText, row.title.fontSize || 10, isBoldLeaf(row.title.el)) * EMU_PER_INCH + 0.06 * EMU_PER_INCH;
    const cur = lb.x + dx;
    const end = cur + lb.w;
    if (end - start < 0.15 * EMU_PER_INCH) { leaderEl.parentNode?.removeChild(leaderEl); return; }
    shiftUnit(leaderEl, start - cur, 0);
    resizeUnit(leaderEl, -(start - cur), 0);
}

function isAncestor(a, b) {
    let cur = b;
    while (cur) { if (cur === a) return true; cur = cur.parentNode; }
    return false;
}

// ─────────────────────────────────────────────────────────────────────────────
// 3-2) 카드형 전체 목차 샘플 인식: 장마다 카드(머리: 로마숫자·장 제목·'n개 절', 몸통: 절 행)
// ─────────────────────────────────────────────────────────────────────────────
function detectCards(slideDoc, layoutXfrms) {
    const spTree = descendants(slideDoc, 'spTree')[0];
    const leaves = collectLeaves(spTree, layoutXfrms);
    const romans = leaves.filter(l => l.hasText && BIG_ROMAN_RE.test(l.text) && (l.fontSize || 0) >= 18);
    if (romans.length < 2) return null;
    const shapes = leaves.filter(l => !l.hasText && l.el.localName === 'sp');
    const inside = (b, p) => p.cx >= b.x && p.cx <= b.x + b.w && p.cy >= b.y && p.cy <= b.y + b.h;
    const cards = [];
    romans.forEach(r => {
        const header = shapes.filter(s => inside(s.box, center(r.box))).sort((a, b) => a.box.w * a.box.h - b.box.w * b.box.h)[0];
        if (!header) return;
        const hb = header.box;
        const body = shapes
            .filter(s => s !== header && Math.abs(s.box.x - hb.x) <= 0.08 * IN && Math.abs(s.box.w - hb.w) <= 0.08 * IN && s.box.y >= hb.y + hb.h - 0.08 * IN)
            .sort((a, b) => a.box.y - b.box.y)[0];
        if (!body) return;
        const x0 = hb.x;
        const x1 = hb.x + hb.w;
        const y0 = hb.y;
        const y1 = body.box.y + body.box.h;
        const members = leaves.filter(l => { const c = center(l.box); return c.cx >= x0 && c.cx <= x1 && c.cy >= y0 && c.cy <= y1; });
        const headTexts = members.filter(l => l.hasText && center(l.box).cy < body.box.y);
        const word = headTexts.find(l => /CHAPTER/i.test(l.text)) || null;
        const count = headTexts.find(l => COUNT_RE.test(l.text)) || null;
        const title = headTexts.filter(l => l !== r && l !== word && l !== count).sort((a, b) => (b.fontSize || 0) - (a.fontSize || 0))[0] || null;
        const nums = members.filter(l => l.hasText && NUM_RE.test(l.text) && l.box.y >= body.box.y).sort((a, b) => a.box.y - b.box.y);
        const rows = nums.map((n, i) => {
            const top = n.box.y - 0.04 * IN;
            const bottom = i < nums.length - 1 ? nums[i + 1].box.y - 0.04 * IN : y1;
            const ls = members.filter(l => l !== body && l !== header && l.box.y >= top && l.box.y < bottom);
            const texts = ls.filter(l => l.hasText && l !== n);
            const page = texts.find(l => PAGE_RE.test(l.text)) || null;
            const ttl = texts.filter(l => l !== page).sort((a, b) => (b.fontSize || 0) - (a.fontSize || 0) || b.box.w - a.box.w)[0] || null;
            return { num: n, title: ttl, page, leaves: ls, top: n.box.y };
        }).filter(row => row.title);
        rows.forEach(row => { row.seps = row.leaves.filter(l => isLine(l) && l.box.y > row.title.box.y); });
        cards.push({ x0, x1, y0, y1, header, body, roman: r, word, count, title, rows, members });
    });
    if (cards.length < 2 || !cards.some(c => c.rows.length)) return null;
    cards.sort((a, b) => a.x0 - b.x0);
    return { cards, leaves, spTree };
}

const median = (arr) => { const v = [...arr].sort((a, b) => a - b); return v.length ? v[Math.floor(v.length / 2)] : null; };

// ─────────────────────────────────────────────────────────────────────────────
// 3-2b) 카드형 전체 목차 → 표로 작성: 장마다 같은 폭의 열(번호 칸 + 제목 칸), 장 사이 간격 열
//       색·글꼴은 샘플 카드(머리 상자·몸통 상자·번호 원·구분선·글상자 서식)에서 가져옴
// ─────────────────────────────────────────────────────────────────────────────
const solidOf = (el) => {
    const spPr = elementChildren(el).find(n => n.localName === 'spPr');
    if (!spPr) return null;
    const sf = elementChildren(spPr).find(n => n.localName === 'solidFill');
    const c = sf && descendants(sf, 'srgbClr')[0];
    return c ? c.getAttribute('val') : null;
};
const lineOf = (el) => {
    const spPr = elementChildren(el).find(n => n.localName === 'spPr');
    const ln = spPr && elementChildren(spPr).find(n => n.localName === 'ln');
    if (!ln || firstChild(ln, 'noFill')) return null;
    const c = descendants(ln, 'srgbClr')[0];
    if (!c) return null;
    const dash = descendants(ln, 'prstDash')[0]?.getAttribute('val') || 'solid';
    return { color: c.getAttribute('val'), w: parseInt(ln.getAttribute('w') || '9525', 10), dash };
};
const rPrXml = (leaf) => {
    const r = leaf && descendants(leaf.el, 'rPr')[0];
    return r ? new XMLSerializer().serializeToString(r).replace(/\sxmlns(:\w+)?="[^"]*"/g, '') : '<a:rPr/>';
};
const xmlEsc = (t) => String(t ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
// rPr 문자열의 글자 크기·색 바꾸기
const tuneRPr = (xml, { scale = 1, color = null } = {}) => {
    let out = xml.replace(/\ssz="(\d+)"/, (m, v) => ` sz="${Math.max(600, Math.round((parseInt(v, 10) * scale) / 50) * 50)}"`);
    if (color) {
        out = out.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/, `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`);
        if (!/<a:solidFill>/.test(out)) out = out.replace(/<a:rPr([^>]*?)\/>/, `<a:rPr$1><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:rPr>`).replace(/<a:rPr([^>]*)>(?!<a:solidFill)/, `<a:rPr$1><a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`);
    }
    return out;
};
const szOf = (xml) => { const m = xml.match(/\ssz="(\d+)"/); return m ? parseInt(m[1], 10) / 100 : 10; };
const para = (rPr, text, algn = 'l') => `<a:p><a:pPr algn="${algn}"/>${text ? `<a:r>${rPr}<a:t>${xmlEsc(text)}</a:t></a:r>` : ''}<a:endParaRPr${(rPr.match(/^<a:rPr([^>]*?)\/?>/) || ['', ''])[1].replace(/\/$/, '')}/></a:p>`;
const lnXml = (tag, ln) => (ln
    ? `<a:${tag} w="${ln.w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="${ln.color}"/></a:solidFill><a:prstDash val="${ln.dash}"/></a:${tag}>`
    : `<a:${tag} w="6350"><a:noFill/></a:${tag}>`);
const cellXml = ({ paras, fillColor = null, borders = {}, mar = {}, anchor = 't', span = 0, hMerge = false }) => {
    const m = { l: 45720, r: 36576, t: 45720, b: 27432, ...mar };
    return `<a:tc${span > 1 ? ` gridSpan="${span}"` : ''}${hMerge ? ' hMerge="1"' : ''}><a:txBody><a:bodyPr/><a:lstStyle/>${paras || '<a:p><a:endParaRPr sz="800"/></a:p>'}</a:txBody>`
        + `<a:tcPr marL="${m.l}" marR="${m.r}" marT="${m.t}" marB="${m.b}" anchor="${anchor}">`
        + `${lnXml('lnL', borders.l)}${lnXml('lnR', borders.r)}${lnXml('lnT', borders.t)}${lnXml('lnB', borders.b)}`
        + `${fillColor ? `<a:solidFill><a:srgbClr val="${fillColor}"/></a:solidFill>` : '<a:noFill/>'}</a:tcPr></a:tc>`;
};

export function fillCardTableSlides({ templateXml, layoutXfrms, slideSize, chapters, showPage = true }) {
    const probe = detectCards(parseXml(templateXml), layoutXfrms, slideSize);
    if (!probe) throw new Error('카드형 전체 목차 샘플에서 장 카드를 찾지 못했습니다.');
    const warnings = [];
    const perSlide = probe.cards.length;
    const nSlides = Math.max(1, Math.ceil(chapters.length / perSlide));
    const per = Math.ceil(chapters.length / nSlides);
    const pages = [];
    for (let i = 0; i < chapters.length; i += per) pages.push(chapters.slice(i, i + per));
    const withSub = probe.cards.some(c => c.count && /항/.test(c.count.text));

    const slides = pages.map(chs => {
        const doc = parseXml(templateXml);
        const det = detectCards(doc, layoutXfrms, slideSize);
        const { cards, spTree } = det;
        // 견본 카드: 가장 흔한 제목 글자 크기를 쓰는 첫 카드
        const fontOf = (c) => c.rows[0]?.title.fontSize || 0;
        const freq = {};
        cards.forEach(c => { freq[fontOf(c)] = (freq[fontOf(c)] || 0) + 1; });
        const modeFont = Number(Object.keys(freq).sort((a, b) => freq[b] - freq[a])[0]);
        const tc = cards.find(c => c.rows.length && fontOf(c) === modeFont) || cards.find(c => c.rows.length);
        const singleH = Math.min(...tc.rows.map(r => r.title.box.h));
        const tr = tc.rows.find(r => r.title.box.h <= singleH * 1.2) || tc.rows[0];
        const pitch = median(tc.rows.slice(1).map((r, i) => r.top - tc.rows[i].top)) || singleH * 2.2;

        // 서식
        const headFill = solidOf(tc.header.el) || '224F77';
        const bodyFill = solidOf(tc.body.el);
        const bodyLine = lineOf(tc.body.el) || { color: '7F8FA9', w: 9525, dash: 'solid' };
        const sepLeaf = tr.seps[0];
        const sep = (sepLeaf && lineOf(sepLeaf.el)) || { color: 'D9D9D9', w: 6350, dash: 'sysDot' };
        const oval = tr.leaves.find(l => !l.hasText && l.el.localName === 'sp' && solidOf(l.el));
        const numColor = (oval && solidOf(oval.el)) || '0072BA';
        const rp = {
            roman: rPrXml(tc.roman), word: rPrXml(tc.word), title: rPrXml(tc.title), count: rPrXml(tc.count),
            num: rPrXml(tr.num), rowTitle: rPrXml(tr.title), page: rPrXml(tr.page)
        };

        // 배치: 샘플 카드 영역 전체 폭을 장 수로 똑같이 나눔
        const X0 = Math.min(...cards.map(c => c.x0));
        const X1 = Math.max(...cards.map(c => c.x1));
        const gap = median(cards.slice(1).map((c, i) => c.x0 - cards[i].x1).filter(g => g >= 0)) ?? 0.1 * IN;
        const n = chs.length;
        const colW = Math.floor((X1 - X0 - gap * (n - 1)) / n);
        const numW = Math.round(Math.min(colW * 0.25, Math.max(0.26 * IN, (tr.title.box.x - tc.x0) * 0.85)));
        const titleW = colW - numW;
        const headH = tc.header.box.h;
        const bodyH = tc.body.box.h * 0.97; // 뷰어별 줄 높이 차이 여유

        // 행 높이 추정 (두 줄 제목 반영), 몸통 높이에 맞게 글자 축소
        const M = Math.max(...chs.map(c => c.sections.filter(s => s.include !== false).length), 1);
        const secsOf = (c) => c.sections.filter(s => s.include !== false);
        const insL = 45720;
        const insR = 36576;
        const lineIn = (pt) => (pt * 1.3) / 72;
        const titlePt = szOf(rp.rowTitle);
        const pagePt = szOf(rp.page);
        const rowNeed = (s) => Array.from({ length: M }, (_, r) => Math.max(...chs.map(c => {
            const sec = secsOf(c)[r];
            if (!sec) return 0;
            const lines = Math.min(3, Math.max(1, Math.ceil(textWidthIn(sec.title, titlePt * s, true) / ((titleW - insL - insR) / IN) - 0.02)));
            return lines * lineIn(titlePt * s) + (showPage ? lineIn(pagePt * s) : 0) + 0.09;
        })) * IN);
        let s = 1;
        let need = rowNeed(1);
        while (need.reduce((a, b) => a + b, 0) > bodyH && s > 0.66) { s -= 0.04; need = rowNeed(s); }
        if (s < 0.99) warnings.push(`절이 많은 장이 있어 표 글자를 ${Math.round((1 - s) * 100)}% 줄였습니다.`);
        if (need.reduce((a, b) => a + b, 0) > bodyH) warnings.push('절이 많아 표가 아래 영역을 조금 넘을 수 있습니다. 일부 절을 빼거나 샘플 카드 높이를 늘려 주세요.');
        // 최소 행 높이(샘플 행 간격)는 전체가 몸통 높이를 넘지 않는 범위로
        const sumWith = (f) => need.reduce((a, h) => a + Math.max(h, f), 0);
        let floor = Math.min(pitch, bodyH / M);
        while (floor > 0 && sumWith(floor) > bodyH) floor -= 0.01 * IN;
        const rowH = need.map(h => Math.round(Math.max(h, floor, 0)));
        const used = rowH.reduce((a, b) => a + b, 0);
        const fillerH = Math.max(0, Math.round(tc.body.box.h - used - 45720));

        // 표 열: [번호, 제목] × 장, 사이에 간격 열
        const grid = [];
        chs.forEach((c, i) => { grid.push(numW, titleW); if (i < n - 1) grid.push(Math.round(gap)); });
        const spacer = (bodyRow) => cellXml({ borders: bodyRow ? { l: bodyLine, r: bodyLine } : {}, mar: { l: 0, r: 0, t: 0, b: 0 } });

        // 머리 행
        let headCells = '';
        chs.forEach((c, i) => {
            const secs = secsOf(c);
            const subs = secs.reduce((m, sx) => m + (sx.subsections || []).filter(x => x.include !== false).length, 0);
            const titleScale = Math.min(1, ((colW - 2 * 0.12 * IN) / IN) / Math.max(0.01, textWidthIn(c.title, szOf(rp.title), true)));
            const paras = (tc.word ? para(rp.word, tc.word.text.trim(), 'r') : '')
                + para(rp.roman, c.roman || String(c.num ?? ''), 'l')
                + para(tuneRPr(rp.title, { scale: Math.max(0.7, titleScale) }), c.title, 'l')
                + (tc.count ? para(rp.count, countLabel(secs.length, subs, withSub), 'l') : '');
            headCells += cellXml({ paras, fillColor: headFill, span: 2, mar: { l: Math.round(0.12 * IN), r: Math.round(0.1 * IN), t: Math.round(0.06 * IN), b: Math.round(0.06 * IN) }, anchor: 't' });
            headCells += cellXml({ hMerge: true, fillColor: headFill });
            if (i < n - 1) headCells += spacer(false);
        });
        let rowsXml = `<a:tr h="${Math.round(headH)}">${headCells}</a:tr>`;

        // 절 행
        for (let r = 0; r < M; r++) {
            let cells = '';
            chs.forEach((c, i) => {
                const secs = secsOf(c);
                const sec = secs[r];
                const lastOfChapter = r === secs.length - 1;
                const bottom = sec && !lastOfChapter ? sep : null;
                const numPara = sec ? para(tuneRPr(rp.num, { scale: s, color: numColor }), String(sec.num ?? r + 1), 'ctr') : '';
                const titleParas = sec
                    ? para(tuneRPr(rp.rowTitle, { scale: s }), sec.title, 'l') + (showPage ? para(tuneRPr(rp.page, { scale: s }), sec.startLabel || '', 'l') : '')
                    : '';
                cells += cellXml({ paras: numPara, fillColor: bodyFill, borders: { l: bodyLine, b: bottom }, mar: { l: 0, r: 0 } });
                cells += cellXml({ paras: titleParas, fillColor: bodyFill, borders: { r: bodyLine, b: bottom } });
                if (i < n - 1) cells += spacer(true);
            });
            rowsXml += `<a:tr h="${rowH[r]}">${cells}</a:tr>`;
        }
        // 카드 아래 끝까지 채우는 빈 행 (몸통 테두리 마감)
        let fillCells = '';
        chs.forEach((c, i) => {
            fillCells += cellXml({ fillColor: bodyFill, borders: { l: bodyLine, b: bodyLine }, mar: { t: 0, b: 0 } });
            fillCells += cellXml({ fillColor: bodyFill, borders: { r: bodyLine, b: bodyLine }, mar: { t: 0, b: 0 } });
            if (i < n - 1) fillCells += spacer(true);
        });
        rowsXml += `<a:tr h="${Math.max(fillerH, 45720)}">${fillCells}</a:tr>`;

        const totalW = grid.reduce((a, b) => a + b, 0);
        const totalH = headH + used + Math.max(fillerH, 45720);
        const frameXml = `<p:graphicFrame xmlns:p="${P_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}">`
            + `<p:nvGraphicFramePr><p:cNvPr id="${nextShapeId(doc)}" name="전체 목차 표"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>`
            + `<p:xfrm><a:off x="${Math.round(X0)}" y="${Math.round(tc.y0)}"/><a:ext cx="${Math.round(totalW)}" cy="${Math.round(totalH)}"/></p:xfrm>`
            + `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/>`
            + `<a:tblGrid>${grid.map(w => `<a:gridCol w="${w}"/>`).join('')}</a:tblGrid>${rowsXml}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;

        // 샘플 카드 지우고 표 넣기
        cards.forEach(c => c.members.forEach(l => { const u = rowUnits({ leaves: [l] })[0]; u.parentNode?.removeChild(u); }));
        const frame = parseXml(frameXml).documentElement;
        appendShape(spTree, doc.importNode ? doc.importNode(frame, true) : frame);
        return serializeXml(doc);
    });
    return { slides, warnings };
}

// ─────────────────────────────────────────────────────────────────────────────
// 장 표시 로마숫자(Ⅰ~Ⅹ) 글꼴: 생성 슬라이드의 모든 글자에서 로마숫자만 따로 떼어 지정 글꼴 적용
// ─────────────────────────────────────────────────────────────────────────────
export const ROMAN_FONT = 'KoPub바탕체 Bold';
const ROMAN_SPLIT = /([ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+)/;
const RPR_TAIL = ['sym', 'hlinkClick', 'hlinkMouseOver', 'rtl', 'extLst'];
function setRunFont(run, face) {
    const d = run.ownerDocument;
    let rPr = firstChild(run, 'rPr');
    if (!rPr) { rPr = d.createElementNS(A_NS, 'a:rPr'); run.insertBefore(rPr, run.firstChild); }
    ['latin', 'ea', 'cs'].forEach(tag => {
        let el = firstChild(rPr, tag);
        if (!el) {
            el = d.createElementNS(A_NS, `a:${tag}`);
            const before = elementChildren(rPr).find(c => RPR_TAIL.includes(c.localName)
                || (tag === 'latin' && ['ea', 'cs'].includes(c.localName)) || (tag === 'ea' && c.localName === 'cs'));
            if (before) rPr.insertBefore(el, before); else rPr.appendChild(el);
        }
        el.setAttribute('typeface', face);
        ['panose', 'pitchFamily', 'charset'].forEach(a => el.removeAttribute(a));
    });
}
function applyRomanFont(doc, face = ROMAN_FONT) {
    descendants(doc, 'r').filter(r => r.namespaceURI === A_NS).forEach(run => {
        const tEl = firstChild(run, 't');
        const text = tEl?.textContent || '';
        if (!/[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]/.test(text)) return;
        const parts = text.split(ROMAN_SPLIT).filter(Boolean);
        parts.forEach(part => {
            const c = run.cloneNode(true);
            firstChild(c, 't').textContent = part;
            if (/^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+$/.test(part)) setRunFont(c, face);
            run.parentNode.insertBefore(c, run);
        });
        run.parentNode.removeChild(run);
    });
}

// ─────────────────────────────────────────────────────────────────────────────
// 3-3) 표형 상세 목차: 표 행(장·절·항)을 복제해 표 2개(좌·우)로 이어 채움, 넘치면 다음 장
// ─────────────────────────────────────────────────────────────────────────────
function detectTables(slideDoc, layoutXfrms, slideSize) {
    const spTree = descendants(slideDoc, 'spTree')[0];
    const leaves = collectLeaves(spTree, layoutXfrms);
    const frames = leaves.filter(l => l.el.localName === 'graphicFrame' && descendants(l.el, 'tbl').length);
    if (!frames.length) return null;
    frames.sort((a, b) => a.box.x - b.box.x || a.box.y - b.box.y);
    const tbls = frames.map(f => {
        const tbl = descendants(f.el, 'tbl')[0];
        const trs = elementChildren(tbl).filter(n => n.localName === 'tr');
        return { frame: f, tbl, trs, cells: trs.map(tr => elementChildren(tr).filter(n => n.localName === 'tc').map(tc => leafText(tc))) };
    });
    const allRows = tbls.flatMap(t => t.trs.map((tr, i) => ({ tr, cells: t.cells[i], h: parseInt(tr.getAttribute('h') || '0', 10) })));
    const ncol = Math.max(...allRows.map(r => r.cells.length));
    if (ncol < 2) return null;
    const score = (re) => Array.from({ length: ncol }, (_, c) => allRows.filter(r => re.test(r.cells[c] || '')).length);
    const numRe = /^\s*([ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]|\d{1,2}(\.\d{1,2})*)\s*[.)]?\s*$/;
    const numScores = score(numRe);
    const pageScores = score(PAGE_RE);
    const numCol = numScores.indexOf(Math.max(...numScores));
    const pageCol = Math.max(...pageScores) > 0 ? pageScores.indexOf(Math.max(...pageScores)) : -1;
    if (numScores[numCol] < 2) return null;
    const lens = Array.from({ length: ncol }, (_, c) => (c === numCol || c === pageCol ? -1 : allRows.reduce((sum, r) => sum + (r.cells[c] || '').length, 0)));
    const titleCol = lens.indexOf(Math.max(...lens));
    const kindOf = (r, i) => {
        const t = (r.cells[numCol] || '').trim();
        if (/^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]\s*[.]?$/.test(t)) return 'ch';
        if (COMPOUND_RE.test(t)) return 'sub';
        if (NUM_RE.test(t)) return 'sec';
        return i === 0 ? 'header' : null;
    };
    const templates = {};
    tbls.forEach(t => t.trs.forEach((tr, i) => {
        const r = { tr, cells: t.cells[i] };
        const kind = kindOf(r, i);
        const cont = /\(계속\)\s*$/.test(r.cells[titleCol] || '');
        if (kind && !templates[kind] && !cont) templates[kind] = tr;
    }));
    if (!templates.sec && !templates.ch) return null;
    const rowH = median(allRows.map(r => r.h).filter(Boolean)) || 0.21 * IN;
    // 표 아래 한계: 표 영역 아래쪽 첫 글상자(주석 등) 또는 슬라이드 92%
    const top = Math.min(...frames.map(f => f.box.y));
    const below = leaves.filter(l => l.hasText && l.box.y > top + 1 * IN).map(l => l.box.y);
    const limitY = Math.min(slideSize.cy * 0.92, ...below);
    const capacity = Math.max(4, Math.floor((limitY - top) / rowH));
    // 샘플의 장 (머리글 언급 치환용)
    let mention = null;
    const chRow = allRows.find(r => kindOf(r, 1) === 'ch');
    if (chRow) {
        const re = chapterMentionRe((chRow.cells[numCol] || '').trim(), (chRow.cells[titleCol] || '').replace(/\s*\(계속\)\s*$/, '').trim());
        const leaf = leaves.find(l => l.hasText && re.test(l.text));
        if (leaf) mention = { leaf, re };
    }
    return { spTree, tbls, numCol, pageCol, titleCol, templates, rowH, capacity, mention, leaves };
}

function setCell(tr, col, text) {
    const tc = elementChildren(tr).filter(n => n.localName === 'tc')[col];
    if (tc) setLeafText(tc, text);
}

export function fillTableSlides({ templateXml, layoutXfrms, slideSize, chapters, showPage = true }) {
    const probe = detectTables(parseXml(templateXml), layoutXfrms, slideSize);
    if (!probe) throw new Error('표형 상세 목차 샘플에서 목차 표를 찾지 못했습니다.');
    const warnings = [];
    const T = probe.tbls.length;
    const content = probe.capacity - (probe.templates.header ? 1 : 0);
    const { templates } = probe;

    // 항목 순서: 장 → 절 → 항 (샘플에 있는 단계만)
    const seq = [];
    chapters.forEach((c, ci) => {
        const chItem = { kind: 'ch', ci, num: c.roman || NUM_TO_ROMAN[c.num] || String(c.num ?? ''), title: c.title, page: c.startLabel };
        if (templates.ch) seq.push(chItem);
        c.sections.filter(s => s.include !== false).forEach(s => {
            const secItem = { kind: 'sec', ci, num: String(s.num), title: s.title, page: s.startLabel, ch: chItem };
            seq.push(templates.sec ? secItem : { ...secItem, kind: 'ch' });
            if (templates.sub) {
                (s.subsections || []).filter(x => x.include !== false).forEach(x => seq.push({ kind: 'sub', ci, num: `${s.num}.${x.num}`, title: x.title, page: x.startLabel, ch: chItem, sec: secItem }));
            }
        });
    });
    if (!seq.length) return { slides: [], warnings };

    const cont = (it) => ({ ...it, title: `${it.title}${CONT_SUFFIX}`, cont: true });
    const contRows = (it) => {
        if (!it || it.kind === 'ch') return [];
        const out = [];
        if (templates.ch && it.ch) out.push(cont(it.ch));
        if (it.kind === 'sub' && it.sec) out.push(cont(it.sec));
        return out;
    };
    // 한 장 채우기 (target: 표당 최대 행 수)
    const simulate = (i, target) => {
        const tables = [];
        let j = i;
        let used = 0;
        let stopped = false;
        for (let t = 0; t < T && j < seq.length && !stopped; t++) {
            const rows = contRows(seq[j]);
            const start = j;
            while (j < seq.length && rows.length < target) {
                const it = seq[j];
                if (it.kind === 'ch' && j > i) {
                    // 새 장은 이 쪽이 절반 이상 찼으면 다음 쪽에서 시작
                    if (used + rows.length > 0.5 * T * content) { stopped = true; break; }
                    // 표 아래쪽(절반 이후)에서 시작하게 되면 오른쪽 표로, 마지막 표면 다음 쪽으로
                    if (rows.length > 0.5 * target) {
                        if (t === T - 1) stopped = true;
                        break;
                    }
                }
                rows.push(it);
                j++;
            }
            if (j === start && j < seq.length && !stopped) { rows.push(seq[j]); j++; }
            // 표 끝에 장·절 제목만 남지 않게
            while (!stopped && j < seq.length && j - start > 1) {
                const last = rows[rows.length - 1];
                if (last.kind === 'ch' || (last.kind === 'sec' && seq[j].kind === 'sub')) { rows.pop(); j--; } else break;
            }
            used += rows.length;
            tables.push(rows);
        }
        return { tables, next: j, stopped };
    };
    const pages = [];
    let i = 0;
    while (i < seq.length) {
        let r = simulate(i, content);
        if (r.next >= seq.length || r.stopped) {
            // 이 장(slide)에서 끝나는 내용은 좌·우 표에 고르게
            const total = r.tables.reduce((n, t) => n + t.length, 0);
            for (let tg = Math.ceil(total / T); tg < content; tg++) {
                const b = simulate(i, tg);
                if (b.next === r.next) { r = b; break; }
            }
        }
        pages.push(r.tables);
        i = r.next;
    }

    // 장별 쪽 수 (제목의 (1/2) 표기)
    const pagesOfCh = {};
    pages.forEach((tables, pi) => {
        new Set(tables.flat().map(it => it.ci)).forEach(ci => { (pagesOfCh[ci] = pagesOfCh[ci] || []).push(pi); });
    });

    const slides = pages.map((tables, pi) => {
        const doc = parseXml(templateXml);
        const det = detectTables(doc, layoutXfrms, slideSize);
        const tplRows = {};
        Object.keys(det.templates).forEach(k => { tplRows[k] = det.templates[k].cloneNode(true); });
        det.tbls.forEach((t, ti) => {
            t.trs.forEach(tr => t.tbl.removeChild(tr));
            const rows = tables[ti] || [];
            if (!rows.length) {
                const u = rowUnits({ leaves: [t.frame] })[0];
                u.parentNode?.removeChild(u);
                return;
            }
            let h = 0;
            const add = (kind, fill) => {
                const tr = tplRows[kind].cloneNode(true);
                fill(tr);
                t.tbl.appendChild(tr);
                h += parseInt(tr.getAttribute('h') || String(det.rowH), 10);
            };
            if (tplRows.header) add('header', () => {});
            rows.forEach(it => add(it.kind, tr => {
                setCell(tr, det.numCol, it.num);
                setCell(tr, det.titleCol, it.title);
                if (det.pageCol >= 0) setCell(tr, det.pageCol, showPage ? (it.page || '') : '');
            }));
            const xf = readXfrm(t.frame.el);
            if (xf) xf.extEl.setAttribute('cy', String(h));
        });
        // 머리글 장 언급: 'Ⅰ. 사업 개요 / Ⅱ. 환경분석', 한 장이 여러 쪽이면 (1/2)
        const cis = [...new Set(tables.flat().map(it => it.ci))];
        const targetChapters = cis.length ? cis.map(ci => chapters[ci]).filter(Boolean) : chapters;
        const label = targetChapters.map((c, idx) => {
            const ci = cis[idx] ?? 0;
            const ps = pagesOfCh[ci] || [];
            const part = ps.length > 1 ? ` (${ps.indexOf(pi) + 1}/${ps.length})` : '';
            return `${c.roman ? `${c.roman}. ` : ''}${c.title}${part}`;
        }).join(' / ');

        if (det.mention) {
            const l = det.mention.leaf;
            const full = l.text;
            const m = full.match(det.mention.re);
            setLeafText(l.el, m ? full.slice(0, m.index) + label : label);
        } else if (chapters.length === 1 && det.leaves) {
            // 장별 세부목차(단일 장)인데 샘플의 장 언급 텍스트를 찾지 못한 경우: 상단 제목 상자를 찾아 장명 반영
            const titleLeaf = det.leaves.find(l => l.hasText && l.box.y < slideSize.cy * 0.25 && (l.text.includes('—') || l.text.includes('-') || TOC_MARKER_RE.test(l.text) || /상세\s*목차|세부\s*목차/.test(l.text)));
            if (titleLeaf) {
                const full = titleLeaf.text;
                const dashIdx = full.search(/[—–-]/);
                const prefix = dashIdx >= 0 ? full.slice(0, dashIdx + 1) + ' ' : '상세 목차 — ';
                setLeafText(titleLeaf.el, `${prefix}${label}`);
            }
        }
        return serializeXml(doc);
    });
    if (pages.length > 1) {
        const pageWarnPrefix = chapters.length === 1 ? `${chapters[0].roman ? `${chapters[0].roman}. ` : ''}${chapters[0].title} 목차` : '상세 목차';
        warnings.push(`${pageWarnPrefix} ${seq.length}행을 ${pages.length}장으로 나눴습니다.`);
    }
    return { slides, warnings };
}

// ─────────────────────────────────────────────────────────────────────────────
// 4) 결과 PPTX 조립: 샘플 파일(마스터·테마 유지)에 생성 슬라이드만 남김
// ─────────────────────────────────────────────────────────────────────────────
const CT_SLIDE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';

// 마무리: 바닥글 '목차-N' 쪽 표기, 주석의 '2026.11.21 … 기준' 날짜
function finalizeSlide(xml, { footer, dateText }) {
    const doc = parseXml(xml);
    descendants(doc, 'sp').forEach(sp => {
        const t = leafText(sp);
        if (!t) return;
        if (footer != null && /^\s*목\s*차\s*[-–]\s*[\dⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩA-Za-z]+([-–]\d+)?\s*$/.test(t)) setLeafText(sp, footer);
        else if (dateText && /기준/.test(t) && /\d{4}\.\s?\d{1,2}\.\s?\d{1,2}/.test(t)) {
            descendants(sp, 't').forEach(n => { n.textContent = n.textContent.replace(/\d{4}\.\s?\d{1,2}\.\s?\d{1,2}/g, dateText); });
        }
    });
    applyRomanFont(doc);
    return serializeXml(doc);
}

const guid = () => `{${'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, ch => {
    const r = Math.floor(Math.random() * 16);
    return (ch === 'x' ? r : (r & 0x3) | 0x8).toString(16).toUpperCase();
})}}`;

/**
 * 목차 PPT 생성 (한 파일): [앞부분 목차] 전체 목차 + 상세 목차 → [장별 목차] 장마다
 * @param {object} p
 * @param {File} p.sampleFile 목차 샘플 PPTX (기본형 또는 사용자 파일)
 * @param {Array} p.chapters extractTocFromDeliverables 결과(편집 반영)
 * @param {object} p.options { overviewIndex, detailIndex, chapterIndex: 샘플 슬라이드 순번(null이면 만들지 않음), showPage, overviewTitle }
 * @returns {Promise<{blob, warnings, slideCount, counts:{front, chapter}}>}
 */
export async function buildTocPptx({ sampleFile, chapters, options = {} }) {
    const { overviewIndex = null, detailIndex = null, chapterIndex = null, showPage = true, overviewTitle = '목차' } = options;
    const zip = await JSZip.loadAsync(await sampleFile.arrayBuffer());
    const slides = await presentationSlides(zip);
    const slideSize = await slideSizeOf(zip);
    const loadTpl = async (index) => {
        const path = slides[index];
        if (!path) throw new Error(`목차 샘플의 ${index + 1}번 슬라이드를 찾을 수 없습니다.`);
        const templateXml = await zip.file(path).async('string');
        const layoutXfrms = await layoutPlaceholderXfrms(zip, path);
        const relsXml = await zip.file(relsPathOf(path))?.async('string') || null;
        const info = classifySlide(parseXml(templateXml), layoutXfrms, slideSize);
        return { path, templateXml, layoutXfrms, relsXml, kind: info.kind };
    };

    // 포함 체크 반영
    const included = chapters.filter(c => c.include).map(c => ({
        ...c,
        sections: c.sections.filter(s => s.include).map(s => ({ ...s, subsections: (s.subsections || []).filter(x => x.include) }))
    }));
    if (!included.length) throw new Error('생성할 장이 없습니다.');
    const countsOf = (c) => ({ sections: c.sections.length, subsections: c.sections.reduce((n, s) => n + s.subsections.length, 0) });

    // 행 반복(list) 샘플용 항목 구성
    const listItems = (tpl, mode, c, i) => {
        const desc = describeTemplate({ ...tpl, slideSize });
        const L = desc.levels;
        const secItems = (ch, level, withChapterParent) => ch.sections.flatMap(s => [
            { level, num: s.num, parentNum: withChapterParent ? ch.num : null, title: s.title, page: s.startLabel },
            ...(level + 1 < L ? s.subsections.map(x => ({ level: level + 1, num: x.num, parentNum: s.num, title: x.title, page: x.startLabel })) : [])
        ]);
        const chapterItem = (ch, k) => ({ level: 0, isChapter: true, num: ch.num ?? (k + 1), roman: ch.roman, title: ch.title, page: ch.startLabel });
        if (mode === 'overview') return included.flatMap((ch, k) => [chapterItem(ch, k), ...(L >= 2 ? secItems(ch, 1, true) : [])]);
        return desc.topIsChapter ? [chapterItem(c, i), ...secItems(c, 1, true)] : secItems(c, 0, false);
    };

    const out = []; // { xml, tpl, section, footer }
    const warnings = [];
    const push = (label, tpl, section, r, footerOf) => {
        r.slides.forEach((xml, k) => out.push({ xml, tpl, section, footer: footerOf ? footerOf(k, r.slides.length) : null }));
        warnings.push(...r.warnings.map(w => `[${label}] ${w}`));
    };
    const FRONT = '앞부분 목차';
    let frontNo = 0;
    const frontFooter = () => `목차-${++frontNo}`;

    // ① 앞부분: 전체 목차
    if (overviewIndex != null) {
        const tpl = await loadTpl(overviewIndex);
        if (tpl.kind === 'cards') {
            push('전체 목차', tpl, FRONT, fillCardTableSlides({ ...tpl, slideSize, chapters: included, showPage }), frontFooter);
        } else if (tpl.kind === 'list') {
            push('전체 목차', tpl, FRONT, fillTemplateSlides({ ...tpl, slideSize, items: listItems(tpl, 'overview'), chapter: { roman: '', title: overviewTitle, num: null }, showPage }), frontFooter);
        } else if (tpl.kind === 'table') {
            push('전체 목차', tpl, FRONT, fillTableSlides({ ...tpl, slideSize, chapters: included.map(c => ({ ...c, sections: c.sections.map(s => ({ ...s, subsections: [] })) })), showPage }), frontFooter);
        } else throw new Error('전체 목차로 고른 샘플 슬라이드에서 목차 구조를 찾지 못했습니다.');
    }
    // ② 앞부분: 상세 목차 (절·항)
    if (detailIndex != null) {
        const tpl = await loadTpl(detailIndex);
        if (tpl.kind === 'table') {
            push('상세 목차', tpl, FRONT, fillTableSlides({ ...tpl, slideSize, chapters: included, showPage }), frontFooter);
        } else if (tpl.kind === 'list') {
            included.forEach((c, i) => {
                const items = listItems(tpl, 'chapter', c, i);
                if (items.length) push(`상세 목차 ${c.roman || ''}`.trim(), tpl, FRONT, fillTemplateSlides({ ...tpl, slideSize, items, chapter: { roman: c.roman, title: c.title, num: c.num, counts: countsOf(c) }, showPage }), frontFooter);
            });
        } else throw new Error('상세 목차로 고른 샘플 슬라이드는 표형 또는 행 반복형이어야 합니다.');
    }
    // ③ 장별 목차 (장 앞에 넣는 장 목차)
    if (chapterIndex != null) {
        const tpl = await loadTpl(chapterIndex);
        if (tpl.kind === 'cards' || !tpl.kind) throw new Error('장별 목차로 고른 샘플 슬라이드는 행 반복형 또는 표형이어야 합니다.');
        included.forEach((c, i) => {
            const label = `${c.roman || ''} ${c.title}`.trim();
            const section = `${label} 목차`;
            const footerOf = (k, n) => `목차-${c.roman || c.num || i + 1}${n > 1 ? `-${k + 1}` : ''}`;
            if (tpl.kind === 'table') {
                push(label, tpl, section, fillTableSlides({ ...tpl, slideSize, chapters: [c], showPage }), footerOf);
                return;
            }
            const items = listItems(tpl, 'chapter', c, i);
            if (!items.length) return;
            push(label, tpl, section, fillTemplateSlides({ ...tpl, slideSize, items, chapter: { roman: c.roman, title: c.title, num: c.num, counts: countsOf(c) }, showPage }), footerOf);
        });
    }
    if (out.length === 0) throw new Error('생성할 목차 슬라이드가 없습니다. 만들 목차 종류를 하나 이상 선택해 주세요.');

    const d = new Date();
    const dateText = `${d.getFullYear()}.${String(d.getMonth() + 1).padStart(2, '0')}.${String(d.getDate()).padStart(2, '0')}`;
    out.forEach(o => { o.xml = finalizeSlide(o.xml, { footer: o.footer, dateText }); });

    // 기존 슬라이드 제거
    let presXml = await zip.file('ppt/presentation.xml').async('string');
    let presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
    let ctXml = await zip.file('[Content_Types].xml').async('string');
    const presRels = await readRels(zip, 'ppt/presentation.xml');
    const slideRelIds = new Set(presRels.filter(r => /\/slide$/.test(r.type)).map(r => r.id));
    presXml = presXml.replace(/<p:sldId\b[^>]*\/>/g, (m) => {
        const rid = (m.match(/r:id="([^"]+)"/) || [])[1];
        return slideRelIds.has(rid) ? '' : m;
    });
    presRelsXml = presRelsXml.replace(/<Relationship\b[^>]*\/>/g, (m) => {
        const id = (m.match(/\bId="([^"]+)"/) || [])[1];
        return slideRelIds.has(id) ? '' : m;
    });
    Object.keys(zip.files).filter(f => /^ppt\/slides\/(_rels\/)?slide\d+\.xml(\.rels)?$/i.test(f)).forEach(f => zip.remove(f));
    ctXml = ctXml.replace(/<Override\b[^>]*PartName="\/ppt\/slides\/slide\d+\.xml"[^>]*\/>/g, '');

    // 새 슬라이드 추가 (슬라이드마다 원본 샘플 슬라이드의 관계 복사)
    let maxRid = 0;
    for (const m of presRelsXml.matchAll(/\bId="rId(\d+)"/g)) maxRid = Math.max(maxRid, parseInt(m[1], 10));
    const cleanRels = (xml) => (xml ? xml.replace(/<Relationship\b[^>]*Type="[^"]*\/(notesSlide|comments|slide)"[^>]*\/>/g, '') : null);
    let sldIdXml = '';
    const sectionIds = new Map();
    out.forEach((o, i) => {
        const n = i + 1;
        const path = `ppt/slides/slide${n}.xml`;
        zip.file(path, o.xml);
        const rels = cleanRels(o.tpl.relsXml);
        if (rels) zip.file(`ppt/slides/_rels/slide${n}.xml.rels`, rels);
        const rid = `rId${++maxRid}`;
        presRelsXml = presRelsXml.replace('</Relationships>', `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${n}.xml"/></Relationships>`);
        sldIdXml += `<p:sldId id="${256 + i}" r:id="${rid}"/>`;
        if (!sectionIds.has(o.section)) sectionIds.set(o.section, []);
        sectionIds.get(o.section).push(256 + i);
        ctXml = ctXml.replace('</Types>', `<Override PartName="/${path}" ContentType="${CT_SLIDE}"/></Types>`);
    });
    if (/<p:sldIdLst\s*\/>/.test(presXml)) presXml = presXml.replace(/<p:sldIdLst\s*\/>/, `<p:sldIdLst>${sldIdXml}</p:sldIdLst>`);
    else if (/<p:sldIdLst>/.test(presXml)) presXml = presXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst>${sldIdXml}</p:sldIdLst>`);
    else presXml = presXml.replace(/(<p:sldMasterIdLst>[\s\S]*?<\/p:sldMasterIdLst>)/, `$1<p:sldIdLst>${sldIdXml}</p:sldIdLst>`);
    // 사용자 지정 쇼·섹션 목록이 옛 슬라이드를 가리키면 손상되므로 제거
    presXml = presXml.replace(/<p:custShowLst>[\s\S]*?<\/p:custShowLst>/, '');
    presXml = presXml.replace(/<p:ext\b[^>]*uri="\{521415D9-36F7-43E2-AB2F-B90AF26B5E84\}"[^>]*>[\s\S]*?<\/p:ext>/, '');
    presXml = presXml.replace(/<p14:sectionLst\b[\s\S]*?<\/p14:sectionLst>/, '');
    // 구역: 앞부분 목차 / 장별 목차 (PowerPoint 구역 목록에서 바로 찾아 옮길 수 있게)
    const escXml = (t) => String(t).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const sectionXml = `<p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">${
        [...sectionIds.entries()].map(([name, ids]) => `<p14:section name="${escXml(name)}" id="${guid()}"><p14:sldIdLst>${ids.map(id => `<p14:sldId id="${id}"/>`).join('')}</p14:sldIdLst></p14:section>`).join('')
    }</p14:sectionLst></p:ext>`;
    presXml = presXml.replace(/<p:extLst>\s*<\/p:extLst>/, '');
    if (/<p:extLst>/.test(presXml)) presXml = presXml.replace(/<p:extLst>/, `<p:extLst>${sectionXml}`);
    else presXml = presXml.replace(/<\/p:presentation>\s*$/, `<p:extLst>${sectionXml}</p:extLst></p:presentation>`);
    zip.file('ppt/presentation.xml', presXml);
    zip.file('ppt/_rels/presentation.xml.rels', presRelsXml);

    // 연결이 끊긴 부품(옛 슬라이드의 노트·이미지·차트 등) 정리
    const reachable = new Set();
    const queue = ['_rels/.rels'];
    const visit = async (part) => {
        if (reachable.has(part)) return;
        reachable.add(part);
        const relsPath = part === '_rels/.rels' ? null : relsPathOf(part);
        const relsFile = part === '_rels/.rels' ? zip.file('_rels/.rels') : zip.file(relsPath);
        if (relsPath && relsFile) reachable.add(relsPath);
        if (!relsFile) return;
        const doc = parseXml(await relsFile.async('string'));
        descendants(doc, 'Relationship').forEach(r => {
            if (r.getAttribute('TargetMode') === 'External') return;
            const base = part === '_rels/.rels' ? '' : partDir(part);
            const full = resolvePath(base, r.getAttribute('Target') || '');
            if (zip.file(full) && !reachable.has(full)) queue.push(full);
        });
    };
    while (queue.length) await visit(queue.shift());
    const removed = [];
    Object.keys(zip.files).forEach(f => {
        if (zip.files[f].dir || f === '[Content_Types].xml' || f === '_rels/.rels') return;
        if (!f.startsWith('ppt/')) return;
        if (reachable.has(f)) return;
        // rels 파일은 소유 부품이 남아 있으면 유지
        const owner = f.replace(/_rels\/([^/]+)\.rels$/, '$1');
        if (f.endsWith('.rels') && reachable.has(owner)) return;
        zip.remove(f);
        removed.push(f);
    });
    removed.forEach(f => {
        const esc = f.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        ctXml = ctXml.replace(new RegExp(`<Override\\b[^>]*PartName="/${esc}"[^>]*/>`, 'g'), '');
    });
    zip.file('[Content_Types].xml', ctXml);

    const app = zip.file('docProps/app.xml');
    if (app) {
        const appXml = (await app.async('string')).replace(/<Slides>\d+<\/Slides>/, `<Slides>${out.length}</Slides>`);
        zip.file('docProps/app.xml', appXml);
    }

    const blob = await zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
    const front = out.filter(o => o.section === FRONT).length;
    return { blob, warnings, slideCount: out.length, counts: { front, chapter: out.length - front } };
}
