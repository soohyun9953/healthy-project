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
        // 절 구간이 없으면 목차 슬라이드의 번호 목차로 보완 (페이지 미확인)
        if (sections.length === 0) {
            (doc.numericToc || []).forEach(t => sections.push({ num: t.num, title: t.title, startLabel: '', pageRange: '', include: true, subsections: [] }));
        }
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
        offEl: off
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
            if (NUM_RE.test(t) || ROMAN_RE.test(t) || TOC_MARKER_RE.test(t) || CHAPTER_WORD_RE.test(t) || PAGE_RE.test(t)) return false;
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
    leaves.forEach(l => {
        if (titleSet.has(l)) return;
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
        if (row) row.leaves.push(l);
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
            if ((NUM_RE.test(l.text) || ROMAN_RE.test(l.text) || /^\s*\d{1,2}(\.\d{1,2})+\s*$/.test(l.text)) && lc < tc && !r.num) r.num = l;
            else if ((PAGE_RE.test(l.text) || NUM_RE.test(l.text)) && lc > tc && !r.page) r.page = l;
        });
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
        lastRowBottom
    };
}

async function slideSizeOf(zip) {
    const pres = parseXml(await zip.file('ppt/presentation.xml').async('string'));
    const sz = descendants(pres, 'sldSz')[0];
    return { cx: parseInt(sz?.getAttribute('cx') || '12192000', 10), cy: parseInt(sz?.getAttribute('cy') || '6858000', 10) };
}

/**
 * 목차 샘플 파일 분석: 목차 슬라이드 후보와 인식 결과 요약
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
        let det = null;
        try { det = detectTemplate(doc, xf, slideSize); } catch (e) { det = null; console.warn('목차 샘플 인식 실패', path, e); }
        const allText = leafText(doc);
        const tocWord = TOC_MARKER_RE.test(allText.split('\n').find(t => TOC_MARKER_RE.test(t)) || '') || /목\s*차|CONTENTS/i.test(allText);
        const rowCount = det ? det.rows.length : 0;
        results.push({
            index: i,
            path,
            preview: allText.split('\n').map(s => s.trim()).filter(Boolean).slice(0, 4).join(' · ').slice(0, 80),
            rowCount,
            columns: det?.groups.length || 0,
            levels: det?.levels || 0,
            topIsChapter: !!det?.topIsChapter,
            mode: det?.mode || null,
            hasNum: !!det?.rows.some(r => r.num) || !!(det && Object.values(det.titlePrefixes).some(Boolean)),
            hasPage: !!det?.rows.some(r => r.page),
            hasChapterRoman: !!det?.chapterRoman,
            hasChapterTitle: !!det?.chapterTitle,
            score: (det?.mode === 'token' ? 1000 : 0) + rowCount * 10 + (tocWord ? 25 : 0) + (det?.chapterRoman ? 5 : 0)
        });
    }
    const usable = results.filter(r => r.rowCount >= 2 || (r.mode === 'token' && r.rowCount >= 1));
    const best = usable.sort((a, b) => b.score - a.score)[0] || null;
    return { slideCount: slides.length, slides: results, bestIndex: best ? best.index : null };
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

    // ── 배치 계산: 열마다 위에서부터 단계별 간격으로, 넘치면 다음 열·다음 장 ──
    const groups = probe.groups;
    const maxRowH = Math.max(...probe.rows.map(r => r.title.box.h));
    const bottomLimit = Math.max(probe.lastRowBottom, slideSize.cy * 0.9) - maxRowH / 2;
    const place = (seq, f) => {
        const pos = [];
        let g = 0;
        let y = groups[0].y0;
        for (let i = 0; i < seq.length; i++) {
            if (i > 0) {
                const key = `${seq[i - 1].level}>${seq[i].level}`;
                const step = probe.pairPitch[key] ?? probe.pitch[seq[i].level];
                const next = y + step * f;
                if (next > bottomLimit) {
                    g += 1;
                    if (g >= groups.length) return null;
                    y = groups[g].y0;
                } else {
                    y = next;
                }
            }
            pos.push({ g, y });
        }
        return pos;
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

        const doc = parseXml(templateXml);
        const det = detectTemplate(doc, layoutXfrms, slideSize);
        const tokenMode = det.mode === 'token';

        // 장 슬롯
        const title = pi > 0 && chapter.title ? `${chapter.title} (계속)` : (chapter.title || '');
        const chapterTokens = { 장번호: chapter.roman || '', 장제목: title };
        if (det.chapterRoman) setLeafText(det.chapterRoman.el, chapter.roman || '', tokenMode ? chapterTokens : null);
        if (det.chapterTitle) setLeafText(det.chapterTitle.el, title, tokenMode ? chapterTokens : null);
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
            tpl[lv] = { row, units, title: locate(row.title), num: locate(row.num), page: locate(row.page), groupX: det.groups[row.group].x };
        }
        det.rows.forEach(r => rowUnits(r).forEach(u => u.parentNode && u.parentNode.removeChild(u)));

        let idCounter = nextShapeId(doc);
        pageItems.forEach((it, i) => {
            const t = tpl[it.level];
            const { g, y } = pos[i];
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
            }
        });
        return serializeXml(doc);
    });
    return { slides, warnings };
}

function isAncestor(a, b) {
    let cur = b;
    while (cur) { if (cur === a) return true; cur = cur.parentNode; }
    return false;
}

// ─────────────────────────────────────────────────────────────────────────────
// 4) 결과 PPTX 조립: 샘플 파일(마스터·테마 유지)에 생성 슬라이드만 남김
// ─────────────────────────────────────────────────────────────────────────────
const CT_SLIDE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';

/**
 * @param {object} p
 * @param {File} p.sampleFile 목차 샘플 PPTX
 * @param {number} p.slideIndex 샘플의 목차 슬라이드 순번(0부터)
 * @param {Array} p.chapters extractTocFromDeliverables 결과(편집 반영)
 * @param {object} p.options { overview: bool, perChapter: bool, showPage: bool, overviewTitle: string }
 * @returns {Promise<{blob, warnings, slideCount}>}
 */
export async function buildTocPptx({ sampleFile, slideIndex, chapters, options = {} }) {
    const { overview = true, perChapter = true, showPage = true, overviewTitle = '목차' } = options;
    const zip = await JSZip.loadAsync(await sampleFile.arrayBuffer());
    const slides = await presentationSlides(zip);
    const templatePath = slides[slideIndex];
    if (!templatePath) throw new Error('목차 샘플 슬라이드를 찾을 수 없습니다.');
    const templateXml = await zip.file(templatePath).async('string');
    const templateRelsXml = await zip.file(relsPathOf(templatePath))?.async('string') || null;
    const layoutXfrms = await layoutPlaceholderXfrms(zip, templatePath);
    const slideSize = await slideSizeOf(zip);

    const included = chapters.filter(c => c.include);
    const desc = describeTemplate({ templateXml, layoutXfrms, slideSize });
    if (!desc) throw new Error('목차 샘플에서 반복되는 목차 항목 칸을 찾지 못했습니다.');
    const L = desc.levels;
    const secItems = (c, level, withChapterParent) => c.sections.filter(s => s.include).flatMap(s => [
        { level, num: s.num, parentNum: withChapterParent ? c.num : null, title: s.title, page: s.startLabel },
        ...(level + 1 < L ? (s.subsections || []).filter(x => x.include).map(x => ({ level: level + 1, num: x.num, parentNum: s.num, title: x.title, page: x.startLabel })) : [])
    ]);
    const chapterItem = (c, i) => ({ level: 0, isChapter: true, num: c.num ?? (i + 1), roman: c.roman, title: c.title, page: c.startLabel });

    const outSlides = [];
    const warnings = [];
    const push = (label, r) => { outSlides.push(...r.slides); warnings.push(...r.warnings.map(w => `[${label}] ${w}`)); };
    if (overview && included.length) {
        // 전체 목차: 샘플 깊이만큼 장 › 절 › 항
        const items = included.flatMap((c, i) => [chapterItem(c, i), ...(L >= 2 ? secItems(c, 1, true) : [])]);
        push('전체 목차', fillTemplateSlides({ templateXml, layoutXfrms, slideSize, items, chapter: { roman: '', title: overviewTitle, num: null }, showPage }));
    }
    if (perChapter) {
        included.forEach((c, i) => {
            // 장별 목차: 샘플 최상위가 장 단계면 장 › 절(› 항), 아니면 절(› 항)
            const items = desc.topIsChapter ? [chapterItem(c, i), ...secItems(c, 1, true)] : secItems(c, 0, false);
            if (!items.some(it => it.level >= (desc.topIsChapter ? 1 : 0))) return;
            push(`${c.roman || ''} ${c.title}`.trim(), fillTemplateSlides({ templateXml, layoutXfrms, slideSize, items, chapter: { roman: c.roman, title: c.title, num: c.num }, showPage }));
        });
    }
    if (outSlides.length === 0) throw new Error('생성할 목차 항목이 없습니다.');

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

    // 새 슬라이드 추가
    let maxRid = 0;
    for (const m of presRelsXml.matchAll(/\bId="rId(\d+)"/g)) maxRid = Math.max(maxRid, parseInt(m[1], 10));
    const cleanRels = templateRelsXml
        ? templateRelsXml.replace(/<Relationship\b[^>]*Type="[^"]*\/(notesSlide|comments|slide)"[^>]*\/>/g, '')
        : null;
    let sldIdXml = '';
    outSlides.forEach((xml, i) => {
        const n = i + 1;
        const path = `ppt/slides/slide${n}.xml`;
        zip.file(path, xml);
        if (cleanRels) zip.file(`ppt/slides/_rels/slide${n}.xml.rels`, cleanRels);
        const rid = `rId${++maxRid}`;
        presRelsXml = presRelsXml.replace('</Relationships>', `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${n}.xml"/></Relationships>`);
        sldIdXml += `<p:sldId id="${256 + i}" r:id="${rid}"/>`;
        ctXml = ctXml.replace('</Types>', `<Override PartName="/${path}" ContentType="${CT_SLIDE}"/></Types>`);
    });
    if (/<p:sldIdLst\s*\/>/.test(presXml)) presXml = presXml.replace(/<p:sldIdLst\s*\/>/, `<p:sldIdLst>${sldIdXml}</p:sldIdLst>`);
    else if (/<p:sldIdLst>/.test(presXml)) presXml = presXml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst>${sldIdXml}</p:sldIdLst>`);
    else presXml = presXml.replace(/(<p:sldMasterIdLst>[\s\S]*?<\/p:sldMasterIdLst>)/, `$1<p:sldIdLst>${sldIdXml}</p:sldIdLst>`);
    // 사용자 지정 쇼·섹션 목록이 옛 슬라이드를 가리키면 손상되므로 제거
    presXml = presXml.replace(/<p:custShowLst>[\s\S]*?<\/p:custShowLst>/, '');
    presXml = presXml.replace(/<p14:sectionLst\b[\s\S]*?<\/p14:sectionLst>/, '');
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
        const appXml = (await app.async('string')).replace(/<Slides>\d+<\/Slides>/, `<Slides>${outSlides.length}</Slides>`);
        zip.file('docProps/app.xml', appXml);
    }

    const blob = await zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
    return { blob, warnings, slideCount: outSlides.length, levels: L, topIsChapter: desc.topIsChapter };
}
