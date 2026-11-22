/**
 * 목차 디자인 유지 갱신 (직접 올린 목차 PPT)
 *
 * 올린 파일의 슬라이드·도형·표·서식은 그대로 두고, 목차 글자(장 제목·절/항 번호·제목)와 시작 쪽만
 * 최종 산출물 기준으로 점검해 바꾼다. 바뀐 내용은 점검 결과(changes)로 돌려준다.
 *
 * 인식하는 목차 구조 (슬라이드마다)
 * - 장 열 표: 머리 행에 장 로마숫자가 2개 이상인 표 (장마다 [번호 칸 + 제목·시작쪽 칸] 열) → 장별로 행 칸을 채움
 * - 장 카드(도형): 로마숫자 머리 상자 + 몸통 행 → 카드별로 행을 채움
 * - 상세 목차 표: 번호|목차|시작쪽 행 표 (장·절·항 행) → 덱 전체의 표 행 칸을 순서대로 채움 ('(계속)' 행 포함)
 * - 장별 목차(행 반복 도형): 장 번호 칸이 있는 행 목록 → 그 장의 절·항으로 채움
 * 칸이 남으면 비우고(도형 행은 삭제), 모자라면 경고로 알린다. 칸 배치·크기·색은 바꾸지 않는다.
 */
import JSZip from 'jszip';
import { tocInternals as T } from './pptTocBuilder';

const {
    A_NS, CONT_SUFFIX, NUM_RE, COMPOUND_RE, COUNT_RE,
    parseXml, serializeXml, elementChildren, firstChild, descendants,
    presentationSlides, slideSizeOf, layoutPlaceholderXfrms, collectLeaves, leafText, readXfrm,
    detectTemplate, detectCards, detectTables, setLeafText, rowUnits, shiftUnit, nextShapeId, pathFrom, followPath, isAncestor,
    numText, titlePrefixOf, countLabel, fitLeader, finalizeSlide, fillTableSlides, relsPathOf
} = T;

const paraText = (p) => descendants(p, 't').map(t => t.textContent || '').join('');
const parasOf = (el) => {
    const body = descendants(el, 'txBody')[0];
    return body ? elementChildren(body).filter(n => n.localName === 'p') : [];
};
const cellsOf = (tr) => elementChildren(tr).filter(n => n.localName === 'tc');
const ROMAN_ONLY = /^\s*([ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ])\s*[.．]?\s*$/;

// 문단 글자 바꾸기 (첫 글자 서식 유지, 빈 글자면 문단 끝 서식만 남김)
function setParaText(p, text) {
    const d = p.ownerDocument;
    const runs = elementChildren(p).filter(n => ['r', 'fld', 'br'].includes(n.localName));
    let run = runs.find(n => n.localName === 'r') || null;
    const end = firstChild(p, 'endParaRPr');
    if (!run) {
        if (!text) return;
        run = d.createElementNS(A_NS, 'a:r');
        if (end) {
            const rPr = d.createElementNS(A_NS, 'a:rPr');
            Array.from(end.attributes).forEach(a => rPr.setAttribute(a.name, a.value));
            elementChildren(end).forEach(c => rPr.appendChild(c.cloneNode(true)));
            run.appendChild(rPr);
        }
        run.appendChild(d.createElementNS(A_NS, 'a:t'));
        if (end) p.insertBefore(run, end); else p.appendChild(run);
    }
    runs.forEach(n => { if (n !== run) p.removeChild(n); });
    let t = firstChild(run, 't');
    if (!t) { t = d.createElementNS(A_NS, 'a:t'); run.appendChild(t); }
    t.textContent = text;
    if (!text) {
        const rPr = firstChild(run, 'rPr');
        if (rPr && !end) {
            const e = d.createElementNS(A_NS, 'a:endParaRPr');
            Array.from(rPr.attributes).forEach(a => { if (a.name !== 'dirty' && a.name !== 'err') e.setAttribute(a.name, a.value); });
            elementChildren(rPr).forEach(c => e.appendChild(c.cloneNode(true)));
            p.appendChild(e);
        }
        p.removeChild(run);
    }
}

// 장 찾기: 로마숫자 → 산출물 장
const chapterByRoman = (chapters, roman) => chapters.find(c => (c.roman || '') === roman) || null;
const codeOf = (c, s, x) => [c?.roman, s?.num, x?.num].filter(v => v != null && v !== '').join('.');

function makeLogger(changes, slide, area) {
    return (where, field, before, after) => {
        const b = String(before ?? '').trim();
        const a = String(after ?? '').trim();
        if (b !== a) changes.push({ slide, area, where, field, before: b, after: a });
    };
}

// ─────────────────────────────────────────────────────────────────────────────
// 1) 장 열 표 (전체 목차): 머리 행에 로마숫자 칸이 2개 이상
// ─────────────────────────────────────────────────────────────────────────────
function detectChapterTables(doc, layoutXfrms) {
    const leaves = collectLeaves(descendants(doc, 'spTree')[0], layoutXfrms);
    const out = [];
    leaves.filter(l => l.el.localName === 'graphicFrame' && descendants(l.el, 'tbl').length).forEach(frame => {
        const tbl = descendants(frame.el, 'tbl')[0];
        const trs = elementChildren(tbl).filter(n => n.localName === 'tr');
        for (let r = 0; r < Math.min(3, trs.length); r++) {
            const groups = [];
            cellsOf(trs[r]).forEach((tc, col) => {
                if (tc.getAttribute('hMerge') === '1') return;
                const romanP = parasOf(tc).find(p => ROMAN_ONLY.test(paraText(p)));
                if (romanP) groups.push({ col, span: parseInt(tc.getAttribute('gridSpan') || '1', 10), tc, roman: paraText(romanP).trim().replace(/[.．]$/, ''), romanP });
            });
            if (groups.length >= 2) { out.push({ frame, tbl, trs, headerIdx: r, groups }); break; }
        }
    });
    return out;
}

function fillChapterTable(t, chapters, { showPage, log, warnings }) {
    const body = t.trs.slice(t.headerIdx + 1);
    // 맨 아래 빈 행(카드 높이 채움)은 칸이 아님
    const groupCells = (tr, g) => cellsOf(tr).slice(g.col, g.col + g.span);
    const isEmptyRow = (tr) => t.groups.every(g => groupCells(tr, g).every(tc => !leafText(tc)));
    while (body.length && isEmptyRow(body[body.length - 1]) && body.length > 1) {
        const last = body[body.length - 1];
        const h = parseInt(last.getAttribute('h') || '0', 10);
        const prevH = parseInt(body[body.length - 2].getAttribute('h') || '0', 10);
        if (h > prevH * 1.6 || h < prevH * 0.5) body.pop(); else break;
    }
    const withSub = t.groups.some(g => parasOf(g.tc).some(p => COUNT_RE.test(paraText(p)) && /항/.test(paraText(p))));
    t.groups.forEach(g => {
        const ch = chapterByRoman(chapters, g.roman);
        if (!ch) { warnings.push(`표의 ${g.roman}장 칸에 해당하는 장이 최종 산출물에 없어 그대로 두었습니다.`); return; }
        // 머리 칸: 장 제목, 'n개 절 · m개 항'
        const ps = parasOf(g.tc);
        const ri = ps.indexOf(g.romanP);
        const titleP = ps.slice(ri + 1).find(p => paraText(p).trim() && !COUNT_RE.test(paraText(p)) && !/^\s*CHAPTER/i.test(paraText(p)));
        if (titleP) { log(ch.roman, '장 제목', paraText(titleP), ch.title); setParaText(titleP, ch.title); }
        const countP = ps.find(p => COUNT_RE.test(paraText(p)));
        const secs = ch.sections;
        if (countP) {
            const label = countLabel(secs.length, secs.reduce((n, s) => n + s.subsections.length, 0), withSub);
            log(ch.roman, '절·항 수', paraText(countP), label);
            setParaText(countP, label);
        }
        body.forEach((tr, k) => {
            const cells = groupCells(tr, g);
            if (!cells.length) return;
            const sec = secs[k];
            const numCell = g.span >= 2 ? cells[0] : null;
            const titleCell = cells[cells.length - 1];
            const tp = parasOf(titleCell);
            const where = sec ? codeOf(ch, sec) : `${ch.roman} ${k + 1}번째 칸`;
            if (numCell) {
                const np = parasOf(numCell)[0];
                if (np) { log(where, '번호', paraText(np), sec ? sec.num : ''); setParaText(np, sec ? String(sec.num) : ''); }
            }
            if (tp[0]) {
                const old = paraText(tp[0]);
                const prefix = !numCell ? titlePrefixOf(old) : null;
                const text = sec ? (prefix ? `${sec.num}${prefix.sep}${sec.title}` : sec.title) : '';
                log(where, '제목', old, text);
                setParaText(tp[0], text);
            }
            if (tp[1]) {
                const old = paraText(tp[1]);
                const text = sec && showPage ? (sec.startLabel || '') : '';
                log(where, '시작쪽', old, text);
                setParaText(tp[1], text);
            }
            tp.slice(2).forEach(p => setParaText(p, ''));
        });
        if (secs.length > body.length) warnings.push(`${ch.roman}장: 표 칸이 ${body.length}개라 절 ${secs.length - body.length}개(${secs.slice(body.length).map(s => s.num).join(', ')})를 넣지 못했습니다. 디자인 파일에 행을 늘려 주세요.`);
    });
}

// ─────────────────────────────────────────────────────────────────────────────
// 2) 장 카드(도형)
// ─────────────────────────────────────────────────────────────────────────────
function fillCards(det, chapters, { showPage, log, warnings }) {
    const removeLeaf = (l) => { const u = rowUnits({ leaves: [l] })[0]; u.parentNode?.removeChild(u); };
    const withSub = det.cards.some(c => c.count && /항/.test(c.count.text));
    det.cards.forEach(card => {
        const roman = card.roman.text.trim();
        const ch = chapterByRoman(chapters, roman);
        if (!ch) { warnings.push(`카드의 ${roman}장에 해당하는 장이 최종 산출물에 없어 그대로 두었습니다.`); return; }
        if (card.title) { log(roman, '장 제목', card.title.text, ch.title); setLeafText(card.title.el, ch.title); }
        if (card.count) {
            const label = countLabel(ch.sections.length, ch.sections.reduce((n, s) => n + s.subsections.length, 0), withSub);
            log(roman, '절·항 수', card.count.text, label);
            setLeafText(card.count.el, label);
        }
        card.rows.forEach((row, k) => {
            const sec = ch.sections[k];
            if (!sec) {
                log(`${roman} ${k + 1}번째 칸`, '제목', row.title.text, '');
                row.leaves.concat([row.num]).forEach(removeLeaf);
                return;
            }
            const where = codeOf(ch, sec);
            log(where, '번호', row.num.text, sec.num); setLeafText(row.num.el, String(sec.num));
            log(where, '제목', row.title.text, sec.title); setLeafText(row.title.el, sec.title);
            if (row.page) { const p = showPage ? (sec.startLabel || '') : ''; log(where, '시작쪽', row.page.text, p); setLeafText(row.page.el, p); }
        });
        if (ch.sections.length > card.rows.length) warnings.push(`${roman}장 카드: 칸이 ${card.rows.length}개라 절 ${ch.sections.length - card.rows.length}개를 넣지 못했습니다.`);
    });
}

// ─────────────────────────────────────────────────────────────────────────────
// 3) 장별 목차 (행 반복 도형): 장 번호 칸의 로마숫자로 장을 정함
// ─────────────────────────────────────────────────────────────────────────────
function fillListSlide(doc, det, chapters, { showPage, log, warnings }) {
    const roman = (det.chapterRoman?.text || '').trim().replace(/[.．]$/, '');
    const ch = chapterByRoman(chapters, roman);
    if (!ch) return false;
    const L = det.levels;
    const items = [];
    const base = det.topIsChapter ? 1 : 0;
    if (det.topIsChapter) items.push({ level: 0, isChapter: true, num: ch.num, roman: ch.roman, title: ch.title, page: ch.startLabel, code: ch.roman });
    ch.sections.forEach(s => {
        items.push({ level: base, num: s.num, parentNum: det.topIsChapter ? ch.num : null, title: s.title, page: s.startLabel, code: codeOf(ch, s) });
        if (base + 1 < L) s.subsections.forEach(x => items.push({ level: base + 1, num: x.num, parentNum: s.num, title: x.title, page: x.startLabel, code: codeOf(ch, s, x) }));
    });

    // 단계별 견본(삭제·교체 전에 복사)
    const tpl = {};
    for (let lv = 0; lv < L; lv++) {
        const row = det.templates[lv];
        if (!row) continue;
        const units = rowUnits(row);
        const locate = (leaf) => {
            if (!leaf) return null;
            const ui = units.findIndex(u => u === leaf.el || isAncestor(u, leaf.el));
            return ui < 0 ? null : { ui, path: units[ui] === leaf.el ? [] : pathFrom(units[ui], leaf.el) };
        };
        tpl[lv] = { row, copies: units.map(u => ({ node: u.cloneNode(true), parent: u.parentNode })), title: locate(row.title), num: locate(row.num), page: locate(row.page), leader: locate(row.leader), groupX: det.groups[row.group].x };
    }
    let idSeq = nextShapeId(doc);
    const writeTexts = (it, els, row, dx) => {
        const numStr = numText(it, det.numStyles[it.level]);
        const prefix = det.titlePrefixes[it.level];
        const titleText = prefix ? `${numText(it, { base: prefix.base, suffix: '' })}${prefix.sep}${it.title}` : it.title;
        if (els.title) setLeafText(els.title, titleText);
        if (els.num) setLeafText(els.num, numStr);
        if (els.page) setLeafText(els.page, showPage ? (it.page || '') : '');
        if (els.leader) fitLeader(els.leader, row, titleText, dx);
        return { numStr, titleText };
    };
    const slots = [...det.rows];
    slots.forEach((slot, k) => {
        const it = items[k];
        const oldTitle = slot.title.text;
        const oldPage = slot.page?.text || '';
        if (!it) {
            log(`${roman} ${k + 1}번째 칸`, '제목', oldTitle, '');
            rowUnits(slot).forEach(u => u.parentNode?.removeChild(u));
            return;
        }
        if (slot.level === it.level || !tpl[it.level]) {
            const r = writeTexts(it, { title: slot.title.el, num: slot.num?.el, page: slot.page?.el, leader: slot.leader?.el }, slot, 0);
            log(it.code, '제목', oldTitle, r.titleText);
            if (slot.page) log(it.code, '시작쪽', oldPage, showPage ? it.page : '');
            return;
        }
        // 칸 단계가 다르면 그 자리에 해당 단계 견본 행을 넣음 (위치는 그대로)
        const t = tpl[it.level];
        rowUnits(slot).forEach(u => u.parentNode?.removeChild(u));
        const dx = det.groups[slot.group].x - t.groupX;
        const dy = slot.center - t.row.center;
        const clones = t.copies.map(c => {
            const n = c.node.cloneNode(true);
            descendants(n, 'cNvPr').forEach(x => x.setAttribute('id', String(idSeq++)));
            c.parent.appendChild(n);
            shiftUnit(n, dx, dy);
            return n;
        });
        const leafOf = (loc) => (loc ? (loc.path.length ? followPath(clones[loc.ui], loc.path) : clones[loc.ui]) : null);
        const r = writeTexts(it, { title: leafOf(t.title), num: leafOf(t.num), page: leafOf(t.page), leader: leafOf(t.leader) }, t.row, dx);
        log(it.code, '제목', oldTitle, r.titleText);
        if (t.page) log(it.code, '시작쪽', oldPage, showPage ? it.page : '');
    });
    if (items.length > slots.length) warnings.push(`${roman}장 장별 목차: 칸이 ${slots.length}개라 ${items.length - slots.length}개 항목을 넣지 못했습니다. 디자인 파일에 행을 늘려 주세요.`);
    // 'n개 절', 머리글 장 언급
    if (det.chapterCount) {
        const label = countLabel(ch.sections.length, ch.sections.reduce((n, s) => n + s.subsections.length, 0), /항/.test(det.chapterCount.text));
        log(roman, '절·항 수', det.chapterCount.text, label);
        setLeafText(det.chapterCount.el, label);
    }
    if (det.chapterTitle) { log(roman, '장 제목', det.chapterTitle.text, ch.title); setLeafText(det.chapterTitle.el, ch.title); }
    det.mentions.forEach(m => {
        const next = m.text.replace(det.mentionRe, `${ch.roman}. ${ch.title}`);
        log(roman, '머리글', m.text, next);
        setLeafText(m.el, next);
    });
    return true;
}

// ─────────────────────────────────────────────────────────────────────────────
// 4) 상세 목차 표 (덱 전체): 번호|목차|시작쪽 행을 순서대로 채움
// ─────────────────────────────────────────────────────────────────────────────
function kindOfRow(tr, det, idx) {
    const t = (leafText(cellsOf(tr)[det.numCol] || tr) || '').trim();
    if (/^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]\s*[.]?$/.test(t)) return 'ch';
    if (COMPOUND_RE.test(t)) return 'sub';
    if (NUM_RE.test(t)) return 'sec';
    return idx === 0 ? 'header' : (t ? null : 'empty');
}

function fillDetailTables(entries, chapters, { showPage, logFor, warnings }) {
    // 단계별 견본 행 (어느 슬라이드든 처음 나온 것)
    const templates = {};
    entries.forEach(e => e.det.tbls.forEach(t => t.trs.forEach((tr, i) => {
        const k = kindOfRow(tr, e.det, i);
        const cont = /\(계속\)\s*$/.test(leafText(cellsOf(tr)[e.det.titleCol] || tr));
        if (k && k !== 'header' && k !== 'empty' && !templates[k] && !cont) templates[k] = tr.cloneNode(true);
    })));
    // 장별 항목 (장 → 절 → 항)
    const chItems = chapters.map((c, ci) => {
        const list = [];
        const chItem = { kind: 'ch', ci, num: c.roman || String(c.num ?? ''), title: c.title, page: c.startLabel, code: c.roman };
        if (templates.ch) list.push(chItem);
        c.sections.forEach(s => {
            const secItem = { kind: 'sec', ci, num: String(s.num), title: s.title, page: s.startLabel, ch: chItem, code: codeOf(c, s) };
            list.push(secItem);
            if (templates.sub) s.subsections.forEach(x => list.push({ kind: 'sub', ci, num: `${s.num}.${x.num}`, title: x.title, page: x.startLabel, ch: chItem, sec: secItem, code: codeOf(c, s, x) }));
        });
        return list;
    });
    const cont = (it) => ({ ...it, title: `${it.title}${CONT_SUFFIX}`, cont: true });
    const contRows = (it) => {
        if (!it || it.kind === 'ch') return [];
        const out = [];
        if (templates.ch && it.ch) out.push(cont(it.ch));
        if (it.kind === 'sub' && it.sec) out.push(cont(it.sec));
        return out;
    };
    // 디자인 슬라이드마다 원래 담긴 장(장 행·'(계속)' 행의 로마숫자) → 그 장들의 내용을 그 슬라이드에 채움
    const romansOf = (e) => {
        const out = [];
        e.det.tbls.forEach(t => t.trs.forEach((tr, idx) => {
            if (kindOfRow(tr, e.det, idx) !== 'ch') return;
            const r = leafText(cellsOf(tr)[e.det.numCol] || tr).trim().replace(/[.．]$/, '');
            if (!out.includes(r)) out.push(r);
        }));
        return out;
    };
    const slideCis = entries.map(e => romansOf(e).map(r => chapters.findIndex(c => c.roman === r)).filter(ci => ci >= 0));
    const placedAny = new Set(slideCis.flat());
    chapters.forEach((c, ci) => { if (!placedAny.has(ci)) warnings.push(`상세 목차 표에 ${c.roman || c.num}장 자리가 없어 넣지 못했습니다. 디자인 파일에 그 장의 표 슬라이드를 추가해 주세요.`); });
    const pointer = chapters.map(() => 0);
    const slideChapters = new Map();
    // 장 시작 행('Ⅲ | 현황분석', '(계속)' 아님)이 있는 슬라이드는 그 장을 처음부터 다시 채움
    // (예: 앞부분 상세 목차 2~6번과 장별 세부목차 7~12번이 같은 장을 각각 담는 경우)
    const startsOf = (e) => {
        const out = [];
        e.det.tbls.forEach(t => t.trs.forEach((tr, idx) => {
            if (kindOfRow(tr, e.det, idx) !== 'ch') return;
            if (/\(계속\)\s*$/.test(leafText(cellsOf(tr)[e.det.titleCol] || tr))) return;
            const ci = chapters.findIndex(c => c.roman === leafText(cellsOf(tr)[e.det.numCol] || tr).trim().replace(/[.．]$/, ''));
            if (ci >= 0) out.push(ci);
        }));
        return out;
    };
    entries.forEach((e, ei) => {
        const log = logFor(e.slideNo);
        const covered = new Set();
        startsOf(e).forEach(ci => { pointer[ci] = 0; });
        const stream = slideCis[ei].flatMap(ci => chItems[ci].slice(pointer[ci]));
        let i = 0;
        e.det.tbls.forEach(t => {
            const slots = t.trs.filter((tr, idx) => kindOfRow(tr, e.det, idx) !== 'header');
            const rows = i < stream.length ? contRows(stream[i]) : [];
            const start = i;
            while (i < stream.length && rows.length < slots.length) rows.push(stream[i++]);
            // 표 끝에 장·절 제목만 남지 않게
            while (i < stream.length && i - start > 1) {
                const last = rows[rows.length - 1];
                if (last.kind === 'ch' || (last.kind === 'sec' && stream[i].kind === 'sub')) { rows.pop(); i--; } else break;
            }
            let h = 0;
            t.trs.forEach((tr, idx) => { if (kindOfRow(tr, e.det, idx) === 'header') h += parseInt(tr.getAttribute('h') || '0', 10); });
            slots.forEach((tr, k) => {
                const it = rows[k];
                const cells = cellsOf(tr);
                const old = { num: leafText(cells[e.det.numCol] || tr), title: leafText(cells[e.det.titleCol] || tr), page: e.det.pageCol >= 0 ? leafText(cells[e.det.pageCol] || tr) : '' };
                let target = tr;
                // 빈 칸은 가장 단순한 행 서식(항 → 절)으로 바꿔 남김
                const plain = templates.sub ? 'sub' : 'sec';
                const want = it ? it.kind : plain;
                if (it) covered.add(it.ci);
                {
                    const k0 = kindOfRow(tr, e.det, 1);
                    if (k0 !== want && templates[want]) {
                        const fresh = templates[want].cloneNode(true);
                        const imported = tr.ownerDocument.importNode ? tr.ownerDocument.importNode(fresh, true) : fresh;
                        tr.parentNode.insertBefore(imported, tr);
                        tr.parentNode.removeChild(tr);
                        target = imported;
                    }
                }
                const c2 = cellsOf(target);
                const set = (col, text) => { const tc = c2[col]; if (tc) parasOf(tc).forEach((p, pi) => setParaText(p, pi === 0 ? text : '')); };
                const where = it ? `${it.code}${it.cont ? ' (계속)' : ''}` : '빈 칸';
                set(e.det.numCol, it ? it.num : '');
                set(e.det.titleCol, it ? it.title : '');
                if (e.det.pageCol >= 0) set(e.det.pageCol, it && showPage ? (it.page || '') : '');
                log(where, '번호', old.num, it ? it.num : '');
                log(where, '제목', old.title, it ? it.title : '');
                if (e.det.pageCol >= 0) log(where, '시작쪽', old.page, it && showPage ? it.page : '');
                h += parseInt(target.getAttribute('h') || '0', 10);
            });
            const xf = readXfrm(t.frame.el);
            if (xf && h) xf.extEl.setAttribute('cy', String(h));
        });
        // 이 슬라이드에서 채운 만큼 장별 진행
        stream.slice(0, i).forEach(it => { pointer[it.ci] += 1; });
        slideChapters.set(e, [...covered]);
        if (!covered.size) warnings.push(`${e.slideNo}번 슬라이드(상세 목차 표): 넣을 내용이 없어 빈 표로 남았습니다. 필요 없으면 삭제해 주세요.`);
    });
    chapters.forEach((c, ci) => {
        const left = chItems[ci].length - pointer[ci];
        if (left > 0 && placedAny.has(ci)) warnings.push(`${c.roman || c.num}장 상세 목차: 표 칸이 모자라 ${left}행(${chItems[ci][pointer[ci]].code}부터)을 넣지 못했습니다. 디자인 파일에 행이나 슬라이드를 늘려 주세요.`);
    });
    // 머리글 장 언급 'Ⅰ. 사업 개요 / Ⅱ. 환경분석 (1/2)'
    const pagesOf = {};
    entries.forEach((e, ei) => (slideChapters.get(e) || []).forEach(ci => { (pagesOf[ci] = pagesOf[ci] || []).push(ei); }));
    entries.forEach((e, ei) => {
        if (!e.det.mention) return;
        const cis = slideChapters.get(e) || [];
        if (!cis.length) return;
        const label = cis.map(ci => {
            const c = chapters[ci];
            const ps = pagesOf[ci] || [];
            return `${c.roman ? `${c.roman}. ` : ''}${c.title}${ps.length > 1 ? ` (${ps.indexOf(ei) + 1}/${ps.length})` : ''}`;
        }).join(' / ');
        const l = e.det.mention.leaf;
        // 제목의 첫 장 표기('Ⅰ.')부터 끝까지를 새 장 목록으로 바꿈
        const at = l.text.search(/[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]\s*[.．]/);
        const next = at >= 0 ? l.text.slice(0, at) + label : label;
        logFor(e.slideNo)('머리글', '장', l.text, next);
        setLeafText(l.el, next);
    });
}

// ─────────────────────────────────────────────────────────────────────────────
// 5) 빈 상세 목차 표 (머리 행 '번호|목차|시작쪽' + 빈 행만 있는 디자인)
//    → 슬라이드 제목의 장('상세 목차 — Ⅰ. 사업 개요')으로 채우고, 넘치면 같은 디자인 슬라이드를 바로 뒤에 추가
// ─────────────────────────────────────────────────────────────────────────────
const HEAD_NUM = /^\s*(번\s*호|No\.?|순\s*번)\s*$/i;
const HEAD_TITLE = /목\s*차|제\s*목|내\s*용|항\s*목/;
const HEAD_PAGE = /시작|쪽|페이지|page/i;
const MENTION_ROMANS = /([ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ])\s*[.．]\s*([^/(（\n]+)/g;

function detectBlankDetailTable(doc, layoutXfrms, slideSize) {
    const leaves = collectLeaves(descendants(doc, 'spTree')[0], layoutXfrms);
    const tbls = leaves.filter(l => l.el.localName === 'graphicFrame' && descendants(l.el, 'tbl').length).map(frame => {
        const tbl = descendants(frame.el, 'tbl')[0];
        const trs = elementChildren(tbl).filter(n => n.localName === 'tr');
        return { frame, tbl, trs, head: trs[0] ? cellsOf(trs[0]).map(tc => leafText(tc).trim()) : [] };
    }).filter(t => t.trs.length >= 2 && t.head.some(h => HEAD_TITLE.test(h)) && t.trs.slice(1).every(tr => cellsOf(tr).every(tc => !leafText(tc).trim())));
    if (!tbls.length) return null;
    tbls.sort((a, b) => a.frame.box.x - b.frame.box.x);
    const head = tbls[0].head;
    const numCol = head.findIndex(h => HEAD_NUM.test(h));
    const titleCol = head.findIndex(h => HEAD_TITLE.test(h));
    const pageCol = head.findIndex((h, i) => i !== titleCol && HEAD_PAGE.test(h));
    if (titleCol < 0 || numCol < 0) return null;
    const mentionLeaf = leaves.find(l => l.hasText && l.box.y < slideSize.cy * 0.2 && /[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]\s*[.．]\s*\S/.test(l.text)) || null;
    const mentions = mentionLeaf ? [...mentionLeaf.text.matchAll(MENTION_ROMANS)].map(m => ({ roman: m[1], title: m[2].trim() })) : [];
    return { tbls, numCol, titleCol, pageCol, ncol: head.length, mentionLeaf, mentions };
}

// 표 행 견본 찾기: 같은 파일의 채워진 상세 목차 표 → 기본형
async function loadRowTemplates(files) {
    for (const f of files) {
        if (!f) continue;
        const zip = f.zip || await JSZip.loadAsync(await f.file.arrayBuffer());
        const slides = await presentationSlides(zip);
        const size = await slideSizeOf(zip);
        for (const path of slides) {
            const doc = parseXml(await zip.file(path).async('string'));
            const xf = await layoutPlaceholderXfrms(zip, path);
            let det = null;
            try { det = detectTables(doc, xf, size); } catch { det = null; }
            if (det && det.templates.sec && det.pageCol >= 0) {
                const ncol = cellsOf(det.templates.sec).length;
                return { templates: det.templates, numCol: det.numCol, titleCol: det.titleCol, pageCol: det.pageCol, ncol };
            }
        }
    }
    return null;
}

// 빈 표에 장·절·항 견본 행을 넣어 '채울 수 있는' 견본 슬라이드로 만듦
function synthesizeTemplate(doc, blank, rowTpl, mention) {
    const fits = rowTpl && rowTpl.ncol === blank.ncol && rowTpl.numCol === blank.numCol && rowTpl.titleCol === blank.titleCol && rowTpl.pageCol === blank.pageCol;
    const roman = mention?.roman || 'Ⅰ';
    const title = mention?.title || '장 제목';
    blank.tbls.forEach(t => {
        const blankRow = t.trs[1].cloneNode(true);
        t.trs.slice(1).forEach(tr => t.tbl.removeChild(tr));
        const mk = (kind, num, text, page) => {
            const src = fits && rowTpl.templates[kind] ? rowTpl.templates[kind] : blankRow;
            const tr = doc.importNode ? doc.importNode(src, true) : src.cloneNode(true);
            const cells = cellsOf(tr);
            const set = (col, v) => { if (col >= 0 && cells[col]) parasOf(cells[col]).forEach((p, i) => setParaText(p, i === 0 ? v : '')); };
            set(blank.numCol, num);
            set(blank.titleCol, text);
            set(blank.pageCol, page);
            t.tbl.appendChild(tr);
        };
        if (!fits || rowTpl.templates.ch) mk('ch', roman, title, `${roman}-1`);
        mk('sec', '1', '절 제목 견본', `${roman}-4`);
        if (!fits || rowTpl.templates.sub) mk('sub', '1.1', '항 제목 견본', `${roman}-5`);
    });
    return { usedDesignRows: !fits };
}

// 슬라이드를 기준 슬라이드 바로 뒤에 추가 (관계·콘텐츠 형식·구역 포함)
async function insertSlidesAfter(zip, basePath, xmls) {
    if (!xmls.length) return [];
    let pres = await zip.file('ppt/presentation.xml').async('string');
    let rels = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
    let ct = await zip.file('[Content_Types].xml').async('string');
    const baseRels = (await zip.file(relsPathOf(basePath))?.async('string')) || null;
    const cleanRels = baseRels ? baseRels.replace(/<Relationship\b[^>]*Type="[^"]*\/(notesSlide|comments)"[^>]*\/>/g, '') : null;
    const baseTarget = basePath.replace(/^ppt\//, '');
    const baseRid = (rels.match(new RegExp(`<Relationship\\b[^>]*Id="([^"]+)"[^>]*Target="/?(?:ppt/)?${baseTarget.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"`)) || rels.match(new RegExp(`<Relationship\\b[^>]*Target="/?(?:ppt/)?${baseTarget.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*Id="([^"]+)"`)) || [])[1];
    const baseTag = baseRid ? (pres.match(new RegExp(`<p:sldId\\b[^>]*r:id="${baseRid}"[^>]*/>`)) || [])[0] : null;
    if (!baseTag) throw new Error('상세 목차 슬라이드를 늘리는 중 원본 슬라이드 정보를 찾지 못했습니다.');
    const baseId = (baseTag.match(/\bid="(\d+)"/) || [])[1];
    let maxNum = Math.max(0, ...Object.keys(zip.files).map(f => (f.match(/^ppt\/slides\/slide(\d+)\.xml$/) || [])[1]).filter(Boolean).map(Number));
    let maxRid = Math.max(0, ...[...rels.matchAll(/\bId="rId(\d+)"/g)].map(m => parseInt(m[1], 10)));
    let maxId = Math.max(255, ...[...pres.matchAll(/<p:sldId\b[^>]*\bid="(\d+)"/g)].map(m => parseInt(m[1], 10)));
    const added = [];
    let newTags = '';
    let newSecIds = '';
    xmls.forEach(xml => {
        const path = `ppt/slides/slide${++maxNum}.xml`;
        zip.file(path, xml);
        if (cleanRels) zip.file(relsPathOf(path), cleanRels);
        const rid = `rId${++maxRid}`;
        const id = ++maxId;
        rels = rels.replace('</Relationships>', `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${maxNum}.xml"/></Relationships>`);
        ct = ct.replace('</Types>', `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`);
        newTags += `<p:sldId id="${id}" r:id="${rid}"/>`;
        newSecIds += `<p14:sldId id="${id}"/>`;
        added.push(path);
    });
    pres = pres.replace(baseTag, baseTag + newTags);
    pres = pres.replace(new RegExp(`(<p14:sldId id="${baseId}"\\s*/>)`), `$1${newSecIds}`);
    zip.file('ppt/presentation.xml', pres);
    zip.file('ppt/_rels/presentation.xml.rels', rels);
    zip.file('[Content_Types].xml', ct);
    const app = zip.file('docProps/app.xml');
    if (app) {
        const total = (pres.match(/<p:sldId\b/g) || []).length;
        zip.file('docProps/app.xml', (await app.async('string')).replace(/<Slides>\d+<\/Slides>/, `<Slides>${total}</Slides>`));
    }
    return added;
}

// ─────────────────────────────────────────────────────────────────────────────
// 전체: 올린 파일의 모든 슬라이드에서 목차를 찾아 내용만 갱신
// ─────────────────────────────────────────────────────────────────────────────
/**
 * @param {object} p
 * @param {File} p.file 직접 올린 목차 PPTX (디자인 유지)
 * @param {Array} p.chapters extractTocFromDeliverables 결과(편집·포함 반영)
 * @param {object} p.options { showPage }
 * @returns {Promise<{blob, changes:Array<{slide,area,where,field,before,after}>, warnings, updatedSlides:number[]}>}
 */
export async function updateTocDeck({ file, chapters, options = {} }) {
    const { showPage = true, fallbackFile = null } = options;
    const included = chapters.filter(c => c.include).map(c => ({
        ...c,
        sections: c.sections.filter(s => s.include).map(s => ({ ...s, subsections: (s.subsections || []).filter(x => x.include) }))
    }));
    if (!included.length) throw new Error('갱신할 장이 없습니다.');
    const zip = await JSZip.loadAsync(await file.arrayBuffer());
    const slides = await presentationSlides(zip);
    const slideSize = await slideSizeOf(zip);
    const changes = [];
    const warnings = [];
    const touched = new Map(); // path → doc
    const detailEntries = [];
    const blankEntries = [];

    for (let si = 0; si < slides.length; si++) {
        const path = slides[si];
        const slideNo = si + 1;
        const doc = parseXml(await zip.file(path).async('string'));
        const xf = await layoutPlaceholderXfrms(zip, path);
        const ctx = (area) => ({ showPage, log: makeLogger(changes, slideNo, area), warnings });

        const colTables = detectChapterTables(doc, xf);
        if (colTables.length) {
            colTables.forEach(t => fillChapterTable(t, included, ctx('전체 목차(표)')));
            touched.set(path, doc);
            continue;
        }
        let cards = null;
        try { cards = detectCards(doc, xf, slideSize); } catch { cards = null; }
        if (cards) {
            fillCards(cards, included, ctx('전체 목차(카드)'));
            touched.set(path, doc);
            continue;
        }
        let tables = null;
        try { tables = detectTables(doc, xf, slideSize); } catch { tables = null; }
        if (tables && tables.pageCol >= 0 && tables.templates.sec) {
            detailEntries.push({ path, doc, det: tables, slideNo });
            touched.set(path, doc);
            continue;
        }
        const blank = detectBlankDetailTable(doc, xf, slideSize);
        if (blank) {
            blankEntries.push({ path, doc, blank, xf, slideNo });
            continue;
        }
        let det = null;
        try { det = detectTemplate(doc, xf, slideSize); } catch { det = null; }
        if (det && det.mode !== 'token' && det.chapterRoman && det.rows.length >= 2) {
            if (fillListSlide(doc, det, included, ctx('장별 목차'))) touched.set(path, doc);
        }
    }
    if (detailEntries.length) {
        fillDetailTables(detailEntries, included, { showPage, warnings, logFor: (slideNo) => makeLogger(changes, slideNo, '상세 목차(표)') });
    }
    // 빈 상세 목차 표: 슬라이드 제목의 장으로 채우고, 다른 상세 목차 슬라이드에 없는 장은 앞 슬라이드에 이어 붙임
    const extraSlides = [];
    const footerFix = new Map(); // path → 원래 바닥글
    if (blankEntries.length) {
        const rowTpl = await loadRowTemplates([{ zip }, fallbackFile ? { file: fallbackFile } : null]);
        const romanIdx = (r) => included.findIndex(c => c.roman === r);
        const assigned = blankEntries.map(b => b.blank.mentions.map(m => romanIdx(m.roman)).filter(i => i >= 0));
        const covered = new Set(assigned.flat());
        detailEntries.forEach(e => e.det.tbls.forEach(t => t.trs.forEach((tr, idx) => {
            if (kindOfRow(tr, e.det, idx) === 'ch') covered.add(romanIdx(leafText(cellsOf(tr)[e.det.numCol] || tr).trim().replace(/[.．]$/, '')));
        })));
        included.forEach((c, ci) => {
            if (covered.has(ci)) return;
            let k = -1;
            assigned.forEach((list, bi) => { if (list.some(x => x < ci)) k = bi; });
            if (k < 0) k = assigned.findIndex(list => list.length === 0) >= 0 ? assigned.findIndex(list => list.length === 0) : 0;
            assigned[k].push(ci);
        });
        // 같은 장을 담은 연속 빈 슬라이드(예: Ⅳ (1/2)·(2/2))는 한 묶음으로 한 번만 작성해 차례로 나눠 넣음
        const groups = [];
        blankEntries.forEach((b, bi) => {
            const cis = [...new Set(assigned[bi])].sort((x, y) => x - y);
            const key = cis.join(',');
            const last = groups[groups.length - 1];
            if (last && last.key === key && cis.length && last.lastIdx === bi - 1 && slides.indexOf(b.path) === slides.indexOf(last.items[last.items.length - 1].path) + 1) {
                last.items.push(b);
                last.lastIdx = bi;
            } else groups.push({ key, cis, items: [b], lastIdx: bi });
        });
        const footerOf = (doc) => {
            const sp = descendants(doc, 'sp').find(x => /^\s*목\s*차\s*[-–]/.test(leafText(x)));
            return sp ? leafText(sp).trim() : null;
        };
        const nextFooter = (f, k) => {
            if (!f) return null;
            const m = f.match(/^(.*?)[-–](\d+)$/);
            return m && /[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]/.test(m[1]) ? `${m[1]}-${parseInt(m[2], 10) + k}` : `${f}-${k + 1}`;
        };
        for (const g of groups) {
            const b = g.items[0];
            const nos = g.items.map(x => x.slideNo).join('·');
            if (!g.cis.length) { warnings.push(`${nos}번 슬라이드(빈 상세 목차 표): 채울 장을 정하지 못했습니다. 슬라이드 제목에 'Ⅰ. 사업 개요'처럼 장을 적어 주세요.`); continue; }
            const chs = g.cis.map(ci => included[ci]);
            const footers = g.items.map(x => footerOf(x.doc));
            const mention = b.blank.mentions.find(m => chs.some(c => c.roman === m.roman)) || { roman: chs[0].roman, title: chs[0].title };
            const { usedDesignRows } = synthesizeTemplate(b.doc, b.blank, rowTpl, mention);
            let r;
            try {
                r = fillTableSlides({ templateXml: serializeXml(b.doc), layoutXfrms: b.xf, slideSize, chapters: chs, showPage });
            } catch (e) {
                warnings.push(`${nos}번 슬라이드(빈 상세 목차 표)를 채우지 못했습니다: ${e.message}`);
                continue;
            }
            if (!r.slides.length) continue;
            if (usedDesignRows) warnings.push(`${nos}번 슬라이드: 장·절·항 행 서식 견본이 없어 표의 빈 행 서식으로 모두 작성했습니다.`);
            r.warnings.forEach(w => warnings.push(`${nos}번 슬라이드: ${w}`));
            const rows = chs.reduce((n, c) => n + 1 + c.sections.length + c.sections.reduce((m, s) => m + s.subsections.length, 0), 0);
            // 묶음의 슬라이드에 차례로 넣고(바닥글은 각 슬라이드 원래 값), 남으면 뒤에 추가, 모자라면 남는 슬라이드 경고
            g.items.forEach((x, k) => {
                if (k < r.slides.length) {
                    footerFix.set(x.path, footers[k]);
                    touched.set(x.path, parseXml(r.slides[k]));
                }
            });
            const extra = r.slides.slice(g.items.length);
            if (extra.length) {
                const lastFooter = footers[footers.length - 1];
                extraSlides.push({ basePath: g.items[g.items.length - 1].path, xmls: extra, footers: extra.map((_, k) => nextFooter(lastFooter, k + 1)) });
            }
            if (r.slides.length < g.items.length) {
                warnings.push(`${g.items.slice(r.slides.length).map(x => x.slideNo).join('·')}번 슬라이드: 내용이 적어 빈 표로 남았습니다. 필요 없으면 삭제해 주세요.`);
            }
            changes.push({ slide: b.slideNo, area: '상세 목차(빈 표)', where: chs.map(c => c.roman).join('·'), field: '작성', before: '(빈 표)', after: `${chs.length}개 장 ${rows}행을 ${r.slides.length}장에 작성${extra.length ? ` (뒤에 ${extra.length}장 추가)` : ''}` });
        }
    }
    if (!touched.size) throw new Error('올린 파일에서 갱신할 목차(장 열 표·장 카드·상세 목차 표·장별 목차)를 찾지 못했습니다.');

    const d = new Date();
    const dateText = `${d.getFullYear()}.${String(d.getMonth() + 1).padStart(2, '0')}.${String(d.getDate()).padStart(2, '0')}`;
    touched.forEach((doc, path) => zip.file(path, finalizeSlide(serializeXml(doc), { footer: footerFix.get(path) ?? null, dateText })));
    for (const x of extraSlides) await insertSlidesAfter(zip, x.basePath, x.xmls.map((xml, k) => finalizeSlide(xml, { footer: x.footers?.[k] ?? null, dateText })));
    const blob = await zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
    const updatedSlides = slides.map((p, i) => (touched.has(p) ? i + 1 : null)).filter(Boolean);
    const addedSlides = extraSlides.reduce((n, x) => n + x.xmls.length, 0);
    return { blob, changes, warnings, updatedSlides, addedSlides };
}
