/**
 * 요구사항추적표(RTM) 근거 수집·판정 엔진
 *
 * 브라우저/Node 공용 순수 함수 모듈 (파일 추출기·엑셀 의존 없음).
 * 근거 우선순위: 수동 확정 > 산출물 내 요구사항 대응표 > 장표 제목의 ID 명시 > 거버닝·본문 ID 언급
 *               > 절 제목 일치 / 동의어 사전 기반 의미 매칭
 * 산출정보(예: 현황 분석서 + 방안)별로 근거를 찾아 충족/부분충족/확인필요/미충족을 판정한다.
 */

// 정렬된 숫자 배열을 ['7~9', '12'] 형태의 구간 목록으로 축약
function compressNumericRanges(nums) {
    if (nums.length === 0) return [];
    if (nums.length >= 10 && (nums[nums.length - 1] - nums[0] + 1) === nums.length) {
        return [`${nums[0]}~${nums[nums.length - 1]}`];
    }
    const ranges = [];
    let start = nums[0];
    let prev = nums[0];
    for (let i = 1; i < nums.length; i++) {
        const cur = nums[i];
        if (cur === prev + 1) {
            prev = cur;
        } else {
            ranges.push(start === prev ? `${start}` : `${start}~${prev}`);
            start = cur;
            prev = cur;
        }
    }
    ranges.push(start === prev ? `${start}` : `${start}~${prev}`);
    return ranges;
}

/**
 * 페이지 배열을 '7~58' 또는 '12, 15~18' 형식으로 스마트 축약
 */
export function formatPageRanges(pages = []) {
    if (!pages || pages.length === 0) return '-';
    const nums = Array.from(new Set(pages.map(p => parseInt(p, 10)).filter(n => !isNaN(n)))).sort((a, b) => a - b);
    if (nums.length === 0) return '전 페이지';
    return compressNumericRanges(nums).join(', ');
}

/**
 * 슬라이드 하단 표시 라벨("장번호-슬라이드번호", 예: Ⅱ-15)을 장별로 묶어 'Ⅱ-7~9, Ⅱ-12, Ⅲ-5' 형식으로 축약.
 * 라벨이 하나도 없으면(PPTX 외 문서 등) 숫자 페이지로 폴백한다.
 */
export function formatPageLabelRanges(labels = [], plainPages = []) {
    const labelList = (labels || []).filter(Boolean);
    if (labelList.length === 0) return formatPageRanges(plainPages);

    const groups = new Map(); // 장번호 -> Set<슬라이드번호> (등장 순서 유지)
    const unlabeledNums = new Set();
    labelList.forEach(label => {
        const m = String(label).match(/^(.+?)-(\d+)$/);
        if (m) {
            if (!groups.has(m[1])) groups.set(m[1], new Set());
            groups.get(m[1]).add(parseInt(m[2], 10));
        } else {
            const n = parseInt(label, 10);
            if (!isNaN(n)) unlabeledNums.add(n);
        }
    });

    const parts = [];
    groups.forEach((numSet, prefix) => {
        const nums = Array.from(numSet).sort((a, b) => a - b);
        compressNumericRanges(nums).forEach(range => parts.push(`${prefix}-${range}`));
    });
    if (unlabeledNums.size > 0) {
        parts.push(...compressNumericRanges(Array.from(unlabeledNums).sort((a, b) => a - b)));
    }
    return parts.length > 0 ? parts.join(', ') : formatPageRanges(plainPages);
}


export const normalizeForMatch = (s = '') => String(s).toLowerCase().replace(/아키텍쳐/g, '아키텍처').replace(/[\s·ㆍ・•\-_/(),.'"「」[\]]/g, '');

// ── 요구사항명/서류명 ↔ 절·항 제목 내용 매칭 (요구사항 ID가 본문에 없는 실제 산출물 대응) ──
export const MATCH_STOPWORDS = new Set(['분석', '수립', '방안', '제시', '검토', '도출', '설계', '작성', '현황', '및', '등', '개요', '추진', '계획', '방향', '관리', '체계', '전략', '결과', '마련', '기반', '대상', '서', '세부', '상세', '정의', '분석서', '설계서', '정의서']);
export const MATCH_GENERIC = new Set(['데이터', '시스템', '정보', '정보화', '업무', '공공의료', 'ai', '사업', '기술', '서비스', '공공', '의료', '보건의료', '목표', '정보시스템']);
const STEM_SUFFIXES = ['분석서', '설계서', '정의서', '분석', '현황', '수립', '방안', '설계', '계획', '전략', '개요'];

export function matchTokens(text = '') {
    return String(text)
        .split(/[\s,/()]+|및/)
        .map(normalizeForMatch)
        .filter(Boolean)
        .map(t => {
            const suf = STEM_SUFFIXES.find(sf => t.endsWith(sf) && t.length - sf.length >= 2);
            return suf ? t.slice(0, -suf.length) : t;
        });
}

function splitStem(token) {
    const suf = STEM_SUFFIXES.find(sf => token.endsWith(sf) && token.length - sf.length >= 2);
    return suf ? [token.slice(0, -suf.length), suf] : [token, ''];
}

// 절·항 제목이 요구사항명/서류명과 같은 주제인지 판정
// partial=true(절 단위): 제목 핵심어의 절반 이상이 일치하고 그중 고유어가 있으면 일치 (예: '이행과제 로드맵' ↔ '…추진 로드맵 수립')
export function titleMatchesText(title = '', text = '', { partial = false } = {}) {
    const normTitle = normalizeForMatch(title);
    const normText = normalizeForMatch(text);
    if (normTitle.length < 2 || normText.length < 2) return false;
    if (normTitle.length >= 4 && (normText.includes(normTitle) || normTitle.includes(normText))) return true;

    const textTokens = new Set(matchTokens(text));
    // 범용어(데이터·업무 등)는 독립 단어로 쓰였을 때만, 고유어(법·제도 등)는 포함 여부로 판정
    const stemHit = (stem) => (MATCH_GENERIC.has(stem) ? textTokens.has(stem) : normText.includes(stem));
    const rawTokens = String(title).split(/[\s,/()]+|및/).map(normalizeForMatch).filter(t => t.length >= 2);
    const parts = rawTokens.map(splitStem).filter(([stem]) => stem.length >= 2 && !MATCH_STOPWORDS.has(stem));
    if (parts.length === 0) return false;

    // 범용어만으로 된 제목('데이터 현황분석', '업무현황 분석')은 꼬리말(현황·분석)까지 모두 들어 있어야 일치
    if (parts.every(([stem]) => MATCH_GENERIC.has(stem))) {
        return rawTokens.every(raw => {
            if (normText.includes(raw)) return true;
            const [stem, suf] = splitStem(raw);
            return suf !== '' && stemHit(stem) && normText.includes(suf);
        });
    }

    const hits = parts.filter(([stem]) => stemHit(stem));
    if (hits.length === parts.length) return true;
    if (!partial) return false;
    const strong = hits.filter(([stem]) => !MATCH_GENERIC.has(stem));
    return strong.length > 0 && hits.length * 2 >= parts.length && strong.reduce((n, [stem]) => n + stem.length, 0) >= 3;
}

// 장 주제어: 요구사항에 다른 장의 주제어(예: '현황')가 있으면 이 장의 절과는 부분 일치로 연결하지 않는다
export const CHAPTER_TOPIC_STEMS = ['환경', '현황', '목표모델', '이행'];
export const chapterTopicOf = (text = '') => {
    const n = normalizeForMatch(text);
    return CHAPTER_TOPIC_STEMS.find(t => n.includes(t)) || null;
};


// ─────────────────────────────────────────────────────────────────────────────
// 공통 상수·유틸
// ─────────────────────────────────────────────────────────────────────────────
export const ROMAN_LIST = ['', 'Ⅰ', 'Ⅱ', 'Ⅲ', 'Ⅳ', 'Ⅴ', 'Ⅵ', 'Ⅶ', 'Ⅷ', 'Ⅸ', 'Ⅹ'];
const ROMAN_CHARS = 'ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ';

export const EVIDENCE_TYPE_LABELS = {
    manual: '수동 확정',
    table: '대응표',
    titleId: 'ID 명시',
    bodyId: 'ID 언급',
    file: '문서명 파일',
    mgmt: '사업관리 문서',
    section: '절 제목 일치',
    semantic: '의미 매칭',
    mgmtFile: '사업관리 문서(파일명)'
};
const TYPE_ORDER = ['manual', 'table', 'titleId', 'bodyId', 'file', 'mgmt', 'section', 'semantic', 'mgmtFile'];

export const RTM_STATUS = {
    MET: '충족',
    PARTIAL: '부분충족',
    CHECK: '확인필요',
    UNMET: '미충족',
    MGMT_CHECK: '사업관리 문서 확인필요'
};

const HEADING_CODE_RE = /^0?(\d{1,2}(?:\.\d{1,2}){1,3})\.?\s+(\S.*)$/;
const SECTION_ONLY_RE = /^0?(\d{1,2})\.\s*([^\d\s.].*)$/;
const REQ_ID_RE = /\b([A-Z]{2,5}-\d{2,4})\b/g;

const uniq = (arr) => Array.from(new Set(arr));
const cleanSlideTitle = (t = '') => String(t).replace(/\s*\(\s*\d+\s*\/\s*\d+\s*\)\s*$/, '').replace(/\s+/g, ' ').trim();

function phraseMatches(phrase, text) {
    const normText = normalizeForMatch(text);
    const normPhrase = normalizeForMatch(phrase);
    if (!normPhrase || !normText) return false;
    if (normText.includes(normPhrase)) return true;
    const tokens = String(phrase).split(/[\s,/()]+/).map(normalizeForMatch).filter(t => t.length >= 2);
    return tokens.length > 1 && tokens.every(t => normText.includes(t));
}

// 요구사항명에 해당하는 동의어 사전의 산출물 쪽 용어
export function expandSynonyms(reqName, config) {
    const out = [];
    (config?.synonyms || []).forEach(entry => {
        if ((entry.from || []).some(f => phraseMatches(f, reqName))) out.push(...(entry.to || []));
    });
    return uniq(out);
}

// ─────────────────────────────────────────────────────────────────────────────
// 설정(동의어 사전·산출정보·기대 절) 병합 / CSV 사전 / 수동 확정 목록 파싱
// ─────────────────────────────────────────────────────────────────────────────
export function mergeMatchingConfig(base = {}, override = null) {
    if (!override) return base;
    const merged = {
        ...base,
        policy: { ...(base.policy || {}), ...(override.policy || {}) },
        synonyms: [...(base.synonyms || []), ...(override.synonyms || [])],
        requirements: { ...(base.requirements || {}) },
        management: {
            ...(base.management || {}),
            ...(override.management || {}),
            fileKeywords: { ...(base.management?.fileKeywords || {}), ...(override.management?.fileKeywords || {}) }
        }
    };
    Object.entries(override.requirements || {}).forEach(([id, cfg]) => {
        merged.requirements[id] = { ...(merged.requirements[id] || {}), ...cfg };
    });
    return merged;
}

function splitCsvLine(line) {
    const cells = [];
    let cur = '';
    let quoted = false;
    for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (quoted) {
            if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; }
            else if (ch === '"') quoted = false;
            else cur += ch;
        } else if (ch === '"') quoted = true;
        else if (ch === ',' || ch === '\t') { cells.push(cur.trim()); cur = ''; }
        else cur += ch;
    }
    cells.push(cur.trim());
    return cells;
}

export function parseCsvRows(text = '') {
    return String(text).replace(/^\uFEFF/, '').split(/\r?\n/).filter(l => l.trim()).map(splitCsvLine);
}

// 동의어 사전 CSV: "요구사항 용어, 산출물 용어1|산출물 용어2" (첫 열에도 | 로 여러 용어 가능)
export function parseSynonymCsv(text = '') {
    const rows = parseCsvRows(text);
    const synonyms = [];
    rows.forEach((cells, i) => {
        if (cells.length < 2) return;
        if (i === 0 && /요구사항|from/i.test(cells[0]) && /산출물|to/i.test(cells[1])) return;
        const split = (v) => v.split(/[|;]/).map(x => x.trim()).filter(Boolean);
        const from = split(cells[0]);
        const to = cells.slice(1).flatMap(split);
        if (from.length && to.length) synonyms.push({ from, to });
    });
    return { synonyms };
}

const MANUAL_HEADER_ALIASES = {
    reqId: ['요구사항id', '요구사항번호', '요구사항 id', 'id', 'requirementid'],
    fileName: ['파일명', '파일', 'file', 'filename'],
    location: ['절번호', '절 번호', '위치', 'location', '절'],
    title: ['장표제목', '장표 제목', 'title', '제목'],
    pages: ['페이지', 'pages', 'page', '해당 페이지'],
    output: ['산출정보', 'output'],
    status: ['판정', 'status', '충족여부'],
    confirm: ['확정', '확정(y)', 'confirm', '확정여부']
};

// 수동 확정 목록 (CSV 행 배열 또는 객체 배열) → [{ reqId, fileName, location, title, pages, output, status }]
export function parseManualEntries(rowsOrObjects = []) {
    if (!Array.isArray(rowsOrObjects) || rowsOrObjects.length === 0) return [];
    let objects = rowsOrObjects;
    if (Array.isArray(rowsOrObjects[0])) {
        const header = rowsOrObjects[0].map(h => String(h || '').trim().toLowerCase());
        objects = rowsOrObjects.slice(1).map(cells => {
            const o = {};
            header.forEach((h, i) => { o[h] = cells[i]; });
            return o;
        });
    }
    const pick = (o, key) => {
        const keys = Object.keys(o);
        const hit = keys.find(k => MANUAL_HEADER_ALIASES[key].includes(String(k).trim().toLowerCase()));
        return hit != null ? String(o[hit] ?? '').trim() : '';
    };
    const hasConfirmColumn = objects.some(o => pick(o, 'confirm') !== '');
    return objects
        .map(o => ({
            reqId: pick(o, 'reqId').toUpperCase(),
            fileName: pick(o, 'fileName'),
            location: pick(o, 'location'),
            title: pick(o, 'title'),
            pages: pick(o, 'pages'),
            output: pick(o, 'output'),
            status: pick(o, 'status'),
            confirm: pick(o, 'confirm')
        }))
        .filter(e => /^[A-Z]{2,5}-\d{2,4}$/.test(e.reqId))
        .filter(e => !hasConfirmColumn || /^(y|yes|o|확정|true|1)$/i.test(e.confirm));
}

// ─────────────────────────────────────────────────────────────────────────────
// 장표 색인 (파일 · 장 · 절 번호 · 장표 제목 · 거버닝 · 표)
// ─────────────────────────────────────────────────────────────────────────────
/**
 * @param {object} p
 * @param {string} p.fileName
 * @param {object} p.content  extractStructuredFileContent 결과 (pageMap, pageLabelMap, slideChapterMap, slideMetaMap)
 * @param {object} p.doc      절 색인 { chapter, chapterTitle, sections:[{num,title,pages,labels,subsections}] }
 */
export function buildDeckIndex({ fileName, content = {}, doc = {} }) {
    const { pageMap = {}, pageLabelMap = {}, slideChapterMap = {}, slideMetaMap = {} } = content;
    const deckRoman = ROMAN_LIST[doc.chapter] || '';
    const secOfSlide = new Map();
    const subOfSlide = new Map();
    (doc.sections || []).forEach(sec => {
        sec.pages.forEach(p => secOfSlide.set(p, sec.num));
        (sec.subsections || []).forEach(sub => sub.pages.forEach(p => subOfSlide.set(p, `${sec.num}.${sub.num}`)));
    });

    const positions = uniq([...Object.keys(pageMap), ...Object.keys(slideMetaMap)].map(Number).filter(n => !isNaN(n))).sort((a, b) => a - b);
    let carry = null;
    const slides = positions.map(pos => {
        const paragraphs = pageMap[pos] || [];
        const meta = slideMetaMap[pos] || {};
        let title = cleanSlideTitle(meta.title || '');
        if (!title) {
            const h = paragraphs.find(p => p.length <= 80 && HEADING_CODE_RE.test(p.trim()));
            if (h) title = cleanSlideTitle(h);
        }
        let code = null;
        const tm = title.match(HEADING_CODE_RE);
        if (tm) code = tm[1];
        if (!code) {
            const codes = (meta.headings || []).map(h => (h.match(HEADING_CODE_RE) || [])[1]).filter(Boolean);
            if (codes.length) code = codes.sort((a, b) => b.split('.').length - a.split('.').length)[0];
        }
        let secNum = secOfSlide.get(pos);
        if (secNum == null && code) secNum = parseInt(code.split('.')[0], 10);
        if (secNum == null) {
            const sh = (meta.headings || []).map(h => h.match(SECTION_ONLY_RE)).find(Boolean);
            if (sh) secNum = parseInt(sh[1], 10);
        }
        if (!code) {
            if (carry && secNum != null && carry.split('.')[0] === String(secNum)) code = carry;
            else code = subOfSlide.get(pos) || (secNum != null ? String(secNum) : null);
        }
        if (code && code.includes('.')) carry = code;

        const label = pageLabelMap[pos] || String(pos);
        const labelRoman = ROMAN_CHARS.includes(String(label)[0]) ? String(label)[0] : '';
        const roman = labelRoman || slideChapterMap[pos]?.roman || deckRoman;
        const governing = meta.governing || '';
        const tables = meta.tables || [];
        const text = paragraphs.join(' ');
        const upper = `${title} ${governing} ${text}`.toUpperCase();
        const ids = uniq(Array.from(upper.matchAll(REQ_ID_RE)).map(m => m[1]));
        return { pos, label, roman, secNum: secNum ?? null, code, title, governing, text, tables, ids, titleUpper: title.toUpperCase() };
    });

    // 절·항별 장표 목록 (요구사항마다 반복 조회하므로 미리 색인)
    const byPos = new Map(slides.map(sl => [sl.pos, sl]));
    const pick = (pages) => pages.map(p => byPos.get(p)).filter(Boolean);
    const sectionSlides = new Map();
    (doc.sections || []).forEach(sec => {
        sectionSlides.set(String(sec.num), pick(sec.pages));
        (sec.subsections || []).forEach(sub => sectionSlides.set(`${sec.num}.${sub.num}`, pick(sub.pages)));
    });
    return { fileName, roman: deckRoman, chapter: doc.chapter || null, chapterTitle: doc.chapterTitle || '', sections: doc.sections || [], slides, sectionSlides };
}

// ─────────────────────────────────────────────────────────────────────────────
// 대응표의 작성 위치 표기 해석: "Ⅲ.5 5.2.1~5.2.9", "Ⅳ.6 6.5.1, 6.5.7", "5.2.x", "Ⅳ-690~692", "Ⅳ.11 11.3.1"
// ─────────────────────────────────────────────────────────────────────────────
const LOCATION_TOKEN_RE = new RegExp(
    `([${ROMAN_CHARS}])\\s*[-–]\\s*(\\d{1,4})(?:\\s*[~∼～]\\s*(?:[${ROMAN_CHARS}]\\s*[-–]\\s*)?(\\d{1,4}))?` + // 1-3: 페이지 라벨
    `|([${ROMAN_CHARS}])\\s*[.．]\\s*(\\d{1,2})(?![.\\d])` + // 4-5: 장.절 (Ⅲ.5)
    `|(?<![\\d.])(\\d{1,2}(?:\\.\\d{1,2}){0,3})\\.[xX*]` + // 6: 와일드카드 (5.2.x)
    `|(?<![\\d.])(\\d{1,2}(?:\\.\\d{1,2}){1,3})(?![\\d])(?:\\s*[~∼～]\\s*(\\d{1,2}(?:\\.\\d{1,2}){1,3}))?`, // 7-8: 절 번호(범위)
    'g'
);

export function parseLocationRefs(text = '', defaultRoman = '') {
    const refs = [];
    let curRoman = defaultRoman;
    let pendingSection = null; // { ref, sec } "Ⅲ.5" 뒤에 "5.2.1" 같은 세부 번호가 오면 장.절 범위 참조는 생략
    for (const m of String(text).matchAll(LOCATION_TOKEN_RE)) {
        if (m[1]) {
            refs.push({ kind: 'label', roman: m[1], from: parseInt(m[2], 10), to: parseInt(m[3] || m[2], 10) });
            curRoman = m[1];
            pendingSection = null;
        } else if (m[4]) {
            curRoman = m[4];
            const ref = { kind: 'code', roman: m[4], code: String(parseInt(m[5], 10)) };
            refs.push(ref);
            pendingSection = { ref, sec: ref.code };
        } else {
            const code = m[6] != null ? m[6] : m[7];
            const end = m[6] != null ? null : (m[8] || null);
            if (!curRoman) continue;
            if (pendingSection && pendingSection.ref.roman === curRoman && code.split('.')[0] === pendingSection.sec) {
                const idx = refs.indexOf(pendingSection.ref);
                if (idx >= 0) refs.splice(idx, 1);
                pendingSection = null;
            }
            refs.push({ kind: 'code', roman: curRoman, code, end });
        }
    }
    return refs;
}

const codeParts = (c) => String(c).split('.').map(n => parseInt(n, 10));

function slideMatchesCodeRef(slide, ref) {
    const slideCode = slide.code || (slide.secNum != null ? String(slide.secNum) : null);
    if (!slideCode || slide.roman !== ref.roman) return false;
    const s = codeParts(slideCode);
    const a = codeParts(ref.code);
    if (!ref.end) return a.every((v, i) => s[i] === v);
    const b = codeParts(ref.end);
    if (a.length !== b.length) return a.every((v, i) => s[i] === v);
    const d = a.length - 1;
    for (let i = 0; i < d; i++) if (s[i] !== a[i]) return false;
    return s[d] >= a[d] && s[d] <= b[d];
}

const refDisplay = (ref) => (ref.kind === 'label'
    ? `${ref.roman}-${ref.from}${ref.to !== ref.from ? `~${ref.to}` : ''}`
    : `${ref.code}${ref.end ? `~${ref.end}` : ''}`);

// 참조 → 근거 묶음 (파일별)
function resolveRef(ref, decks) {
    const out = [];
    decks.forEach(deck => {
        const hits = deck.slides.filter(sl => (ref.kind === 'label'
            ? sl.roman === ref.roman && (() => { const n = parseInt(String(sl.label).split('-').pop(), 10); return n >= ref.from && n <= ref.to; })()
            : slideMatchesCodeRef(sl, ref)));
        if (hits.length) out.push({ deck, slides: hits, codeText: refDisplay(ref) });
    });
    return out;
}

// ─────────────────────────────────────────────────────────────────────────────
// 근거 객체
// ─────────────────────────────────────────────────────────────────────────────
function sectionTitleOf(deck, secNum) {
    const sec = deck.sections.find(s => s.num === secNum);
    return sec ? sec.title : '';
}

function makeEvidence(type, deck, slides, { codeText = null, title = null } = {}) {
    const first = slides[0];
    const secNums = uniq(slides.map(s => s.secNum).filter(n => n != null));
    const roman = first?.roman || deck.roman;
    const codes = uniq(slides.map(s => s.code).filter(Boolean));
    const code = codeText || (codes.length === 1 ? codes[0] : (codes.length ? `${codes[0]}~${codes[codes.length - 1]}` : ''));
    const titles = uniq(slides.map(s => s.title.replace(HEADING_CODE_RE, '$2').trim()).filter(Boolean));
    return {
        type,
        fileName: deck.fileName,
        roman,
        sections: secNums.map(n => `${roman}.${n}`),
        sectionCode: secNums.length === 1 ? `${roman}.${secNums[0]}` : secNums.map(n => `${roman}.${n}`).join(', '),
        code,
        title: title ?? (titles.length === 1 ? titles[0] : (titles[0] ? `${titles[0]} 외` : sectionTitleOf(deck, secNums[0]))),
        pages: slides.map(s => s.pos),
        labels: slides.map(s => s.label)
    };
}

// 장표들을 (파일, 절 번호) 단위 근거로 묶는다
function groupSlidesAsEvidence(type, deck, slides) {
    const byCode = new Map();
    slides.forEach(sl => {
        const key = sl.code || String(sl.secNum ?? '');
        if (!byCode.has(key)) byCode.set(key, []);
        byCode.get(key).push(sl);
    });
    return Array.from(byCode.values()).map(group => makeEvidence(type, deck, group));
}

const isMappingSlide = (slide, masterIds) => {
    const idsOnSlide = slide.ids.filter(id => masterIds.has(id));
    if (idsOnSlide.length < 2) return false;
    return /대응표|대응\s*현황|추적표|매핑표|요구사항\s*(대응|추적|매핑)|산출정보\s*대응/.test(`${slide.title} ${slide.text}`);
};

// ─────────────────────────────────────────────────────────────────────────────
// 요구사항별 근거 수집
// ─────────────────────────────────────────────────────────────────────────────
function collectDeckEvidence(req, decks, ctx, expected = []) {
    const id = req.id.toUpperCase();
    const evidences = [];
    const { masterIds, config } = ctx;
    const listingThreshold = config.policy?.idListingThreshold || 4;

    decks.forEach(deck => {
        const titleIdSlides = [];
        const bodyIdSlides = [];
        deck.slides.forEach(slide => {
            if (!slide.ids.includes(id)) return;
            // ① 대응표: 해당 행(또는 문단)에 적힌 작성 위치를 근거로 채택
            if (isMappingSlide(slide, masterIds)) {
                const rows = slide.tables.filter(cells => cells.join(' ').toUpperCase().includes(id));
                const sources = rows.length ? rows.map(cells => cells.join(' ')) : [slide.text];
                sources.forEach(src => {
                    const refs = parseLocationRefs(src.replace(new RegExp(id, 'gi'), ' '), deck.roman);
                    refs.forEach(ref => {
                        resolveRef(ref, decks).forEach(({ deck: d, slides, codeText }) => {
                            evidences.push({ ...makeEvidence('table', d, slides, { codeText }), via: `${deck.fileName} ${slide.label} 대응표` });
                        });
                    });
                });
                return;
            }
            // ② 장표 제목에 ID 명시 / ③ 거버닝·본문 ID 언급 (ID를 여러 개 나열한 장표는 제외)
            if (slide.titleUpper.includes(id)) titleIdSlides.push(slide);
            else if (slide.ids.filter(x => masterIds.has(x)).length < listingThreshold) bodyIdSlides.push(slide);
        });
        evidences.push(...groupSlidesAsEvidence('titleId', deck, titleIdSlides));
        evidences.push(...groupSlidesAsEvidence('bodyId', deck, bodyIdSlides));
    });

    // ④ 절 제목 일치 / 동의어 사전 기반 의미 매칭
    const synonyms = expandSynonyms(req.name || '', config);
    const normName = normalizeForMatch(req.name || '');
    const nameTopics = CHAPTER_TOPIC_STEMS.filter(t => normName.includes(t));
    const titleHit = (title, partial) => !!title && (titleMatchesText(title, req.name || '', { partial }) || synonyms.some(ph => phraseMatches(ph, title)));

    decks.forEach(deck => {
        const deckTopic = chapterTopicOf(deck.chapterTitle) || chapterTopicOf(deck.fileName);
        const allowPartial = nameTopics.length === 0 || (deckTopic && nameTopics.includes(deckTopic));
        deck.sections.forEach(sec => {
            const secSlides = deck.sectionSlides.get(String(sec.num)) || [];
            if (titleHit(sec.title, allowPartial)) {
                if (secSlides.length) evidences.push(makeEvidence('section', deck, secSlides, { codeText: String(sec.num), title: sec.title }));
                return;
            }
            (sec.subsections || []).forEach(sub => {
                if (!titleHit(sub.title, false)) return;
                const subSlides = deck.sectionSlides.get(`${sec.num}.${sub.num}`) || [];
                if (subSlides.length) evidences.push(makeEvidence('section', deck, subSlides, { codeText: `${sec.num}.${sub.num}`, title: sub.title }));
            });
        });
        // 장표 제목(세부 절 번호가 있는 장표) 의미 매칭
        // 기대 절(RFP상 산출물 범위) 안의 장표는 핵심어 절반 이상 일치도 의미 매칭으로 인정 (단어 불일치만으로 미충족 처리하지 않음)
        const inExpectedSlide = (sl) => expected.includes(`${sl.roman}.${sl.secNum}`);
        const semanticSlides = deck.slides.filter(sl => sl.code && sl.code.includes('.')
            && titleHit(sl.title.replace(HEADING_CODE_RE, '$2').replace(REQ_ID_RE, ''), inExpectedSlide(sl)));
        evidences.push(...groupSlidesAsEvidence('semantic', deck, semanticSlides));
    });
    return evidences;
}

export function documentNameTokens(text = '') {
    return String(text)
        .split(/[\s,/()]+|및/)
        .map(normalizeForMatch)
        .filter(t => t.length >= 4 && /(서|표)$/.test(t) && !MATCH_GENERIC.has(t));
}

function collectFileEvidence(keywords, files) {
    const keys = (keywords || []).map(normalizeForMatch).filter(Boolean);
    return files
        .filter(f => keys.some(k => normalizeForMatch(f.fileName).includes(k)))
        .map(f => ({ type: 'file', fileName: f.fileName, roman: '', sections: [], sectionCode: '', code: '', title: '파일 전체', pages: [], labels: [] }));
}

function collectManagementEvidence(req, mgmtDocs, config) {
    const id = req.id.toUpperCase();
    const prefix = id.split('-')[0];
    const nameNorm = normalizeForMatch(req.name || '');
    const evidences = [];
    mgmtDocs.forEach(doc => {
        const pages = [];
        Object.entries(doc.pageMap || {}).forEach(([p, lines]) => {
            const t = (Array.isArray(lines) ? lines.join(' ') : String(lines));
            if (t.toUpperCase().includes(id) || (nameNorm.length >= 4 && normalizeForMatch(t).includes(nameNorm))) pages.push(parseInt(p, 10));
        });
        if (pages.length) {
            evidences.push({ type: 'mgmt', fileName: doc.fileName, roman: '', sections: [], sectionCode: '', code: '', title: '요구사항 ID·명칭 언급', pages, labels: pages.map(String) });
            return;
        }
        const kws = (config.management?.fileKeywords?.[prefix] || []).map(normalizeForMatch);
        if (kws.some(k => normalizeForMatch(doc.fileName).includes(k))) {
            evidences.push({ type: 'mgmtFile', fileName: doc.fileName, roman: '', sections: [], sectionCode: '', code: '', title: '파일명 기준', pages: [], labels: [] });
        }
    });
    return evidences;
}

// ─────────────────────────────────────────────────────────────────────────────
// 판정
// ─────────────────────────────────────────────────────────────────────────────
function expectedSectionsFor(id, reqCfg, deliverableItems) {
    const set = new Set(reqCfg.expectedSections || []);
    (reqCfg.deliverables || []).forEach(seq => {
        const item = deliverableItems.find(d => d.seq === seq);
        if (!item) return;
        const re = /([ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ])\.(\d+)/g;
        const tokens = Array.from(String(item.note || '').matchAll(re));
        tokens.forEach((t, i) => {
            set.add(`${t[1]}.${parseInt(t[2], 10)}`);
            const prev = tokens[i - 1];
            if (prev && prev[1] === t[1] && String(item.note).slice(prev.index, t.index).includes('~')) {
                for (let n = parseInt(prev[2], 10) + 1; n < parseInt(t[2], 10); n++) set.add(`${t[1]}.${n}`);
            }
        });
    });
    return Array.from(set);
}

const isStrongType = (type, policy) => ['manual', 'table', 'titleId', 'bodyId', 'file', 'mgmt'].includes(type) || (type === 'section' && policy.sectionTitleMatchIsMet !== false);

const evidenceLocationText = (e) => {
    const loc = [e.sectionCode, e.code && e.code !== String(e.sectionCode).split('.').pop() ? e.code : ''].filter(Boolean).join(' ');
    const pages = e.labels.length ? formatPageLabelRanges(e.labels, e.pages) : (e.pagesText || '전 페이지');
    return [e.fileName, loc || '-', e.title || '-', pages].join(' / ');
};
export { evidenceLocationText };

/**
 * 요구사항별 근거 수집 및 판정
 * @returns {Array<{id,name,status,typeLabels,evidences,outputs,missingOutputs,warnings,notes,files,pageText,isManagement}>}
 */
export function evaluateRequirements({
    requirements = [],
    decks = [],
    otherFiles = [],      // 검수 대상 중 PPTX 외 파일 [{ fileName }]
    managementDocs = [],  // 사업관리 근거 문서 [{ fileName, pageMap }]
    manualEntries = [],
    config = {},
    deliverableItems = []
}) {
    const policy = config.policy || {};
    const masterIds = new Set(requirements.map(r => (r.id || '').trim().toUpperCase()).filter(Boolean));
    const ctx = { masterIds, config };
    const mgmtPrefixes = config.management?.prefixes || [];
    const allFiles = [...decks.map(d => ({ fileName: d.fileName })), ...otherFiles, ...managementDocs.map(d => ({ fileName: d.fileName }))];

    return requirements.map(req => {
        const id = (req.id || '').trim().toUpperCase();
        const reqCfg = config.requirements?.[id] || {};
        const isManagement = mgmtPrefixes.includes(id.split('-')[0]);
        const expected = expectedSectionsFor(id, reqCfg, deliverableItems);
        const expectedChapters = uniq(expected.map(s => s.split('.')[0]));
        const outputs = (reqCfg.outputs && reqCfg.outputs.length ? reqCfg.outputs : [{ key: '본문', label: req.name, chapters: expectedChapters }])
            .map(o => ({ ...o, chapters: o.chapters || [], grade: 'none' }));
        const warnings = [];
        const notes = [];

        // 수동 확정: PM이 확정한 연결을 그대로 반영
        const manual = manualEntries.filter(m => m.reqId === id);
        if (manual.length) {
            const evidences = manual.map(m => ({
                type: 'manual', fileName: m.fileName || '-', roman: '', sections: [], sectionCode: m.location || '', code: '',
                title: m.title || '', pages: [], labels: [], pagesText: m.pages || '-', output: m.output || ''
            }));
            // PM 확정 = 충족. 단, 판정 열에 부분충족·미충족을 직접 적은 경우는 그대로 반영
            const explicit = manual.map(m => m.status).find(st => st === RTM_STATUS.PARTIAL || st === RTM_STATUS.UNMET);
            const status = explicit || RTM_STATUS.MET;
            outputs.forEach(o => { o.grade = 'strong'; });
            return finalize(req, status, evidences, outputs, [], [], ['수동 확정 목록 반영'], isManagement);
        }

        let evidences = [];
        if (isManagement) evidences.push(...collectManagementEvidence(req, managementDocs, config));
        evidences.push(...collectDeckEvidence(req, decks, ctx, expected));
        // 요구사항명칭 속 문서명(제안요청서·사업계획서 등 '~서/~표')이 파일명에 있으면 그 파일 전체
        if (!reqCfg.outputs) evidences.push(...collectFileEvidence(documentNameTokens(req.name), allFiles));

        // 기대 위치(RFP상 산출물 장·절 범위) 점검
        // 절을 알 수 없는 장표(절 구조를 못 읽은 파일)는 범위를 판단할 수 없으므로 유지
        const inExpected = (e) => expected.length === 0 || e.sections.length === 0 || e.sections.some(s => expected.includes(s));
        const kept = [];
        evidences.forEach(e => {
            if (!['table', 'titleId', 'bodyId', 'section', 'semantic'].includes(e.type) || inExpected(e)) { kept.push(e); return; }
            if (e.type === 'table' || e.type === 'titleId') {
                kept.push(e);
                warnings.push(`근거 위치 확인필요: ${e.sectionCode || e.code} (${formatPageLabelRanges(e.labels, e.pages)})는 기대 위치 ${expected.join(', ')} 밖`);
            } else if (e.type === 'bodyId') {
                notes.push(`범위 밖 ID 언급 제외: ${e.fileName} ${formatPageLabelRanges(e.labels, e.pages)} (${e.sectionCode})`);
            }
        });
        evidences = kept;

        // 산출정보별 대응
        outputs.forEach(o => {
            let own;
            if (o.fileKeywords) {
                own = collectFileEvidence(o.fileKeywords, allFiles);
                evidences.push(...own);
            } else {
                own = evidences.filter(e => o.chapters.length === 0 || !e.roman || o.chapters.includes(e.roman));
            }
            own.forEach(e => { e.outputs = uniq([...(e.outputs || []), o.key]); });
            if (own.some(e => isStrongType(e.type, policy))) o.grade = 'strong';
            else if (own.length) o.grade = 'weak';
        });
        // 파일 기반 산출정보만 있는 요구사항: 장표 근거는 참고로만 남김
        if (outputs.every(o => o.fileKeywords)) {
            const deckEv = evidences.filter(e => e.type !== 'file');
            if (deckEv.length) notes.push(`별도 파일 산출물 - 장표 내 언급(${deckEv.map(e => `${e.fileName} ${formatPageLabelRanges(e.labels, e.pages)}`).join(', ')})은 근거에서 제외`);
            evidences = evidences.filter(e => e.type === 'file');
        }
        evidences = evidences.filter(e => (e.outputs || []).length > 0 || outputs.length === 0);
        evidences = pruneRedundantEvidence(evidences);

        const grades = outputs.map(o => o.grade);
        let status;
        if (grades.every(g => g === 'strong')) status = RTM_STATUS.MET;
        else if (grades.every(g => g === 'none')) status = isManagement ? RTM_STATUS.MGMT_CHECK : RTM_STATUS.UNMET;
        else if (grades.some(g => g === 'none')) status = RTM_STATUS.PARTIAL;
        else status = RTM_STATUS.CHECK;
        if (isManagement && status === RTM_STATUS.CHECK && evidences.every(e => e.type === 'mgmtFile')) status = RTM_STATUS.MGMT_CHECK;

        const isDefaultOutput = !(reqCfg.outputs && reqCfg.outputs.length);
        const missing = isDefaultOutput ? [] : outputs.filter(o => o.grade === 'none').map(o => `${o.label || o.key}${o.chapters.length ? `(${o.chapters.join('·')})` : ''}`);
        const weakOnly = outputs.filter(o => o.grade === 'weak').map(o => o.label || o.key);
        if (status !== RTM_STATUS.MET && weakOnly.length) {
            if (evidences.length && evidences.every(e => e.type === 'mgmtFile')) notes.push('사업관리 문서 파일명만 일치 - 본문에서 요구사항 대응 내용 확인 필요');
            else notes.push(isDefaultOutput ? '의미 매칭 근거만 있음 - 내용 확인 필요' : `의미 매칭 근거만 있음: ${weakOnly.join(', ')}`);
        }
        if (status === RTM_STATUS.MGMT_CHECK && evidences.length === 0) notes.push(managementDocs.length ? '지정한 사업관리 문서에서 근거를 찾지 못함' : '사업관리 근거 문서 미지정');
        return finalize(req, status, evidences, outputs, missing, warnings, notes, isManagement);
    });
}

// 더 강한 근거(대응표·ID 등)가 이미 있는 산출정보의 의미 매칭·절 제목 근거, 그리고 절 근거에 포함되는 장표 의미 매칭은 생략
const DIRECT_TYPES = new Set(['manual', 'table', 'titleId', 'bodyId', 'file', 'mgmt']);
function pruneRedundantEvidence(evidences) {
    const directOutputs = new Set(evidences.filter(e => DIRECT_TYPES.has(e.type)).flatMap(e => e.outputs || []));
    const covered = (e) => (e.outputs || []).length > 0 && e.outputs.every(o => directOutputs.has(o));
    const sectionSlides = new Map(); // fileName -> Set(pos)
    evidences.filter(e => e.type === 'section').forEach(e => {
        if (!sectionSlides.has(e.fileName)) sectionSlides.set(e.fileName, new Set());
        e.pages.forEach(p => sectionSlides.get(e.fileName).add(p));
    });
    const seen = new Set();
    return evidences.filter(e => {
        if ((e.type === 'semantic' || e.type === 'section') && covered(e)) return false;
        if (e.type === 'semantic' && e.pages.length && e.pages.every(p => sectionSlides.get(e.fileName)?.has(p))) return false;
        const key = `${e.type}|${e.fileName}|${e.code}|${e.pages.join(',')}`;
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
    });
}

function finalize(req, status, evidences, outputs, missing, warnings, notes, isManagement) {
    evidences.sort((a, b) => TYPE_ORDER.indexOf(a.type) - TYPE_ORDER.indexOf(b.type));
    const typeLabels = uniq(evidences.map(e => EVIDENCE_TYPE_LABELS[e.type]));
    const files = uniq(evidences.map(e => e.fileName));
    const labels = evidences.flatMap(e => e.labels);
    const pages = evidences.flatMap(e => e.pages);
    const pageText = labels.length ? formatPageLabelRanges(labels, pages) : (evidences.length ? (evidences.find(e => e.pagesText)?.pagesText || '전 페이지') : '-');
    return {
        id: req.id,
        name: req.name,
        status,
        typeLabels,
        evidences,
        outputs,
        missingOutputs: missing,
        warnings: uniq(warnings),
        notes: uniq(notes),
        files,
        pageText,
        isManagement
    };
}

// ─────────────────────────────────────────────────────────────────────────────
// RFP상 산출물 ↔ 요구사항 연계 (추적표 판정에서 산출물 현황 도출)
// ─────────────────────────────────────────────────────────────────────────────
export function buildDeliverableLinks(config = {}) {
    const links = new Map(); // seq -> [{ reqId, outputKey|null }]
    Object.entries(config.requirements || {}).forEach(([reqId, cfg]) => {
        (cfg.deliverables || []).forEach(seq => {
            if (!links.has(seq)) links.set(seq, []);
            links.get(seq).push({ reqId, outputKey: null });
        });
        (cfg.outputs || []).forEach(o => {
            if (o.deliverable == null) return;
            if (!links.has(o.deliverable)) links.set(o.deliverable, []);
            links.get(o.deliverable).push({ reqId, outputKey: o.key });
        });
    });
    return links;
}

export function gradeOfLink(link, resultsById) {
    const r = resultsById.get(link.reqId);
    if (!r) return { grade: 'none', result: null, evidences: [] };
    if (link.outputKey) {
        const o = r.outputs.find(x => x.key === link.outputKey);
        const evidences = r.evidences.filter(e => (e.outputs || []).includes(link.outputKey) || e.type === 'manual');
        return { grade: r.status === RTM_STATUS.MET && r.evidences.some(e => e.type === 'manual') ? 'strong' : (o ? o.grade : 'none'), result: r, evidences };
    }
    const grade = r.status === RTM_STATUS.MET ? 'strong' : (r.evidences.length ? (r.outputs.some(o => o.grade === 'strong') ? 'strong' : 'weak') : 'none');
    return { grade, result: r, evidences: r.evidences };
}
