/**
 * pptxMasterEngine.js
 * OpenXML 파싱 기반 PPTX 슬라이드 마스터 표준 규격 추출, 서식 정합성 검수 및 원본 보존형 자동 교정 엔진
 * 외부 라이브러리 추가 없이 jszip, exceljs, file-saver, DOMParser, XMLSerializer만 사용
 */
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { loadDesignTokens } from './designTokens.js';

const STORAGE_KEY = 'healthy_project_pptx_master_spec';
export const EMU_PER_INCH = 914400;
export const EMU_PER_PT = 12700;

// ── 0. 공공의료/ISP 표준 골든 마스터 규격 기본 프리셋 ──────────────
export const DEFAULT_MASTER_SPEC = {
    fileName: '공공의료 AI ISP 표준 골든마스터 (기본 프리셋)',
    slideSize: {
        cx: 12192000, // 16:9 widescreen: 13.333 inches
        cy: 6858000   // 7.5 inches
    },
    allowedFonts: ['KoPub돋움체 Bold', 'KoPub돋움체 Medium', 'KoPub돋움체', '맑은 고딕'],
    themeColors: ['1E3A8A', '2563EB', '3B82F6', '60A5FA', 'F1F5F9', '0F172A', 'CBD5E1', 'DC2626', '16A34A'],
    toleranceRatio: 0.08, // 허용오차 ±8%
    standardZones: {
        sectionNoAndTitle: {
            name: '좌측 상단 섹션 번호 및 제목',
            x: 1097280,  // ~1.2 inch
            y: 320040,   // ~0.35 inch
            cx: 7315200, // ~8.0 inch
            cy: 411480,  // ~0.45 inch
            sz: 2400,    // 24pt
            b: true,
            color: '1E3A8A'
        },
        governingMessage: {
            name: '상단 헤드메시지 요약문 박스',
            x: 839072,   // ~0.92 inch
            y: 1079537,  // ~1.18 inch
            cx: 10515600,// ~11.5 inch
            cy: 685800,  // ~0.75 inch
            sz: 1600,    // 16pt (또는 18pt)
            b: true,
            color: '0F172A'
        },
        breadcrumb: {
            name: '상단 대분류·중분류 네비게이션',
            x: 1097280,  // ~1.2 inch
            y: 137160,   // ~0.15 inch
            cx: 5000000, // ~5.46 inch
            cy: 228600,  // ~0.25 inch
            sz: 950,     // 9.5pt
            b: true,
            color: '64748B'
        },
        keySummaryLabels: {
            name: '고정 라벨 박스 (주요 내용, 시사점 등)',
            cx: 1787868, // ~1.95 inch
            cy: 475638,  // ~0.52 inch
            sz: 1400,    // 14pt
            b: true,
            color: '1E3A8A'
        },
        citationFooter: {
            name: '출처 표기 및 하단 푸터',
            y: 6247400,  // ~6.83 inch
            cx: 9882514, // ~10.8 inch
            cy: 365760,  // ~0.4 inch
            sz: 900,     // 9pt
            b: false,
            color: '64748B',
            prefix: '※ 출처 : '
        }
    }
};

// ── 1. LocalStorage 저장 및 로드 ──────────────────────────────────
export function loadMasterSpec() {
    try {
        const stored = localStorage.getItem(STORAGE_KEY);
        if (stored) {
            const parsed = JSON.parse(stored);
            if (parsed && parsed.allowedFonts && parsed.standardZones) {
                return parsed;
            }
        }
    } catch (e) {
        console.warn('마스터 규격 로드 실패, 기본값 사용:', e);
    }
    return DEFAULT_MASTER_SPEC;
}

export function saveMasterSpec(spec) {
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(spec));
    } catch (e) {
        console.error('마스터 규격 저장 실패:', e);
    }
}

// ── 헬퍼: XML 엘리먼트 검색 (localName 기반 브라우저 완전 호환) ─────
function findElementsByLocalName(root, targetName) {
    const list = root.getElementsByTagName('*');
    const matched = [];
    for (let i = 0; i < list.length; i++) {
        const el = list[i];
        if (el.nodeType === 1) {
            const local = el.localName || el.tagName.split(':').pop();
            if (local === targetName) {
                matched.push(el);
            }
        }
    }
    return matched;
}

function getShapeText(shapeEl) {
    const tNodes = findElementsByLocalName(shapeEl, 't');
    return tNodes.map(t => t.textContent || '').join(' ').trim();
}

function getShapeXfrm(shapeEl) {
    const xfrm = findElementsByLocalName(shapeEl, 'xfrm')[0];
    if (!xfrm) return null;
    const off = findElementsByLocalName(xfrm, 'off')[0];
    const ext = findElementsByLocalName(xfrm, 'ext')[0];
    return {
        x: off ? parseInt(off.getAttribute('x') || '0', 10) : 0,
        y: off ? parseInt(off.getAttribute('y') || '0', 10) : 0,
        cx: ext ? parseInt(ext.getAttribute('cx') || '0', 10) : 0,
        cy: ext ? parseInt(ext.getAttribute('cy') || '0', 10) : 0
    };
}

function getShapeFontSize(shapeEl) {
    const rPrList = findElementsByLocalName(shapeEl, 'rPr');
    for (const rPr of rPrList) {
        const sz = rPr.getAttribute('sz');
        if (sz) return parseInt(sz, 10);
    }
    return null;
}

function isDividerSlide(slideDoc, slideText) {
    const shapes = findElementsByLocalName(slideDoc, 'sp');
    if (shapes.length <= 3) return true;
    const lower = slideText.toLowerCase();
    if (lower.includes('chapter') || lower.includes('목차') || lower.includes('간지')) {
        return true;
    }
    return false;
}

// ── 2. 표준 양식(Golden Master) 규격 자동 추출 함수 ─────────────────
export async function extractMasterSpec(pptxFile) {
    const zip = new JSZip();
    await zip.loadAsync(await pptxFile.arrayBuffer());
    const parser = new DOMParser();

    // 1) 슬라이드 크기 추출 (ppt/presentation.xml)
    let slideSize = { cx: 12192000, cy: 6858000 };
    const presFile = zip.file('ppt/presentation.xml');
    if (presFile) {
        const presXml = await presFile.async('text');
        const presDoc = parser.parseFromString(presXml, 'application/xml');
        const sldSz = findElementsByLocalName(presDoc, 'sldSz')[0];
        if (sldSz) {
            slideSize = {
                cx: parseInt(sldSz.getAttribute('cx') || '12192000', 10),
                cy: parseInt(sldSz.getAttribute('cy') || '6858000', 10)
            };
        }
    }

    // 2) 테마 폰트 및 색상 추출 (ppt/theme/theme1.xml 또는 ppt/themes/theme1.xml)
    const allowedFontsSet = new Set(['KoPub돋움체 Bold', 'KoPub돋움체 Medium', 'KoPub돋움체']);
    const themeColorsSet = new Set(['1E3A8A', '2563EB', '0F172A', 'F1F5F9']);

    const themePath = Object.keys(zip.files).find(k => k.match(/^ppt\/theme(s)?\/theme1\.xml$/i));
    if (themePath) {
        const themeXml = await zip.file(themePath).async('text');
        const themeDoc = parser.parseFromString(themeXml, 'application/xml');

        // 테마 폰트
        const fontElements = [
            ...findElementsByLocalName(themeDoc, 'latin'),
            ...findElementsByLocalName(themeDoc, 'ea'),
            ...findElementsByLocalName(themeDoc, 'cs')
        ];
        fontElements.forEach(el => {
            const tf = el.getAttribute('typeface');
            if (tf && tf.trim() && !tf.startsWith('+') && tf.length > 1) {
                allowedFontsSet.add(tf.trim());
            }
        });

        // 테마 컬러
        const clrElements = findElementsByLocalName(themeDoc, 'srgbClr');
        clrElements.forEach(el => {
            const val = el.getAttribute('val');
            if (val && /^[0-9A-Fa-f]{6}$/.test(val)) {
                themeColorsSet.add(val.toUpperCase());
            }
        });
    }

    // 3) 전체 슬라이드 순회: 폰트 빈도수 카운트 및 5대 표준 영역 좌표 샘플링
    const slideFiles = Object.keys(zip.files)
        .filter(k => k.match(/^ppt\/slides\/slide\d+\.xml$/i))
        .sort((a, b) => {
            const na = parseInt(a.match(/slide(\d+)\.xml/)[1], 10);
            const nb = parseInt(b.match(/slide(\d+)\.xml/)[1], 10);
            return na - nb;
        });

    const fontFreqMap = {};
    const sectionTitleSamples = [];
    const governingSamples = [];
    const breadcrumbSamples = [];
    const keySummarySamples = [];
    const citationSamples = [];

    for (const slidePath of slideFiles) {
        const slideXml = await zip.file(slidePath).async('text');
        const slideDoc = parser.parseFromString(slideXml, 'application/xml');

        // 폰트 수집
        const fontTags = [
            ...findElementsByLocalName(slideDoc, 'latin'),
            ...findElementsByLocalName(slideDoc, 'ea'),
            ...findElementsByLocalName(slideDoc, 'cs')
        ];
        fontTags.forEach(el => {
            const tf = el.getAttribute('typeface');
            if (tf && tf.trim() && !tf.startsWith('+') && tf.length > 1) {
                const cleanFont = tf.trim();
                fontFreqMap[cleanFont] = (fontFreqMap[cleanFont] || 0) + 1;
            }
        });

        const slideText = findElementsByLocalName(slideDoc, 't').map(t => t.textContent || '').join(' ');
        if (isDividerSlide(slideDoc, slideText)) continue; // 간지 슬라이드는 좌표 샘플링에서 제외

        // 도형 순회하여 5대 영역 후보군 수집
        const shapes = findElementsByLocalName(slideDoc, 'sp');
        for (const sp of shapes) {
            const text = getShapeText(sp);
            const xfrm = getShapeXfrm(sp);
            const sz = getShapeFontSize(sp);
            if (!text || !xfrm) continue;

            // ① 섹션 번호 및 제목: 좌측 상단, sz >= 2000, 또는 정규식 매칭
            if ((text.match(/^\d+\.\d+(\.\d+)?\s+.+/) || (xfrm.y < slideSize.cy * 0.15 && xfrm.x < slideSize.cx * 0.35 && sz >= 2000))) {
                sectionTitleSamples.push({ ...xfrm, sz: sz || 2400 });
            }
            // ② 헤드메시지 요약문 박스: y가 상단 10~22%, cx가 가로 폭의 60% 이상
            else if (xfrm.y >= slideSize.cy * 0.10 && xfrm.y <= slideSize.cy * 0.25 && xfrm.cx >= slideSize.cx * 0.60) {
                governingSamples.push({ ...xfrm, sz: sz || 1600 });
            }
            // ③ 브레드크럼: 최상단 (y < 12%), '>' 포함 또는 분류 번호 포함
            else if (xfrm.y < slideSize.cy * 0.12 && (text.includes('>') || text.includes('환경분석') || text.includes('추진계획'))) {
                breadcrumbSamples.push({ ...xfrm, sz: sz || 950 });
            }
            // ④ 키 서머리 고정 라벨: 주요 내용, 시사점, 수행 활동
            else if (text.includes('주요 내용') || text.includes('주요내용') || text.includes('시사점') || text.includes('수행 활동')) {
                keySummarySamples.push({ ...xfrm, sz: sz || 1400 });
            }
            // ⑤ 출처/푸터: 하단 85% 이상 위치, '출처' 포함
            else if (xfrm.y > slideSize.cy * 0.82 && text.includes('출처')) {
                citationSamples.push({ ...xfrm, sz: sz || 900 });
            }
        }
    }

    // 최빈값 상위 폰트 2개 추출
    const sortedFonts = Object.entries(fontFreqMap)
        .sort((a, b) => b[1] - a[1])
        .map(entry => entry[0]);

    if (sortedFonts.length > 0) {
        sortedFonts.slice(0, 3).forEach(f => allowedFontsSet.add(f));
    }

    // 각 영역별 평균 좌표 계산 헬퍼
    const calcAvgZone = (samples, fallback) => {
        if (!samples || samples.length === 0) return fallback;
        const sum = samples.reduce((acc, s) => ({
            x: acc.x + s.x,
            y: acc.y + s.y,
            cx: acc.cx + s.cx,
            cy: acc.cy + s.cy,
            sz: acc.sz + (s.sz || fallback.sz)
        }), { x: 0, y: 0, cx: 0, cy: 0, sz: 0 });
        const len = samples.length;
        return {
            ...fallback,
            x: Math.round(sum.x / len),
            y: Math.round(sum.y / len),
            cx: Math.round(sum.cx / len),
            cy: Math.round(sum.cy / len),
            sz: Math.round(sum.sz / len)
        };
    };

    const finalSpec = {
        fileName: pptxFile.name,
        slideSize,
        allowedFonts: Array.from(allowedFontsSet),
        themeColors: Array.from(themeColorsSet).slice(0, 12),
        toleranceRatio: 0.08,
        standardZones: {
            sectionNoAndTitle: calcAvgZone(sectionTitleSamples, DEFAULT_MASTER_SPEC.standardZones.sectionNoAndTitle),
            governingMessage: calcAvgZone(governingSamples, DEFAULT_MASTER_SPEC.standardZones.governingMessage),
            breadcrumb: calcAvgZone(breadcrumbSamples, DEFAULT_MASTER_SPEC.standardZones.breadcrumb),
            keySummaryLabels: calcAvgZone(keySummarySamples, DEFAULT_MASTER_SPEC.standardZones.keySummaryLabels),
            citationFooter: calcAvgZone(citationSamples, DEFAULT_MASTER_SPEC.standardZones.citationFooter)
        }
    };

    saveMasterSpec(finalSpec);
    return finalSpec;
}

// ── 3. 산출물 표준 양식 정합성 검수 함수 ────────────────────────────
export async function inspectPptxCompliance(targetFile, masterSpec = DEFAULT_MASTER_SPEC) {
    const zip = new JSZip();
    await zip.loadAsync(await targetFile.arrayBuffer());
    const parser = new DOMParser();

    const slideFiles = Object.keys(zip.files)
        .filter(k => k.match(/^ppt\/slides\/slide\d+\.xml$/i))
        .sort((a, b) => {
            const na = parseInt(a.match(/slide(\d+)\.xml/)[1], 10);
            const nb = parseInt(b.match(/slide(\d+)\.xml/)[1], 10);
            return na - nb;
        });

    const totalSlides = slideFiles.length;
    const violations = [];
    const allowedFontsList = (masterSpec.allowedFonts || DEFAULT_MASTER_SPEC.allowedFonts).map(f => f.trim().toLowerCase());
    const tolerance = masterSpec.toleranceRatio || 0.08;
    const stdZones = masterSpec.standardZones || DEFAULT_MASTER_SPEC.standardZones;
    const slideSize = masterSpec.slideSize || DEFAULT_MASTER_SPEC.slideSize;

    // 슬라이드별 순회 검사
    for (let sIdx = 0; sIdx < slideFiles.length; sIdx++) {
        const slideNum = sIdx + 1;
        const slidePath = slideFiles[sIdx];
        const slideXml = await zip.file(slidePath).async('text');
        const slideDoc = parser.parseFromString(slideXml, 'application/xml');

        const allText = findElementsByLocalName(slideDoc, 't').map(t => t.textContent || '').join(' ');
        const isDivider = isDividerSlide(slideDoc, allText);

        // ── 1) [FONT_VIOLATION] 비표준 폰트 사용 검출 ──
        const fontElements = [
            ...findElementsByLocalName(slideDoc, 'latin'),
            ...findElementsByLocalName(slideDoc, 'ea'),
            ...findElementsByLocalName(slideDoc, 'cs')
        ];

        const reportedFontsOnSlide = new Set();
        fontElements.forEach(el => {
            const tf = el.getAttribute('typeface');
            if (tf && tf.trim() && !tf.startsWith('+') && tf.length > 1) {
                const clean = tf.trim();
                const lower = clean.toLowerCase();
                if (!allowedFontsList.includes(lower) && !reportedFontsOnSlide.has(clean)) {
                    reportedFontsOnSlide.add(clean);
                    // 인접 텍스트 발췌
                    const parentRun = el.parentNode;
                    const runText = parentRun?.parentNode ? getShapeText(parentRun.parentNode) : '';
                    violations.push({
                        id: `V-FONT-${slideNum}-${violations.length + 1}`,
                        slideNum,
                        type: 'FONT_VIOLATION',
                        typeLabel: '비표준 폰트 사용',
                        severity: 'HIGH',
                        currentValue: clean,
                        standardValue: masterSpec.allowedFonts[0] || 'KoPub돋움체 Bold',
                        snippet: runText.slice(0, 40) || '텍스트 요소',
                        reason: `표준 허용 폰트(${masterSpec.allowedFonts.slice(0, 2).join(', ')})가 아닌 비표준 폰트('${clean}')가 사용되었습니다.`,
                        canAutoFix: true
                    });
                }
            }
        });

        // ── 2) [CITATION_FORMAT_ERROR] 출처 표기 포맷 불일치 ──
        const tNodes = findElementsByLocalName(slideDoc, 't');
        tNodes.forEach(tNode => {
            const text = tNode.textContent || '';
            if (text.includes('출처')) {
                const citationMatch = text.match(/([*※\[]?[ ]*출처[ ]*[:\]]?[ ]*)/);
                if (citationMatch) {
                    const foundPrefix = citationMatch[1];
                    const standardPrefix = '※ 출처 : ';
                    if (foundPrefix !== standardPrefix && !foundPrefix.includes('※ 출처 : ')) {
                        violations.push({
                            id: `V-CITE-${slideNum}-${violations.length + 1}`,
                            slideNum,
                            type: 'CITATION_FORMAT_ERROR',
                            typeLabel: '출처 표기 불일치',
                            severity: 'MEDIUM',
                            currentValue: foundPrefix,
                            standardValue: standardPrefix,
                            snippet: text.slice(0, 50),
                            reason: `출처 표기 접두사가 표준 규격('${standardPrefix}')과 일치하지 않습니다 (현재: '${foundPrefix}').`,
                            canAutoFix: true
                        });
                    }
                }
            }
        });

        // ── 간지(Divider) 슬라이드는 구조/레이아웃 검사 제외 ──
        if (isDivider) continue;

        // ── 3) [HEADER_BREADCRUMB_ERROR] 섹션 번호 및 마침표 규칙 불일치 ──
        let foundSectionTitle = false;
        let foundBreadcrumb = false;
        const shapes = findElementsByLocalName(slideDoc, 'sp');

        shapes.forEach(sp => {
            const text = getShapeText(sp);
            if (!text) return;

            // 섹션 번호 검출 (예: 2.2, 2.2.1, 2.4 등)
            const secMatch = text.match(/^(\d+\.\d+(\.\d+)?)([\s.]+)(.+)/);
            if (secMatch) {
                foundSectionTitle = true;
                const fullSecNo = secMatch[1];
                const separator = secMatch[3];
                // 마침표 표기 규칙: "2.4."처럼 끝에 점이 있어야 하는데 없는 경우 적발
                if (!separator.includes('.')) {
                    violations.push({
                        id: `V-DOT-${slideNum}-${violations.length + 1}`,
                        slideNum,
                        type: 'HEADER_BREADCRUMB_ERROR',
                        typeLabel: '섹션 마침표 규칙 누락',
                        severity: 'MEDIUM',
                        currentValue: `${fullSecNo} `,
                        standardValue: `${fullSecNo}. `,
                        snippet: text.slice(0, 45),
                        reason: `섹션 번호 표기에서 표준 마침표 규칙("${fullSecNo}.")이 누락되어 공백만 기재되었습니다.`,
                        canAutoFix: true
                    });
                }
            }

            // 브레드크럼 검출
            if (text.includes('>') || text.includes('환경분석') || text.includes('전략수립') || text.includes('이행계획')) {
                const xfrm = getShapeXfrm(sp);
                if (xfrm && xfrm.y < slideSize.cy * 0.15) {
                    foundBreadcrumb = true;
                }
            }
        });

        if (!foundSectionTitle && slideNum > 2) {
            violations.push({
                id: `V-NOTITLE-${slideNum}-${violations.length + 1}`,
                slideNum,
                type: 'HEADER_BREADCRUMB_ERROR',
                typeLabel: '섹션 번호 및 제목 누락',
                severity: 'HIGH',
                currentValue: '미식별',
                standardValue: '표준 섹션 번호 (예: 2.2.1 ...)',
                snippet: allText.slice(0, 40),
                reason: '본문 슬라이드에서 섹션 번호 체계(예: 2.2.1 제목)를 확인할 수 없습니다.',
                canAutoFix: false
            });
        }

        // ── 4) [LAYOUT_OFFSET_ERROR] 제목 및 헤드메시지 위치/폰트크기 이탈 검사 ──
        shapes.forEach(sp => {
            const text = getShapeText(sp);
            const xfrm = getShapeXfrm(sp);
            const sz = getShapeFontSize(sp);
            if (!text || !xfrm) return;

            // 섹션 제목 박스 위치 이탈 검사
            if (text.match(/^\d+\.\d+/) || (xfrm.y < slideSize.cy * 0.15 && xfrm.x < slideSize.cx * 0.35 && sz >= 2000)) {
                const std = stdZones.sectionNoAndTitle;
                const maxDiffX = slideSize.cx * tolerance;
                const maxDiffY = slideSize.cy * tolerance;

                if (Math.abs(xfrm.x - std.x) > maxDiffX || Math.abs(xfrm.y - std.y) > maxDiffY) {
                    violations.push({
                        id: `V-OFF-TITLE-${slideNum}-${violations.length + 1}`,
                        slideNum,
                        type: 'LAYOUT_OFFSET_ERROR',
                        typeLabel: '제목 위치 이탈',
                        severity: 'LOW',
                        currentValue: `X: ${(xfrm.x / EMU_PER_INCH).toFixed(2)}" / Y: ${(xfrm.y / EMU_PER_INCH).toFixed(2)}"`,
                        standardValue: `X: ${(std.x / EMU_PER_INCH).toFixed(2)}" / Y: ${(std.y / EMU_PER_INCH).toFixed(2)}"`,
                        snippet: text.slice(0, 35),
                        reason: `섹션 제목 위치가 표준 마스터 좌표 오차(±8%)를 벗어났습니다.`,
                        canAutoFix: true
                    });
                }
            }

            // 헤드메시지 요약문 박스 이탈 검사
            if (xfrm.y >= slideSize.cy * 0.08 && xfrm.y <= slideSize.cy * 0.28 && xfrm.cx >= slideSize.cx * 0.60) {
                const std = stdZones.governingMessage;
                const maxDiffY = slideSize.cy * tolerance;
                if (Math.abs(xfrm.y - std.y) > maxDiffY) {
                    violations.push({
                        id: `V-OFF-GOV-${slideNum}-${violations.length + 1}`,
                        slideNum,
                        type: 'LAYOUT_OFFSET_ERROR',
                        typeLabel: '헤드메시지 위치 이탈',
                        severity: 'LOW',
                        currentValue: `Y: ${(xfrm.y / EMU_PER_INCH).toFixed(2)}"`,
                        standardValue: `Y: ${(std.y / EMU_PER_INCH).toFixed(2)}"`,
                        snippet: text.slice(0, 35),
                        reason: `헤드메시지 박스 수직 위치가 표준 규격(Y: ${(std.y / EMU_PER_INCH).toFixed(2)}")을 벗어났습니다.`,
                        canAutoFix: true
                    });
                }
            }
        });

        // ── 5) [KEY_FRAME_MISSING] 필수 분석 프레임 누락 검사 ──
        if (allText.includes('환경분석') || allText.includes('현황분석') || allText.includes('이슈') || allText.includes('문제점')) {
            const hasMainContent = allText.includes('주요 내용') || allText.includes('주요내용');
            const hasImplication = allText.includes('시사점');
            if (!hasMainContent && !hasImplication) {
                violations.push({
                    id: `V-FRAME-${slideNum}-${violations.length + 1}`,
                    slideNum,
                    type: 'KEY_FRAME_MISSING',
                    typeLabel: '필수 분석 프레임 누락',
                    severity: 'MEDIUM',
                    currentValue: '주요 내용/시사점 미포함',
                    standardValue: '‘주요 내용’ 또는 ‘시사점’ 라벨 박스',
                    snippet: allText.slice(0, 45),
                    reason: '환경분석/현황분석 슬라이드에는 표준 프레임워크(주요 내용 또는 시사점)가 명시되어야 합니다.',
                    canAutoFix: false
                });
            }
        }

        // ── 6) [TABLE_STYLE_VIOLATION] 표 서식 불일치 검사 ──
        const tables = findElementsByLocalName(slideDoc, 'tbl');
        if (tables.length > 0) {
            tables.forEach((tbl, tIdx) => {
                const cellFonts = findElementsByLocalName(tbl, 'ea');
                cellFonts.forEach(cf => {
                    const tf = cf.getAttribute('typeface');
                    if (tf && !allowedFontsList.includes(tf.toLowerCase()) && !tf.startsWith('+')) {
                        violations.push({
                            id: `V-TBL-FONT-${slideNum}-${tIdx + 1}`,
                            slideNum,
                            type: 'TABLE_STYLE_VIOLATION',
                            typeLabel: '표 내부 비표준 폰트',
                            severity: 'MEDIUM',
                            currentValue: tf,
                            standardValue: masterSpec.allowedFonts[0] || 'KoPub돋움체',
                            snippet: `표 #${tIdx + 1} 셀 텍스트`,
                            reason: `표 내부 셀 텍스트에 비표준 폰트('${tf}')가 적용되었습니다.`,
                            canAutoFix: true
                        });
                    }
                });
            });
        }
    }

    // 종합 서식 준수율(%) 산출 (100점 만점 기준 위반 페널티 차감)
    const penaltyTotal = violations.reduce((acc, v) => {
        if (v.severity === 'HIGH') return acc + 4;
        if (v.severity === 'MEDIUM') return acc + 2;
        return acc + 1;
    }, 0);
    const complianceScore = Math.max(0, Math.min(100, Math.round(100 - (penaltyTotal / Math.max(1, totalSlides)) * 10)));

    const violationsSummary = {
        fontViolations: violations.filter(v => v.type === 'FONT_VIOLATION').length,
        layoutOffsetErrors: violations.filter(v => v.type === 'LAYOUT_OFFSET_ERROR').length,
        headerBreadcrumbErrors: violations.filter(v => v.type === 'HEADER_BREADCRUMB_ERROR').length,
        citationErrors: violations.filter(v => v.type === 'CITATION_FORMAT_ERROR').length,
        keyFrameMissing: violations.filter(v => v.type === 'KEY_FRAME_MISSING').length,
        tableStyleViolations: violations.filter(v => v.type === 'TABLE_STYLE_VIOLATION').length,
        totalViolations: violations.length
    };

    return {
        filename: targetFile.name,
        fileSize: targetFile.size,
        totalSlides,
        complianceScore,
        violationsSummary,
        violations,
        checkedAt: new Date().toISOString()
    };
}

// ── 4. 원본 보존형 원클릭 서식 자동 교정 함수 (XML 직접 수정) ────────
export async function autoFixPptxFormatting(targetFile, masterSpec = DEFAULT_MASTER_SPEC, fixOptions = {}) {
    const {
        fixFonts = true,
        fixCitations = true,
        fixPositions = true,
        fixBreadcrumbDots = true,
        selectedViolationIds = null // null이면 전체 자동 교정 대상 적용
    } = fixOptions;

    const zip = new JSZip();
    await zip.loadAsync(await targetFile.arrayBuffer());
    const parser = new DOMParser();
    const serializer = new XMLSerializer();

    const slideFiles = Object.keys(zip.files)
        .filter(k => k.match(/^ppt\/slides\/slide\d+\.xml$/i));

    const defaultFontBold = masterSpec.allowedFonts[0] || 'KoPub돋움체 Bold';
    const defaultFontMedium = masterSpec.allowedFonts[1] || masterSpec.allowedFonts[0] || 'KoPub돋움체 Medium';
    const stdZones = masterSpec.standardZones || DEFAULT_MASTER_SPEC.standardZones;
    const slideSize = masterSpec.slideSize || DEFAULT_MASTER_SPEC.slideSize;

    let fixedCount = 0;

    for (const slidePath of slideFiles) {
        const slideXml = await zip.file(slidePath).async('text');
        const slideDoc = parser.parseFromString(slideXml, 'application/xml');
        let modified = false;

        // 1) 폰트 일괄 교정 (fixFonts)
        if (fixFonts) {
            const fontTags = [
                ...findElementsByLocalName(slideDoc, 'latin'),
                ...findElementsByLocalName(slideDoc, 'ea'),
                ...findElementsByLocalName(slideDoc, 'cs')
            ];

            const allowedLower = (masterSpec.allowedFonts || []).map(f => f.trim().toLowerCase());

            fontTags.forEach(el => {
                const tf = el.getAttribute('typeface');
                if (tf && tf.trim() && !tf.startsWith('+') && !allowedLower.includes(tf.trim().toLowerCase())) {
                    // 상위 rPr에서 bold 여부 확인
                    const rPr = el.parentNode;
                    const isBold = rPr && (rPr.getAttribute('b') === '1' || rPr.getAttribute('b') === 'true');
                    const targetFont = isBold ? defaultFontBold : defaultFontMedium;

                    el.setAttribute('typeface', targetFont);
                    modified = true;
                    fixedCount++;
                }
            });
        }

        // 2) 출처 표기 일괄 교정 (fixCitations)
        if (fixCitations) {
            const tNodes = findElementsByLocalName(slideDoc, 't');
            tNodes.forEach(tNode => {
                const raw = tNode.textContent || '';
                if (raw.includes('출처')) {
                    // *출처:, ※출처 :, [출처], 출처 : 등을 "※ 출처 : "로 정규식 치환
                    const replaced = raw.replace(/([*※\[]?[ ]*출처[ ]*[:\]]?[ ]*)/g, '※ 출처 : ');
                    if (replaced !== raw) {
                        tNode.textContent = replaced;
                        modified = true;
                        fixedCount++;
                    }
                }
            });
        }

        // 3) 섹션 마침표 규칙 일괄 교정 (fixBreadcrumbDots)
        if (fixBreadcrumbDots) {
            const tNodes = findElementsByLocalName(slideDoc, 't');
            tNodes.forEach(tNode => {
                const raw = tNode.textContent || '';
                // 예: "2.4 내부환경 분석" ➔ "2.4. 내부환경 분석"
                if (/^(\d+\.\d+)(\s+[가-힣A-Za-z])/.test(raw.trim())) {
                    const replaced = raw.replace(/^(\d+\.\d+)(\s+)/, '$1.$2');
                    if (replaced !== raw) {
                        tNode.textContent = replaced;
                        modified = true;
                        fixedCount++;
                    }
                }
            });
        }

        // 4) 위치 및 폰트 크기 정밀 스냅 정렬 (fixPositions)
        if (fixPositions) {
            const shapes = findElementsByLocalName(slideDoc, 'sp');
            shapes.forEach(sp => {
                const text = getShapeText(sp);
                const xfrm = getShapeXfrm(sp);
                if (!text || !xfrm) return;

                // 제목 박스 스냅
                if (text.match(/^\d+\.\d+/) || (xfrm.y < slideSize.cy * 0.15 && xfrm.x < slideSize.cx * 0.35)) {
                    const std = stdZones.sectionNoAndTitle;
                    // 근접 범위 (±20% 이내)에 있는 경우 표준 좌표값으로 정밀 스냅
                    if (Math.abs(xfrm.x - std.x) < slideSize.cx * 0.20 && Math.abs(xfrm.y - std.y) < slideSize.cy * 0.20) {
                        const off = findElementsByLocalName(findElementsByLocalName(sp, 'xfrm')[0] || sp, 'off')[0];
                        if (off) {
                            off.setAttribute('x', String(std.x));
                            off.setAttribute('y', String(std.y));
                            modified = true;
                            fixedCount++;
                        }
                    }
                }

                // 헤드메시지 박스 스냅
                if (xfrm.y >= slideSize.cy * 0.08 && xfrm.y <= slideSize.cy * 0.28 && xfrm.cx >= slideSize.cx * 0.60) {
                    const std = stdZones.governingMessage;
                    if (Math.abs(xfrm.y - std.y) < slideSize.cy * 0.20) {
                        const off = findElementsByLocalName(findElementsByLocalName(sp, 'xfrm')[0] || sp, 'off')[0];
                        if (off) {
                            off.setAttribute('y', String(std.y));
                            modified = true;
                            fixedCount++;
                        }
                    }
                }
            });
        }

        if (modified) {
            const newSlideXml = serializer.serializeToString(slideDoc);
            zip.file(slidePath, newSlideXml);
        }
    }

    const fixedBlob = await zip.generateAsync({ type: 'blob' });
    const cleanOrigName = targetFile.name.replace(/\.[^/.]+$/, '');
    const outName = `표준교정완료_${cleanOrigName}.pptx`;
    saveAs(fixedBlob, outName);

    return {
        fixedCount,
        blob: fixedBlob,
        fileName: outName
    };
}

// ── 5. 서식 검수 결과 리포트 엑셀 내보내기 (exceljs) ────────────────
export async function exportMasterComplianceReportToExcel(inspectionResult, masterSpec = DEFAULT_MASTER_SPEC) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = '건강한 프로젝트 (Master Template Inspector)';
    workbook.created = new Date();
    const tokens = loadDesignTokens();
    const colors = tokens?.colors || { primary: '1E3A8A', tableHeaderBg: 'F1F5F9', borderColor: 'CBD5E1' };

    const headerStyle = {
        font: { name: '맑은 고딕', size: 10, bold: true, color: { argb: 'FFFFFFFF' } },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${colors.primary}` } },
        alignment: { vertical: 'middle', horizontal: 'center' }
    };

    const borderStyle = {
        top: { style: 'thin', color: { argb: `FF${colors.borderColor}` } },
        left: { style: 'thin', color: { argb: `FF${colors.borderColor}` } },
        bottom: { style: 'thin', color: { argb: `FF${colors.borderColor}` } },
        right: { style: 'thin', color: { argb: `FF${colors.borderColor}` } }
    };

    // 시트 1: 검수 종합 요약표
    const s1 = workbook.addWorksheet('①_마스터_서식_검수_요약');
    s1.columns = [
        { header: '검수 항목', width: 28 },
        { header: '검수 결과 및 기준값', width: 50 },
        { header: '상태 / 조치', width: 20 }
    ];
    s1.getRow(1).height = 26;
    s1.getRow(1).eachCell(c => { c.style = headerStyle; });

    const summaryData = [
        ['검수 대상 파일명', inspectionResult.filename, '검수 완료'],
        ['총 슬라이드 수', `${inspectionResult.totalSlides} 장`, '-'],
        ['서식 준수율(%)', `${inspectionResult.complianceScore} 점`, inspectionResult.complianceScore >= 80 ? '양호' : '개선 필요'],
        ['총 위반 건수', `${inspectionResult.violationsSummary.totalViolations} 건`, inspectionResult.violationsSummary.totalViolations === 0 ? '적합' : '교정 권장'],
        ['비표준 폰트 위반', `${inspectionResult.violationsSummary.fontViolations} 건`, '원클릭 자동교정 가능'],
        ['제목/헤드메시지 좌표 이탈', `${inspectionResult.violationsSummary.layoutOffsetErrors} 건`, '표준 좌표 스냅 가능'],
        ['섹션 번호 및 마침표 오류', `${inspectionResult.violationsSummary.headerBreadcrumbErrors} 건`, '마침표 일괄 통일 가능'],
        ['출처 표기 불일치', `${inspectionResult.violationsSummary.citationErrors} 건`, '※ 출처 : 일괄 교정 가능'],
        ['필수 분석 프레임 누락', `${inspectionResult.violationsSummary.keyFrameMissing} 건`, '내용 보완 필요'],
        ['표 서식 불일치', `${inspectionResult.violationsSummary.tableStyleViolations} 건`, '표 폰트 교정 가능'],
        ['적용된 표준 폰트 규격', masterSpec.allowedFonts.join(', '), '골든 마스터'],
        ['검수 일시', new Date(inspectionResult.checkedAt).toLocaleString(), '-']
    ];

    summaryData.forEach(row => {
        const r = s1.addRow(row);
        r.height = 24;
        r.alignment = { vertical: 'middle' };
        r.eachCell(c => { c.border = borderStyle; });
    });

    // 시트 2: 슬라이드별 상세 위반 내역표
    const s2 = workbook.addWorksheet('②_슬라이드별_상세_위반_내역');
    s2.columns = [
        { header: 'No.', width: 6 },
        { header: '슬라이드', width: 12 },
        { header: '위반 유형', width: 20 },
        { header: '심각도', width: 10 },
        { header: '현재 기재값 (발췌)', width: 35 },
        { header: '표준 기준값', width: 30 },
        { header: '세부 판정 사유', width: 45 },
        { header: '자동교정', width: 12 }
    ];
    s2.getRow(1).height = 26;
    s2.getRow(1).eachCell(c => { c.style = headerStyle; });

    inspectionResult.violations.forEach((v, idx) => {
        const r = s2.addRow([
            idx + 1,
            `p.${v.slideNum}`,
            v.typeLabel || v.type,
            v.severity,
            v.currentValue,
            v.standardValue,
            v.reason,
            v.canAutoFix ? '가능' : '수동확인'
        ]);
        r.height = 24;
        r.alignment = { vertical: 'middle', wrapText: true };
        r.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };
        r.getCell(2).alignment = { vertical: 'middle', horizontal: 'center' };
        r.getCell(4).alignment = { vertical: 'middle', horizontal: 'center' };
        r.getCell(8).alignment = { vertical: 'middle', horizontal: 'center' };
        r.eachCell(c => { c.border = borderStyle; });
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const cleanName = inspectionResult.filename.replace(/\.[^/.]+$/, '');
    saveAs(blob, `마스터_서식검수리포트_${cleanName}_${new Date().toISOString().slice(0, 10)}.xlsx`);
}
