import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

/**
 * 두 문자열 사이의 유사도 계산 (Levenshtein Distance 기반 0.0 ~ 1.0)
 */
export function calculateSimilarity(s1, s2) {
    if (!s1 && !s2) return 1.0;
    if (!s1 || !s2) return 0.0;
    if (s1 === s2) return 1.0;

    const longer = s1.length > s2.length ? s1 : s2;
    const shorter = s1.length > s2.length ? s2 : s1;
    if (longer.length === 0) return 1.0;

    // 긴 문자열 길이 차이가 너무 크면 조기 반환
    if (shorter.length / longer.length < 0.2) return 0.1;

    // 간이 Levenshtein (길이 제한 2000자)
    const maxLen = 1500;
    const str1 = s1.slice(0, maxLen);
    const str2 = s2.slice(0, maxLen);

    const costs = [];
    for (let i = 0; i <= str1.length; i++) {
        let lastValue = i;
        for (let j = 0; j <= str2.length; j++) {
            if (i === 0) {
                costs[j] = j;
            } else if (j > 0) {
                let newValue = costs[j - 1];
                if (str1.charAt(i - 1) !== str2.charAt(j - 1)) {
                    newValue = Math.min(Math.min(newValue, lastValue), costs[j]) + 1;
                }
                costs[j - 1] = lastValue;
                lastValue = newValue;
            }
        }
        if (i > 0) costs[str2.length] = lastValue;
    }

    const distance = costs[str2.length];
    return Math.max(0, (longer.length - distance) / longer.length);
}

/**
 * 단어 단위 인라인 Diff 토큰 생성
 * [{ type: 'equal' | 'delete' | 'insert', text: string }]
 */
export function computeInlineDiff(textA, textB) {
    if (textA === textB) {
        return [{ type: 'equal', text: textA }];
    }
    if (!textA) {
        return [{ type: 'insert', text: textB }];
    }
    if (!textB) {
        return [{ type: 'delete', text: textA }];
    }

    // 단어 및 공백 분리
    const wordsA = textA.split(/(\s+|[.,!?;:()[\]{}'"])/).filter(Boolean);
    const wordsB = textB.split(/(\s+|[.,!?;:()[\]{}'"])/).filter(Boolean);

    // LCS (Longest Common Subsequence) 테이블 구성
    const n = Math.min(wordsA.length, 300);
    const m = Math.min(wordsB.length, 300);
    const matrix = Array.from({ length: n + 1 }, () => new Uint16Array(m + 1));

    for (let i = 1; i <= n; i++) {
        for (let j = 1; j <= m; j++) {
            if (wordsA[i - 1] === wordsB[j - 1]) {
                matrix[i][j] = matrix[i - 1][j - 1] + 1;
            } else {
                matrix[i][j] = Math.max(matrix[i - 1][j], matrix[i][j - 1]);
            }
        }
    }

    // 역추적으로 diff 토큰 생성
    let i = n;
    let j = m;
    const rawTokens = [];

    while (i > 0 || j > 0) {
        if (i > 0 && j > 0 && wordsA[i - 1] === wordsB[j - 1]) {
            rawTokens.unshift({ type: 'equal', text: wordsA[i - 1] });
            i--;
            j--;
        } else if (j > 0 && (i === 0 || matrix[i][j - 1] >= matrix[i - 1][j])) {
            rawTokens.unshift({ type: 'insert', text: wordsB[j - 1] });
            j--;
        } else if (i > 0 && (j === 0 || matrix[i][j - 1] < matrix[i - 1][j])) {
            rawTokens.unshift({ type: 'delete', text: wordsA[i - 1] });
            i--;
        }
    }

    // 남은 뒷부분 처리 (길이 제한 300 초과분)
    if (wordsA.length > n) {
        rawTokens.push({ type: 'delete', text: wordsA.slice(n).join('') });
    }
    if (wordsB.length > m) {
        rawTokens.push({ type: 'insert', text: wordsB.slice(m).join('') });
    }

    // 연속된 동일 타입 토큰 병합
    const merged = [];
    for (const t of rawTokens) {
        if (merged.length > 0 && merged[merged.length - 1].type === t.type) {
            merged[merged.length - 1].text += t.text;
        } else {
            merged.push({ ...t });
        }
    }

    return merged;
}

/**
 * 인라인 diff 토큰에서 삭제/추가 요약 텍스트 생성
 */
export function summarizeDiffTokens(tokens) {
    const deletedParts = tokens.filter(t => t.type === 'delete').map(t => t.text.trim()).filter(Boolean);
    const insertedParts = tokens.filter(t => t.type === 'insert').map(t => t.text.trim()).filter(Boolean);

    const summaries = [];
    if (deletedParts.length > 0) {
        const dText = deletedParts.join(', ');
        summaries.push(`삭제: "${dText.length > 60 ? dText.slice(0, 57) + '...' : dText}"`);
    }
    if (insertedParts.length > 0) {
        const iText = insertedParts.join(', ');
        summaries.push(`추가: "${iText.length > 60 ? iText.slice(0, 57) + '...' : iText}"`);
    }

    return summaries.length > 0 ? summaries.join(' | ') : '내용 부분 수정';
}

/**
 * 파일 확장자 기반 페이지/위치 레이블 생성기
 */
export function formatPageLabel(ext, pageNum, rawText = '') {
    const normalizedExt = (ext || '').toLowerCase();
    if (normalizedExt === 'pptx') {
        return `슬라이드 ${pageNum}`;
    }
    if (normalizedExt === 'xlsx' || normalizedExt === 'xls') {
        const sheetMatch = rawText.match(/\[시트:\s*([^\]]+)\]/);
        if (sheetMatch) {
            return `시트: ${sheetMatch[1].trim()}`;
        }
        return `시트 ${pageNum}`;
    }
    if (normalizedExt === 'hwpx') {
        return `${pageNum}페이지`;
    }
    if (normalizedExt === 'pdf') {
        return `${pageNum}페이지`;
    }
    if (normalizedExt === 'docx') {
        return `${pageNum}페이지`;
    }
    return `${pageNum}페이지`;
}

/**
 * 두 개의 파싱된 구조화 문서 데이터({ text, pages, pageMap, fileName, ext })를 비교
 */
export function compareDocuments(docA, docB) {
    const changes = [];
    let counter = 1;

    // 문서 A의 모든 단락 수집
    const itemsA = [];
    const pageKeysA = Object.keys(docA.pageMap || {}).map(Number).sort((a, b) => a - b);
    for (const pNum of pageKeysA) {
        const paragraphs = docA.pageMap[pNum] || [];
        paragraphs.forEach((p, idx) => {
            const trimmed = String(p || '').trim();
            if (trimmed.length > 0) {
                itemsA.push({
                    page: pNum,
                    pageLabel: formatPageLabel(docA.ext, pNum, trimmed),
                    lineIndex: idx + 1,
                    text: trimmed,
                    used: false
                });
            }
        });
    }

    // 문서 B의 모든 단락 수집
    const itemsB = [];
    const pageKeysB = Object.keys(docB.pageMap || {}).map(Number).sort((a, b) => a - b);
    for (const pNum of pageKeysB) {
        const paragraphs = docB.pageMap[pNum] || [];
        paragraphs.forEach((p, idx) => {
            const trimmed = String(p || '').trim();
            if (trimmed.length > 0) {
                itemsB.push({
                    page: pNum,
                    pageLabel: formatPageLabel(docB.ext, pNum, trimmed),
                    lineIndex: idx + 1,
                    text: trimmed,
                    used: false
                });
            }
        });
    }

    // 1단계: 완전 일치 단락 마킹 (무변경 항목)
    for (const itemA of itemsA) {
        // 동일 페이지 또는 근처 페이지에서 완전 일치 항목 먼저 탐색
        let matchB = itemsB.find(b => !b.used && b.page === itemA.page && b.text === itemA.text);
        if (!matchB) {
            matchB = itemsB.find(b => !b.used && b.text === itemA.text);
        }
        if (matchB) {
            itemA.used = true;
            matchB.used = true;
        }
    }

    // 2단계: 변경된 단락(수정) 매칭 (유사도 0.35 이상)
    for (const itemA of itemsA) {
        if (itemA.used) continue;

        let bestMatch = null;
        let bestScore = 0;

        for (const itemB of itemsB) {
            if (itemB.used) continue;

            // 페이지 차이가 너무 크지 않은 경우 우선순위 가산점
            const pageDiff = Math.abs(itemA.page - itemB.page);
            if (pageDiff > 5 && itemsA.length > 20) continue; // 큰 문서에서 너무 먼 페이지는 스킵

            const sim = calculateSimilarity(itemA.text, itemB.text);
            const pageBonus = (pageDiff === 0) ? 0.08 : (pageDiff <= 1 ? 0.04 : 0);
            const totalScore = sim + pageBonus;

            if (sim >= 0.35 && totalScore > bestScore) {
                bestScore = totalScore;
                bestMatch = itemB;
            }
        }

        if (bestMatch && bestScore >= 0.40) {
            itemA.used = true;
            bestMatch.used = true;

            const inlineDiff = computeInlineDiff(itemA.text, bestMatch.text);
            const summary = summarizeDiffTokens(inlineDiff);

            changes.push({
                id: `diff-${counter++}`,
                type: 'modified',
                typeLabel: '수정',
                fileA: {
                    fileName: docA.fileName,
                    page: itemA.page,
                    pageLabel: itemA.pageLabel,
                    lineIndex: itemA.lineIndex,
                    text: itemA.text
                },
                fileB: {
                    fileName: docB.fileName,
                    page: bestMatch.page,
                    pageLabel: bestMatch.pageLabel,
                    lineIndex: bestMatch.lineIndex,
                    text: bestMatch.text
                },
                inlineDiff,
                summary
            });
        }
    }

    // 3단계: 매칭되지 않은 문서 A의 단락 -> '삭제'로 분류
    for (const itemA of itemsA) {
        if (itemA.used) continue;
        changes.push({
            id: `diff-${counter++}`,
            type: 'deleted',
            typeLabel: '삭제',
            fileA: {
                fileName: docA.fileName,
                page: itemA.page,
                pageLabel: itemA.pageLabel,
                lineIndex: itemA.lineIndex,
                text: itemA.text
            },
            fileB: {
                fileName: docB.fileName,
                page: null,
                pageLabel: '-',
                lineIndex: null,
                text: '-'
            },
            inlineDiff: [{ type: 'delete', text: itemA.text }],
            summary: `내용 삭제됨 (${itemA.text.length > 50 ? itemA.text.slice(0, 47) + '...' : itemA.text})`
        });
    }

    // 4단계: 매칭되지 않은 문서 B의 단락 -> '추가'로 분류
    for (const itemB of itemsB) {
        if (itemB.used) continue;
        changes.push({
            id: `diff-${counter++}`,
            type: 'added',
            typeLabel: '추가',
            fileA: {
                fileName: docA.fileName,
                page: null,
                pageLabel: '-',
                lineIndex: null,
                text: '-'
            },
            fileB: {
                fileName: docB.fileName,
                page: itemB.page,
                pageLabel: itemB.pageLabel,
                lineIndex: itemB.lineIndex,
                text: itemB.text
            },
            inlineDiff: [{ type: 'insert', text: itemB.text }],
            summary: `신규 내용 추가됨 (${itemB.text.length > 50 ? itemB.text.slice(0, 47) + '...' : itemB.text})`
        });
    }

    // 정렬: 기준 문서 페이지 번호 및 비교 문서 페이지 번호 순으로 안정적 정렬
    changes.sort((a, b) => {
        const pageA1 = a.fileA.page || a.fileB.page || 0;
        const pageB1 = b.fileA.page || b.fileB.page || 0;
        if (pageA1 !== pageB1) return pageA1 - pageB1;
        const lineA = a.fileA.lineIndex || a.fileB.lineIndex || 0;
        const lineB = b.fileA.lineIndex || b.fileB.lineIndex || 0;
        return lineA - lineB;
    });

    // 재정렬 후 순번 재할당
    changes.forEach((c, idx) => {
        c.index = idx + 1;
    });

    // 통계 계산
    const stats = {
        total: changes.length,
        modified: changes.filter(c => c.type === 'modified').length,
        added: changes.filter(c => c.type === 'added').length,
        deleted: changes.filter(c => c.type === 'deleted').length,
        affectedPagesA: new Set(changes.filter(c => c.fileA.page).map(c => c.fileA.page)).size,
        affectedPagesB: new Set(changes.filter(c => c.fileB.page).map(c => c.fileB.page)).size
    };

    return {
        docA: {
            fileName: docA.fileName,
            ext: docA.ext,
            pages: docA.pages,
            totalItems: itemsA.length
        },
        docB: {
            fileName: docB.fileName,
            ext: docB.ext,
            pages: docB.pages,
            totalItems: itemsB.length
        },
        stats,
        changes,
        comparedAt: new Date().toISOString()
    };
}

/**
 * 비교 결과를 서식화된 엑셀(.xlsx) 파일로 내보내기
 */
export async function exportDiffToExcel(diffResult) {
    const wb = new ExcelJS.Workbook();
    wb.creator = '건강한 프로젝트 AI 교정교열 시스템';
    wb.created = new Date();

    const { docA, docB, stats, changes } = diffResult;

    // ── 시트 1: 문서 변경 내역 비교표 ────────────────────────────────
    const ws1 = wb.addWorksheet('문서_변경내역_비교표', {
        views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }]
    });

    ws1.columns = [
        { header: '순번', key: 'index', width: 7 },
        { header: '변경 구분', key: 'typeLabel', width: 12 },
        { header: '원본(A) 파일명', key: 'fileAName', width: 42 }, // 💡 300픽셀 상당 ((300-5)/7 ≈ 42.1)
        { header: '원본(A) 페이지/위치', key: 'fileAPage', width: 18 },
        { header: '원본(A) 내용 (수정 전)', key: 'fileAText', width: 45 },
        { header: '비교본(B) 파일명', key: 'fileBName', width: 42 }, // 💡 300픽셀 상당
        { header: '비교본(B) 페이지/위치', key: 'fileBPage', width: 18 },
        { header: '비교본(B) 변경 내용 (수정 후)', key: 'fileBText', width: 45 },
        { header: '상세 변경 차이점 (요약)', key: 'summary', width: 38 }
    ];

    // 헤더 행 스타일링
    const headerRow = ws1.getRow(1);
    headerRow.height = 30;
    headerRow.font = { name: 'Malgun Gothic', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF1E293B' } // 다크 슬레이트 블루
    };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' };

    // 데이터 행 추가
    changes.forEach((c) => {
        const row = ws1.addRow({
            index: c.index,
            typeLabel: c.typeLabel,
            fileAName: docA.fileName,
            fileAPage: c.fileA.pageLabel,
            fileAText: c.fileA.text,
            fileBName: docB.fileName,
            fileBPage: c.fileB.pageLabel,
            fileBText: c.fileB.text,
            summary: c.summary
        });

        row.font = { name: 'Malgun Gothic', size: 10 };
        row.alignment = { vertical: 'top', wrapText: true };

        // 순번, 변경구분, 페이지 중앙 정렬
        row.getCell('index').alignment = { vertical: 'top', horizontal: 'center' };
        row.getCell('typeLabel').alignment = { vertical: 'top', horizontal: 'center' };
        row.getCell('fileAPage').alignment = { vertical: 'top', horizontal: 'center' };
        row.getCell('fileBPage').alignment = { vertical: 'top', horizontal: 'center' };

        // 변경 구분별 셀 색상 강조
        const typeCell = row.getCell('typeLabel');
        typeCell.font = { name: 'Malgun Gothic', size: 10, bold: true };
        if (c.type === 'modified') {
            typeCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } }; // 연한 주황
            typeCell.font.color = { argb: 'FFB45309' };
        } else if (c.type === 'added') {
            typeCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } }; // 연한 초록
            typeCell.font.color = { argb: 'FF15803D' };
        } else if (c.type === 'deleted') {
            typeCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } }; // 연한 빨강
            typeCell.font.color = { argb: 'FFB91C1C' };
        }

        // 테두리 적용
        row.eachCell({ includeEmpty: true }, (cell) => {
            cell.border = {
                top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
            };
        });
    });

    // 자동 필터 적용
    ws1.autoFilter = {
        from: { row: 1, column: 1 },
        to: { row: 1, column: 9 }
    };

    // ── 시트 2: 비교_요약_보고서 ────────────────────────────────────
    const ws2 = wb.addWorksheet('비교_요약_보고서');
    ws2.columns = [
        { header: '구분 항목', key: 'item', width: 25 },
        { header: '내용 및 수치', key: 'value', width: 60 }
    ];

    const summaryHeader = ws2.getRow(1);
    summaryHeader.height = 28;
    summaryHeader.font = { name: 'Malgun Gothic', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    summaryHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
    summaryHeader.alignment = { vertical: 'middle', horizontal: 'center' };

    const summaryRows = [
        { item: '비교 실행 일시', value: new Date().toLocaleString('ko-KR') },
        { item: '원본(A) 문서명', value: `${docA.fileName} (${docA.ext.toUpperCase()}, 총 ${docA.pages}페이지/슬라이드/시트)` },
        { item: '비교본(B) 문서명', value: `${docB.fileName} (${docB.ext.toUpperCase()}, 총 ${docB.pages}페이지/슬라이드/시트)` },
        { item: '총 변경 건수', value: `${stats.total}건` },
        { item: ' - 수정 (Modified)', value: `${stats.modified}건 (기존 문구/수치 변경)` },
        { item: ' - 추가 (Added)', value: `${stats.added}건 (비교본에 신규 추가된 내용)` },
        { item: ' - 삭제 (Deleted)', value: `${stats.deleted}건 (원본에서 제거된 내용)` },
        { item: '원본 영향 페이지 수', value: `${stats.affectedPagesA}개 페이지/슬라이드` },
        { item: '비교본 영향 페이지 수', value: `${stats.affectedPagesB}개 페이지/슬라이드` },
        { item: '시스템 엔진', value: '건강한 프로젝트 AI 교정교열 문서 변경점 추적 엔진 (v2.16.0)' }
    ];

    summaryRows.forEach(sr => {
        const row = ws2.addRow(sr);
        row.height = 24;
        row.font = { name: 'Malgun Gothic', size: 10 };
        row.getCell('item').font = { name: 'Malgun Gothic', size: 10, bold: true };
        row.getCell('item').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        row.eachCell({ includeEmpty: true }, (cell) => {
            cell.border = {
                top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
                right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
            };
            cell.alignment = { vertical: 'middle' };
        });
    });

    const buffer = await wb.xlsx.writeBuffer();
    const cleanNameA = (docA.fileName || '원본').replace(/\.[^.]+$/, '');
    const cleanNameB = (docB.fileName || '비교본').replace(/\.[^.]+$/, '');
    const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    const exportFileName = `문서변경비교_${cleanNameA}_vs_${cleanNameB}_${dateStr}.xlsx`;

    saveAs(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), exportFileName);
}

/**
 * Gemini 또는 LLM을 통해 변경점에 대한 고차원 요약 브리핑 생성
 */
export async function generateAiDiffSummary(diffResult, apiKey, llmProvider = 'gemini', omniRouteModel = 'auto') {
    if (!apiKey) {
        return null;
    }

    const { docA, docB, stats, changes } = diffResult;
    if (changes.length === 0) {
        return '두 문서 사이에 텍스트 변경점이 발견되지 않았습니다. 내용이 100% 동일합니다.';
    }

    // 변경점 중 핵심 25개 샘플링
    const sampleChanges = changes.slice(0, 25).map((c, i) => {
        return `${i + 1}. [${c.typeLabel}] 원본(${c.fileA.pageLabel}): "${c.fileA.text.slice(0, 80)}" -> 비교본(${c.fileB.pageLabel}): "${c.fileB.text.slice(0, 80)}"`;
    }).join('\n');

    const prompt = `당신은 공공/기업 IT 및 보건의료 문서 전문 감리 및 검수 전문가입니다.
두 문서([원본] ${docA.fileName} vs [비교본] ${docB.fileName})의 변경 내역을 분석하여, 실무진과 의사결정권자가 한눈에 파악할 수 있도록 주요 변경점 브리핑 보고서를 작성해 주세요.

[변경 통계]
- 총 변경 건수: ${stats.total}건 (수정: ${stats.modified}건, 추가: ${stats.added}건, 삭제: ${stats.deleted}건)
- 원본 영향 페이지: ${stats.affectedPagesA}개, 비교본 영향 페이지: ${stats.affectedPagesB}개

[주요 변경 항목 샘플]
${sampleChanges}

[작성 요구사항]
1. 핵심 변경 사항 요약 (사업 범위, 수치/일정/금액, 추가된 주요 정책이나 기능, 삭제된 항목 등)
2. 3~5개의 핵심 불렛포인트로 작성
3. 간결하고 전문적인 한국어 비즈니스 보고서 톤으로 작성
4. 마크다운 형식으로 작성`;

    try {
        let endpoint = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent';
        let headers = { 'Content-Type': 'application/json' };
        let body = {};

        if (llmProvider === 'omniroute') {
            endpoint = 'https://api.omniroute.ai/v1/chat/completions';
            headers['Authorization'] = `Bearer ${apiKey}`;
            body = {
                model: omniRouteModel === 'auto' ? 'gemini-2.5-flash' : omniRouteModel,
                messages: [{ role: 'user', content: prompt }],
                temperature: 0.2
            };
        } else {
            endpoint = `${endpoint}?key=${apiKey}`;
            body = {
                contents: [{ parts: [{ text: prompt }] }],
                generationConfig: { temperature: 0.2, maxOutputTokens: 1200 }
            };
        }

        const res = await fetch(endpoint, {
            method: 'POST',
            headers,
            body: JSON.stringify(body)
        });

        if (!res.ok) {
            console.warn('AI Diff Summary API 호출 실패:', res.statusText);
            return null;
        }

        const data = await res.json();
        let summaryText = '';
        if (llmProvider === 'omniroute') {
            summaryText = data?.choices?.[0]?.message?.content || '';
        } else {
            summaryText = data?.candidates?.[0]?.content?.parts?.[0]?.text || '';
        }
        return summaryText.trim() || null;
    } catch (e) {
        console.warn('generateAiDiffSummary 에러:', e);
        return null;
    }
}
