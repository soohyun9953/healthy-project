import JSZip from 'jszip';
import * as docx from 'docx';
import ExcelJS from 'exceljs';
import PptxGenJS from 'pptxgenjs';
import { saveAs } from 'file-saver';
import { extractStructuredFileContent, getFileExtension } from './fileExtractor.js';
import { loadDesignTokens } from './designTokens.js';
import { loadSSOTData, saveSSOTData } from './ssotValidator.js';

// ── 1. 파일명 메타데이터 자동 추출기 ────────────────────────────────
export function parseInterviewFileMeta(filename) {
    const cleanName = filename.replace(/\.[^/.]+$/, ''); // 확장자 제거

    // 날짜 감지 (예: 20260415, 2026.04.15, 2026-04-15, 260415)
    let date = '';
    const dateMatch = cleanName.match(/\b(20\d{2}[.\-_]?\d{2}[.\-_]?\d{2})\b/);
    if (dateMatch) {
        const d = dateMatch[1].replace(/[.\-_]/g, '');
        date = `${d.slice(0, 4)}.${d.slice(4, 6)}.${d.slice(6, 8)}`;
    } else {
        date = new Date().toISOString().slice(0, 10).replace(/-/g, '.');
    }

    // 대상기관 / 부서 감지
    let department = '관련부서/기관';
    const deptKeywords = ['의료원', '병원', '정보팀', '진료팀', '의무기록', '기획', '전산', '원무', '간호', '응급', '약제', '연구원', '보건소', '재활', '행정'];
    for (const kw of deptKeywords) {
        if (cleanName.includes(kw)) {
            // 해당 키워드 주변 텍스트 추출
            const parts = cleanName.split(/[_ \-]/);
            const foundPart = parts.find(p => p.includes(kw));
            if (foundPart) {
                department = foundPart;
                break;
            }
        }
    }

    // 인터뷰이 (면담자 / 직책) 감지
    let interviewee = '면담자';
    const positionKeywords = ['원장', '부장', '팀장', '과장', '실장', '차장', '책임', '선임', '주임', '교수', '의사', '간호사', '약사', '담당', '선생님'];
    for (const pos of positionKeywords) {
        if (cleanName.includes(pos)) {
            const parts = cleanName.split(/[_ \-]/);
            const foundPart = parts.find(p => p.includes(pos));
            if (foundPart) {
                interviewee = foundPart;
                break;
            }
        }
    }

    return {
        filename,
        cleanName,
        date,
        department,
        interviewee
    };
}

// ── 2. 다중 파일 일괄 텍스트 추출 ───────────────────────────────────
export async function extractMultipleInterviewFiles(fileList) {
    const results = [];

    for (let i = 0; i < fileList.length; i++) {
        const file = fileList[i];
        const meta = parseInterviewFileMeta(file.name);
        try {
            const content = await extractStructuredFileContent(file);
            results.push({
                id: `int-file-${i + 1}`,
                file,
                meta,
                text: content.text || '',
                pages: content.pages || 1
            });
        } catch (err) {
            console.error(`파일 [${file.name}] 추출 오류:`, err);
            results.push({
                id: `int-file-${i + 1}`,
                file,
                meta,
                text: `[파일 읽기 실패: ${err.message}]`,
                pages: 1,
                error: err.message
            });
        }
    }

    return results;
}

// ── 3. AI 1-Click 연쇄 데이터 추출 파이프라인 (Cascade Engine) ────────
export async function runInterviewCascadePipeline({
    interviewFiles,
    apiKey,
    modelId = 'gemini-2.5-flash',
    onProgress = () => {}
}) {
    if (!interviewFiles || interviewFiles.length === 0) {
        throw new Error('분석할 인터뷰 회의록 파일이 없습니다.');
    }

    onProgress({ percent: 15, message: '인터뷰 회의록 원문 취합 및 컨텍스트 구성 중...' });

    // 원문들을 구조화된 텍스트 블록으로 병합
    const documentsContext = interviewFiles.map((item, idx) => {
        const snippet = item.text.length > 5000 ? item.text.substring(0, 5000) + '...(이하 중략)' : item.text;
        return `[인터뷰 문서 #${idx + 1}]
- 파일명: ${item.meta.filename}
- 면담일자: ${item.meta.date}
- 면담기관/부서: ${item.meta.department}
- 면담대상(인터뷰이): ${item.meta.interviewee}
[회의록 본문]
${snippet}
----------------------------------------`;
    }).join('\n\n');

    onProgress({ percent: 35, message: 'Gemini AI 연쇄 분석 및 4대 산출물 동시 도출 중...' });

    const systemPrompt = `당신은 공공기관 정보화전략계획(ISP) 및 IT 컨설팅 전문 수석 컨설턴트입니다.
제시된 복수의 인터뷰 회의록 원문을 정밀 분석하여 현업의 생생한 목소리를 체계적인 정보화 산출물로 변환하는 'OSMU(One-Source Multi-Use) 연쇄 추출'을 수행하십시오.

[분석 지침]:
반드시 아래 4가지 핵심 영역을 심층 도출하여 단일 JSON 객체로 반환하십시오:
1. interviewSummary: 각 인터뷰별 일시, 대상, 핵심 논의 안건(3~4줄 요약), 후속 요청자료(Action Items)
2. asIsIssues: 분류(업무/응용시스템/데이터/인프라/거버넌스), 대상 업무/시스템명, 현황(As-Is), 핵심 문제점 및 Pain Point, 실제 발언 인용 근거(Quote)
3. toBeDirections: 문제점으로부터 도출된 시사점, 혁신 목표모델(To-Be 개선방향), 연관 정보화 추진과제명
4. derivedRequirements: 표준 요구사항 정의 - 고유 ID(사용자 요구: USR-001~, 시스템기능: SFR-001~, 데이터연계: DAR-001~ 등 체계적 부여), 요구사항명, 상세 요건 설명, 우선순위(상/중/하), 출처 회의록 파일명

[출력 요구사항]:
마크다운 백틱(\`\`\`json)이나 부가 설명 없이 반드시 순수한 JSON 형식으로만 응답하십시오.
{
  "interviewSummary": [
    {
      "sourceFile": "파일명",
      "date": "2026.04.15",
      "department": "부서명",
      "interviewee": "면담자",
      "summary": "핵심 논의 내용 요약",
      "actionItems": "향후 요청 및 후속 조치 사항"
    }
  ],
  "asIsIssues": [
    {
      "id": "ISSUE-01",
      "category": "업무" | "응용시스템" | "데이터" | "인프라" | "거버넌스",
      "targetSystem": "대상 업무/시스템명",
      "asIsStatus": "현황 요약",
      "painPoint": "핵심 문제점 및 병목",
      "quote": "실제 인터뷰이 발언 원문 인용",
      "sourceFile": "출처 파일명"
    }
  ],
  "toBeDirections": [
    {
      "issueId": "ISSUE-01",
      "implication": "핵심 시사점",
      "toBeDirection": "To-Be 개선 방향 및 목표모델",
      "relatedProject": "연관 추진과제명"
    }
  ],
  "derivedRequirements": [
    {
      "id": "SFR-001",
      "name": "요구사항 명칭",
      "type": "기능" | "비기능" | "데이터" | "연계" | "보안",
      "description": "구체적인 요구사항 정의 설명",
      "priority": "상" | "중" | "하",
      "sourceFile": "출처 파일명"
    }
  ]
}`;

    onProgress({ percent: 65, message: 'AI 응답 수신 및 구조화 데이터 정규화 중...' });

    let responseData = null;
    const contents = [{ role: 'user', parts: [{ text: `${systemPrompt}\n\n[입력 회의록 원문 데이터]:\n${documentsContext}` }] }];
    const generationConfig = { temperature: 0.1, responseMimeType: 'application/json' };

    // 1차: 서버 프록시
    try {
        const proxyRes = await fetch('/api/gemini-proxy', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ apiKey, modelId, contents, generationConfig })
        });
        if (proxyRes.ok) {
            responseData = await proxyRes.json();
        }
    } catch (_) {}

    // 2차: 직접 호출 폴백
    if (!responseData && apiKey) {
        const directUrl = `https://generativelanguage.googleapis.com/v1beta/${modelId}:generateContent`;
        const directRes = await fetch(directUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'x-goog-api-key': apiKey },
            body: JSON.stringify({ contents, generationConfig })
        });
        if (directRes.ok) {
            responseData = await directRes.json();
        }
    }

    onProgress({ percent: 90, message: '연쇄 산출물 파이프라인 데이터 결합 중...' });

    const rawText = responseData?.candidates?.[0]?.content?.parts?.[0]?.text || '';
    const cleaned = rawText.replace(/```json/gi, '').replace(/```/g, '').trim();

    let parsed = null;
    try {
        parsed = JSON.parse(cleaned);
    } catch (e) {
        console.error('JSON 파싱 실패:', rawText);
        throw new Error('AI 응답을 JSON 형식으로 해석하지 못했습니다. 다시 시도해 주세요.');
    }

    onProgress({ percent: 100, message: '완료!' });

    return {
        interviewSummary: Array.isArray(parsed.interviewSummary) ? parsed.interviewSummary : [],
        asIsIssues: Array.isArray(parsed.asIsIssues) ? parsed.asIsIssues : [],
        toBeDirections: Array.isArray(parsed.toBeDirections) ? parsed.toBeDirections : [],
        derivedRequirements: Array.isArray(parsed.derivedRequirements) ? parsed.derivedRequirements : []
    };
}

// ── 4. 1순위 SSOT Validator 요구사항 마스터로 1클릭 동기화 ────────
export function syncToSSOTRequirements(derivedRequirements) {
    const current = loadSSOTData();
    const existingReqs = current.requirements || [];
    const existingIds = new Set(existingReqs.map(r => r.id));

    let addedCount = 0;
    const newItems = [];

    derivedRequirements.forEach(req => {
        if (!existingIds.has(req.id)) {
            newItems.push({
                id: req.id,
                name: req.name,
                keywords: req.type ? `${req.type}, ${req.priority || '상'}` : '인터뷰도출'
            });
            existingIds.add(req.id);
            addedCount++;
        }
    });

    if (addedCount > 0) {
        const updated = {
            ...current,
            requirements: [...existingReqs, ...newItems]
        };
        saveSSOTData(updated);
    }

    return addedCount;
}

// ── 5. Excel 4개 시트 종합표 생성 (exceljs) ────────────────────────
export async function generateInterviewExcel(cascadeData, tokens) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = '건강한 프로젝트 (Interview Cascade Pipeline)';
    workbook.created = new Date();
    const colors = tokens?.colors || { primary: '1E3A8A', tableHeaderBg: 'F1F5F9', borderColor: 'CBD5E1' };

    const headerStyle = {
        font: { name: '맑은 고딕', size: 10.5, bold: true, color: { argb: 'FFFFFFFF' } },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${colors.primary}` } },
        alignment: { vertical: 'middle', horizontal: 'center' }
    };

    const borderStyle = {
        top: { style: 'thin', color: { argb: `FF${colors.borderColor}` } },
        left: { style: 'thin', color: { argb: `FF${colors.borderColor}` } },
        bottom: { style: 'thin', color: { argb: `FF${colors.borderColor}` } },
        right: { style: 'thin', color: { argb: `FF${colors.borderColor}` } }
    };

    // 시트 1: 인터뷰 결과 요약
    const s1 = workbook.addWorksheet('①_인터뷰결과요약');
    s1.columns = [
        { header: 'No.', width: 8 },
        { header: '면담일자', width: 14 },
        { header: '대상기관/부서', width: 22 },
        { header: '면담대상자', width: 16 },
        { header: '핵심 논의 안건 요약', width: 50 },
        { header: '후속 조치 및 요청자료(Action Items)', width: 35 },
        { header: '출처 파일명', width: 25 }
    ];
    s1.getRow(1).height = 28;
    s1.getRow(1).eachCell(c => { c.style = headerStyle; });
    cascadeData.interviewSummary.forEach((r, idx) => {
        const row = s1.addRow([idx + 1, r.date, r.department, r.interviewee, r.summary, r.actionItems, r.sourceFile]);
        row.height = 26;
        row.alignment = { vertical: 'middle', wrapText: true };
        row.eachCell(c => { c.border = borderStyle; });
    });

    // 시트 2: 현황 및 문제점 (As-Is)
    const s2 = workbook.addWorksheet('②_현황및문제점(As-Is)');
    s2.columns = [
        { header: '이슈ID', width: 12 },
        { header: '분류', width: 14 },
        { header: '대상 업무/시스템', width: 22 },
        { header: '현황 (As-Is)', width: 35 },
        { header: '핵심 문제점 & Pain Point', width: 45 },
        { header: '원문 발언 근거 (Quote)', width: 45 },
        { header: '출처 파일명', width: 22 }
    ];
    s2.getRow(1).height = 28;
    s2.getRow(1).eachCell(c => { c.style = headerStyle; });
    cascadeData.asIsIssues.forEach(r => {
        const row = s2.addRow([r.id, r.category, r.targetSystem, r.asIsStatus, r.painPoint, r.quote, r.sourceFile]);
        row.height = 26;
        row.alignment = { vertical: 'middle', wrapText: true };
        row.eachCell(c => { c.border = borderStyle; });
    });

    // 시트 3: 개선방향 및 시사점 (To-Be)
    const s3 = workbook.addWorksheet('③_개선방향및시사점(To-Be)');
    s3.columns = [
        { header: '연계 이슈ID', width: 14 },
        { header: '도출 시사점', width: 45 },
        { header: '목표모델 및 개선방향 (To-Be)', width: 50 },
        { header: '연관 정보화 추진과제명', width: 35 }
    ];
    s3.getRow(1).height = 28;
    s3.getRow(1).eachCell(c => { c.style = headerStyle; });
    cascadeData.toBeDirections.forEach(r => {
        const row = s3.addRow([r.issueId, r.implication, r.toBeDirection, r.relatedProject]);
        row.height = 26;
        row.alignment = { vertical: 'middle', wrapText: true };
        row.eachCell(c => { c.border = borderStyle; });
    });

    // 시트 4: 도출 요구사항정의서
    const s4 = workbook.addWorksheet('④_도출요구사항정의서');
    s4.columns = [
        { header: '요구사항 ID', width: 15 },
        { header: '요구사항명', width: 35 },
        { header: '구분', width: 12 },
        { header: '상세 요건 정의', width: 55 },
        { header: '우선순위', width: 12 },
        { header: '출처 회의록', width: 25 }
    ];
    s4.getRow(1).height = 28;
    s4.getRow(1).eachCell(c => { c.style = headerStyle; });
    cascadeData.derivedRequirements.forEach(r => {
        const row = s4.addRow([r.id, r.name, r.type || '기능', r.description, r.priority || '상', r.sourceFile]);
        row.height = 26;
        row.alignment = { vertical: 'middle', wrapText: true };
        row.eachCell(c => { c.border = borderStyle; });
    });

    const buffer = await workbook.xlsx.writeBuffer();
    return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

// ── 6. PPTX 컨설팅 비교 보고서 장표 생성 (pptxgenjs) ───────────────
export async function generateInterviewPptx(cascadeData, tokens) {
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_16x9';
    const colors = tokens?.colors || { primary: '1E3A8A', secondary: '2563EB', keyPointBg: 'EFF6FF', danger: 'DC2626', success: '16A34A', borderColor: 'CBD5E1' };
    const typography = tokens?.typography || { fontFamily: '맑은 고딕', titleSize: 18 };

    // 슬라이드 1: 인터뷰 총괄 요약
    const s1 = pptx.addSlide();
    s1.addText('Ⅱ. 현황 분석 및 시사점 > 1. 현업 인터뷰 종합', { x: 0.6, y: 0.35, w: 8, h: 0.3, fontFace: typography.fontFamily, fontSize: 9.5, color: '64748B', bold: true });
    s1.addText('현업 부서별 인터뷰 결과 및 주요 이슈 종합 총괄표', { x: 0.6, y: 0.68, w: 12, h: 0.45, fontFace: typography.fontFamily, fontSize: typography.titleSize, color: colors.primary, bold: true });

    s1.addShape(pptx.ShapeType.rect, { x: 0.6, y: 1.18, w: 12.13, h: 0.75, fill: { color: colors.keyPointBg }, line: { color: colors.secondary, width: 1.2 } });
    s1.addText('💡 현업 실무진 인터뷰 결과, 레거시 EMR 시스템 노후화와 데이터 연계 부재로 인한 업무 병목이 가장 시급한 개선 과제로 식별되었습니다.\n표준화된 연계 체계 구축 및 중앙 집중식 데이터 통합 요구가 전 부서에 걸쳐 공통적으로 도출되었습니다.', {
        x: 0.8, y: 1.22, w: 11.7, h: 0.67, fontFace: typography.fontFamily, fontSize: 11, bold: true, color: '0F172A', valign: 'middle'
    });

    const sumTableRows = [
        ['면담일자', '대상 기관/부서', '인터뷰이', '핵심 논의 및 Pain Point 요약'].map(h => ({
            text: h, options: { fill: { color: colors.primary }, color: 'FFFFFF', bold: true, align: 'center', fontFace: typography.fontFamily, fontSize: 10 }
        })),
        ...cascadeData.interviewSummary.slice(0, 5).map((item, idx) => [
            { text: item.date, options: { align: 'center', fontSize: 9, fill: { color: idx % 2 === 1 ? 'F8FAFC' : 'FFFFFF' } } },
            { text: item.department, options: { align: 'center', fontSize: 9, bold: true, fill: { color: idx % 2 === 1 ? 'F8FAFC' : 'FFFFFF' } } },
            { text: item.interviewee, options: { align: 'center', fontSize: 9, fill: { color: idx % 2 === 1 ? 'F8FAFC' : 'FFFFFF' } } },
            { text: item.summary, options: { align: 'left', fontSize: 9, fill: { color: idx % 2 === 1 ? 'F8FAFC' : 'FFFFFF' } } }
        ])
    ];
    s1.addTable(sumTableRows, { x: 0.6, y: 2.1, w: 12.13, rowH: 0.8, border: { pt: 0.5, color: colors.borderColor } });

    // 슬라이드 2: As-Is 문제점 vs To-Be 시사점 대칭 비교 (TYPE_B 템플릿)
    const s2 = pptx.addSlide();
    s2.addText('Ⅱ. 현황 분석 및 시사점 > 2. As-Is 문제점 vs To-Be 개선방향', { x: 0.6, y: 0.35, w: 8, h: 0.3, fontFace: typography.fontFamily, fontSize: 9.5, color: '64748B', bold: true });
    s2.addText('현행 이슈(As-Is) 기반 목표 개선모델(To-Be) 도출 매핑', { x: 0.6, y: 0.68, w: 12, h: 0.45, fontFace: typography.fontFamily, fontSize: typography.titleSize, color: colors.primary, bold: true });

    s2.addShape(pptx.ShapeType.rect, { x: 0.6, y: 1.18, w: 12.13, h: 0.75, fill: { color: colors.keyPointBg }, line: { color: colors.secondary, width: 1.2 } });
    s2.addText('💡 현업 인터뷰에서 적발된 구체적 문제점과 원문 발언을 토대로 실효성 있는 To-Be 혁신 방안을 1:1로 매핑하여 도출했습니다.\n추진 과제별 명확한 기대효과를 정의하여 예산 타당성 및 정책적 당위성을 확보합니다.', {
        x: 0.8, y: 1.22, w: 11.7, h: 0.67, fontFace: typography.fontFamily, fontSize: 11, bold: true, color: '0F172A', valign: 'middle'
    });

    const cardW = 5.7;
    const cardH = 4.4;

    // 좌측 As-Is 카드
    s2.addShape(pptx.ShapeType.roundRect, { x: 0.6, y: 2.1, w: cardW, h: cardH, rectRadius: 0.08, fill: { color: 'FFF5F5' }, line: { color: colors.danger, width: 1.5 } });
    s2.addText('현행 핵심 문제점 및 Pain Point (As-Is)', { x: 0.8, y: 2.25, w: cardW - 0.4, h: 0.35, fontFace: typography.fontFamily, fontSize: 12, bold: true, color: colors.danger });

    const asIsBullets = cascadeData.asIsIssues.slice(0, 4).map(iss => `[${iss.category}] ${iss.painPoint}\n  💬 "${iss.quote || '관련 발언 발췌'}"`);
    s2.addText(asIsBullets.map(b => ({ text: `• ${b}\n\n`, options: { fontSize: 9, color: '1E293B', fontFace: typography.fontFamily } })), {
        x: 0.8, y: 2.65, w: cardW - 0.4, h: cardH - 0.7, valign: 'top'
    });

    // 중앙 화살표
    s2.addText('➔', { x: 0.6 + cardW, y: 2.1 + cardH / 2 - 0.3, w: 0.73, h: 0.6, fontSize: 24, color: colors.secondary, align: 'center', valign: 'middle' });

    // 우측 To-Be 카드
    s2.addShape(pptx.ShapeType.roundRect, { x: 0.6 + cardW + 0.73, y: 2.1, w: cardW, h: cardH, rectRadius: 0.08, fill: { color: 'F0FDF4' }, line: { color: colors.success, width: 1.5 } });
    s2.addText('목표 혁신 방향 및 시사점 (To-Be)', { x: 0.6 + cardW + 0.93, y: 2.25, w: cardW - 0.4, h: 0.35, fontFace: typography.fontFamily, fontSize: 12, bold: true, color: colors.success });

    const toBeBullets = cascadeData.toBeDirections.slice(0, 4).map(dir => `[과제: ${dir.relatedProject || '정보화과제'}]\n  ➔ ${dir.toBeDirection}`);
    s2.addText(toBeBullets.map(b => ({ text: `• ${b}\n\n`, options: { fontSize: 9.5, color: '1E293B', fontFace: typography.fontFamily } })), {
        x: 0.6 + cardW + 0.93, y: 2.65, w: cardW - 0.4, h: cardH - 0.7, valign: 'top'
    });

    const blob = await pptx.write({ outputType: 'blob' });
    return blob;
}

// ── 7. Word 요구사항정의서 초안 생성 (docx) ───────────────────────
export async function generateInterviewDocx(cascadeData, tokens) {
    const colors = tokens?.colors || { primary: '1E3A8A', tableHeaderBg: 'F1F5F9', borderColor: 'CBD5E1' };

    const children = [
        new docx.Paragraph({
            text: '인터뷰 기반 정보시스템 요구사항정의서 (초안)',
            heading: docx.HeadingLevel.TITLE,
            alignment: docx.AlignmentType.CENTER,
            spacing: { after: 300 }
        }),
        new docx.Paragraph({
            children: [
                new docx.TextRun({ text: '문서 개요: ', bold: true }),
                new docx.TextRun({ text: '본 문서는 현업 실무진 인터뷰 및 회의록 분석을 통해 자동 도출된 초기 요구사항 명세서입니다.' })
            ],
            spacing: { after: 200 }
        })
    ];

    // 요구사항별 상세 명세 테이블 블록 생성
    cascadeData.derivedRequirements.forEach((req, idx) => {
        children.push(
            new docx.Paragraph({
                text: `${idx + 1}. [${req.id}] ${req.name}`,
                heading: docx.HeadingLevel.HEADING_2,
                spacing: { before: 240, after: 120 }
            }),
            new docx.Table({
                width: { size: 100, type: docx.WidthType.PERCENTAGE },
                rows: [
                    new docx.TableRow({
                        children: [
                            new docx.TableCell({
                                children: [new docx.Paragraph({ text: '요구사항 ID', bold: true, alignment: docx.AlignmentType.CENTER })],
                                shading: { fill: colors.tableHeaderBg },
                                width: { size: 20, type: docx.WidthType.PERCENTAGE }
                            }),
                            new docx.TableCell({
                                children: [new docx.Paragraph({ text: req.id, bold: true })],
                                width: { size: 30, type: docx.WidthType.PERCENTAGE }
                            }),
                            new docx.TableCell({
                                children: [new docx.Paragraph({ text: '우선순위 / 구분', bold: true, alignment: docx.AlignmentType.CENTER })],
                                shading: { fill: colors.tableHeaderBg },
                                width: { size: 25, type: docx.WidthType.PERCENTAGE }
                            }),
                            new docx.TableCell({
                                children: [new docx.Paragraph({ text: `${req.priority || '상'} / ${req.type || '기능요구사항'}` })],
                                width: { size: 25, type: docx.WidthType.PERCENTAGE }
                            })
                        ]
                    }),
                    new docx.TableRow({
                        children: [
                            new docx.TableCell({
                                children: [new docx.Paragraph({ text: '상세 요건 정의', bold: true, alignment: docx.AlignmentType.CENTER })],
                                shading: { fill: colors.tableHeaderBg }
                            }),
                            new docx.TableCell({
                                children: [new docx.Paragraph({ text: req.description || '-' })],
                                columnSpan: 3
                            })
                        ]
                    }),
                    new docx.TableRow({
                        children: [
                            new docx.TableCell({
                                children: [new docx.Paragraph({ text: '출처 회의록', bold: true, alignment: docx.AlignmentType.CENTER })],
                                shading: { fill: colors.tableHeaderBg }
                            }),
                            new docx.TableCell({
                                children: [new docx.Paragraph({ text: req.sourceFile || '인터뷰 회의록' })],
                                columnSpan: 3
                            })
                        ]
                    })
                ]
            }),
            new docx.Paragraph({ text: '', spacing: { after: 180 } })
        );
    });

    const doc = new docx.Document({
        sections: [{
            properties: { page: { margin: { top: 1200, bottom: 1200, left: 1200, right: 1200 } } },
            children
        }]
    });

    const blob = await docx.Packer.toBlob(doc);
    return blob;
}

// ── 8. JSZip 3종 산출물 통합 패키징 빌더 ─────────────────────────────
export async function generateInterviewCascadeZip(cascadeData, tokens) {
    const zip = new JSZip();

    // 1) 엑셀 종합표
    const excelBlob = await generateInterviewExcel(cascadeData, tokens);
    zip.file('01_인터뷰결과_및_이슈도출_종합표.xlsx', excelBlob);

    // 2) PPTX 보고서 장표
    const pptxBlob = await generateInterviewPptx(cascadeData, tokens);
    zip.file('02_현황분석_주요문제점_및_시사점_보고서장표.pptx', pptxBlob);

    // 3) Word 요구사항정의서
    const docxBlob = await generateInterviewDocx(cascadeData, tokens);
    zip.file('03_요구사항정의서_초안.docx', docxBlob);

    const zipBlob = await zip.generateAsync({ type: 'blob' });
    const downloadName = `인터뷰_연쇄추출_산출물_패키지_${new Date().toISOString().slice(0, 10)}.zip`;
    saveAs(zipBlob, downloadName);
}

// ── 9. 공공의료/ISP 인터뷰 회의록 기본 샘플 프리셋 ────────────────────
export const DEFAULT_SAMPLE_INTERVIEWS = [
    {
        id: 'sample-int-1',
        meta: {
            filename: '20260415_지방의료원_진료지원팀_원장_면담록.txt',
            cleanName: '20260415_지방의료원_진료지원팀_원장_면담록',
            date: '2026.04.15',
            department: '진료지원팀/원장실',
            interviewee: '김원장(원장)'
        },
        text: `[면담 개요]
- 일시: 2026년 4월 15일 14:00 ~ 15:30
- 장소: A도립의료원 원장실
- 참석자: 김원장(의료원장), 이부장(진료부장), ISP 사업단 총괄PM, 수석컨설턴트

[주요 발언 및 현황 분석]
1. 전자의무기록(EMR) 시스템 노후화
- 2012년 구축된 EMR이 14년째 운영 중으로 진료 피크시간대(오전 10~11시) 처방 전송 지연이 5초 이상 발생하여 외래 진료 대기시간이 가중됨.
- "월요일 오전마다 진료 차트가 멈춰서 환자분들이 진료실 앞에서 고성을 지르는 일이 다반사입니다. 진료 중단 방지를 위한 성능 개선이 절실합니다."

2. 타 의료기관 및 권역책임의료기관 간 진료정보 교류 부재
- 급성기 환자 발생 시 인근 대학병원(권역책임의료기관)으로 긴급 전원 의뢰를 보내는데, 진료의뢰서와 PACS 영상 CD를 환자 보호자가 직접 복사해서 들고 이동함.
- "디지털 시대에 응급환자 CT/MRI 영상을 CD로 구워서 들려보내는 실정입니다. 전원 병원 도착 전 실시간 영상 및 의무기록 조회가 반드시 가능해야 환자 골든타임을 지킬 수 있습니다."

3. DUR 및 마약류 처방 연계 미흡
- 지역 내 타 의료기관 투약 이력이 실시간 연동되지 않아 중복 처방 및 특정 약물 부작용 이력 파악에 한계 존재.

[후속 조치 (Action Items)]
- 최근 3개년 외래/입원 EMR 처방 건수 및 피크타임 트래픽 로그 제출
- 권역책임의료기관 전원 환자 발생 통계 및 의뢰 서식 샘플 제공`,
        pages: 1
    },
    {
        id: 'sample-int-2',
        meta: {
            filename: '20260417_의료정보팀_전산인프라_팀장_인터뷰.txt',
            cleanName: '20260417_의료정보팀_전산인프라_팀장_인터뷰',
            date: '2026.04.17',
            department: '의료정보팀',
            interviewee: '박팀장(정보팀장)'
        },
        text: `[면담 개요]
- 일시: 2026년 4월 17일 10:00 ~ 11:30
- 장소: 전산교육장 회의실
- 참석자: 박팀장(의료정보팀장), 최선임(서버인프라), ISP 인프라 컨설턴트

[주요 발언 및 현황 분석]
1. 온프레미스 하드웨어 내용연수 초과 및 장애 위험
- 원내 전산실 서버 24대 중 18대가 내구연수(5년)를 초과하여 작년에도 스토리지 컨트롤러 장애로 2시간 동안 병원 전산망이 다운됨.
- "부품 단종으로 중고 부품을 구해 겨우 가동 중입니다. 한 번 더 대형 장애가 나면 백업 스토리지 복구까지 24시간 이상 걸립니다."

2. 보안 및 개인정보보호 규제 미준수 위험
- 의료데이터 접근제어 감사 로그를 2년 이상 보관해야 하나 스토리지 용량 부족으로 6개월 분량만 남기고 강제 삭제하고 있어 감사 지적 우려가 매우 큼.
- "전산 인력이 3명뿐이라 24시간 관제는 불가능합니다. 보안관제와 인프라 유지보수를 공공 클라우드 통합센터(CSAP 인증)로 전환하지 않으면 버티기 어렵습니다."

[후속 조치 (Action Items)]
- 전산 장비 도입연도별 자산 실사 대장 제출
- 유지보수 SLA 계약서 및 최근 1년 장애 일지 제공`,
        pages: 1
    },
    {
        id: 'sample-int-3',
        meta: {
            filename: '20260419_원무과_전원관리_실장_회의록.txt',
            cleanName: '20260419_원무과_전원관리_실장_회의록',
            date: '2026.04.19',
            department: '원무과/공공의료사업단',
            interviewee: '정실장(원무실장)'
        },
        text: `[면담 개요]
- 일시: 2026년 4월 19일 15:00 ~ 16:20
- 장소: 원무과 회의실
- 참석자: 정실장(원무실장), 한주임(전원코디네이터), ISP 업무분석 컨설턴트

[주요 발언 및 현황 분석]
1. 취약계층 퇴원 환자 지역사회 사후관리 단절
- 독거노인, 만성질환자 등 취약계층 환자가 퇴원 후 관할 보건소나 방문간호센터로 연계되어야 하나 시스템 연계가 전무하여 전화와 팩스로 서류를 주고받음.
- "퇴원 후 복약 관리가 안 되어 2주 만에 응급실로 재입원하는 어르신들이 너무 많습니다. 보건소 지역보건의료정보시스템(PHIS)과 병원 간 케어플랜이 공유되어야 합니다."

2. 수납 대기시간 및 키오스크 이용 접근성 취약
- 평균 수납 대기시간이 25분 이상 소요되어 환자 불만 급증.
- "원내 키오스크 2대가 있으나 폰트가 작고 UI가 복잡해 고령 환자들이 전혀 쓰지 못하고 결국 수기 번호표 창구로 몰려 대기 줄이 길어집니다."

[후속 조치 (Action Items)]
- 보건소 연계 퇴원환자 현황 및 전원 관리 대장 제출
- 원무 수납 창구별 평균 대기시간 및 키오스크 이용률 통계 공유`,
        pages: 1
    }
];

export const DEFAULT_SAMPLE_CASCADE_DATA = {
    interviewSummary: [
        {
            sourceFile: '20260415_지방의료원_진료지원팀_원장_면담록.txt',
            date: '2026.04.15',
            department: '진료지원팀/원장실',
            interviewee: '김원장(원장)',
            summary: '14년 경과된 레거시 EMR 처방 지연(피크타임 5초 이상) 해소 및 상급병원 전원 시 종이/CD 복사 탈피를 위한 국가 표준(FHIR) 기반 실시간 진료정보교류 체계 도입 필요성 제기.',
            actionItems: '최근 3개년 외래/입원 처방 트래픽 로그 및 권역책임의료기관 전원 의뢰 통계 제출'
        },
        {
            sourceFile: '20260417_의료정보팀_전산인프라_팀장_인터뷰.txt',
            date: '2026.04.17',
            department: '의료정보팀',
            interviewee: '박팀장(정보팀장)',
            summary: '원내 서버 75%가 내구연수 초과로 장애 위험 상존. 스토리지 용량 부족으로 법정 감사로그 보관 미흡. 공공 클라우드(CSAP 인증) 통합 전환 및 24시간 보안관제 외주 체계 요구.',
            actionItems: '전산 하드웨어 자산실사표, 유지보수 계약서 및 연간 장애 일지 제공'
        },
        {
            sourceFile: '20260419_원무과_전원관리_실장_회의록.txt',
            date: '2026.04.19',
            department: '원무과/공공의료사업단',
            interviewee: '정실장(원무실장)',
            summary: '취약계층 퇴원 환자의 지역 보건소(PHIS) 사후 연계 부재로 재입원율 증가 문제 식별. 평균 25분 소요되는 원무 수납 대기 해소를 위해 고령자 친화 배리어프리 키오스크 및 모바일 사전수납 요구.',
            actionItems: '보건소 전원관리 대장 및 창구별 대기시간/키오스크 이용률 통계 공유'
        }
    ],
    asIsIssues: [
        {
            id: 'ISSUE-01',
            category: '응용시스템',
            targetSystem: '외래/입원 처방전달시스템(OCS/EMR)',
            asIsStatus: '2012년 도입된 레거시 클라이언트-서버(C/S) EMR 구조 유지',
            painPoint: '오전 진료 피크타임 처방 입력 지연(5초+) 및 차트 멈춤 현상 빈발',
            quote: '월요일 오전마다 진료 차트가 멈춰서 환자분들이 진료실 앞에서 고성을 지르는 일이 다반사입니다.',
            sourceFile: '20260415_지방의료원_진료지원팀_원장_면담록.txt'
        },
        {
            id: 'ISSUE-02',
            category: '데이터',
            targetSystem: '진료정보교류 및 영상전송시스템(PACS)',
            asIsStatus: '상급종합병원 전원 시 진료의뢰서 출력 및 영상 PACS CD 수기 복사 전달',
            painPoint: '골든타임 내 응급 전원 환자 데이터 실시간 사전 조회 불가 및 중복 검사 초래',
            quote: '디지털 시대에 응급환자 CT/MRI 영상을 CD로 구워서 들려보내는 실정입니다.',
            sourceFile: '20260415_지방의료원_진료지원팀_원장_면담록.txt'
        },
        {
            id: 'ISSUE-03',
            category: '인프라',
            targetSystem: '원내 전산실 서버 및 백업 스토리지',
            asIsStatus: '물리 서버 24대 중 18대(75%) 내구연수(5년) 초과 운영 중',
            painPoint: '부품 단종 및 스토리지 컨트롤러 장애 발생 시 복구 24시간 이상 소요 위험',
            quote: '부품 단종으로 중고 부품을 구해 겨우 가동 중입니다. 대형 장애가 나면 24시간 이상 걸립니다.',
            sourceFile: '20260417_의료정보팀_전산인프라_팀장_인터뷰.txt'
        },
        {
            id: 'ISSUE-04',
            category: '거버넌스',
            targetSystem: '보안관제 및 개인정보 감사로그',
            asIsStatus: '전산 인력 3인 체계로 야간/주말 보안 관제 공백 및 스토리지 용량 한계',
            painPoint: '의료데이터 접속기록 2년 법정 보관 의무 미준수(용량 부족으로 6개월치만 보관)',
            quote: '전산 인력이 3명뿐이라 24시간 관제는 불가능하며 용량 부족으로 6개월 분량만 남기고 강제 삭제합니다.',
            sourceFile: '20260417_의료정보팀_전산인프라_팀장_인터뷰.txt'
        },
        {
            id: 'ISSUE-05',
            category: '업무',
            targetSystem: '공공보건의료 연계 및 원무 수납 창구',
            asIsStatus: '보건소(PHIS)와 병원 간 퇴원 환자 연계 수기(전화/팩스) 처리 및 복잡한 키오스크 UI',
            painPoint: '퇴원 후 복약 관리 단절로 인한 2주 내 재입원 빈발 및 원무 창구 평균 대기 25분 초과',
            quote: '퇴원 후 복약 관리가 안 되어 재입원하는 어르신들이 너무 많고, 키오스크가 어려워 번호표 창구로 몰립니다.',
            sourceFile: '20260419_원무과_전원관리_실장_회의록.txt'
        }
    ],
    toBeDirections: [
        {
            issueId: 'ISSUE-01',
            implication: '의료진 진료 생산성 제고를 위해 MSA 기반 고성능 클라우드 차세대 EMR 전면 전환 시급',
            toBeDirection: '웹 표준 기반 클라우드 네이티브 EMR 도입 및 진료 차트 1초 이내 로딩 응답성 확보',
            relatedProject: '차세대 클라우드 전자의무기록(EMR) 고도화 사업'
        },
        {
            issueId: 'ISSUE-02',
            implication: '응급 및 중증 환자의 골든타임 사수를 위한 범국가 진료정보교류(EHR) 표준 연동 필수',
            toBeDirection: '국가 표준(FHIR) 인터페이스 및 마이헬스웨이(건강정보 고속도로) 직접 연계 체계 구축',
            relatedProject: '공공보건의료 스마트 진료정보교류(FHIR) 연계망 구축'
        },
        {
            issueId: 'ISSUE-03',
            implication: '노후 온프레미스 인프라 한계 극복을 위한 공공보건의료 특화 클라우드(IaaS/PaaS) 이관 필요',
            toBeDirection: 'CSAP 인증 공공 클라우드 전환 및 99.99% 가용성 보장 이중화 재해복구(DR) 체계 구축',
            relatedProject: '공공의료 통합 클라우드 인프라 전환 및 재해복구센터 구축'
        },
        {
            issueId: 'ISSUE-04',
            implication: '의료법 및 개인정보보호법 완전 준수를 위한 통합 SIEM 및 24×365 원격 보안관제 필수',
            toBeDirection: '개인정보 접속기록 2년 무결성 보관 스토리지 확충 및 전문 보안관제 센터 위탁 운영',
            relatedProject: '공공의료 사이버보안 및 의료데이터 감사로그 통합 관제 시스템'
        },
        {
            issueId: 'ISSUE-05',
            implication: '취약계층 필수의료 연속성 확보를 위한 병원-보건소 환자 케어플랜 양방향 연동 체계 구축',
            toBeDirection: 'PHIS 연계 퇴원환자 통합 케어플랫폼 구축 및 고령자 배리어프리 간편 키오스크/모바일 수납 보급',
            relatedProject: '지역사회 필수의료 연계 퇴원환자 케어플랜 및 스마트 원무 서비스 구축'
        }
    ],
    derivedRequirements: [
        {
            id: 'SFR-001',
            name: '클라우드 기반 차세대 EMR 처방전달 성능 보장',
            type: '기능',
            description: '외래 및 입원 피크시간대(오전 9시~12시) 동시 접속자 300인 이상 환경에서 진료 차트 조회 및 처방 전송 트랜잭션 응답속도를 1.5초 이내로 보장해야 한다.',
            priority: '상',
            sourceFile: '20260415_지방의료원_진료지원팀_원장_면담록.txt'
        },
        {
            id: 'DAR-001',
            name: '국가 표준 FHIR 기반 진료정보 및 영상 연계',
            type: '데이터',
            description: '상급종합병원 및 권역책임의료기관 전원 의뢰 시 HL7 FHIR 표준 프로파일을 적용하여 환자 기본정보, 진료소견, 투약이력 및 PACS DICOM 영상 링크를 암호화하여 실시간 자동 전송해야 한다.',
            priority: '상',
            sourceFile: '20260415_지방의료원_진료지원팀_원장_면담록.txt'
        },
        {
            id: 'INF-001',
            name: 'CSAP 인증 공공 클라우드 전환 및 이중화 DR 구성',
            type: '인프라',
            description: '병원 핵심 주전산기 시스템을 CSAP 인증 민관협력형 공공 클라우드로 100% 이전하고, RTO 2시간, RPO 15분 이내의 실시간 원격 백업 재해복구(DR) 체계를 구현해야 한다.',
            priority: '상',
            sourceFile: '20260417_의료정보팀_전산인프라_팀장_인터뷰.txt'
        },
        {
            id: 'SER-001',
            name: '의료정보 개인정보 접속기록 2년 보관 및 감사 자동화',
            type: '보안',
            description: '환자 개인정보 및 전자의무기록 열람, 수정, 다운로드 이력을 WORM(Write Once Read Many) 스토리지에 최소 2년간 위변조 없이 보관하고 비정상 대량 조회 발생 시 실시간 경보를 발송해야 한다.',
            priority: '상',
            sourceFile: '20260417_의료정보팀_전산인프라_팀장_인터뷰.txt'
        },
        {
            id: 'DAR-002',
            name: '지역보건의료정보시스템(PHIS) 연계 퇴원환자 케어플랜',
            type: '연계',
            description: '취약계층 만성질환 퇴원 환자의 진료 요약지 및 복약 지도 계획을 관할 보건소 PHIS 시스템과 OpenAPI로 양방향 동기화하여 지역사회 연속 돌봄을 지원해야 한다.',
            priority: '중',
            sourceFile: '20260419_원무과_전원관리_실장_회의록.txt'
        },
        {
            id: 'USR-001',
            name: '고령자 배리어프리 키오스크 및 모바일 사전수납',
            type: '기능',
            description: '시니어 환자를 위한 24인치 대화면 고대비 큰글씨 UI 키오스크를 도입하고, 스마트폰 알림톡 기반 1클릭 모바일 사전수납 및 전자영수증 발급 기능을 제공해야 한다.',
            priority: '중',
            sourceFile: '20260419_원무과_전원관리_실장_회의록.txt'
        }
    ]
};
