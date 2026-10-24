/**
 * batchQueueRunner.js
 * 동시성 제어 기반 비동기 작업 큐 및 Rate Limit(429) 자동 재시도 배치 검수 엔진
 * 10~30개 산출물 일괄 파싱 ➔ SSOT 1단계 Rule 검증 ➔ RAG 결합 2단계 AI 문맥 검수 ➔ 통합 엑셀 & Word 교정본 ZIP 빌드
 */
import JSZip from 'jszip';
import * as docx from 'docx';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { extractStructuredFileContent, getFileExtension } from './fileExtractor.js';
import {
    loadSSOTData,
    extractTablesFromFile,
    runStage1Validation
} from './ssotValidator.js';
import { loadDesignTokens } from './designTokens.js';
import { searchRagHybrid, formatTopChunksForPrompt } from './clientRagIndexer.js';

export const BATCH_ITEM_STATUS = {
    PENDING: 'PENDING',         // 대기중
    PARSING: 'PARSING',         // 파싱중
    RULE_CHECK: 'RULE_CHECK',   // SSOT 규칙검증
    AI_CHECK: 'AI_CHECK',       // AI 문맥검수
    COMPLETED: 'COMPLETED',     // 완료
    ERROR: 'ERROR'              // 오류
};

export const STATUS_LABELS = {
    PENDING: '대기중',
    PARSING: '파싱중',
    RULE_CHECK: 'SSOT 규칙검증',
    AI_CHECK: 'AI 문맥검수',
    COMPLETED: '교정본 생성 완료',
    ERROR: '오류 발생'
};

export const STATUS_COLORS = {
    PENDING: '#94a3b8',
    PARSING: '#3b82f6',
    RULE_CHECK: '#8b5cf6',
    AI_CHECK: '#ec4899',
    COMPLETED: '#10b981',
    ERROR: '#ef4444'
};

// ── 1. Exponential Backoff 재시도 헬퍼 (429 Rate Limit 극복) ──────
export async function executeWithBackoff(asyncFn, { maxRetries = 4, baseDelayMs = 2000, onRetry = () => {} } = {}) {
    let attempt = 0;
    while (true) {
        try {
            return await asyncFn();
        } catch (err) {
            attempt++;
            const status = err?.status || err?.statusCode || 0;
            const errMsg = String(err?.message || '');
            const isRateLimit = status === 429 ||
                errMsg.includes('429') ||
                errMsg.includes('quota') ||
                errMsg.includes('Quota') ||
                errMsg.includes('RESOURCE_EXHAUSTED') ||
                errMsg.includes('Rate limit');

            if (isRateLimit && attempt <= maxRetries) {
                const jitter = Math.floor(Math.random() * 800);
                const delay = Math.min(baseDelayMs * Math.pow(2, attempt - 1) + jitter, 16000);
                onRetry({
                    attempt,
                    maxRetries,
                    delayMs: delay,
                    message: `Gemini API 할당량(429) 초과 감지. ${Math.round(delay / 1000)}초 후 자동 재시도 (${attempt}/${maxRetries})...`
                });
                await new Promise(r => setTimeout(r, delay));
            } else {
                throw err;
            }
        }
    }
}

// ── 2. AI 문맥 검수 단일 호출 헬퍼 (RAG 결합) ────────────────────
async function runSingleFileAiCheck({ fileContent, masterData, apiKey, modelId, ragContext, onRetry }) {
    const textSnippet = fileContent.text.length > 4000
        ? fileContent.text.substring(0, 4000) + '...(이하 중략)'
        : fileContent.text;

    const systemPrompt = `당신은 공공기관 ISP/ISMP 산출물 전문 감리원 및 수석 컨설턴트입니다.
제시된 산출물 본문에서 [맞춤법/오탈자], [문맥 모순], [행정용어 부적절], [수치/단위 표기 오류]를 정밀하게 적발하십시오.
${ragContext ? `\n반드시 아래 제공된 [지침 기준문서 / RAG 지식베이스]의 표준 절차 및 기술 요건을 기준으로 적합성을 대조 검증하십시오.\n${ragContext}\n` : ''}

[출력 요구사항]:
마크다운 백틱(\`\`\`json) 없이 순수한 JSON 배열([]) 형식으로만 응답하십시오. 오류가 없으면 빈 배열 []을 반환하십시오.
[
  {
    "errorType": "맞춤법" | "문맥모순" | "행정용어부적절" | "수치불일치" | "지침불일치",
    "originalText": "원문에서 결함이 있는 정확한 문장 또는 어구 발췌",
    "suggestedText": "권장하는 수정안",
    "reason": "결함으로 판단한 명확한 근거",
    "severity": "HIGH" | "MEDIUM" | "LOW"
  }
]`;

    const contents = [{
        role: 'user',
        parts: [{ text: `${systemPrompt}\n\n[검수 대상 산출물 본문 발췌]:\n${textSnippet}` }]
    }];

    const generationConfig = { temperature: 0.1, responseMimeType: 'application/json' };

    return await executeWithBackoff(async () => {
        let responseData = null;

        // 1차 프록시 호출
        try {
            const proxyRes = await fetch('/api/gemini-proxy', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ apiKey, modelId, contents, generationConfig })
            });
            if (proxyRes.ok) {
                responseData = await proxyRes.json();
            } else if (proxyRes.status === 429) {
                const err = new Error('HTTP 429 Too Many Requests');
                err.status = 429;
                throw err;
            }
        } catch (e) {
            if (e?.status === 429) throw e;
        }

        // 2차 직접 엔드포인트 호출
        if (!responseData && apiKey) {
            const directUrl = `https://generativelanguage.googleapis.com/v1beta/${modelId}:generateContent`;
            const directRes = await fetch(directUrl, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json', 'x-goog-api-key': apiKey },
                body: JSON.stringify({ contents, generationConfig })
            });
            if (!directRes.ok) {
                const err = new Error(`Gemini API 에러: ${directRes.status}`);
                err.status = directRes.status;
                throw err;
            }
            responseData = await directRes.json();
        }

        const rawText = responseData?.candidates?.[0]?.content?.parts?.[0]?.text || '[]';
        const cleaned = rawText.replace(/```json/gi, '').replace(/```/g, '').trim();

        try {
            const parsed = JSON.parse(cleaned);
            return Array.isArray(parsed) ? parsed : [];
        } catch (_) {
            return [];
        }
    }, { maxRetries: 4, baseDelayMs: 2000, onRetry });
}

// ── 3. Batch Queue Runner 클래스 (동시성 제어 큐 엔진) ───────────
export class BatchQueueRunner {
    constructor(options = {}) {
        this.concurrency = options.concurrency || 2;
        this.apiKey = options.apiKey || '';
        this.modelId = options.modelId || 'gemini-2.5-flash';
        this.masterData = options.masterData || loadSSOTData();
        this.useAiCheck = options.useAiCheck ?? true;
        this.useRagContext = options.useRagContext ?? true;

        this.items = []; // 작업 항목 리스트
        this.activeCount = 0;
        this.isPaused = false;
        this.isCancelled = false;

        this.onItemChange = options.onItemChange || (() => {});
        this.onProgress = options.onProgress || (() => {});
        this.onComplete = options.onComplete || (() => {});
    }

    addFiles(files) {
        const newItems = Array.from(files).map((file, idx) => ({
            id: `batch_${Date.now()}_${idx}_${Math.random().toString(36).substring(2, 6)}`,
            file,
            filename: file.name,
            fileSize: file.size,
            ext: getFileExtension(file.name),
            status: BATCH_ITEM_STATUS.PENDING,
            statusMessage: '대기중',
            progress: 0,
            pages: 1,
            ruleDefectsCount: 0,
            aiDefectsCount: 0,
            totalDefectsCount: 0,
            ruleIssues: [],
            aiIssues: [],
            allIssues: [],
            ragSnippets: [],
            extractedContent: null,
            error: null,
            startTime: null,
            endTime: null,
            durationMs: 0
        }));

        this.items = [...this.items, ...newItems];
        this._notifyProgress();
        return newItems;
    }

    start() {
        this.isPaused = false;
        this.isCancelled = false;
        this._drain();
    }

    pause() {
        this.isPaused = true;
    }

    resume() {
        if (!this.isPaused) return;
        this.isPaused = false;
        this._drain();
    }

    cancel() {
        this.isCancelled = true;
        this.isPaused = true;
    }

    retryFailed() {
        this.items.forEach(item => {
            if (item.status === BATCH_ITEM_STATUS.ERROR) {
                item.status = BATCH_ITEM_STATUS.PENDING;
                item.statusMessage = '대기중 (재시도)';
                item.progress = 0;
                item.error = null;
                this.onItemChange(item);
            }
        });
        this.start();
    }

    _drain() {
        if (this.isPaused || this.isCancelled) return;

        while (this.activeCount < this.concurrency) {
            const nextItem = this.items.find(item => item.status === BATCH_ITEM_STATUS.PENDING);
            if (!nextItem) break;

            this.activeCount++;
            this._processItem(nextItem).finally(() => {
                this.activeCount--;
                this._notifyProgress();
                this._drain();

                // 모든 작업 완료 여부 검사
                const allDone = this.items.every(
                    item => item.status === BATCH_ITEM_STATUS.COMPLETED || item.status === BATCH_ITEM_STATUS.ERROR
                );
                if (allDone && this.activeCount === 0) {
                    this.onComplete(this.items);
                }
            });
        }
    }

    async _processItem(item) {
        item.startTime = Date.now();
        try {
            // [1단계: 파일 파싱]
            item.status = BATCH_ITEM_STATUS.PARSING;
            item.statusMessage = '텍스트 및 표 구조 추출 중...';
            item.progress = 20;
            this.onItemChange(item);

            const fileContent = await extractStructuredFileContent(item.file);
            item.extractedContent = fileContent;
            item.pages = fileContent.pages || 1;

            const tables = await extractTablesFromFile(item.file, fileContent);

            if (this.isCancelled) return;

            // [2단계: SSOT Rule 결정론적 검증]
            item.status = BATCH_ITEM_STATUS.RULE_CHECK;
            item.statusMessage = '표준 용어·요구사항ID·수치 정합성 검증 중...';
            item.progress = 50;
            this.onItemChange(item);

            const ruleIssues = await runStage1Validation(fileContent, tables, this.masterData);
            item.ruleIssues = ruleIssues;
            item.ruleDefectsCount = ruleIssues.length;

            if (this.isCancelled) return;

            // [3단계: RAG 가이드라인 검색 및 AI 문맥 검수]
            if (this.useAiCheck && this.apiKey) {
                item.status = BATCH_ITEM_STATUS.AI_CHECK;
                item.statusMessage = '가이드라인 RAG 검색 및 AI 문맥 정밀 검수 중...';
                item.progress = 75;
                this.onItemChange(item);

                let ragPromptContext = '';
                if (this.useRagContext) {
                    try {
                        const topChunks = await searchRagHybrid(
                            `${item.filename} ${fileContent.text.slice(0, 400)}`,
                            { topK: 5 }
                        );
                        item.ragSnippets = topChunks;
                        ragPromptContext = formatTopChunksForPrompt(topChunks);
                    } catch (ragErr) {
                        console.warn('RAG 검색 폴백:', ragErr);
                    }
                }

                const aiResults = await runSingleFileAiCheck({
                    fileContent,
                    masterData: this.masterData,
                    apiKey: this.apiKey,
                    modelId: this.modelId,
                    ragContext: ragPromptContext,
                    onRetry: (retryInfo) => {
                        item.statusMessage = retryInfo.message;
                        this.onItemChange(item);
                    }
                });

                item.aiIssues = aiResults.map((r, i) => ({
                    id: `AI-${item.id}-${i + 1}`,
                    stage: 'Stage 2 (AI 문맥검증)',
                    category: r.errorType || 'AI 문맥/지침',
                    errorType: r.errorType || 'AI 검출 결함',
                    severity: r.severity || 'MEDIUM',
                    page: Math.max(1, Math.min(item.pages, Math.ceil((i + 1) * (item.pages / (aiResults.length || 1))))),
                    matchedTerm: r.originalText || '',
                    suggestedText: r.suggestedText || '',
                    originalText: r.originalText || '',
                    reason: r.reason || 'AI 문맥 및 지침 정밀 검증'
                }));
                item.aiDefectsCount = item.aiIssues.length;
            }

            // 전체 결함 취합 및 완료
            item.allIssues = [...item.ruleIssues, ...item.aiIssues];
            item.totalDefectsCount = item.allIssues.length;
            item.status = BATCH_ITEM_STATUS.COMPLETED;
            item.statusMessage = `완료 (결함 ${item.totalDefectsCount}건 식별)`;
            item.progress = 100;
        } catch (err) {
            console.error(`[${item.filename}] 배치 검수 에러:`, err);
            item.status = BATCH_ITEM_STATUS.ERROR;
            item.statusMessage = `오류: ${err.message}`;
            item.error = err.message;
        } finally {
            item.endTime = Date.now();
            item.durationMs = item.endTime - item.startTime;
            this.onItemChange(item);
        }
    }

    _notifyProgress() {
        const total = this.items.length;
        const completed = this.items.filter(i => i.status === BATCH_ITEM_STATUS.COMPLETED).length;
        const error = this.items.filter(i => i.status === BATCH_ITEM_STATUS.ERROR).length;
        const running = this.items.filter(i => [BATCH_ITEM_STATUS.PARSING, BATCH_ITEM_STATUS.RULE_CHECK, BATCH_ITEM_STATUS.AI_CHECK].includes(i.status)).length;
        const pending = this.items.filter(i => i.status === BATCH_ITEM_STATUS.PENDING).length;
        const percent = total > 0 ? Math.round(((completed + error) / total) * 100) : 0;

        this.onProgress({ total, completed, error, running, pending, percent });
    }
}

// ── 4. 통합 검수 대시보드 엑셀 빌더 (exceljs) ─────────────────────
export async function buildBatchSummaryWorkbook(batchItems, tokens) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = '건강한 프로젝트 (Batch Validator)';
    workbook.created = new Date();
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

    // 시트 1: 산출물별 검수 요약
    const s1 = workbook.addWorksheet('①_산출물별_검수_요약');
    s1.columns = [
        { header: 'No.', width: 6 },
        { header: '산출물 파일명', width: 34 },
        { header: '포맷', width: 10 },
        { header: '용량', width: 12 },
        { header: '페이지', width: 10 },
        { header: '총 결함수', width: 12 },
        { header: 'Rule 검출', width: 12 },
        { header: 'AI 검출', width: 12 },
        { header: '용어위반', width: 12 },
        { header: '검수 소요시간', width: 14 },
        { header: '검수 상태', width: 14 }
    ];
    s1.getRow(1).height = 26;
    s1.getRow(1).eachCell(c => { c.style = headerStyle; });

    batchItems.forEach((item, idx) => {
        const typoCount = (item.allIssues || []).filter(i => i.category === '용어사전' || i.errorType?.includes('용어')).length;
        const sec = (item.durationMs / 1000).toFixed(1) + 's';
        const row = s1.addRow([
            idx + 1,
            item.filename,
            (item.ext || '').toUpperCase(),
            (item.fileSize / 1024).toFixed(1) + ' KB',
            item.pages || 1,
            item.totalDefectsCount || 0,
            item.ruleDefectsCount || 0,
            item.aiDefectsCount || 0,
            typoCount,
            sec,
            item.status === BATCH_ITEM_STATUS.COMPLETED ? '정상 완료' : '오류'
        ]);
        row.height = 24;
        row.alignment = { vertical: 'middle', horizontal: 'center' };
        row.getCell(2).alignment = { vertical: 'middle', horizontal: 'left' };
        row.eachCell(c => { c.border = borderStyle; });
    });

    // 시트 2: 전체 결함 상세 통계표
    const s2 = workbook.addWorksheet('②_전체_결함_상세_통계표');
    s2.columns = [
        { header: 'No.', width: 6 },
        { header: '대상 파일명', width: 28 },
        { header: '검수 단계', width: 14 },
        { header: '결함 유형', width: 16 },
        { header: '심각도', width: 10 },
        { header: '페이지', width: 10 },
        { header: '원문 발췌 (오표기/결함)', width: 45 },
        { header: '수정 권장안 (Approved)', width: 35 },
        { header: '판정 근거 및 사유', width: 40 }
    ];
    s2.getRow(1).height = 26;
    s2.getRow(1).eachCell(c => { c.style = headerStyle; });

    let defectIdx = 1;
    batchItems.forEach(item => {
        (item.allIssues || []).forEach(issue => {
            const row = s2.addRow([
                defectIdx++,
                item.filename,
                issue.stage?.includes('Stage 1') ? 'Rule 검증' : 'AI 문맥검증',
                issue.category || issue.errorType || '기타',
                issue.severity || 'MEDIUM',
                issue.page || 1,
                issue.originalText || issue.matchedTerm || '',
                issue.suggestedText || '',
                issue.reason || ''
            ]);
            row.height = 24;
            row.alignment = { vertical: 'middle', wrapText: true };
            row.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };
            row.getCell(5).alignment = { vertical: 'middle', horizontal: 'center' };
            row.getCell(6).alignment = { vertical: 'middle', horizontal: 'center' };
            row.eachCell(c => { c.border = borderStyle; });
        });
    });

    const buffer = await workbook.xlsx.writeBuffer();
    return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

// ── 5. 개별 파일별 Word 교정본 (Redline 비교표) 생성 (docx) ───────
export async function buildSingleFileDocxReport(item, tokens) {
    const colors = tokens?.colors || { primary: '1E3A8A', tableHeaderBg: 'F1F5F9', borderColor: 'CBD5E1' };

    const children = [
        new docx.Paragraph({
            text: `산출물 교정 리포트 - ${item.filename}`,
            heading: docx.HeadingLevel.TITLE,
            alignment: docx.AlignmentType.CENTER,
            spacing: { after: 200 }
        }),
        new docx.Paragraph({
            children: [
                new docx.TextRun({ text: '검수 파일명: ', bold: true }),
                new docx.TextRun({ text: `${item.filename} (${(item.ext || '').toUpperCase()}, ${item.pages || 1}페이지)  |  ` }),
                new docx.TextRun({ text: '검출된 총 결함: ', bold: true }),
                new docx.TextRun({ text: `${item.totalDefectsCount || 0}건 (Rule: ${item.ruleDefectsCount || 0}건, AI: ${item.aiDefectsCount || 0}건)`, color: 'DC2626', bold: true })
            ],
            spacing: { after: 240 }
        }),
        new docx.Paragraph({
            text: '1. 원문 대비 수정 전/후 (Redline) 상세 비교표',
            heading: docx.HeadingLevel.HEADING_2,
            spacing: { before: 180, after: 120 }
        })
    ];

    const issues = item.allIssues || [];
    if (issues.length === 0) {
        children.push(
            new docx.Paragraph({
                text: '✓ 본 산출물에서는 표준 용어 위반 및 정합성 결함이 검출되지 않았습니다.',
                spacing: { after: 200 }
            })
        );
    } else {
        const tableRows = [
            new docx.TableRow({
                children: [
                    new docx.TableCell({ children: [new docx.Paragraph({ text: 'No.', bold: true, alignment: docx.AlignmentType.CENTER })], shading: { fill: colors.tableHeaderBg }, width: { size: 6, type: docx.WidthType.PERCENTAGE } }),
                    new docx.TableCell({ children: [new docx.Paragraph({ text: '위치', bold: true, alignment: docx.AlignmentType.CENTER })], shading: { fill: colors.tableHeaderBg }, width: { size: 10, type: docx.WidthType.PERCENTAGE } }),
                    new docx.TableCell({ children: [new docx.Paragraph({ text: '분류', bold: true, alignment: docx.AlignmentType.CENTER })], shading: { fill: colors.tableHeaderBg }, width: { size: 14, type: docx.WidthType.PERCENTAGE } }),
                    new docx.TableCell({ children: [new docx.Paragraph({ text: '원문 (오표기/결함)', bold: true, alignment: docx.AlignmentType.CENTER })], shading: { fill: colors.tableHeaderBg }, width: { size: 35, type: docx.WidthType.PERCENTAGE } }),
                    new docx.TableCell({ children: [new docx.Paragraph({ text: '수정 권장안 (Approved)', bold: true, alignment: docx.AlignmentType.CENTER })], shading: { fill: colors.tableHeaderBg }, width: { size: 35, type: docx.WidthType.PERCENTAGE } })
                ]
            }),
            ...issues.map((iss, i) => new docx.TableRow({
                children: [
                    new docx.TableCell({ children: [new docx.Paragraph({ text: String(i + 1), alignment: docx.AlignmentType.CENTER })] }),
                    new docx.TableCell({ children: [new docx.Paragraph({ text: `p.${iss.page || 1}`, alignment: docx.AlignmentType.CENTER })] }),
                    new docx.TableCell({ children: [new docx.Paragraph({ text: iss.category || iss.errorType || '-' })] }),
                    new docx.TableCell({ children: [new docx.Paragraph({ text: iss.originalText || iss.matchedTerm || '-', color: 'DC2626', strike: true })] }),
                    new docx.TableCell({ children: [new docx.Paragraph({ text: iss.suggestedText || '-', color: '16A34A', bold: true })] })
                ]
            }))
        ];

        children.push(
            new docx.Table({
                width: { size: 100, type: docx.WidthType.PERCENTAGE },
                rows: tableRows
            }),
            new docx.Paragraph({ text: '', spacing: { after: 200 } })
        );
    }

    const doc = new docx.Document({
        sections: [{
            properties: { page: { margin: { top: 1200, bottom: 1200, left: 1200, right: 1200 } } },
            children
        }]
    });

    return await docx.Packer.toBlob(doc);
}

// ── 6. 일괄 ZIP 패키징 빌더 (jszip) ───────────────────────────────
export async function generateBatchReportZip(batchItems, tokens) {
    const zip = new JSZip();

    // 1) 전체산출물 통합검수 대시보드.xlsx
    const excelBlob = await buildBatchSummaryWorkbook(batchItems, tokens);
    zip.file('00_전체산출물_통합검수_대시보드.xlsx', excelBlob);

    // 2) 교정리포트_파일명.docx 개별 파일 생성
    const reportsFolder = zip.folder('개별_교정리포트');
    for (const item of batchItems) {
        if (item.status === BATCH_ITEM_STATUS.COMPLETED) {
            const docxBlob = await buildSingleFileDocxReport(item, tokens);
            const cleanName = (item.filename || 'doc').replace(/\.[^/.]+$/, '');
            reportsFolder.file(`교정리포트_${cleanName}.docx`, docxBlob);
        }
    }

    const zipBlob = await zip.generateAsync({ type: 'blob' });
    const downloadName = `산출물_일괄검수_패키지_${new Date().toISOString().slice(0, 10)}.zip`;
    saveAs(zipBlob, downloadName);
}
