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

    for (const fileName of slideFiles) {
        const content = await zip.files[fileName].async('string');
        const slideParagraphs = [];

        try {
            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(content, 'text/xml');
            const paragraphs = xmlDoc.getElementsByTagNameNS('*', 'p');

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

        if (slideParagraphs.length > 0) {
            const slideNum = fileName.match(/\d+/)[0];
            textBlocks.push(`[슬라이드 ${slideNum}]\n` + slideParagraphs.join('\n'));
        }
    }
    
    return {
        text: textBlocks.join('\n\n').trim(),
        pages: slideFiles.length
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

    if (ext === 'pptx') {
        const res = await extractTextFromPPTX(file);
        // PPTX는 text 안에 [슬라이드 N]이 들어있으므로 pageMap 구성
        const pageMap = {};
        const slideBlocks = res.text.split(/\[슬라이드 \d+\]\n/).filter(b => b.trim().length > 0);
        slideBlocks.forEach((block, idx) => {
            pageMap[idx + 1] = block.split('\n').map(l => l.trim()).filter(l => l.length > 0);
        });
        return { text: res.text, pages: res.pages || Object.keys(pageMap).length, pageMap };
    } else if (ext === 'hwpx') {
        return await extractTextFromHWPX(file);
    } else if (ext === 'docx') {
        return await extractTextFromDOCX(file);
    } else if (ext === 'pdf') {
        return await extractTextFromPDF(file);
    } else if (ext === 'xlsx' || ext === 'xls') {
        return await extractTextFromExcel(file);
    } else if (ext === 'md' || ext === 'markdown' || ext === 'txt' || ext === 'csv' || ext === 'json' || ext === 'html' || ext === 'xml') {
        return await extractTextFromMD(file);
    } else {
        const text = await file.text();
        return {
            text: text.trim(),
            pages: 1,
            pageMap: { 1: text.split('\n').map(l => l.trim()).filter(l => l.length > 0) }
        };
    }
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
