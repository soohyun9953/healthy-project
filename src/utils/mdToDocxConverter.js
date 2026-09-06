/**
 * Markdown (.md) to MS Word (.docx) Converter Utility
 * 
 * 마크다운 문서를 표준 MS Word(.docx) 문서로 완벽하게 변환하는 엔진입니다.
 * - 제목(H1~H6) 계층 스타일 및 여백
 * - 인라인 서식 (굵게, 기울임, 취소선, 인라인 코드, 하이퍼링크)
 * - 리스트 (순서 있는 목록, 순서 없는 목록, 다단계 목록)
 * - 인용 블록 (Blockquote)
 * - 표 (Table) - 헤더 배경, 테두리, 정렬, 자동 크기 조정
 * - 코드 블록 (Code Block) - 모노스페이스 폰트, 배경 박스
 * - 수평선 (Horizontal Rule)
 * - 머리글/바닥글(페이지 번호), 커스텀 테마 색상, 폰트 및 여백 설정 지원
 * - 단일/다중 파일 일괄 변환, ZIP 압축 다운로드, 로컬 폴더 직접 저장(showDirectoryPicker) 지원
 */

import { marked } from 'marked';
import {
    Document,
    Packer,
    Paragraph,
    TextRun,
    Table,
    TableRow,
    TableCell,
    HeadingLevel,
    AlignmentType,
    BorderStyle,
    WidthType,
    ShadingType,
    convertMillimetersToTwip,
    PageNumber,
    Footer,
    Header,
    ExternalHyperlink
} from 'docx';
import saveAs from 'file-saver';
import JSZip from 'jszip';

// ── 테마 색상 프리셋 ──────────────────────────────
export const DOCX_THEMES = {
    blue: {
        id: 'blue',
        name: '모던 블루',
        primary: '2563EB',      // 헤더 및 주요 강조색
        secondary: '3B82F6',
        tableHeaderBg: 'DBEAFE', // 표 헤더 배경색
        tableHeaderText: '1E3A8A',// 표 헤더 글자색
        tableBorder: '93C5FD',   // 표 테두리색
        codeBg: 'F1F5F9',        // 코드 블록 배경색
        quoteBorder: '3B82F6',   // 인용구 테두리색
        quoteBg: 'EFF6FF',       // 인용구 배경색
    },
    charcoal: {
        id: 'charcoal',
        name: '비즈니스 차콜',
        primary: '1E293B',
        secondary: '475569',
        tableHeaderBg: 'F1F5F9',
        tableHeaderText: '0F172A',
        tableBorder: 'CBD5E1',
        codeBg: 'F8FAFC',
        quoteBorder: '475569',
        quoteBg: 'F8FAFC',
    },
    emerald: {
        id: 'emerald',
        name: '에메랄드 그린',
        primary: '059669',
        secondary: '10B981',
        tableHeaderBg: 'D1FAE5',
        tableHeaderText: '065F46',
        tableBorder: '6EE7B7',
        codeBg: 'F0FDF4',
        quoteBorder: '10B981',
        quoteBg: 'ECFDF5',
    },
    purple: {
        id: 'purple',
        name: '로열 퍼플',
        primary: '7C3AED',
        secondary: '8B5CF6',
        tableHeaderBg: 'EDE9FE',
        tableHeaderText: '5B21B6',
        tableBorder: 'C4B5FD',
        codeBg: 'FAF5FF',
        quoteBorder: '8B5CF6',
        quoteBg: 'F5F3FF',
    },
    monochrome: {
        id: 'monochrome',
        name: '클린 블랙',
        primary: '111827',
        secondary: '374151',
        tableHeaderBg: 'E5E7EB',
        tableHeaderText: '111827',
        tableBorder: 'D1D5DB',
        codeBg: 'F3F4F6',
        quoteBorder: '9CA3AF',
        quoteBg: 'F9FAFB',
    }
};

// ── 폰트 프리셋 ────────────────────────────────────
export const DOCX_FONTS = [
    { id: 'Malgun Gothic', name: '맑은 고딕 (기본)' },
    { id: 'NanumGothic', name: '나눔고딕' },
    { id: 'Noto Sans KR', name: '본고딕 (Noto Sans KR)' },
    { id: 'Pretendard', name: '프리텐다드 (Pretendard)' },
    { id: 'Batang', name: '바탕체 / 명조' },
    { id: 'Arial', name: 'Arial (영문/글로벌)' },
    { id: 'Times New Roman', name: 'Times New Roman (학술/공문)' },
    { id: 'Calibri', name: 'Calibri (MS Office 기본)' }
];

// ── 기본 설정 옵션 ──────────────────────────────────
export const DEFAULT_DOCX_OPTIONS = {
    fontFamily: 'Malgun Gothic',
    fontSizePt: 10.5,           // 본문 폰트 크기 (pt)
    lineSpacing: 276,           // 줄 간격 (240 = 1.0배, 276 = 1.15배, 312 = 1.3배, 360 = 1.5배)
    marginMm: 25.4,             // 페이지 여백 (mm) - 상하좌우 25.4mm = 1인치 (표준)
    themeId: 'blue',            // 테마 ID
    includeHeaderFooter: true,  // 머리글(제목) 및 바닥글(페이지 번호) 포함 여부
    documentTitle: '',          // 머리글에 표시할 문서명 (비어있으면 파일명 또는 H1 사용)
    prefix: '변환_',             // 출력 파일명 접두사
};

/**
 * 인라인 토큰 배열을 TextRun 및 ExternalHyperlink 배열로 변환
 */
function parseInlineTokens(tokens, options, theme) {
    if (!tokens || !Array.isArray(tokens)) return [];
    
    const runs = [];
    const font = options.fontFamily || 'Malgun Gothic';
    const baseSizeHalfPt = Math.round((options.fontSizePt || 10.5) * 2);

    for (const token of tokens) {
        if (token.type === 'text') {
            // plain text
            runs.push(new TextRun({
                text: token.text,
                font: font,
                size: baseSizeHalfPt,
                color: '333333',
            }));
        } else if (token.type === 'strong') {
            // 굵게
            const innerRuns = parseInlineTokens(token.tokens, options, theme);
            for (const r of innerRuns) {
                if (r instanceof TextRun) {
                    runs.push(new TextRun({
                        ...r.root[1],
                        text: token.text || r.root[1]?.text,
                        bold: true,
                        font: font,
                        size: baseSizeHalfPt,
                        color: '111827',
                    }));
                } else {
                    runs.push(r);
                }
            }
            if (innerRuns.length === 0 && token.text) {
                runs.push(new TextRun({
                    text: token.text,
                    bold: true,
                    font: font,
                    size: baseSizeHalfPt,
                    color: '111827',
                }));
            }
        } else if (token.type === 'em') {
            // 기울임
            runs.push(new TextRun({
                text: token.text,
                italics: true,
                font: font,
                size: baseSizeHalfPt,
            }));
        } else if (token.type === 'del') {
            // 취소선
            runs.push(new TextRun({
                text: token.text,
                strike: true,
                font: font,
                size: baseSizeHalfPt,
                color: '6B7280',
            }));
        } else if (token.type === 'codespan') {
            // 인라인 코드
            runs.push(new TextRun({
                text: ` ${token.text} `,
                font: 'Consolas',
                size: Math.max(16, baseSizeHalfPt - 2),
                color: 'BE185D',
                shading: {
                    type: ShadingType.CLEAR,
                    fill: 'F3F4F6',
                    color: 'auto',
                },
            }));
        } else if (token.type === 'link') {
            // 하이퍼링크
            try {
                runs.push(new ExternalHyperlink({
                    children: [
                        new TextRun({
                            text: token.text || token.href,
                            font: font,
                            size: baseSizeHalfPt,
                            color: theme.primary,
                            underline: {},
                        }),
                    ],
                    link: token.href,
                }));
            } catch (e) {
                runs.push(new TextRun({
                    text: token.text || token.href,
                    font: font,
                    size: baseSizeHalfPt,
                    color: theme.primary,
                    underline: {},
                }));
            }
        } else if (token.type === 'br') {
            runs.push(new TextRun({
                text: '\n',
                break: 1,
            }));
        } else if (token.type === 'escape') {
            runs.push(new TextRun({
                text: token.text,
                font: font,
                size: baseSizeHalfPt,
            }));
        } else {
            // fallback
            if (token.text) {
                runs.push(new TextRun({
                    text: token.text,
                    font: font,
                    size: baseSizeHalfPt,
                }));
            }
        }
    }
    
    return runs;
}

/**
 * marked 토큰 목록을 docx 엘리먼트 배열(Paragraph, Table 등)로 변환
 */
function convertTokensToDocxElements(tokens, options, theme) {
    const elements = [];
    const font = options.fontFamily || 'Malgun Gothic';
    const baseSizeHalfPt = Math.round((options.fontSizePt || 10.5) * 2);
    const lineSpacing = options.lineSpacing || 276;

    for (let i = 0; i < tokens.length; i++) {
        const token = tokens[i];

        switch (token.type) {
            case 'heading': {
                const depth = token.depth; // 1 ~ 6
                let headingLevel = HeadingLevel.HEADING_1;
                let fontSize = 32; // 16pt
                let spaceBefore = 280;
                let spaceAfter = 120;
                let color = theme.primary;
                let isBold = true;

                if (depth === 1) {
                    headingLevel = HeadingLevel.HEADING_1;
                    fontSize = 36; // 18pt
                    spaceBefore = 400;
                    spaceAfter = 160;
                    color = theme.primary;
                } else if (depth === 2) {
                    headingLevel = HeadingLevel.HEADING_2;
                    fontSize = 30; // 15pt
                    spaceBefore = 320;
                    spaceAfter = 120;
                    color = theme.secondary;
                } else if (depth === 3) {
                    headingLevel = HeadingLevel.HEADING_3;
                    fontSize = 26; // 13pt
                    spaceBefore = 240;
                    spaceAfter = 80;
                    color = '1F2937';
                } else if (depth === 4) {
                    headingLevel = HeadingLevel.HEADING_4;
                    fontSize = 24; // 12pt
                    spaceBefore = 200;
                    spaceAfter = 60;
                    color = '374151';
                } else {
                    headingLevel = HeadingLevel.HEADING_5;
                    fontSize = 22; // 11pt
                    spaceBefore = 160;
                    spaceAfter = 40;
                    color = '4B5563';
                }

                const runs = token.tokens && token.tokens.length > 0 
                    ? parseInlineTokens(token.tokens, options, theme)
                    : [new TextRun({ text: token.text, font: font })];

                // 헤딩 서식 덮어쓰기
                const formattedRuns = runs.map(r => {
                    if (r instanceof TextRun) {
                        return new TextRun({
                            ...r.root[1],
                            font: font,
                            size: fontSize,
                            bold: isBold,
                            color: color,
                        });
                    }
                    return r;
                });

                elements.push(new Paragraph({
                    heading: headingLevel,
                    children: formattedRuns,
                    spacing: {
                        before: spaceBefore,
                        after: spaceAfter,
                        line: lineSpacing,
                    },
                }));
                break;
            }

            case 'paragraph': {
                const runs = token.tokens && token.tokens.length > 0 
                    ? parseInlineTokens(token.tokens, options, theme)
                    : [new TextRun({ text: token.text, font: font, size: baseSizeHalfPt })];

                elements.push(new Paragraph({
                    children: runs,
                    spacing: {
                        before: 60,
                        after: 100,
                        line: lineSpacing,
                    },
                }));
                break;
            }

            case 'list': {
                const isOrdered = token.ordered;
                const start = token.start || 1;

                token.items.forEach((item, itemIdx) => {
                    const prefix = isOrdered ? `${start + itemIdx}. ` : '• ';
                    const runs = item.tokens && item.tokens.length > 0
                        ? parseInlineTokens(item.tokens, options, theme)
                        : [new TextRun({ text: item.text, font: font, size: baseSizeHalfPt })];

                    // 앞 접두사 Run 추가
                    const listRuns = [
                        new TextRun({
                            text: prefix,
                            bold: isOrdered,
                            color: isOrdered ? theme.primary : '4B5563',
                            font: font,
                            size: baseSizeHalfPt,
                        }),
                        ...runs
                    ];

                    elements.push(new Paragraph({
                        children: listRuns,
                        indent: {
                            left: 360, // 0.25 inch
                            hanging: 240,
                        },
                        spacing: {
                            before: 30,
                            after: 40,
                            line: lineSpacing,
                        },
                    }));
                });
                break;
            }

            case 'blockquote': {
                // 인용 블록
                const quoteText = token.text || '';
                const lines = quoteText.split('\n');

                lines.forEach((line, lineIdx) => {
                    const cleanLine = line.replace(/^>\s*/, '');
                    elements.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: cleanLine,
                                italics: true,
                                color: '4B5563',
                                font: font,
                                size: baseSizeHalfPt,
                            }),
                        ],
                        border: {
                            left: {
                                color: theme.quoteBorder,
                                size: 24, // 3pt
                                style: BorderStyle.SINGLE,
                                space: 15,
                            },
                        },
                        shading: {
                            type: ShadingType.CLEAR,
                            fill: theme.quoteBg,
                            color: 'auto',
                        },
                        indent: {
                            left: 360,
                            right: 360,
                        },
                        spacing: {
                            before: lineIdx === 0 ? 120 : 40,
                            after: lineIdx === lines.length - 1 ? 120 : 40,
                            line: lineSpacing,
                        },
                    }));
                });
                break;
            }

            case 'code': {
                // 코드 블록
                const codeLines = (token.text || '').split('\n');
                
                // 상단 헤더 라벨(언어명 표시)
                if (token.lang) {
                    elements.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: `[${token.lang.toUpperCase()}]`,
                                bold: true,
                                size: 16,
                                color: '6B7280',
                                font: 'Consolas',
                            })
                        ],
                        shading: {
                            type: ShadingType.CLEAR,
                            fill: 'E5E7EB',
                            color: 'auto',
                        },
                        indent: { left: 240, right: 240 },
                        spacing: { before: 120, after: 20 },
                    }));
                }

                codeLines.forEach((line, lineIdx) => {
                    elements.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: line || ' ',
                                font: 'Consolas',
                                size: 18, // 9pt
                                color: '1F2937',
                            }),
                        ],
                        shading: {
                            type: ShadingType.CLEAR,
                            fill: theme.codeBg,
                            color: 'auto',
                        },
                        border: {
                            left: { color: 'CBD5E1', size: 8, style: BorderStyle.SINGLE, space: 10 },
                            right: { color: 'CBD5E1', size: 8, style: BorderStyle.SINGLE, space: 10 },
                            ...(lineIdx === 0 && !token.lang ? { top: { color: 'CBD5E1', size: 8, style: BorderStyle.SINGLE } } : {}),
                            ...(lineIdx === codeLines.length - 1 ? { bottom: { color: 'CBD5E1', size: 8, style: BorderStyle.SINGLE } } : {}),
                        },
                        indent: {
                            left: 240,
                            right: 240,
                        },
                        spacing: {
                            before: 10,
                            after: 10,
                            line: 220,
                        },
                    }));
                });
                break;
            }

            case 'table': {
                // 표 (Table)
                const tableRows = [];
                const colAlignments = token.align || [];

                // 1. 헤더 행
                if (token.header && token.header.length > 0) {
                    const headerCells = token.header.map((cellToken, colIdx) => {
                        const cellRuns = cellToken.tokens && cellToken.tokens.length > 0
                            ? parseInlineTokens(cellToken.tokens, options, theme)
                            : [new TextRun({ text: cellToken.text, font: font, size: baseSizeHalfPt })];

                        // 헤더는 굵게 & 테마 헤더 글자색
                        const formattedRuns = cellRuns.map(r => {
                            if (r instanceof TextRun) {
                                return new TextRun({
                                    ...r.root[1],
                                    font: font,
                                    bold: true,
                                    color: theme.tableHeaderText,
                                    size: baseSizeHalfPt,
                                });
                            }
                            return r;
                        });

                        const align = colAlignments[colIdx] === 'center' ? AlignmentType.CENTER
                            : colAlignments[colIdx] === 'right' ? AlignmentType.RIGHT
                            : AlignmentType.LEFT;

                        return new TableCell({
                            children: [
                                new Paragraph({
                                    children: formattedRuns,
                                    alignment: align,
                                    spacing: { before: 80, after: 80 },
                                })
                            ],
                            shading: {
                                type: ShadingType.CLEAR,
                                fill: theme.tableHeaderBg,
                                color: 'auto',
                            },
                            borders: {
                                top: { style: BorderStyle.SINGLE, size: 8, color: theme.tableBorder },
                                bottom: { style: BorderStyle.SINGLE, size: 14, color: theme.tableBorder },
                                left: { style: BorderStyle.SINGLE, size: 4, color: theme.tableBorder },
                                right: { style: BorderStyle.SINGLE, size: 4, color: theme.tableBorder },
                            },
                            margins: { top: 120, bottom: 120, left: 140, right: 140 },
                        });
                    });

                    tableRows.push(new TableRow({
                        children: headerCells,
                        tableHeader: true,
                    }));
                }

                // 2. 데이터 행
                if (token.rows && token.rows.length > 0) {
                    token.rows.forEach((row, rowIdx) => {
                        const isEven = rowIdx % 2 === 1;
                        const rowCells = row.map((cellToken, colIdx) => {
                            const cellRuns = cellToken.tokens && cellToken.tokens.length > 0
                                ? parseInlineTokens(cellToken.tokens, options, theme)
                                : [new TextRun({ text: cellToken.text, font: font, size: baseSizeHalfPt })];

                            const align = colAlignments[colIdx] === 'center' ? AlignmentType.CENTER
                                : colAlignments[colIdx] === 'right' ? AlignmentType.RIGHT
                                : AlignmentType.LEFT;

                            return new TableCell({
                                children: [
                                    new Paragraph({
                                        children: cellRuns,
                                        alignment: align,
                                        spacing: { before: 60, after: 60, line: lineSpacing },
                                    })
                                ],
                                shading: isEven ? {
                                    type: ShadingType.CLEAR,
                                    fill: 'F9FAFB',
                                    color: 'auto',
                                } : undefined,
                                borders: {
                                    top: { style: BorderStyle.SINGLE, size: 4, color: theme.tableBorder },
                                    bottom: { style: BorderStyle.SINGLE, size: 4, color: theme.tableBorder },
                                    left: { style: BorderStyle.SINGLE, size: 4, color: theme.tableBorder },
                                    right: { style: BorderStyle.SINGLE, size: 4, color: theme.tableBorder },
                                },
                                margins: { top: 100, bottom: 100, left: 140, right: 140 },
                            });
                        });

                        tableRows.push(new TableRow({
                            children: rowCells,
                        }));
                    });
                }

                if (tableRows.length > 0) {
                    elements.push(new Table({
                        rows: tableRows,
                        width: {
                            size: 100,
                            type: WidthType.PERCENTAGE,
                        },
                    }));

                    // 테이블 뒤 여백용 빈 단락
                    elements.push(new Paragraph({
                        spacing: { before: 80, after: 120 },
                    }));
                }
                break;
            }

            case 'hr': {
                // 수평선
                elements.push(new Paragraph({
                    border: {
                        bottom: {
                            color: theme.tableBorder || 'CBD5E1',
                            size: 12,
                            style: BorderStyle.SINGLE,
                            space: 1,
                        },
                    },
                    spacing: { before: 160, after: 160 },
                }));
                break;
            }

            case 'space':
                // 공백
                break;

            default:
                // 기타 HTML 등
                if (token.text) {
                    elements.push(new Paragraph({
                        children: [
                            new TextRun({
                                text: token.text,
                                font: font,
                                size: baseSizeHalfPt,
                            })
                        ],
                        spacing: { before: 40, after: 80 },
                    }));
                }
                break;
        }
    }

    return elements;
}

/**
 * 마크다운 텍스트를 DOCX Blob으로 변환
 * @param {string} markdownText 마크다운 원본 텍스트
 * @param {object} customOptions 옵션 객체
 * @returns {Promise<Blob>} 생성된 .docx Blob
 */
export async function convertMdToDocxBlob(markdownText, customOptions = {}) {
    const options = { ...DEFAULT_DOCX_OPTIONS, ...customOptions };
    const theme = DOCX_THEMES[options.themeId] || DOCX_THEMES.blue;
    const font = options.fontFamily || 'Malgun Gothic';
    const marginTwip = convertMillimetersToTwip(options.marginMm || 25.4);

    // 1. 마크다운 토큰화
    const tokens = marked.lexer(markdownText || '');

    // 2. 문서 제목 추출 (H1 첫 번째 또는 옵션 documentTitle)
    let docTitle = options.documentTitle || '';
    if (!docTitle) {
        const firstH1 = tokens.find(t => t.type === 'heading' && t.depth === 1);
        if (firstH1) {
            docTitle = firstH1.text;
        }
    }

    // 3. 토큰을 docx 엘리먼트로 변환
    const docxElements = convertTokensToDocxElements(tokens, options, theme);

    // 4. 머리글 / 바닥글 설정
    const headers = options.includeHeaderFooter && docTitle ? {
        default: new Header({
            children: [
                new Paragraph({
                    children: [
                        new TextRun({
                            text: docTitle,
                            font: font,
                            size: 18, // 9pt
                            color: '9CA3AF',
                        }),
                    ],
                    alignment: AlignmentType.RIGHT,
                    border: {
                        bottom: { style: BorderStyle.SINGLE, size: 4, color: 'E5E7EB' },
                    },
                    spacing: { after: 120 },
                }),
            ],
        }),
    } : undefined;

    const footers = options.includeHeaderFooter ? {
        default: new Footer({
            children: [
                new Paragraph({
                    children: [
                        new TextRun({
                            text: 'Page ',
                            font: font,
                            size: 18,
                            color: '9CA3AF',
                        }),
                        new TextRun({
                            children: [PageNumber.CURRENT],
                            font: font,
                            size: 18,
                            color: '9CA3AF',
                        }),
                    ],
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 120 },
                }),
            ],
        }),
    } : undefined;

    // 5. Document 객체 빌드
    const doc = new Document({
        styles: {
            default: {
                document: {
                    run: {
                        font: font,
                        size: Math.round((options.fontSizePt || 10.5) * 2),
                        color: '333333',
                    },
                },
            },
        },
        sections: [
            {
                properties: {
                    page: {
                        margin: {
                            top: marginTwip,
                            bottom: marginTwip,
                            left: marginTwip,
                            right: marginTwip,
                        },
                    },
                },
                headers: headers,
                footers: footers,
                children: docxElements.length > 0 ? docxElements : [
                    new Paragraph({
                        children: [new TextRun({ text: '내용이 없습니다.', font: font })]
                    })
                ],
            },
        ],
    });

    // 6. Packer로 Blob 생성
    const blob = await Packer.toBlob(doc);
    return blob;
}

/**
 * 단일 마크다운 File 객체를 DOCX로 변환
 */
export async function convertSingleMdFile(file, options = {}) {
    const text = await file.text();
    const blob = await convertMdToDocxBlob(text, {
        ...options,
        documentTitle: options.documentTitle || file.name.replace(/\.(md|markdown|txt)$/i, ''),
    });

    const prefix = options.prefix !== undefined ? options.prefix : '변환_';
    const baseName = file.name.replace(/\.(md|markdown|txt)$/i, '');
    const outFileName = `${prefix}${baseName}.docx`;

    return {
        blob,
        fileName: outFileName,
        originalName: file.name,
        size: blob.size,
    };
}

/**
 * 복수 마크다운 파일 일괄 변환
 */
export async function convertBatchMdFiles(files, options = {}, onProgress = null) {
    const results = [];
    const total = files.length;

    for (let i = 0; i < total; i++) {
        const file = files[i];
        if (onProgress) {
            onProgress({ current: i + 1, total, currentFile: file.name, status: 'converting' });
        }

        try {
            const res = await convertSingleMdFile(file, options);
            results.push({
                file,
                blob: res.blob,
                fileName: res.fileName,
                status: 'success',
                error: null,
            });
        } catch (err) {
            console.error(`Error converting ${file.name}:`, err);
            results.push({
                file,
                blob: null,
                fileName: file.name.replace(/\.[^/.]+$/, "") + ".docx",
                status: 'error',
                error: err.message || '변환 실패',
            });
        }
    }

    if (onProgress) {
        onProgress({ current: total, total, currentFile: '', status: 'completed' });
    }

    return results;
}

/**
 * 전체 변환 결과를 ZIP 압축 파일로 다운로드
 */
export async function downloadAllAsZip(results, zipName = '변환완료_Word문서목록.zip') {
    const zip = new JSZip();
    let count = 0;

    for (const r of results) {
        if (r.status === 'success' && r.blob) {
            zip.file(r.fileName, r.blob);
            count++;
        }
    }

    if (count === 0) {
        throw new Error('다운로드할 성공 변환 파일이 없습니다.');
    }

    const zipBlob = await zip.generateAsync({ type: 'blob' });
    saveAs(zipBlob, zipName);
    return count;
}

/**
 * 사용자가 선택한 로컬 폴더에 일괄 저장 (File System Access API)
 */
export async function saveFilesToDirectory(results) {
    if (!window.showDirectoryPicker) {
        throw new Error('해당 브라우저는 폴더 직접 저장(File System Access API)을 지원하지 않습니다. Chrome/Edge 최신 버전을 권장하며, 전체 ZIP 다운로드를 이용해 주세요.');
    }

    const dirHandle = await window.showDirectoryPicker();
    let savedCount = 0;

    for (const r of results) {
        if (r.status === 'success' && r.blob) {
            const fileHandle = await dirHandle.getFileHandle(r.fileName, { create: true });
            const writable = await fileHandle.createWritable();
            await writable.write(r.blob);
            await writable.close();
            savedCount++;
        }
    }

    return savedCount;
}

/**
 * 단일 파일 다운로드 헬퍼
 */
export function downloadSingleFile(blob, fileName) {
    saveAs(blob, fileName);
}
