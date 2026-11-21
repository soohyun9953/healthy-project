/**
 * 목차(장·절·항) 엑셀 내보내기
 * - '목차' 시트: 장·절·항 한 줄씩 (번호·제목·시작 페이지·페이지 범위·출처 파일), 단계별 서식·들여쓰기
 * - '장별 요약' 시트: 장마다 절·항 수, 시작 페이지, 출처 파일
 * 미리보기에서 체크 해제한 항목은 빼고, 고친 제목을 그대로 씁니다.
 */
import ExcelJS from 'exceljs';

const BLUE = 'FF0072BA';
const NAVY = 'FF224F77';
const GRAY = 'FFF2F2F2';
const LINE = 'FFD9D9D9';
const FONT = '맑은 고딕';

const border = { top: { style: 'thin', color: { argb: LINE } }, left: { style: 'thin', color: { argb: LINE } }, bottom: { style: 'thin', color: { argb: LINE } }, right: { style: 'thin', color: { argb: LINE } } };
const fill = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });

function styleHeader(row) {
    row.height = 22;
    row.eachCell(c => {
        c.font = { name: FONT, bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
        c.fill = fill(BLUE);
        c.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        c.border = border;
    });
}

/**
 * @param {Array} chapters extractTocFromDeliverables 결과(편집·포함 체크 반영)
 * @param {{ includeSubsections?: boolean }} opts
 * @returns {Promise<{ blob: Blob, counts: { chapters: number, sections: number, subsections: number } }>}
 */
export async function buildTocWorkbook(chapters, { includeSubsections = true } = {}) {
    const included = chapters.filter(c => c.include).map(c => ({
        ...c,
        sections: c.sections.filter(s => s.include).map(s => ({ ...s, subsections: includeSubsections ? (s.subsections || []).filter(x => x.include) : [] }))
    }));
    const wb = new ExcelJS.Workbook();
    wb.created = new Date();

    // ── 목차 시트 ──
    const ws = wb.addWorksheet('목차', { views: [{ state: 'frozen', ySplit: 1 }] });
    ws.columns = [
        { header: '구분', key: 'kind', width: 7 },
        { header: '번호', key: 'code', width: 10 },
        { header: '장', key: 'ch', width: 6 },
        { header: '절', key: 'sec', width: 6 },
        { header: '항', key: 'sub', width: 6 },
        { header: '제목', key: 'title', width: 52 },
        { header: '시작 페이지', key: 'page', width: 12 },
        { header: '페이지 범위', key: 'range', width: 26 },
        { header: '출처 파일', key: 'file', width: 34 }
    ];
    styleHeader(ws.getRow(1));

    const addRow = (values, level) => {
        const row = ws.addRow(values);
        row.height = 18;
        row.eachCell({ includeEmpty: true }, (c, col) => {
            c.border = border;
            c.alignment = { vertical: 'middle', horizontal: col === 6 || col >= 8 ? 'left' : 'center', indent: col === 6 ? level : 0, wrapText: col === 6 };
            if (level === 0) {
                c.font = { name: FONT, bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
                c.fill = fill(NAVY);
            } else if (level === 1) {
                c.font = { name: FONT, bold: true, size: 10, color: { argb: col === 7 ? BLUE : 'FF000000' } };
                c.fill = fill(GRAY);
            } else {
                c.font = { name: FONT, size: 9.5, color: { argb: col === 1 || col === 2 ? 'FF595959' : 'FF000000' } };
            }
        });
        return row;
    };

    let sections = 0;
    let subsections = 0;
    included.forEach(c => {
        const ch = c.roman || (c.num != null ? String(c.num) : '');
        addRow({ kind: '장', code: ch, ch: c.num ?? '', sec: '', sub: '', title: c.title, page: c.startLabel || '', range: '', file: c.fileName || '' }, 0);
        c.sections.forEach(s => {
            sections += 1;
            addRow({ kind: '절', code: ch ? `${ch}.${s.num}` : String(s.num), ch: c.num ?? '', sec: s.num, sub: '', title: s.title, page: s.startLabel || '페이지 미확인', range: s.pageRange || '', file: c.fileName || '' }, 1);
            s.subsections.forEach(x => {
                subsections += 1;
                addRow({ kind: '항', code: `${s.num}.${x.num}`, ch: c.num ?? '', sec: s.num, sub: x.num, title: x.title, page: x.startLabel || '', range: '', file: c.fileName || '' }, 2);
            });
        });
    });
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: Math.max(1, ws.rowCount), column: ws.columnCount } };

    // ── 장별 요약 시트 ──
    const sum = wb.addWorksheet('장별 요약', { views: [{ state: 'frozen', ySplit: 1 }] });
    sum.columns = [
        { header: '장', key: 'ch', width: 7 },
        { header: '장 제목', key: 'title', width: 30 },
        { header: '절 수', key: 'secs', width: 8 },
        { header: '항 수', key: 'subs', width: 8 },
        { header: '시작 페이지', key: 'page', width: 12 },
        { header: '출처 파일', key: 'file', width: 34 }
    ];
    styleHeader(sum.getRow(1));
    included.forEach(c => {
        const row = sum.addRow({
            ch: c.roman || c.num || '',
            title: c.title,
            secs: c.sections.length,
            subs: c.sections.reduce((n, s) => n + s.subsections.length, 0),
            page: c.startLabel || '',
            file: c.fileName || ''
        });
        row.eachCell({ includeEmpty: true }, (cell, col) => {
            cell.border = border;
            cell.font = { name: FONT, size: 10, bold: col <= 2 };
            cell.alignment = { vertical: 'middle', horizontal: col === 2 || col === 6 ? 'left' : 'center' };
        });
    });
    const total = sum.addRow({ ch: '합계', title: `${included.length}개 장`, secs: sections, subs: subsections, page: '', file: '' });
    total.eachCell({ includeEmpty: true }, (cell, col) => {
        cell.border = border;
        cell.fill = fill(GRAY);
        cell.font = { name: FONT, size: 10, bold: true };
        cell.alignment = { vertical: 'middle', horizontal: col === 2 ? 'left' : 'center' };
    });

    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    return { blob, counts: { chapters: included.length, sections, subsections } };
}
