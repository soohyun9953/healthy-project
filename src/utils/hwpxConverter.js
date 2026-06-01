import JSZip from 'jszip';
import { HWPXBuilder, write } from 'hwpx-js';

/**
 * PPTX 파일의 내용을 파싱하여 표준 한글 규격인 HWPX 파일 Blob으로 변환합니다.
 * @param {File} pptxFile 
 * @returns {Promise<Blob>} 변환된 HWPX Blob (통계 프로퍼티가 포함됨)
 */
export async function convertPptxToHwpx(pptxFile) {
    if (!pptxFile) throw new Error('변환할 파워포인트(PPTX) 파일이 존재하지 않습니다.');

    const arrayBuffer = await pptxFile.arrayBuffer();
    const zip = new JSZip();
    await zip.loadAsync(arrayBuffer);

    // 1. 슬라이드 XML 파일 탐색 및 숫자순 정렬
    const slideFiles = Object.keys(zip.files)
        .filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml'))
        .sort((a, b) => {
            const numA = parseInt(a.match(/\d+/)[0]);
            const numB = parseInt(b.match(/\d+/)[0]);
            return numA - numB;
        });

    if (slideFiles.length === 0) {
        throw new Error('파워포인트 파일 내에 슬라이드 XML 데이터가 존재하지 않습니다.');
    }

    const builder = new HWPXBuilder();
    builder.setPageSettings({
        pageWidth: 210, // A4 세로 규격 (mm)
        pageHeight: 297,
        pageMargin: { left: 20, right: 20, top: 20, bottom: 20, header: 15, footer: 15 }
    });

    const parser = new DOMParser();
    const nsA = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    const nsP = 'http://schemas.openxmlformats.org/presentationml/2006/main';
    const nsR = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

    let totalParagraphs = 0;
    let totalTables = 0;
    let totalImages = 0;

    // 슬라이드별 루프 가동
    for (let idx = 0; idx < slideFiles.length; idx++) {
        const slidePath = slideFiles[idx];
        const slideXmlStr = await zip.files[slidePath].async('text');
        const xmlDoc = parser.parseFromString(slideXmlStr, 'application/xml');

        if (xmlDoc.getElementsByTagName('parsererror').length > 0) {
            console.error(`Slide ${idx + 1} 파싱 에러 건너뜀`);
            continue;
        }

        // 해당 슬라이드의 Relationships (.rels) 파일을 읽어 이미지 리소스 경로 해독 준비
        const slideNum = slidePath.match(/\d+/)[0];
        const relsPath = `ppt/slides/_rels/slide${slideNum}.xml.rels`;
        const relsMap = new Map();
        if (zip.files[relsPath]) {
            const relsStr = await zip.files[relsPath].async('text');
            const relsDoc = parser.parseFromString(relsStr, 'application/xml');
            const relationships = relsDoc.getElementsByTagName('Relationship');
            for (let r = 0; r < relationships.length; r++) {
                const rel = relationships[r];
                const rId = rel.getAttribute('Id');
                const target = rel.getAttribute('Target');
                relsMap.set(rId, target);
            }
        }

        const elements = [];

        // 💡 [조치 1]: 도형 및 텍스트 상자 (<p:sp>) 추출
        const spList = xmlDoc.getElementsByTagNameNS(nsP, 'sp');
        for (let i = 0; i < spList.length; i++) {
            const sp = spList[i];
            
            // 좌표 (off x, y) 추출
            let x = 0, y = 0;
            const off = sp.getElementsByTagNameNS(nsA, 'off')[0];
            if (off) {
                x = parseInt(off.getAttribute('x') || '0');
                y = parseInt(off.getAttribute('y') || '0');
            }

            // 텍스트 단락 수집
            const pList = sp.getElementsByTagNameNS(nsA, 'p');
            const paragraphs = [];
            for (let j = 0; j < pList.length; j++) {
                const p = pList[j];
                const tList = p.getElementsByTagNameNS(nsA, 't');
                let paraText = '';
                for (let k = 0; k < tList.length; k++) {
                    paraText += tList[k].textContent || '';
                }
                const trimmed = paraText.trim();
                if (trimmed) {
                    paragraphs.push(trimmed);
                }
            }

            if (paragraphs.length > 0) {
                elements.push({
                    type: 'text',
                    x,
                    y,
                    data: paragraphs
                });
            }
        }

        // 💡 [조치 2]: 표 (<a:tbl>) 추출
        const tblList = xmlDoc.getElementsByTagNameNS(nsA, 'tbl');
        for (let i = 0; i < tblList.length; i++) {
            const tbl = tblList[i];
            
            // 표의 상위 graphicFrame에서 좌표 획득
            let x = 0, y = 0;
            let parent = tbl.parentNode;
            while (parent) {
                const pName = parent.localName || parent.tagName.split(':').pop();
                if (pName === 'graphicFrame') {
                    const off = parent.getElementsByTagNameNS(nsA, 'off')[0];
                    if (off) {
                        x = parseInt(off.getAttribute('x') || '0');
                        y = parseInt(off.getAttribute('y') || '0');
                    }
                    break;
                }
                parent = parent.parentNode;
            }

            // 2차원 표 셀 데이터 생성
            const trList = tbl.getElementsByTagNameNS(nsA, 'tr');
            const tableData = [];
            for (let r = 0; r < trList.length; r++) {
                const tr = trList[r];
                const tcList = tr.getElementsByTagNameNS(nsA, 'tc');
                const rowData = [];
                for (let c = 0; c < tcList.length; c++) {
                    const tc = tcList[c];
                    const tList = tc.getElementsByTagNameNS(nsA, 't');
                    let cellText = '';
                    for (let k = 0; k < tList.length; k++) {
                        cellText += tList[k].textContent || '';
                    }
                    rowData.push(cellText.trim());
                }
                tableData.push(rowData);
            }

            if (tableData.length > 0) {
                elements.push({
                    type: 'table',
                    x,
                    y,
                    data: tableData
                });
            }
        }

        // 💡 [조치 3]: 이미지 (<p:pic>) 추출
        const picList = xmlDoc.getElementsByTagNameNS(nsP, 'pic');
        for (let i = 0; i < picList.length; i++) {
            const pic = picList[i];
            
            let x = 0, y = 0;
            const off = pic.getElementsByTagNameNS(nsA, 'off')[0];
            if (off) {
                x = parseInt(off.getAttribute('x') || '0');
                y = parseInt(off.getAttribute('y') || '0');
            }

            // 그림 리소스 rId 탐색
            const blip = pic.getElementsByTagNameNS(nsA, 'blip')[0];
            if (blip) {
                const rId = blip.getAttributeNS(nsR, 'embed') || blip.getAttribute('r:embed');
                if (rId && relsMap.has(rId)) {
                    let targetPath = relsMap.get(rId);
                    // 상대 경로 보정 (../media/image1.png -> ppt/media/image1.png)
                    if (targetPath.startsWith('../')) {
                        targetPath = 'ppt/' + targetPath.replace('../', '');
                    } else if (!targetPath.startsWith('ppt/')) {
                        targetPath = 'ppt/' + targetPath;
                    }

                    if (zip.files[targetPath]) {
                        const imgData = await zip.files[targetPath].async('uint8array');
                        const ext = targetPath.split('.').pop().toLowerCase();
                        elements.push({
                            type: 'image',
                            x,
                            y,
                            ext: ext === 'jpeg' ? 'jpg' : ext,
                            data: imgData
                        });
                    }
                }
            }
        }

        // 💡 [조치 4]: 위 ➜ 아래, 왼쪽 ➜ 오른쪽 흐름 정렬 알고리즘 가동
        // EMU 좌표 기준, y좌표가 유사(대략 1cm = 360,000 EMU 이내)한 경우 수평 정렬로 간주해 x좌표 기준으로 나열합니다.
        elements.sort((a, b) => {
            const yDiff = a.y - b.y;
            if (Math.abs(yDiff) < 360000) {
                return a.x - b.x;
            }
            return yDiff;
        });

        // 💡 [조치 5]: 슬라이드 헤더 추가 및 HWPXBuilder 빌드 적용
        builder.addParagraph(`■ 슬라이드 ${idx + 1}`, { fontSize: 13, bold: true, color: '#a855f7' });
        builder.addEmptyParagraph();

        for (const el of elements) {
            if (el.type === 'text') {
                el.data.forEach(text => {
                    builder.addParagraph(text, { fontSize: 11 });
                    totalParagraphs++;
                });
                builder.addEmptyParagraph();
            } else if (el.type === 'table') {
                builder.addTable(el.data, {
                    width: 170, // A4 용지 폭 감안
                    borderStyle: 'SOLID',
                    cellPadding: 1.2
                });
                totalTables++;
                builder.addEmptyParagraph();
            } else if (el.type === 'image') {
                try {
                    // hwpx-js binary 이미지 객체 임베딩
                    builder.addImage(el.data, el.ext, { width: 120, height: 90 });
                    totalImages++;
                    builder.addEmptyParagraph();
                } catch (imgErr) {
                    console.error("그림 주입 실패 건너뜀:", imgErr);
                }
            }
        }
    }

    const doc = builder.build();
    const uint8 = await write(doc);
    
    // 최종 결과 파일 생성
    const blob = new Blob([uint8], { type: 'application/x-hwp-hwpx' });
    
    // 통계 메타데이터 바인딩
    blob.totalSlidesCount = slideFiles.length;
    blob.totalParagraphsCount = totalParagraphs;
    blob.totalTablesCount = totalTables;
    blob.totalImagesCount = totalImages;

    return blob;
}
