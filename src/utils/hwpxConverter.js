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

/**
 * PPTX의 표 데이터를 파싱하여 HWPX 양식 문서에 주입 및 병합/분할 생성합니다.
 * @param {File} pptxFile 
 * @param {File} hwpxTemplateFile 
 * @param {boolean} mergeMode 하나의 파일로 병합할지 여부 (true: 단일 HWPX 병합, false: 개별 HWPX들을 담은 ZIP 압축)
 * @returns {Promise<Blob>} 최종 생성된 HWPX Blob 혹은 ZIP Blob
 */
export async function fusePptToHwpxTemplate(pptxFile, hwpxTemplateFile, mergeMode = true) {
    if (!pptxFile || !hwpxTemplateFile) {
        throw new Error('PPTX 요건기술서 파일과 HWPX 양식 템플릿 파일을 모두 등록해주세요.');
    }

    const pptxBuffer = await pptxFile.arrayBuffer();
    const hwpxBuffer = await hwpxTemplateFile.arrayBuffer();

    const pptxZip = new JSZip();
    await pptxZip.loadAsync(pptxBuffer);

    // 1. PPTX 슬라이드별 표 데이터 파싱
    const slideFiles = Object.keys(pptxZip.files)
        .filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml'))
        .sort((a, b) => {
            const numA = parseInt(a.match(/\d+/)[0]);
            const numB = parseInt(b.match(/\d+/)[0]);
            return numA - numB;
        });

    const parser = new DOMParser();
    const nsA = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    const nsP = 'http://schemas.openxmlformats.org/presentationml/2006/main';

    const requirements = [];
    const idPattern = /^[A-Za-z0-9]+[\-_][A-Za-z0-9]+$/;

    for (const slidePath of slideFiles) {
        const slideXmlStr = await pptxZip.files[slidePath].async('text');
        const xmlDoc = parser.parseFromString(slideXmlStr, 'application/xml');
        
        const tblList = xmlDoc.getElementsByTagNameNS(nsA, 'tbl');
        for (let i = 0; i < tblList.length; i++) {
            const tbl = tblList[i];
            const trList = tbl.getElementsByTagNameNS(nsA, 'tr');
            
            for (let j = 0; j < trList.length; j++) {
                const tr = trList[j];
                const tcList = tr.getElementsByTagNameNS(nsA, 'tc');
                const rowCells = [];
                
                for (let k = 0; k < tcList.length; k++) {
                    const tc = tcList[k];
                    const tList = tc.getElementsByTagNameNS(nsA, 't');
                    let cellText = '';
                    for (let tIdx = 0; tIdx < tList.length; tIdx++) {
                        cellText += tList[tIdx].textContent || '';
                    }
                    rowCells.push(cellText.trim());
                }
                
                // 데이터 열 개수가 4개이고 첫 번째 셀에 고유번호가 존재하는지 감지
                if (rowCells.length >= 4) {
                    const col0 = rowCells[0];
                    const col1 = rowCells[1];
                    const col2 = rowCells[2];
                    const col3 = rowCells[3];
                    
                    const isHeader = col0.includes('고유번호') || col0.includes('요구사항') || col0.includes('분류') || col0.includes('No');
                    const hasValidId = idPattern.test(col0) || (col0.length >= 3 && col0.length <= 15 && !isHeader);
                    
                    if (hasValidId && !isHeader) {
                        requirements.push({
                            id: col0,
                            name: col1,
                            desc: col2,
                            detail: col3
                        });
                    }
                }
            }
        }
    }

    if (requirements.length === 0) {
        throw new Error('PPTX 파일 내에서 규격에 맞는 요구사항 기술서 표 데이터를 추출할 수 없습니다. (첫 번째 열이 ECR-XXX 같은 고유번호인 4개 열 구조의 표를 확인해주세요)');
    }

    const hpNS = 'http://www.hancom.co.kr/hwpml/2011/paragraph';

    // 셀 내용 치환 헬퍼 함수
    function fillHwpxCell(tc, text) {
        const subList = tc.getElementsByTagNameNS(hpNS, 'subList')[0] || tc.getElementsByTagName('hp:subList')[0];
        if (!subList) return;
        
        // 템플릿 문단 복제용 확보
        const existingPs = subList.getElementsByTagNameNS(hpNS, 'p');
        const templateP = existingPs.length > 0 ? existingPs[0].cloneNode(true) : null;
        
        // subList 안의 기존 단락 전수 제거
        while (subList.firstChild) {
            subList.removeChild(subList.firstChild);
        }
        
        if (!text) {
            if (templateP) {
                const t = templateP.getElementsByTagNameNS(hpNS, 't')[0];
                if (t) t.textContent = '';
                subList.appendChild(templateP);
            }
            return;
        }
        
        // 줄바꿈 단위로 쪼개서 개별 문단으로 복제 삽입
        const lines = text.split('\n');
        lines.forEach(line => {
            if (templateP) {
                const newP = templateP.cloneNode(true);
                newP.removeAttribute('id'); // ID 중복 방지
                const t = newP.getElementsByTagNameNS(hpNS, 't')[0];
                if (t) t.textContent = line;
                subList.appendChild(newP);
            } else {
                const newP = document.createElementNS(hpNS, 'hp:p');
                const run = document.createElementNS(hpNS, 'hp:run');
                const t = document.createElementNS(hpNS, 'hp:t');
                t.textContent = line;
                run.appendChild(t);
                newP.appendChild(run);
                subList.appendChild(newP);
            }
        });
    }

    // 템플릿 표 치환 핵심 함수 (지능형 플레이스홀더 감지 치환)
    function applyRequirementToP(pNode, req) {
        const tbl = pNode.getElementsByTagNameNS(hpNS, 'tbl')[0] || pNode.getElementsByTagName('hp:tbl')[0];
        if (!tbl) return;
        
        const tcList = tbl.getElementsByTagNameNS(hpNS, 'tc') || tbl.getElementsByTagName('hp:tc');
        for (let i = 0; i < tcList.length; i++) {
            const tc = tcList[i];
            const tNodes = tc.getElementsByTagNameNS(hpNS, 't') || tc.getElementsByTagName('hp:t');
            let cellText = '';
            for (let j = 0; j < tNodes.length; j++) {
                cellText += tNodes[j].textContent || '';
            }
            cellText = cellText.trim();
            
            if (cellText.includes('{고유번호}') || cellText.includes('{요구사항 고유번호}')) {
                fillHwpxCell(tc, req.id);
            } else if (cellText.includes('{요구사항 명칭}')) {
                fillHwpxCell(tc, req.name);
            } else if (cellText.includes('{정의}')) {
                fillHwpxCell(tc, req.desc);
            } else if (cellText.includes('{세부내용}')) {
                fillHwpxCell(tc, req.detail);
            }
        }
    }

    const serializer = new XMLSerializer();

    // 💡 하나의 HWPX 파일로 병합(페이지 나누기 연속 적용)
    if (mergeMode) {
        const hwpxZip = new JSZip();
        await hwpxZip.loadAsync(hwpxBuffer);
        
        const sectionXmlStr = await hwpxZip.files['Contents/section0.xml'].async('text');
        const sectionDoc = parser.parseFromString(sectionXmlStr, 'application/xml');
        
        // 본문 리스트에서 표를 가진 최외곽 hp:p 찾기
        const allPs = sectionDoc.getElementsByTagNameNS(hpNS, 'p');
        let templateP = null;
        for (let i = 0; i < allPs.length; i++) {
            const p = allPs[i];
            const tbls = p.getElementsByTagNameNS(hpNS, 'tbl');
            if (tbls.length > 0) {
                templateP = p;
                break;
            }
        }
        
        if (!templateP) {
            throw new Error('HWPX 양식 템플릿 파일에서 요구사항 기술서 표 양식을 찾을 수 없습니다.');
        }
        
        const parentNode = templateP.parentNode;
        
        // 기존 템플릿 표는 삭제하고, 고유번호만큼 복제하여 주입
        const fragment = sectionDoc.createDocumentFragment();
        
        requirements.forEach((req, idx) => {
            const clonedP = templateP.cloneNode(true);
            clonedP.removeAttribute('id'); // 중복 ID 리셋
            
            // 병합 시 두 번째 요구사항부터 강제 페이지 나누기 적용해 가독성 극대화
            if (idx > 0) {
                clonedP.setAttribute('pageBreak', '1');
            } else {
                clonedP.setAttribute('pageBreak', '0');
            }
            
            applyRequirementToP(clonedP, req);
            fragment.appendChild(clonedP);
            
            // 표 간의 살짝의 빈 여백을 주기 위해 페이지 나누기가 아닌 한 칸의 빈 단락 삽입
            if (idx < requirements.length - 1) {
                const spacerP = document.createElementNS(hpNS, 'hp:p');
                spacerP.setAttribute('pageBreak', '0');
                const run = document.createElementNS(hpNS, 'hp:run');
                const t = document.createElementNS(hpNS, 'hp:t');
                t.textContent = '';
                run.appendChild(t);
                spacerP.appendChild(run);
                fragment.appendChild(spacerP);
            }
        });
        
        parentNode.replaceChild(fragment, templateP);
        
        // XML 문자열 복원 및 덮어쓰기
        const finalXmlStr = serializer.serializeToString(sectionDoc);
        hwpxZip.file('Contents/section0.xml', finalXmlStr);
        
        const finalBuffer = await hwpxZip.generateAsync({ type: 'uint8array' });
        const finalBlob = new Blob([finalBuffer], { type: 'application/x-hwp-hwpx' });
        
        // 통계 바인딩
        finalBlob.fusedCount = requirements.length;
        finalBlob.requirementsList = requirements;
        finalBlob.mode = 'merge';
        return finalBlob;
    } 
    // 💡 개별 HWPX 파일로 분할 생성 후 하나의 ZIP 파일로 압축
    else {
        const outZip = new JSZip();
        
        for (const req of requirements) {
            const hwpxZip = new JSZip();
            await hwpxZip.loadAsync(hwpxBuffer);
            
            const sectionXmlStr = await hwpxZip.files['Contents/section0.xml'].async('text');
            const sectionDoc = parser.parseFromString(sectionXmlStr, 'application/xml');
            
            const allPs = sectionDoc.getElementsByTagNameNS(hpNS, 'p');
            let templateP = null;
            for (let i = 0; i < allPs.length; i++) {
                const p = allPs[i];
                const tbls = p.getElementsByTagNameNS(hpNS, 'tbl');
                if (tbls.length > 0) {
                    templateP = p;
                    break;
                }
            }
            
            if (templateP) {
                const clonedP = templateP.cloneNode(true);
                clonedP.removeAttribute('id');
                clonedP.setAttribute('pageBreak', '0');
                
                applyRequirementToP(clonedP, req);
                templateP.parentNode.replaceChild(clonedP, templateP);
            }
            
            const finalXmlStr = serializer.serializeToString(sectionDoc);
            hwpxZip.file('Contents/section0.xml', finalXmlStr);
            
            const fileBuffer = await hwpxZip.generateAsync({ type: 'uint8array' });
            
            // 파일명에 특수문자 제거 후 안전한 파일명 생성
            const safeId = req.id.replace(/[\\/:*?"<>|]/g, '_');
            const safeName = req.name.replace(/[\\/:*?"<>|]/g, '_').substring(0, 20);
            const fileName = `요구사항_${safeId}_${safeName}.hwpx`;
            
            outZip.file(fileName, fileBuffer);
        }
        
        const zipBuffer = await outZip.generateAsync({ type: 'uint8array' });
        const finalBlob = new Blob([zipBuffer], { type: 'application/zip' });
        
        finalBlob.fusedCount = requirements.length;
        finalBlob.requirementsList = requirements;
        finalBlob.mode = 'zip';
        return finalBlob;
    }
}
