import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { extractStructuredFileContent, getFileExtension } from './fileExtractor.js';

// LocalStorage 영속화 키
export const SSOT_STORAGE_KEY = 'healthy_project_ssot_v1';

// ── 기본 프로젝트 마스터 메타데이터 프리셋 ────────────────────────
export const DEFAULT_PROJECT_METADATA = {
    projectName: '공공보건의료 AI 기반 지능형 플랫폼 구축 ISP',
    clientAgency: '국립중앙의료원',
    period: '2026.04 ~ 2026.11 (8개월)',
    budget: '500,000,000원 (금오억원)'
};

// ── 공공의료 / ISP 자주 틀리는 표준 용어사전 기본 프리셋 ──────────────
export const DEFAULT_GLOSSARY = [
    {
        id: 'glo-1',
        category: '기관명',
        approved: '지방의료원',
        forbidden: '지역의료원, 도립병원, 시립의료원',
        note: '공공보건의료법 제2조 및 지방의료원법 공식 법정명칭'
    },
    {
        id: 'glo-2',
        category: '시스템명',
        approved: '건강정보고속도로',
        forbidden: '건강정보 고속도로, 보건의료고속도로, PHR고속도로, 의료데이터고속도로',
        note: '보건복지부 공식 서비스명 (띄어쓰기 없는 단일 고유명사)'
    },
    {
        id: 'glo-3',
        category: '행정용어',
        approved: '책무',
        forbidden: '의무사항, 필수책임, 법적의무',
        note: '공공 표준 행정법률 용어'
    },
    {
        id: 'glo-4',
        category: '약어',
        approved: 'SSOT',
        forbidden: '단일진실공급원, SSO, SOT, 단일진실원천',
        note: 'Single Source of Truth 표준 정보화 약어'
    },
    {
        id: 'glo-5',
        category: '시스템명',
        approved: '진료정보교류시스템',
        forbidden: '진료기록교류, 의료정보교류망, 환자정보교류시스템',
        note: '보건복지부 및 한국보건의료정보원 공식 사업명'
    },
    {
        id: 'glo-6',
        category: '행정용어',
        approved: '소요예산',
        forbidden: '필요예산, 투입비용, 사업예산액, 소요자금',
        note: '정부 공공 ISP 표준 서식 템플릿 항목명'
    },
    {
        id: 'glo-7',
        category: '기관명',
        approved: '보건복지부',
        forbidden: '복지부, 보건부',
        note: '정부조직법상 중앙행정기관 공식 명칭 풀네임 기재'
    },
    {
        id: 'glo-8',
        category: '행정용어',
        approved: '과제정의서',
        forbidden: '과제설명서, 단위과제서, 세부과제서',
        note: 'ISP/ISMP 표준 이행계획 산출물 공식 명칭'
    },
    {
        id: 'glo-9',
        category: '시스템명',
        approved: '마이차트',
        forbidden: 'MyChart, 나의차트, 마이헬스차트',
        note: '공공 포털 서비스 공식 한글 표기'
    },
    {
        id: 'glo-10',
        category: '행정용어',
        approved: '이행계획',
        forbidden: '실행계획, 추진로드맵, 전개계획',
        note: '공공 정보화전략계획(ISP) 가이드라인 공식 단계명'
    },
    {
        id: 'glo-11',
        category: '기관명',
        approved: '국립중앙의료원',
        forbidden: '중앙의료원, 국립의료원, NMC',
        note: '발주기관 공식 명칭 (본문 내 한글 정식명칭 원칙)'
    },
    {
        id: 'glo-12',
        category: '약어',
        approved: 'CDSS',
        forbidden: 'CDS, 임상결정시스템, 임상의사결정지원',
        note: 'Clinical Decision Support System 표준 약어'
    }
];

// ── 요구사항 마스터 리스트 기본 프리셋 ────────────────────────────
export const DEFAULT_REQUIREMENTS = [
    {
        id: 'CNR-001',
        name: '공공보건의료 데이터 통합 연계 및 표준화 체계 수립',
        keywords: '데이터, 연계, 표준화, EMR, 연계체계'
    },
    {
        id: 'SFR-001',
        name: 'AI 기반 진료지원(CDSS) 알고리즘 실증 및 고도화',
        keywords: 'AI, 진료지원, CDSS, 알고리즘, 실증'
    },
    {
        id: 'SFR-002',
        name: '실시간 응급의료 관제 및 이송 모니터링 대시보드 구축',
        keywords: '응급의료, 관제, 모니터링, 대시보드, 이송'
    },
    {
        id: 'SFR-003',
        name: '지방의료원 간 진료정보교류 및 협진 지원 인터페이스 개발',
        keywords: '지방의료원, 진료정보교류, 협진, 인터페이스'
    },
    {
        id: 'PMR-001',
        name: '개인정보보호 및 의료법 규제 준수 거버넌스 수립',
        keywords: '개인정보, 의료법, 거버넌스, 가명처리, 규제'
    },
    {
        id: 'TER-001',
        name: '민간·공공 하이브리드 클라우드 인프라 표준 아키텍처 설계',
        keywords: '클라우드, 인프라, 아키텍처, 하이브리드, 가용성'
    },
    {
        id: 'DSR-001',
        name: '데이터 레이크 및 AI 학습용 합성데이터셋 파이프라인 구축',
        keywords: '데이터 레이크, 학습데이터, 합성데이터, 파이프라인'
    },
    {
        id: 'COR-001',
        name: '장애 대응 이중화 및 재해복구(DR) 체계 구축 방안',
        keywords: '이중화, 재해복구, DR, 장애대응, 무중단'
    }
];

// ── LocalStorage 로드 / 저장 유틸 ──────────────────────────────
export function loadSSOTData() {
    try {
        const raw = localStorage.getItem(SSOT_STORAGE_KEY);
        if (raw) {
            const parsed = JSON.parse(raw);
            return {
                metadata: { ...DEFAULT_PROJECT_METADATA, ...(parsed.metadata || {}) },
                glossary: Array.isArray(parsed.glossary) && parsed.glossary.length > 0 ? parsed.glossary : DEFAULT_GLOSSARY,
                requirements: Array.isArray(parsed.requirements) && parsed.requirements.length > 0 ? parsed.requirements : DEFAULT_REQUIREMENTS
            };
        }
    } catch (e) {
        console.warn('SSOT LocalStorage 로드 실패, 기본값 사용:', e);
    }
    return {
        metadata: { ...DEFAULT_PROJECT_METADATA },
        glossary: [...DEFAULT_GLOSSARY],
        requirements: [...DEFAULT_REQUIREMENTS]
    };
}

export function saveSSOTData(data) {
    try {
        localStorage.setItem(SSOT_STORAGE_KEY, JSON.stringify(data));
        return true;
    } catch (e) {
        console.error('SSOT LocalStorage 저장 실패:', e);
        return false;
    }
}

// ── 표(Table) 전용 추출 엔진 ──────────────────────────────────
// DOCX, PPTX, HWPX, XLSX, TXT 파일 내의 표 구조를 2D 배열 형태(행x열)로 추출
export async function extractTablesFromFile(file) {
    const ext = getFileExtension(file.name);
    const tables = []; // { source: string, page: number, rows: string[][] }

    try {
        if (ext === 'xlsx' || ext === 'xls') {
            const arrayBuffer = await file.arrayBuffer();
            const wb = XLSX.read(arrayBuffer, { type: 'array' });
            let sheetIdx = 1;
            for (const sheetName of wb.SheetNames) {
                const sheet = wb.Sheets[sheetName];
                if (!sheet) continue;
                const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
                if (rows.length > 0) {
                    tables.push({
                        source: `엑셀 시트 [${sheetName}]`,
                        page: sheetIdx++,
                        rows: rows.map(r => r.map(c => String(c ?? '').trim()))
                    });
                }
            }
        } else if (ext === 'pptx') {
            const arrayBuffer = await file.arrayBuffer();
            const zip = await JSZip.loadAsync(arrayBuffer);
            const slideFiles = Object.keys(zip.files).filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name));
            
            for (const fileName of slideFiles) {
                const slideNum = parseInt(fileName.match(/\d+/)[0], 10);
                const content = await zip.files[fileName].async('string');
                const parser = new DOMParser();
                const xmlDoc = parser.parseFromString(content, 'text/xml');
                const tblList = Array.from(xmlDoc.getElementsByTagNameNS('*', 'tbl'));

                tblList.forEach((tbl, tIdx) => {
                    const trList = Array.from(tbl.getElementsByTagNameNS('*', 'tr'));
                    const tableRows = [];
                    trList.forEach(tr => {
                        const tcList = Array.from(tr.getElementsByTagNameNS('*', 'tc'));
                        const rowCells = tcList.map(tc => {
                            const tElements = Array.from(tc.getElementsByTagNameNS('*', 't'));
                            return tElements.map(t => t.textContent || '').join(' ').trim();
                        });
                        if (rowCells.some(c => c.length > 0)) {
                            tableRows.push(rowCells);
                        }
                    });
                    if (tableRows.length > 1) {
                        tables.push({
                            source: `슬라이드 ${slideNum} 표 #${tIdx + 1}`,
                            page: slideNum,
                            rows: tableRows
                        });
                    }
                });
            }
        } else if (ext === 'docx') {
            const arrayBuffer = await file.arrayBuffer();
            const zip = await JSZip.loadAsync(arrayBuffer);
            const docFile = zip.file('word/document.xml');
            if (docFile) {
                const content = await docFile.async('string');
                const parser = new DOMParser();
                const xmlDoc = parser.parseFromString(content, 'text/xml');
                const tblList = Array.from(xmlDoc.getElementsByTagNameNS('*', 'tbl'));

                tblList.forEach((tbl, tIdx) => {
                    const trList = Array.from(tbl.getElementsByTagNameNS('*', 'tr'));
                    const tableRows = [];
                    trList.forEach(tr => {
                        const tcList = Array.from(tr.getElementsByTagNameNS('*', 'tc'));
                        const rowCells = tcList.map(tc => {
                            const tElements = Array.from(tc.getElementsByTagNameNS('*', 't'));
                            return tElements.map(t => t.textContent || '').join(' ').trim();
                        });
                        if (rowCells.some(c => c.length > 0)) {
                            tableRows.push(rowCells);
                        }
                    });
                    if (tableRows.length > 1) {
                        tables.push({
                            source: `워드 표 #${tIdx + 1}`,
                            page: 1,
                            rows: tableRows
                        });
                    }
                });
            }
        } else if (ext === 'hwpx') {
            const arrayBuffer = await file.arrayBuffer();
            const zip = await JSZip.loadAsync(arrayBuffer);
            const secFiles = Object.keys(zip.files).filter(name => /^Contents\/section\d+\.xml$/.test(name));

            for (let sIdx = 0; sIdx < secFiles.length; sIdx++) {
                const content = await zip.files[secFiles[sIdx]].async('string');
                const parser = new DOMParser();
                const xmlDoc = parser.parseFromString(content, 'text/xml');
                const tblList = Array.from(xmlDoc.getElementsByTagNameNS('*', 'tbl'));

                tblList.forEach((tbl, tIdx) => {
                    const trList = Array.from(tbl.getElementsByTagNameNS('*', 'tr'));
                    const tableRows = [];
                    trList.forEach(tr => {
                        const tcList = Array.from(tr.getElementsByTagNameNS('*', 'tc'));
                        const rowCells = tcList.map(tc => {
                            const tElements = Array.from(tc.getElementsByTagNameNS('*', 't'));
                            return tElements.map(t => t.textContent || '').join(' ').trim();
                        });
                        if (rowCells.some(c => c.length > 0)) {
                            tableRows.push(rowCells);
                        }
                    });
                    if (tableRows.length > 1) {
                        tables.push({
                            source: `HWPX 섹션 ${sIdx + 1} 표 #${tIdx + 1}`,
                            page: sIdx + 1,
                            rows: tableRows
                        });
                    }
                });
            }
        }
    } catch (err) {
        console.warn('표 전용 추출 중 일부 오류(텍스트 파싱으로 폴백):', err);
    }

    return tables;
}

// ── Stage 1: 결정론적 Rule 기반 검증 엔진 ─────────────────────────
/**
 * 1) 금지 동의어/오표기 검출 (표준어 보호 가드 및 단어 경계 검사)
 * 2) 요구사항 ID 추적성(RTM) 검증 (단위 산출물 vs 종합보고서 검수 범위 모드 지원)
 * 3) 수치 및 표 산술 검증 (금액단위 불일치, 합계/총계 산술 오류)
 */
export async function runStage1Validation(fileContent, tables, masterData, options = {}) {
    const {
        inspectionScope = 'unit' // 'unit' (단위 산출물/특정 장표) | 'full' (종합 제안서/최종보고서)
    } = options;

    const { metadata, glossary, requirements } = masterData;
    const { text, pageMap } = fileContent;

    const issues = [];
    let issueSeq = 1;

    // A. 1) 금지 동의어 / 오표기 검출 (오탐 0건 보장 로직)
    glossary.forEach(item => {
        const approved = (item.approved || '').trim();
        const forbiddenList = (item.forbidden || '')
            .split(',')
            .map(s => s.trim())
            .filter(s => s.length > 0 && s !== approved);

        forbiddenList.forEach(forbiddenTerm => {
            if (forbiddenTerm.length < 2) return;

            const isAscii = /^[A-Za-z0-9_-]+$/.test(forbiddenTerm);

            // 페이지별 텍스트 검색
            Object.entries(pageMap || {}).forEach(([pageNum, lines]) => {
                const pageLines = Array.isArray(lines) ? lines : [lines];
                pageLines.forEach(line => {
                    let searchIdx = 0;
                    while (true) {
                        const foundIdx = line.indexOf(forbiddenTerm, searchIdx);
                        if (foundIdx === -1) break;

                        // ── [오탐 방지 핵심 가드 1: 표준어 내부 포함 검사] ──
                        // 만약 이 금지어가 이미 올바르게 작성된 표준어(approved)의 일부라면 무조건 스킵!
                        // 예: 표준어 '국립중앙의료원' 안에 포함된 '국립중앙'은 정상 단어의 일부이므로 결함 아님
                        let isInsideApproved = false;
                        if (approved && approved.includes(forbiddenTerm)) {
                            let appIdx = 0;
                            while (true) {
                                const foundApp = line.indexOf(approved, appIdx);
                                if (foundApp === -1) break;
                                if (foundIdx >= foundApp && (foundIdx + forbiddenTerm.length) <= (foundApp + approved.length)) {
                                    isInsideApproved = true;
                                    break;
                                }
                                appIdx = foundApp + approved.length;
                            }
                        }

                        if (isInsideApproved) {
                            searchIdx = foundIdx + forbiddenTerm.length;
                            continue;
                        }

                        // ── [오탐 방지 핵심 가드 2: 영문/숫자 약어 단어 경계 검사] ──
                        // 예: 'CDS' 검사 시 'CDSS'나 'LESSON' 내부에 매칭되지 않도록 앞뒤 영숫자 확인
                        if (isAscii) {
                            const prevChar = foundIdx > 0 ? line[foundIdx - 1] : ' ';
                            const nextChar = (foundIdx + forbiddenTerm.length < line.length) ? line[foundIdx + forbiddenTerm.length] : ' ';
                            if (/[A-Za-z0-9_]/.test(prevChar) || /[A-Za-z0-9_]/.test(nextChar)) {
                                searchIdx = foundIdx + forbiddenTerm.length;
                                continue;
                            }
                        }

                        // 문맥 발췌 (앞뒤 25자)
                        const start = Math.max(0, foundIdx - 25);
                        const end = Math.min(line.length, foundIdx + forbiddenTerm.length + 25);
                        const snippet = (start > 0 ? '...' : '') + line.substring(start, end) + (end < line.length ? '...' : '');

                        issues.push({
                            id: `S1-GLO-${issueSeq++}`,
                            stage: 'Stage 1 (결정론적 룰)',
                            category: '용어 불일치',
                            errorType: '금지동의어/오표기',
                            severity: 'HIGH',
                            page: parseInt(pageNum, 10) || 1,
                            matchedTerm: forbiddenTerm,
                            suggestedText: approved,
                            originalText: snippet,
                            reason: `[${item.category || '용어'}] 표준용어 '${approved}' 대신 금지/오표기 '${forbiddenTerm}'이(가) 사용되었습니다.${item.note ? ` (참고: ${item.note})` : ''}`
                        });

                        searchIdx = foundIdx + forbiddenTerm.length;
                    }
                });
            });
        });
    });

    // B. 2) 요구사항 ID 추적성(RTM) 검증
    // 1단계: 마스터에 등록된 요구사항 ID의 문서 내 실제 출현 여부를 100% 정밀 대조
    const matchedIdsInDoc = new Set();
    const idLocations = {}; // { [id]: number[] pages }

    requirements.forEach(req => {
        const cleanId = (req.id || '').trim();
        if (!cleanId) return;
        const upperId = cleanId.toUpperCase();

        Object.entries(pageMap || {}).forEach(([pageNum, lines]) => {
            const pageText = (Array.isArray(lines) ? lines.join(' ') : String(lines)).toUpperCase();
            if (pageText.includes(upperId)) {
                matchedIdsInDoc.add(cleanId);
                if (!idLocations[cleanId]) idLocations[cleanId] = [];
                const p = parseInt(pageNum, 10) || 1;
                if (!idLocations[cleanId].includes(p)) {
                    idLocations[cleanId].push(p);
                }
            }
        });
    });

    // 2단계: 본문 정규식 전수 검색 (미등록 유령 Ghost ID 탐색용)
    const idRegex = /\b([A-Z]{2,6}[-_]\d{2,4}(?:[-_]\d{1,3})?)\b/g;
    Object.entries(pageMap || {}).forEach(([pageNum, lines]) => {
        const pageText = Array.isArray(lines) ? lines.join(' ') : String(lines);
        let m;
        while ((m = idRegex.exec(pageText)) !== null) {
            const detectedId = m[1].toUpperCase();
            // 공통 표준 기술 약어(ISO-9001, UTF-8 등) 제외
            if (!['ISO-', 'UTF-', 'RFC-', 'IPV4-', 'IPV6-'].some(prefix => detectedId.startsWith(prefix))) {
                matchedIdsInDoc.add(detectedId);
                if (!idLocations[detectedId]) idLocations[detectedId] = [];
                const p = parseInt(pageNum, 10) || 1;
                if (!idLocations[detectedId].includes(p)) {
                    idLocations[detectedId].push(p);
                }
            }
        }
    });

    const masterIdMap = new Map();
    requirements.forEach(req => {
        const cleanId = (req.id || '').trim();
        if (cleanId) {
            masterIdMap.set(cleanId.toUpperCase(), req);
        }
    });

    // B-1) 누락 ID (Missing ID): 마스터에 있는데 산출물에 없음
    const rtmMatrix = [];
    requirements.forEach(req => {
        const cleanId = (req.id || '').trim();
        const isPresent = matchedIdsInDoc.has(cleanId) || matchedIdsInDoc.has(cleanId.toUpperCase());
        const pages = idLocations[cleanId] || idLocations[cleanId.toUpperCase()] || [];

        // 필수 키워드 검사
        let keywordsMatched = true;
        let missingKeywords = [];
        if (req.keywords && isPresent) {
            const kwList = req.keywords.split(',').map(k => k.trim()).filter(Boolean);
            missingKeywords = kwList.filter(k => !text.includes(k));
            if (missingKeywords.length > 0) {
                keywordsMatched = false;
            }
        }

        const status = isPresent 
            ? (keywordsMatched ? '정상' : '키워드 일부 누락')
            : (inspectionScope === 'unit' ? '미출현 (단위산출물 범위 외)' : '누락(Missing)');

        rtmMatrix.push({
            id: cleanId,
            name: req.name || '',
            keywords: req.keywords || '',
            isPresent,
            pages,
            status,
            missingKeywords
        });

        if (!isPresent) {
            // 종합 검수(full) 모드일 때만 치명적 결함(HIGH)으로 보고서에 등록
            // 단위 산출물(unit) 모드일 때는 RTM 표에만 미출현으로 표기하고 결함에는 등록하지 않음 (오탐 방지)
            if (inspectionScope === 'full') {
                issues.push({
                    id: `S1-RTM-MIS-${issueSeq++}`,
                    stage: 'Stage 1 (결정론적 룰)',
                    category: '요구사항 누락',
                    errorType: '요구사항 미출현(Missing ID)',
                    severity: 'HIGH',
                    page: 1,
                    matchedTerm: cleanId,
                    suggestedText: `요구사항 [${cleanId}] ${req.name || ''} 본문 반영`,
                    originalText: `(종합 산출물 본문에서 [${cleanId}] 미발견)`,
                    reason: `마스터 요구사항 ID '${cleanId}' [${req.name || ''}]이(가) 종합 산출물 내에서 단 한 번도 언급되지 않았습니다.`
                });
            }
        } else if (!keywordsMatched && missingKeywords.length > 0) {
            issues.push({
                id: `S1-RTM-KW-${issueSeq++}`,
                stage: 'Stage 1 (결정론적 룰)',
                category: '요구사항 불완전',
                errorType: '요구사항 필수 키워드 누락',
                severity: 'LOW',
                page: pages[0] || 1,
                matchedTerm: cleanId,
                suggestedText: `누락 키워드 [${missingKeywords.join(', ')}] 본문 보강`,
                originalText: `요구사항 ${cleanId} 출현 위치 (p.${pages.join(', ')})`,
                reason: `요구사항 '${cleanId}'은 기재되었으나, 필수 지정 키워드 [${missingKeywords.join(', ')}]이(가) 본문에 충분히 반영되지 않았습니다.`
            });
        }
    });

    // B-2) 유령 ID (Ghost ID): 산출물에 있는데 마스터에 정의되지 않은 ID
    matchedIdsInDoc.forEach(detectedId => {
        if (!masterIdMap.has(detectedId) && !masterIdMap.has(detectedId.toUpperCase())) {
            const pages = idLocations[detectedId] || [1];
            issues.push({
                id: `S1-RTM-GHO-${issueSeq++}`,
                stage: 'Stage 1 (결정론적 룰)',
                category: '요구사항 유령ID',
                errorType: '미등록 유령 ID(Ghost ID)',
                severity: 'MEDIUM',
                page: pages[0] || 1,
                matchedTerm: detectedId,
                suggestedText: '공식 요구사항 마스터 ID로 정정 또는 마스터 리스트 신규 등록',
                originalText: `본문 내 출현: ${detectedId} (p.${pages.join(', ')})`,
                reason: `산출물에 기재된 ID '${detectedId}'은(는) 요구사항 마스터 리스트에 등록되어 있지 않은 미등록/오타 의심 ID입니다.`
            });
        }
    });

    // C. 3) 수치 및 표 산술 검증
    // C-1) 프로젝트 총사업비 표기 대조 검증
    if (metadata.budget) {
        // 숫자만 추출 (예: '500,000,000' -> 500000000)
        const budgetNums = metadata.budget.match(/\d[\d,.]*/g);
        if (budgetNums && budgetNums.length > 0) {
            const rawBudgetStr = budgetNums[0].replace(/,/g, '');
            const budgetVal = parseFloat(rawBudgetStr);

            // 본문 내 금액 불일치 패턴 (예: "50,000,000원", "5,000,000,000원" 등 자릿수 오기)
            const moneyRegex = /([0-9]{1,3}(?:,[0-9]{3})+)\s*(?:원|백만원|천원|억원)/g;
            Object.entries(pageMap || {}).forEach(([pageNum, lines]) => {
                const pageText = Array.isArray(lines) ? lines.join(' ') : String(lines);
                let m;
                while ((m = moneyRegex.exec(pageText)) !== null) {
                    const foundMoneyStr = m[1].replace(/,/g, '');
                    const foundVal = parseFloat(foundMoneyStr);
                    // 총사업비와 자릿수가 10배/100배 차이나는데 주변에 '총사업비' 또는 '사업비' 단어가 있는 경우
                    const windowText = pageText.substring(Math.max(0, m.index - 30), Math.min(pageText.length, m.index + 50));
                    if ((windowText.includes('총사업비') || windowText.includes('총 사업비') || windowText.includes('예산액')) && foundVal !== budgetVal) {
                        issues.push({
                            id: `S1-NUM-BDG-${issueSeq++}`,
                            stage: 'Stage 1 (결정론적 룰)',
                            category: '수치 오류',
                            errorType: '총사업비 수치 불일치',
                            severity: 'HIGH',
                            page: parseInt(pageNum, 10) || 1,
                            matchedTerm: m[0],
                            suggestedText: metadata.budget,
                            originalText: windowText.trim(),
                            reason: `마스터에 등록된 총사업비 '${metadata.budget}'와 본문 표기 '${m[0]}'이(가) 서로 일치하지 않습니다.`
                        });
                    }
                }
            });
        }
    }

    // C-2) 표(Table) 산술 덧셈(합계/총계/계) 검증
    tables.forEach((table, tIdx) => {
        const { rows, source, page } = table;
        if (!rows || rows.length < 3) return;

        // 합계 행 탐색 ('합계', '총계', '계', '소계', 'Total', '합 계')
        const totalRowKeywords = ['합계', '총계', '소계', '합 계', '총 계', '계', 'total', 'sum'];

        rows.forEach((row, rowIdx) => {
            const hasTotalKeyword = row.some(cell => {
                const clean = cell.replace(/\s+/g, '').toLowerCase();
                return totalRowKeywords.includes(clean);
            });

            if (!hasTotalKeyword || rowIdx === 0) return; // 첫 행(헤더)은 제외

            // 합계 행의 각 열에 대해 상위 행들의 숫자 합산 검증
            row.forEach((cellVal, colIdx) => {
                const cleanCell = cellVal.replace(/[,%원\s]/g, '');
                if (!cleanCell || isNaN(cleanCell) || cleanCell.length === 0) return;

                const recordedTotal = parseFloat(cleanCell);
                if (isNaN(recordedTotal) || recordedTotal === 0) return;

                // 상위 행들의 동일 컬럼 숫자 수집
                let calculatedSum = 0;
                let numericCount = 0;
                const summandSnippets = [];

                for (let prevR = 0; prevR < rowIdx; prevR++) {
                    const prevCell = (rows[prevR][colIdx] || '').replace(/[,%원\s]/g, '');
                    if (prevCell && !isNaN(prevCell) && prevCell.length > 0) {
                        const num = parseFloat(prevCell);
                        calculatedSum += num;
                        numericCount++;
                        summandSnippets.push(rows[prevR][colIdx]);
                    }
                }

                // 2개 이상의 항목이 더해진 합계 행인 경우 오차 확인
                if (numericCount >= 2) {
                    const diff = Math.abs(calculatedSum - recordedTotal);
                    // 부동소수점 오차 감안
                    if (diff > 0.01) {
                        issues.push({
                            id: `S1-NUM-TBL-${issueSeq++}`,
                            stage: 'Stage 1 (결정론적 룰)',
                            category: '수치 오류',
                            errorType: '표 합계 산술 불일치',
                            severity: 'HIGH',
                            page: page || 1,
                            matchedTerm: `기재된 합계: ${cellVal}`,
                            suggestedText: `계산된 실제 합계: ${calculatedSum.toLocaleString()}`,
                            originalText: `[${source}] ${rowIdx + 1}행 ${colIdx + 1}열 (항목값: ${summandSnippets.slice(0, 4).join(' + ')}${summandSnippets.length > 4 ? '...' : ''} = 계산치 ${calculatedSum.toLocaleString()} vs 기재치 ${cellVal})`,
                            reason: `${source}에서 상위 ${numericCount}개 행의 합계는 ${calculatedSum.toLocaleString()}이어야 하나, 표에는 ${cellVal}(으)로 잘못 기재되어 산술 오차가 발생했습니다.`
                        });
                    }
                }
            });
        });
    });

    // RTM 매핑률 계산
    const totalMasterReqCount = requirements.length;
    const coveredReqCount = rtmMatrix.filter(r => r.isPresent).length;
    const rtmMappingRate = totalMasterReqCount > 0 
        ? Math.round((coveredReqCount / totalMasterReqCount) * 100) 
        : 100;

    return {
        issues,
        rtmMatrix,
        rtmMappingRate,
        coveredReqCount,
        totalMasterReqCount
    };
}

// ── Stage 2: Gemini AI 문맥·오탈자 정밀 검증 ──────────────────────
/**
 * 긴 문서를 3,000자 단위 청크로 분할하여 /api/gemini-proxy로 호출
 * JSON Schema:
 * [{ "chunkIndex": 1, "errorType": "맞춤법|문맥모순|행정용어부적절|수치불일치", "originalText": "원문발췌", "suggestedText": "수정안", "reason": "판단근거", "severity": "HIGH|MEDIUM|LOW" }]
 */
export async function runStage2AIValidation({
    fileContent,
    masterData,
    apiKey,
    modelId = 'gemini-2.5-flash',
    onProgress = () => {}
}) {
    const { text } = fileContent;
    const { metadata, glossary } = masterData;

    if (!text || text.trim().length === 0) {
        return [];
    }

    // 3,000자 단위 청크 분할
    const CHUNK_SIZE = 3000;
    const chunks = [];
    let idx = 0;
    while (idx < text.length) {
        let end = idx + CHUNK_SIZE;
        if (end < text.length) {
            // 자연스러운 줄바꿈 또는 마침표 위치에서 분할
            const nextNL = text.lastIndexOf('\n', end);
            if (nextNL > idx + CHUNK_SIZE * 0.7) {
                end = nextNL + 1;
            }
        }
        chunks.push(text.substring(idx, end));
        idx = end;
    }

    const aiIssues = [];
    const totalChunks = chunks.length;

    // 공통 시스템 프롬프트 준비
    const glossaryHints = glossary.map(g => `- 표준: ${g.approved} (오표기/금지: ${g.forbidden})`).slice(0, 15).join('\n');
    const projectInfo = `[프로젝트 마스터 정보]\n- 사업명: ${metadata.projectName}\n- 발주기관: ${metadata.clientAgency}\n- 수행기간: ${metadata.period}\n- 총사업비: ${metadata.budget}\n\n[주요 표준용어]\n${glossaryHints}`;

    for (let i = 0; i < chunks.length; i++) {
        const chunkNum = i + 1;
        const chunkText = chunks[i];
        const progressPercent = Math.round((chunkNum / totalChunks) * 100);
        onProgress({
            current: chunkNum,
            total: totalChunks,
            percent: progressPercent,
            message: `Gemini AI 문맥 정밀 검증 중 (${chunkNum}/${totalChunks} 청크)...`
        });

        const prompt = `${projectInfo}

당신은 공공기관 정보화전략계획(ISP) 및 IT 공공제안서 전문 감리위원입니다.
제공된 아래 텍스트 청크를 정밀 검토하여 다음 결함을 찾아내십시오:
1. 맞춤법 및 오탈자 (의학·행정 전문용어 띄어쓰기 및 철자 오류)
2. 문맥 모순 (앞뒤 내용이 상충되거나 논리적으로 맞지 않는 문장)
3. 행정용어 부적절 (공공제안서/공공보고서에 어울리지 않는 구어체나 비표준 용어)
4. 수치 불일치 (단위 모순, 문맥 내 수치 불일치)

[검토할 텍스트 (청크 #${chunkNum})]:
${chunkText}

[출력 요구사항]:
반드시 아래 JSON 스키마를 만족하는 JSON Array 포맷으로만 응답하십시오.
응답에 백틱(\`\`\`json)이나 부가적인 설명 없이 오직 순수한 JSON 배열만 출력해야 합니다.
오류가 없으면 빈 배열 []을 반환하십시오.

[
  {
    "chunkIndex": ${chunkNum},
    "errorType": "맞춤법" | "문맥모순" | "행정용어부적절" | "수치불일치",
    "originalText": "원문에서 결함이 있는 정확한 문장 또는 어구 발췌",
    "suggestedText": "권장하는 수정안",
    "reason": "결함으로 판단한 명확한 근거",
    "severity": "HIGH" | "MEDIUM" | "LOW"
  }
]`;

        try {
            const resultItems = await callGeminiProxyWithJson({
                apiKey,
                modelId,
                prompt,
                chunkIndex: chunkNum
            });

            if (Array.isArray(resultItems)) {
                resultItems.forEach((item, itemIdx) => {
                    aiIssues.push({
                        id: `S2-AI-${chunkNum}-${itemIdx + 1}`,
                        stage: 'Stage 2 (Gemini AI 정밀검증)',
                        category: item.errorType || 'AI 문맥/오탈자',
                        errorType: item.errorType || 'AI 검출 결함',
                        severity: ['HIGH', 'MEDIUM', 'LOW'].includes(item.severity) ? item.severity : 'MEDIUM',
                        page: Math.max(1, Math.ceil((chunkNum / totalChunks) * (fileContent.pages || 1))),
                        matchedTerm: item.originalText || '',
                        suggestedText: item.suggestedText || '',
                        originalText: item.originalText || '',
                        reason: item.reason || 'AI 문맥 및 오탈자 정밀 검증 결과'
                    });
                });
            }
        } catch (chunkErr) {
            console.error(`청크 #${chunkNum} AI 검증 중 오류:`, chunkErr);
        }
    }

    return aiIssues;
}

// ── Gemini Proxy 호출 헬퍼 (/api/gemini-proxy) ──────────────────
async function callGeminiProxyWithJson({ apiKey, modelId, prompt, chunkIndex }) {
    const contents = [{ role: 'user', parts: [{ text: prompt }] }];
    const generationConfig = {
        temperature: 0.1,
        responseMimeType: 'application/json'
    };

    let responseData = null;

    // 1차: 서버 프록시 (/api/gemini-proxy) 시도
    try {
        const proxyRes = await fetch('/api/gemini-proxy', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                apiKey,
                modelId,
                contents,
                generationConfig
            })
        });

        if (proxyRes.ok) {
            responseData = await proxyRes.json();
        } else if (proxyRes.status !== 404) {
            const errText = await proxyRes.text();
            console.warn(`프록시 응답 에러 (${proxyRes.status}):`, errText);
        }
    } catch (_proxyErr) {
        // 프록시 미배포 환경 시 직접 호출 폴백
    }

    // 2차: 직접 엔드포인트 폴백
    if (!responseData && apiKey) {
        const directUrl = `https://generativelanguage.googleapis.com/v1beta/${modelId}:generateContent`;
        const directRes = await fetch(directUrl, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'x-goog-api-key': apiKey
            },
            body: JSON.stringify({
                contents,
                generationConfig
            })
        });

        if (!directRes.ok) {
            throw new Error(`Gemini API 호출 실패 (HTTP ${directRes.status})`);
        }
        responseData = await directRes.json();
    }

    if (!responseData) {
        return [];
    }

    const textResponse = responseData.candidates?.[0]?.content?.parts?.[0]?.text || '';
    if (!textResponse.trim()) return [];

    try {
        const parsed = JSON.parse(textResponse);
        return Array.isArray(parsed) ? parsed : [parsed];
    } catch (parseErr) {
        // 백틱 등 마크다운이 포함된 경우 정제
        const cleaned = textResponse
            .replace(/```json/gi, '')
            .replace(/```/g, '')
            .trim();
        try {
            const parsed = JSON.parse(cleaned);
            return Array.isArray(parsed) ? parsed : [parsed];
        } catch (_) {
            console.warn(`청크 #${chunkIndex} JSON 파싱 실패, 응답 원문:`, textResponse);
            return [];
        }
    }
}

// ── ExcelJS 기반 통합 정합성 검수 리포트 (.xlsx) 다운로드 ─────────
export async function exportSSOTReportToExcel({
    fileName,
    masterData,
    stage1Result,
    allIssues
}) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = '건강한 프로젝트 (SSOT Validator)';
    workbook.lastModifiedBy = 'SSOT 정합성 검증 엔진';
    workbook.created = new Date();
    workbook.modified = new Date();

    const { metadata } = masterData;
    const { rtmMatrix, rtmMappingRate } = stage1Result;

    // 헤더 공통 스타일
    const headerStyle = {
        font: { name: '맑은 고딕', size: 11, bold: true, color: { argb: 'FFFFFFFF' } },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } }, // Slate 800
        alignment: { vertical: 'middle', horizontal: 'center' }
    };

    const borderStyle = {
        top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        right: { style: 'thin', color: { argb: 'FFCBD5E1' } }
    };

    // ── 시트 1: 요약 및 총괄 보고 ────────────────────────────────
    const summarySheet = workbook.addWorksheet('요약 및 총괄 보고', {
        views: [{ showGridLines: true }]
    });

    summarySheet.columns = [
        { width: 22 },
        { width: 35 },
        { width: 22 },
        { width: 35 }
    ];

    summarySheet.addRow(['프로젝트 표준 정합성 검수 종합 결과 보고서', '', '', '']);
    summarySheet.mergeCells('A1:D1');
    const titleCell = summarySheet.getCell('A1');
    titleCell.font = { name: '맑은 고딕', size: 16, bold: true, color: { argb: 'FF1E293B' } };
    titleCell.alignment = { vertical: 'middle', horizontal: 'left' };
    summarySheet.getRow(1).height = 36;

    summarySheet.addRow(['검수 일시', new Date().toLocaleString('ko-KR'), '검수 대상 파일', fileName || '미지정']);
    summarySheet.addRow(['사업명', metadata.projectName || '-', '발주기관', metadata.clientAgency || '-']);
    summarySheet.addRow(['수행기간', metadata.period || '-', '총사업비', metadata.budget || '-']);

    for (let r = 2; r <= 4; r++) {
        summarySheet.getRow(r).height = 24;
        summarySheet.getCell(`A${r}`).font = { name: '맑은 고딕', bold: true };
        summarySheet.getCell(`C${r}`).font = { name: '맑은 고딕', bold: true };
        summarySheet.getCell(`A${r}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
        summarySheet.getCell(`C${r}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
    }

    summarySheet.addRow([]); // 빈 줄

    // 통계 요약 섹션
    summarySheet.addRow(['정합성 검증 핵심 지표 요약', '', '', '']);
    summarySheet.mergeCells('A6:D6');
    summarySheet.getCell('A6').font = { name: '맑은 고딕', size: 13, bold: true, color: { argb: 'FF0F172A' } };

    summarySheet.addRow(['지표명', '측정 결과', '판정 상태', '비고']);
    const statHeaderRow = summarySheet.getRow(7);
    statHeaderRow.height = 26;
    ['A7', 'B7', 'C7', 'D7'].forEach(c => {
        summarySheet.getCell(c).style = headerStyle;
    });

    const totalCount = allIssues.length;
    const missingReqCount = allIssues.filter(i => i.category === '요구사항 누락').length;
    const glossaryCount = allIssues.filter(i => i.category === '용어 불일치').length;
    const numErrorCount = allIssues.filter(i => i.category === '수치 오류').length;
    const aiCount = allIssues.filter(i => i.stage.includes('Stage 2')).length;

    const summaryStats = [
        ['총 결함 검출 건수', `${totalCount} 건`, totalCount === 0 ? '양호 (CLEAN)' : '조치 필요', '결정론적 룰 및 AI 검출 전수 합계'],
        ['요구사항 ID 매핑률 (RTM Coverage)', `${rtmMappingRate}%`, rtmMappingRate === 100 ? '완전 준수' : '누락 발생 (경고)', `마스터 ${masterData.requirements.length}개 중 ${stage1Result.coveredReqCount}개 발견`],
        ['요구사항 누락(Missing ID) 건수', `${missingReqCount} 건`, missingReqCount === 0 ? '적합' : '치명적 결함 (HIGH)', '제안요청서/과업지시서 필수요구사항'],
        ['표준용어 위반 (금지어/오표기)', `${glossaryCount} 건`, glossaryCount === 0 ? '적합' : '수정 권고', '표준 용어사전 대비 불일치'],
        ['수치 및 표 산술 오류', `${numErrorCount} 건`, numErrorCount === 0 ? '적합' : '오류 정정 필요', '총사업비 불일치 및 표 덧셈 오류'],
        ['Gemini AI 문맥·오탈자 검출', `${aiCount} 건`, aiCount === 0 ? '적합' : '검토 권장', '문맥 모순, 비표준 행정용어 등']
    ];

    summaryStats.forEach(stat => {
        const r = summarySheet.addRow(stat);
        r.height = 22;
        r.alignment = { vertical: 'middle' };
    });

    for (let r = 7; r <= 13; r++) {
        for (let col = 1; col <= 4; col++) {
            const cell = summarySheet.getRow(r).getCell(col);
            cell.border = borderStyle;
            if (col === 2 || col === 3) {
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
            }
        }
    }

    // ── 시트 2: 요구사항 추적 매트릭스(RTM) ──────────────────────────
    const rtmSheet = workbook.addWorksheet('요구사항 추적 매트릭스(RTM)', {
        views: [{ showGridLines: true }]
    });

    rtmSheet.columns = [
        { header: '순번', width: 8 },
        { header: '요구사항 ID', width: 16 },
        { header: '요구사항명', width: 45 },
        { header: '필수 키워드', width: 35 },
        { header: '본문 출현 여부', width: 15 },
        { header: '출현 위치(페이지)', width: 18 },
        { header: '정합성 상태', width: 20 }
    ];

    const rtmHead = rtmSheet.getRow(1);
    rtmHead.height = 28;
    for (let c = 1; c <= 7; c++) {
        rtmHead.getCell(c).style = headerStyle;
    }

    rtmMatrix.forEach((r, idx) => {
        const row = rtmSheet.addRow([
            idx + 1,
            r.id,
            r.name,
            r.keywords,
            r.isPresent ? '출현 (O)' : '누락 (X)',
            r.pages.length > 0 ? `p.${r.pages.join(', ')}` : '-',
            r.status
        ]);
        row.height = 24;
        row.alignment = { vertical: 'middle' };

        // 상태별 하이라이트
        const statusCell = row.getCell(7);
        const presentCell = row.getCell(5);

        if (!r.isPresent) {
            statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } }; // Red 100
            statusCell.font = { name: '맑은 고딕', color: { argb: 'FFDC2626' }, bold: true };
            presentCell.font = { name: '맑은 고딕', color: { argb: 'FFDC2626' }, bold: true };
        } else if (r.status.includes('누락')) {
            statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } }; // Yellow 100
            statusCell.font = { name: '맑은 고딕', color: { argb: 'FFD97706' }, bold: true };
        } else {
            statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } }; // Green 100
            statusCell.font = { name: '맑은 고딕', color: { argb: 'FF16A34A' }, bold: true };
        }

        for (let col = 1; col <= 7; col++) {
            row.getCell(col).border = borderStyle;
            if (col === 1 || col === 2 || col === 5 || col === 6 || col === 7) {
                row.getCell(col).alignment = { vertical: 'middle', horizontal: 'center' };
            }
        }
    });

    // ── 시트 3: 상세 오류 내역 ────────────────────────────────────
    const issueSheet = workbook.addWorksheet('상세 결함 내역', {
        views: [{ showGridLines: true }]
    });

    issueSheet.columns = [
        { header: 'No.', width: 8 },
        { header: '검증 단계', width: 22 },
        { header: '분류', width: 18 },
        { header: '오류 유형', width: 24 },
        { header: '심각도', width: 12 },
        { header: '위치', width: 12 },
        { header: '원문 발췌', width: 45 },
        { header: '수정 권고안', width: 35 },
        { header: '판단 근거 및 상세 사유', width: 50 }
    ];

    const issueHead = issueSheet.getRow(1);
    issueHead.height = 28;
    for (let c = 1; c <= 9; c++) {
        issueHead.getCell(c).style = headerStyle;
    }

    allIssues.forEach((issue, idx) => {
        const row = issueSheet.addRow([
            idx + 1,
            issue.stage || 'Stage 1',
            issue.category || '-',
            issue.errorType || '-',
            issue.severity || 'MEDIUM',
            issue.page ? `p.${issue.page}` : '-',
            issue.originalText || issue.matchedTerm || '-',
            issue.suggestedText || '-',
            issue.reason || '-'
        ]);
        row.height = 28;
        row.alignment = { vertical: 'middle', wrapText: true };

        const sevCell = row.getCell(5);
        if (issue.severity === 'HIGH') {
            sevCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
            sevCell.font = { name: '맑은 고딕', color: { argb: 'FFB91C1C' }, bold: true };
        } else if (issue.severity === 'MEDIUM') {
            sevCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
            sevCell.font = { name: '맑은 고딕', color: { argb: 'FFB45309' }, bold: true };
        } else {
            sevCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
            sevCell.font = { name: '맑은 고딕', color: { argb: 'FF475569' } };
        }

        for (let col = 1; col <= 9; col++) {
            row.getCell(col).border = borderStyle;
            if (col === 1 || col === 2 || col === 3 || col === 5 || col === 6) {
                row.getCell(col).alignment = { vertical: 'middle', horizontal: 'center' };
            }
        }
    });

    // 엑셀 파일 생성 및 다운로드
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const downloadName = `통합_정합성_검수_리포트_${(fileName || '산출물').replace(/\.[^/.]+$/, '')}_${new Date().toISOString().slice(0, 10)}.xlsx`;
    saveAs(blob, downloadName);
}

// ── 7. RFP(제안요청서) 파싱 및 마스터 정보 추출 엔진 ────────────────
/**
 * 룰 기반 결정론적 RFP 분석기 (초고속, 오프라인)
 */
export function extractRfpWithRules(text, tables = []) {
    let projectName = '';
    let clientAgency = '';
    let period = '';
    let budget = '';

    const cleanText = text || '';

    // 1) 사업명 탐색
    const nameMatch = cleanText.match(/(?:사업명|과업명|용역명|프로젝트명|입찰건명|건\s*명)\s*[:：\t|]\s*([^\n\r]+)/i)
        || cleanText.match(/(?:사업명|과업명|용역명|입찰건명)\s*\n\s*([^\n\r]+)/i);
    if (nameMatch) {
        projectName = nameMatch[1].replace(/^[\[<"']|[\]>"']$/g, '').trim();
    } else {
        const lines = cleanText.split(/\r?\n/).map(l => l.trim()).filter(Boolean).slice(0, 15);
        for (const l of lines) {
            if (l.includes('사업') || l.includes('플랫폼') || l.includes('시스템') || l.includes('ISP') || l.includes('ISMP') || l.includes('구축')) {
                if (l.length >= 8 && l.length <= 60 && !l.includes('목차') && !l.includes('개요')) {
                    projectName = l.replace(/^[\[<"']|[\]>"']$/g, '').trim();
                    break;
                }
            }
        }
    }

    // 2) 발주기관 탐색
    const agencyMatch = cleanText.match(/(?:발주기관|수요기관|주관기관|발주처|수요기관명|발주기관명)\s*[:：\t|]\s*([^\n\r]+)/i)
        || cleanText.match(/(?:발주기관|수요기관|주관기관|발주처)\s*\n\s*([^\n\r]+)/i);
    if (agencyMatch) {
        clientAgency = agencyMatch[1].replace(/^[\[<"']|[\]>"']$/g, '').trim();
    } else {
        const agencyRegex = /([가-힣]{2,20}(?:의료원|병원|지원센터|정보원|공단|진흥원|재단|보건복지부|질병관리청|과학기술정보통신부|행정안전부|청|부|처|원))/;
        const am = cleanText.slice(0, 3000).match(agencyRegex);
        if (am) {
            clientAgency = am[1].trim();
        }
    }

    // 3) 사업기간 탐색
    const periodMatch = cleanText.match(/(?:사업기간|과업기간|수행기간|용역기간|계약기간)\s*[:：\t|]\s*([^\n\r]+)/i)
        || cleanText.match(/(?:사업기간|과업기간|수행기간)\s*\n\s*([^\n\r]+)/i)
        || cleanText.match(/(?:계약(?:체결)?일로부터\s*[^.\n\r]+(?:개월|일간|일|이내))/i)
        || cleanText.match(/(?:\d{4}[.\-/]\d{1,2}\s*~\s*\d{4}[.\-/]\d{1,2}(?:\s*\([^)]+\))?)/);
    if (periodMatch) {
        period = (periodMatch[1] || periodMatch[0]).replace(/^[\[<"']|[\]>"']$/g, '').trim();
    }

    // 4) 총사업비 탐색
    const budgetMatch = cleanText.match(/(?:총\s*사업비|사업\s*예산|배정\s*예산|추정\s*가격|기초\s*금액|소요\s*예산|사업비|소요예산액)\s*[:：\t|]\s*([^\n\r]+)/i)
        || cleanText.match(/(?:사업예산|총\s*사업비)\s*\n\s*([^\n\r]+)/i)
        || cleanText.match(/(\d{1,3}(?:,\d{3})+(?:원|천원|백만원|억원)(?:\s*\([^)]+\))?)/);
    if (budgetMatch) {
        budget = (budgetMatch[1] || budgetMatch[0]).replace(/^[\[<"']|[\]>"']$/g, '').trim();
    }

    // 5) 요구사항 목록 추출 (표 우선 -> 텍스트 보완)
    const requirements = [];
    const seenReqIds = new Set();

    if (Array.isArray(tables)) {
        for (const tbl of tables) {
            const rows = tbl.rows || [];
            if (rows.length < 2) continue;

            const header = rows[0].map(c => (c || '').replace(/\s+/g, '').toLowerCase());
            let idCol = header.findIndex(h => h.includes('id') || h.includes('번호') || h.includes('식별자') || h.includes('코드'));
            let nameCol = header.findIndex(h => h.includes('명') || h.includes('요구사항명') || h.includes('명칭') || h.includes('제목'));
            let descCol = header.findIndex(h => h.includes('내용') || h.includes('설명') || h.includes('정의') || h.includes('키워드'));

            for (let r = 1; r < rows.length; r++) {
                const row = rows[r];
                let reqId = '';
                let reqName = '';
                let reqDesc = '';

                if (idCol >= 0 && row[idCol]) {
                    const m = row[idCol].match(/[A-Z]{2,5}-\d{2,4}/i);
                    if (m) reqId = m[0].toUpperCase();
                }

                if (!reqId) {
                    for (let c = 0; c < row.length; c++) {
                        const m = (row[c] || '').match(/^[A-Z]{2,5}-\d{2,4}$/i);
                        if (m) {
                            reqId = m[0].toUpperCase();
                            if (nameCol < 0 && row[c + 1]) reqName = row[c + 1];
                            break;
                        }
                    }
                }

                if (reqId && !seenReqIds.has(reqId)) {
                    if (!reqName && nameCol >= 0 && row[nameCol]) {
                        reqName = row[nameCol].trim();
                    }
                    if (descCol >= 0 && row[descCol]) {
                        reqDesc = row[descCol].trim();
                    }

                    let keywords = '';
                    if (reqDesc) {
                        const words = reqDesc
                            .replace(/[()[\]{}.,~;:_!?"'`/\\|]/g, ' ')
                            .split(/\s+/)
                            .map(w => w.trim())
                            .filter(w => w.length >= 2 && !['체계', '구축', '방안', '수립', '및', '관한', '위한', '지원', '제공', '등', '에', '의', '수행', '대상'].includes(w));
                        keywords = Array.from(new Set(words)).slice(0, 5).join(', ');
                    }

                    seenReqIds.add(reqId);
                    requirements.push({
                        id: reqId,
                        name: reqName || `요구사항 ${reqId}`,
                        keywords: keywords || reqName.split(/\s+/).filter(w => w.length >= 2).slice(0, 4).join(', ')
                    });
                }
            }
        }
    }

    // 텍스트 본문 정규식 보완
    const lineRegex = /(?:^|[•\s*-]+)([A-Z]{2,5}-\d{2,4})\s*[:：|/–—\-]\s*([^\n\r]+)/g;
    let lm;
    while ((lm = lineRegex.exec(cleanText)) !== null) {
        const id = lm[1].trim().toUpperCase();
        if (!seenReqIds.has(id)) {
            let fullStr = lm[2].trim();
            let keywords = '';
            const bracketMatch = fullStr.match(/\[([^\]]+)\]/);
            if (bracketMatch) {
                keywords = bracketMatch[1].trim();
                fullStr = fullStr.replace(/\[[^\]]+\]/, '').trim();
            }
            seenReqIds.add(id);
            requirements.push({
                id,
                name: fullStr,
                keywords: keywords || fullStr.split(/\s+/).filter(w => w.length >= 2).slice(0, 4).join(', ')
            });
        }
    }

    // 6) 표준 용어 추출 (RFP 내 빈출 주요 약어 및 공공 의료/행정 고유명사)
    const glossary = [];
    const seenTerms = new Set();

    const acronymMatches = cleanText.match(/\b[A-Z]{2,6}\b/g) || [];
    const acronymCounts = {};
    for (const a of acronymMatches) {
        if (!['RFP', 'HWP', 'PPT', 'DOC', 'PDF', 'URL', 'ZIP', 'JPG', 'PNG', 'VAT'].includes(a)) {
            acronymCounts[a] = (acronymCounts[a] || 0) + 1;
        }
    }

    const sortedAcronyms = Object.entries(acronymCounts)
        .filter(([_, count]) => count >= 2)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 8);

    for (const [acronym] of sortedAcronyms) {
        if (!seenTerms.has(acronym.toLowerCase())) {
            seenTerms.add(acronym.toLowerCase());
            glossary.push({
                id: `glo-rfp-${Date.now()}-${glossary.length + 1}`,
                category: '약어',
                approved: acronym,
                forbidden: acronym.toLowerCase(),
                note: `RFP 본문 다빈도 출현 표준 약어`
            });
        }
    }

    if (clientAgency && !seenTerms.has(clientAgency.toLowerCase())) {
        seenTerms.add(clientAgency.toLowerCase());
        let autoForbidden = '';
        if (clientAgency.includes('국립중앙의료원')) {
            autoForbidden = '중앙의료원, 국립의료원, NMC, 국립중앙병원';
        } else if (clientAgency.includes('보건복지부')) {
            autoForbidden = '복지부, 보건부';
        } else if (clientAgency.includes('정보원')) {
            autoForbidden = clientAgency.replace(/정보원$/, '센터');
        } else {
            autoForbidden = `${clientAgency}재단, ${clientAgency}센터`;
        }

        glossary.unshift({
            id: `glo-rfp-agency`,
            category: '기관명',
            approved: clientAgency,
            forbidden: autoForbidden,
            note: 'RFP 발주기관 공식 명칭'
        });
    }

    if (projectName) {
        const sysMatch = projectName.match(/([가-힣A-Za-z0-9]+(?:플랫폼|시스템|포털|고속도로|데이터셋|체계))/);
        if (sysMatch && !seenTerms.has(sysMatch[1].toLowerCase())) {
            seenTerms.add(sysMatch[1].toLowerCase());
            glossary.push({
                id: `glo-rfp-sys`,
                category: '시스템명',
                approved: sysMatch[1],
                forbidden: sysMatch[1].split('').join(' '), // 띄어쓰기 오표기 방지 (예: 건강정보고속도로 -> 건 강 정 보...)
                note: 'RFP 대상 핵심 시스템 명칭'
            });
        }
    }

    return {
        metadata: {
            projectName: projectName || '제안요청서 대상 사업',
            clientAgency: clientAgency || '발주기관',
            period: period || '계약체결일로부터 8개월',
            budget: budget || '사업예산액'
        },
        requirements,
        glossary
    };
}

/**
 * Gemini AI 기반 심층 RFP 분석기
 */
export async function extractRfpWithAI({
    text,
    tables = [],
    apiKey,
    modelId = 'gemini-2.5-flash',
    onProgress = () => {}
}) {
    onProgress({ percent: 20, message: 'RFP 핵심 요약 및 요구사항 표 데이터 선별 중...' });

    const textSnippet = (text || '').slice(0, 15000);

    const tableSnippets = (tables || []).slice(0, 6).map((t, idx) => {
        const rowsStr = (t.rows || []).slice(0, 25).map(r => r.join(' | ')).join('\n');
        return `[표 #${idx + 1} (${t.source || '표'})]:\n${rowsStr}`;
    }).join('\n\n');

    const prompt = `당신은 공공기관 정보화전략계획(ISP/ISMP) 및 IT 공공제안요청서(RFP) 분석 전문 수석 감리위원입니다.
제공된 제안요청서(RFP) 문서 원문과 표 데이터를 정밀 분석하여, 향후 제안서 및 산출물 정합성 검증의 기준이 될 [단일 진실 원천(SSOT) 마스터 정보]를 추출하십시오.

[제안요청서(RFP) 텍스트 발췌]:
${textSnippet}

${tableSnippets ? `[제안요청서 표 발췌]:\n${tableSnippets}\n` : ''}

[추출 지침]:
1. metadata:
   - projectName: 공식 사업명 (예: 2026년 공공보건의료 AI 기반 지능형 플랫폼 구축 ISP)
   - clientAgency: 발주기관/수요기관 공식 명칭 (예: 국립중앙의료원)
   - period: 사업기간 (예: 계약체결일로부터 8개월간 (2026.04 ~ 2026.11))
   - budget: 총사업비/배정예산 (숫자 및 한글 표기, 부가세 포함 여부)
2. requirements:
   - RFP 내에 명시된 요구사항 고유번호(ID: 예 CNR-001, SFR-001, PMR-001 등), 요구사항명, 필수 핵심키워드(쉼표 구분)
3. glossary:
   - RFP에서 정의하거나 자주 등장하는 핵심 기관명, 시스템명, 행정용어, 약어, 기술용어 8~15개
   - category: '기관명' | '시스템명' | '행정용어' | '약어' | '기술용어'
   - approved: 표준 공식 표기
   - forbidden: 제안서 작성자가 흔히 범하는 오표기, 띄어쓰기 오류, 금지동의어 (쉼표 구분)
   - note: 선정 사유 또는 RFP 출처

반드시 아래 JSON 스키마를 만족하는 순수 JSON 객체 포맷으로만 응답하십시오. 백틱(\`\`\`json) 없이 순수 JSON만 출력하십시오.
{
  "metadata": {
    "projectName": "사업명",
    "clientAgency": "발주기관명",
    "period": "사업기간",
    "budget": "총사업비"
  },
  "requirements": [
    { "id": "CNR-001", "name": "요구사항명", "keywords": "키워드1, 키워드2" }
  ],
  "glossary": [
    { "category": "기관명", "approved": "표준어", "forbidden": "오표기, 금지어", "note": "비고" }
  ]
}
`;

    onProgress({ percent: 50, message: 'Gemini AI로 마스터 메타데이터, 요구사항, 표준용어 정밀 추출 중...' });

    const aiRes = await callGeminiProxyWithJson({
        apiKey,
        modelId,
        prompt,
        chunkIndex: 0
    });

    onProgress({ percent: 90, message: '추출 데이터 구조화 및 정합성 검증 중...' });

    let parsed = null;
    if (aiRes) {
        if (Array.isArray(aiRes) && aiRes.length > 0 && aiRes[0]?.metadata) {
            parsed = aiRes[0];
        } else if (aiRes.metadata) {
            parsed = aiRes;
        }
    }

    return parsed;
}

/**
 * RFP 통합 분석 실행기 (규칙 기반 + AI 하이브리드 지원)
 */
export async function parseRfpContent({
    text,
    tables = [],
    apiKey = '',
    modelId = 'gemini-2.5-flash',
    useAI = false,
    onProgress = () => {}
}) {
    onProgress({ percent: 10, message: '결정론적 룰 기반 RFP 1차 파싱 중...' });
    const ruleResult = extractRfpWithRules(text, tables);

    if (!useAI || !apiKey) {
        onProgress({ percent: 100, message: 'RFP 룰 기반 분석 완료' });
        return {
            ...ruleResult,
            source: 'RULES',
            ruleReqCount: ruleResult.requirements.length,
            ruleGlossaryCount: ruleResult.glossary.length
        };
    }

    try {
        const aiResult = await extractRfpWithAI({
            text,
            tables,
            apiKey,
            modelId,
            onProgress
        });

        if (!aiResult) {
            return { ...ruleResult, source: 'RULES (AI 응답 없음)' };
        }

        // AI 결과와 룰 결과를 상호 보완 병합
        const mergedMetadata = {
            projectName: aiResult.metadata?.projectName || ruleResult.metadata.projectName,
            clientAgency: aiResult.metadata?.clientAgency || ruleResult.metadata.clientAgency,
            period: aiResult.metadata?.period || ruleResult.metadata.period,
            budget: aiResult.metadata?.budget || ruleResult.metadata.budget
        };

        // 요구사항: AI 결과 우선, 룰 기반 요구사항 중 누락된 ID 보완
        const reqMap = new Map();
        (ruleResult.requirements || []).forEach(r => reqMap.set(r.id.toUpperCase(), r));
        (aiResult.requirements || []).forEach(r => {
            if (r.id) {
                reqMap.set(r.id.toUpperCase(), {
                    id: r.id.toUpperCase(),
                    name: r.name || `요구사항 ${r.id}`,
                    keywords: r.keywords || ''
                });
            }
        });
        const mergedReqs = Array.from(reqMap.values());

        // 용어사전: AI 결과 + 룰 결과 병합
        const gloMap = new Map();
        (ruleResult.glossary || []).forEach(g => gloMap.set(g.approved.toLowerCase(), g));
        (aiResult.glossary || []).forEach((g, idx) => {
            if (g.approved) {
                const key = g.approved.toLowerCase();
                gloMap.set(key, {
                    id: `glo-ai-${Date.now()}-${idx + 1}`,
                    category: g.category || '행정용어',
                    approved: g.approved,
                    forbidden: g.forbidden || '',
                    note: g.note || 'Gemini AI 추출 표준용어'
                });
            }
        });
        const mergedGlossary = Array.from(gloMap.values());

        onProgress({ percent: 100, message: 'Gemini AI 심층 분석 및 마스터 추출 완료' });

        return {
            metadata: mergedMetadata,
            requirements: mergedReqs,
            glossary: mergedGlossary,
            source: 'HYBRID_AI',
            ruleReqCount: ruleResult.requirements.length,
            aiReqCount: (aiResult.requirements || []).length
        };
    } catch (err) {
        console.warn('AI RFP 분석 실패, 룰 결과로 폴백:', err);
        return {
            ...ruleResult,
            source: 'RULES (AI 폴백)',
            ruleReqCount: ruleResult.requirements.length
        };
    }
}

/**
 * 추출된 RFP 데이터를 프로젝트 SSOT 마스터 정보에 적용하는 헬퍼
 * mode: 'replace' (전체 대체) | 'merge' (기존 데이터에 병합)
 */
export function applyRfpToMasterData(currentMaster, rfpData, mode = 'replace') {
    const newMetadata = {
        projectName: rfpData.metadata?.projectName || currentMaster.metadata?.projectName || DEFAULT_PROJECT_METADATA.projectName,
        clientAgency: rfpData.metadata?.clientAgency || currentMaster.metadata?.clientAgency || DEFAULT_PROJECT_METADATA.clientAgency,
        period: rfpData.metadata?.period || currentMaster.metadata?.period || DEFAULT_PROJECT_METADATA.period,
        budget: rfpData.metadata?.budget || currentMaster.metadata?.budget || DEFAULT_PROJECT_METADATA.budget
    };

    let newRequirements = [];
    if (mode === 'replace') {
        newRequirements = Array.isArray(rfpData.requirements) && rfpData.requirements.length > 0
            ? rfpData.requirements
            : (currentMaster.requirements || DEFAULT_REQUIREMENTS);
    } else {
        const reqMap = new Map();
        (currentMaster.requirements || []).forEach(r => reqMap.set(r.id.toUpperCase(), { ...r }));
        (rfpData.requirements || []).forEach(r => {
            if (r.id) {
                const key = r.id.toUpperCase();
                reqMap.set(key, { ...(reqMap.get(key) || {}), ...r });
            }
        });
        newRequirements = Array.from(reqMap.values());
    }

    let newGlossary = [];
    if (mode === 'replace') {
        newGlossary = Array.isArray(rfpData.glossary) && rfpData.glossary.length > 0
            ? rfpData.glossary
            : (currentMaster.glossary || DEFAULT_GLOSSARY);
    } else {
        const gloMap = new Map();
        (currentMaster.glossary || []).forEach(g => gloMap.set(g.approved.toLowerCase(), { ...g }));
        (rfpData.glossary || []).forEach((g, idx) => {
            if (g.approved) {
                const key = g.approved.toLowerCase();
                const existing = gloMap.get(key);
                gloMap.set(key, {
                    id: existing?.id || `glo-${Date.now()}-${idx}`,
                    category: g.category || existing?.category || '행정용어',
                    approved: g.approved,
                    forbidden: g.forbidden || existing?.forbidden || '',
                    note: g.note || existing?.note || ''
                });
            }
        });
        newGlossary = Array.from(gloMap.values());
    }

    const updated = {
        metadata: newMetadata,
        requirements: newRequirements,
        glossary: newGlossary
    };

    saveSSOTData(updated);
    return updated;
}
