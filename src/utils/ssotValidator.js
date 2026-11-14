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

// ── 공공보건의료 AI ISP 공식 44개 마스터 요구사항 표준 프리셋 ────────────
export const PUBLIC_HEALTHCARE_AI_ISP_REQUIREMENTS = [
    {
        id: 'CNR-001',
        name: '공공의료 정보화 환경 및 법·제도 분석',
        keywords: '공공의료, 정보화, 환경, 법·제도'
    },
    {
        id: 'CNR-002',
        name: '보건의료 AI 활용 동향 및 적용 가능성 분석',
        keywords: '보건의료, AI, 활용, 동향'
    },
    {
        id: 'CNR-003',
        name: '공공의료 정보시스템 업무/정보기술(IT) 현황 분석',
        keywords: '시스템, 구성, 기술, 환경, 분석'
    },
    {
        id: 'CNR-004',
        name: '선진사례 벤치마킹 및 차이 분석',
        keywords: '선진사례, 벤치마킹, 차이, 분석'
    },
    {
        id: 'CNR-005',
        name: '과업 분야별 업무 개선방향 및 개선과제 도출',
        keywords: '과업, 분야별, 업무, 개선방향'
    },
    {
        id: 'CNR-006',
        name: '공공의료 AI 정보화 비전 및 전략 수립',
        keywords: '공공의료, AI, 정보화, 비전'
    },
    {
        id: 'CNR-007',
        name: 'AI 기반 공공의료 정보시스템 목표모델(To-Be) 설계',
        keywords: 'AI, 기반, 공공의료, 정보시스템'
    },
    {
        id: 'CNR-008',
        name: '차세대 정보시스템/데이터 플랫폼 구축 전략 및 기술 아키텍쳐 설계',
        keywords: '차세대, 정보시스템/데이터, 플랫폼, 구축'
    },
    {
        id: 'CNR-009',
        name: '공공의료 AI 거버넌스 및 데이터 관리체계 설계',
        keywords: '공공의료, AI, 거버넌스, 데이터'
    },
    {
        id: 'CNR-010',
        name: '공공병원 AI 서비스 확산 및 스마트병원 활성화 추진 전략',
        keywords: '공공병원, AI, 서비스, 확산'
    },
    {
        id: 'CNR-011',
        name: '재난·재해 대응, 개인정보보호 및 정보보안 강화 전략 수립',
        keywords: '재난·재해, 대응,, 개인정보보호, 정보보안'
    },
    {
        id: 'CNR-012',
        name: '공공병원 AX 단계별 추진 로드맵 수립',
        keywords: '공공병원, AX, 단계별, 추진'
    },
    {
        id: 'CNR-013',
        name: 'AI 기반 차세대 시스템 구축에 따른 기대효과 및 성과지표 도출',
        keywords: 'AI, 기반, 차세대, 시스템'
    },
    {
        id: 'CNR-014',
        name: '사업계획 수립 및 소요예산 산출, 개선과제 이행방안 수립',
        keywords: '사업계획, 수립, 소요예산, 산출,'
    },
    {
        id: 'CNR-015',
        name: '세부 사업계획서 및 제안요청서 작성',
        keywords: '세부, 사업계획서, 제안요청서, 작성'
    },
    {
        id: 'CNR-016',
        name: '공통기반 시스템 연계·활용 방안 검토',
        keywords: '공통기반, 시스템, 연계·활용, 방안'
    },
    {
        id: 'CNR-017',
        name: 'OpenAPI 기능의 활용성 확장 및 타 시스템 간 연계',
        keywords: 'OpenAPI, 기능의, 활용성, 확장'
    },
    {
        id: 'DAR-001',
        name: '데이터 표준화 방안 제시',
        keywords: '데이터, 표준화, 방안, 제시'
    },
    {
        id: 'DAR-002',
        name: '데이터 표준관리 방안 제시',
        keywords: '데이터, 표준관리, 방안, 제시'
    },
    {
        id: 'DAR-003',
        name: '데이터 구조 설계 방안 제시',
        keywords: '데이터, 구조, 설계, 방안'
    },
    {
        id: 'DAR-004',
        name: '데이터 구조 검증 및 개선방안 제시',
        keywords: '데이터, 구조, 검증, 개선방안'
    },
    {
        id: 'DAR-005',
        name: '데이터 구조 관리 방안 제시',
        keywords: '데이터, 구조, 관리, 방안'
    },
    {
        id: 'DAR-006',
        name: '데이터 값 검증 방안 제시',
        keywords: '데이터, 검증, 방안, 제시'
    },
    {
        id: 'DAR-007',
        name: '이관 데이터 값 검증 방안 제시',
        keywords: '이관, 데이터, 검증, 방안'
    },
    {
        id: 'DAR-008',
        name: '데이터 관리체계 현황분석 및 목표모델 제시',
        keywords: '데이터, 관리체계, 현황분석, 목표모델'
    },
    {
        id: 'DAR-009',
        name: '데이터 개방 및 메타데이터 관리체계',
        keywords: '데이터, 개방, 메타데이터, 관리체계'
    },
    {
        id: 'DAR-010',
        name: '보건의료데이터 표준화',
        keywords: '보건의료데이터, 표준화'
    },
    {
        id: 'SER-001',
        name: '보안관리 준수사항',
        keywords: '보안관리, 준수사항'
    },
    {
        id: 'SER-002',
        name: '누출금지 대상정보',
        keywords: '누출금지, 대상정보'
    },
    {
        id: 'SER-003',
        name: '사업자 보안위규 처리기준',
        keywords: '사업자, 보안위규, 처리기준'
    },
    {
        id: 'SER-004',
        name: '사업자 보안위약금 부과기준',
        keywords: '사업자, 보안위약금, 부과기준'
    },
    {
        id: 'QUR-001',
        name: '품질관리 준수',
        keywords: '품질관리, 준수'
    },
    {
        id: 'COR-001',
        name: '정보화기반 표준화 지침 및 가이드라인 준수',
        keywords: '정보화기반, 표준화, 지침, 가이드라인'
    },
    {
        id: 'COR-002',
        name: '저작권 및 지식재산권',
        keywords: '저작권, 지식재산권'
    },
    {
        id: 'COR-003',
        name: '정보시스템 운영 성과측정 기능활용도 측정 방안 마련',
        keywords: '정보시스템, 운영, 성과측정, 기능활용도'
    },
    {
        id: 'PMR-001',
        name: '업무보고',
        keywords: '업무보고'
    },
    {
        id: 'PMR-002',
        name: '프로젝트 일정관리 방안 제시',
        keywords: '프로젝트, 일정관리, 방안, 제시'
    },
    {
        id: 'PMR-003',
        name: '사업수행 산출물 작성 및 제출',
        keywords: '사업수행, 산출물, 작성, 제출'
    },
    {
        id: 'PMR-004',
        name: '프로젝트 조직구성(인력관리 등)',
        keywords: '프로젝트, 조직구성(인력관리, 등)'
    },
    {
        id: 'PSR-001',
        name: '관련 기술 이전 및 지원',
        keywords: '관련, 기술, 이전, 지원'
    },
    {
        id: 'PSR-002',
        name: '전산장비 및 각종 사무기기 확보',
        keywords: '전산장비, 각종, 사무기기, 확보'
    },
    {
        id: 'PSR-003',
        name: '하자보수 일반',
        keywords: '하자보수, 일반'
    },
    {
        id: 'PSR-004',
        name: '행사(보고회) 및 워크숍 지원',
        keywords: '행사(보고회), 워크숍, 지원'
    },
    {
        id: 'PSR-005',
        name: '추가제안',
        keywords: '추가제안'
    },
];

export const DEFAULT_REQUIREMENTS = PUBLIC_HEALTHCARE_AI_ISP_REQUIREMENTS;

// ── LocalStorage 로드 / 저장 유틸 ──────────────────────────────
export function loadSSOTData() {
    try {
        const raw = localStorage.getItem(SSOT_STORAGE_KEY);
        if (raw) {
            const parsed = JSON.parse(raw);
            let reqs = Array.isArray(parsed.requirements) && parsed.requirements.length > 0
                ? parsed.requirements
                : DEFAULT_REQUIREMENTS;

            // 1) 구버전 8개 샘플만 남아있는 경우 (SFR-001 등이 포함된 초기 데모 데이터) -> 공식 44개 마스터로 자동 갱신
            const isLegacy8Preset = reqs.length === 8 && reqs.some(r => r.id === 'SFR-001');
            if (isLegacy8Preset) {
                reqs = [...PUBLIC_HEALTHCARE_AI_ISP_REQUIREMENTS];
            } else {
                // 2) 만약 요구사항명이 '요구사항 CNR-xxx' 등 플레이스홀더이거나 비어있는 경우, 공식 카탈로그에서 명칭 자동 복원
                reqs = reqs.map(r => {
                    const cleanId = (r.id || '').trim().toUpperCase();
                    const cleanName = (r.name || '').trim();
                    const isDummy = !cleanName || cleanName === `요구사항 ${cleanId}` || cleanName === cleanId || /^요구사항\s*[A-Z]{2,6}[-_]\d{2,4}$/i.test(cleanName);
                    if (isDummy) {
                        const catalogItem = PUBLIC_HEALTHCARE_AI_ISP_REQUIREMENTS.find(c => c.id === cleanId);
                        if (catalogItem) {
                            return {
                                ...r,
                                name: catalogItem.name,
                                keywords: r.keywords || catalogItem.keywords
                            };
                        }
                    }
                    return r;
                });
            }

            return {
                metadata: { ...DEFAULT_PROJECT_METADATA, ...(parsed.metadata || {}) },
                glossary: Array.isArray(parsed.glossary) && parsed.glossary.length > 0 ? parsed.glossary : DEFAULT_GLOSSARY,
                requirements: reqs
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
    const { text, pageMap, pageLabelMap = {} } = fileContent;

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
    const idLabelLocations = {}; // { [id]: string[] } 슬라이드 하단 표시 라벨(예: Ⅱ-15), 없으면 페이지 숫자
    const addLocation = (id, pageNum) => {
        const p = parseInt(pageNum, 10) || 1;
        if (!idLocations[id]) idLocations[id] = [];
        if (idLocations[id].includes(p)) return;
        idLocations[id].push(p);
        if (!idLabelLocations[id]) idLabelLocations[id] = [];
        idLabelLocations[id].push(pageLabelMap[p] || String(p));
    };

    requirements.forEach(req => {
        const cleanId = (req.id || '').trim();
        if (!cleanId) return;
        const upperId = cleanId.toUpperCase();

        Object.entries(pageMap || {}).forEach(([pageNum, lines]) => {
            const pageText = (Array.isArray(lines) ? lines.join(' ') : String(lines)).toUpperCase();
            if (pageText.includes(upperId)) {
                matchedIdsInDoc.add(cleanId);
                addLocation(cleanId, pageNum);
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
                addLocation(detectedId, pageNum);
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
        const pageLabels = idLabelLocations[cleanId] || idLabelLocations[cleanId.toUpperCase()] || [];

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
            pageLabels,
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
            r.pages.length > 0 ? formatPageLabelRanges(r.pageLabels, r.pages) : '-',
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

/**
 * 6-2. 다중 파일 일괄 검수 종합 엑셀 리포트 내보내기 (ExcelJS 활용)
 */

/**
 * 다중 산출물 일괄 검수 결과 프로젝트 레벨 통합 집계 엔진
 * - 개별 파일 단위 가짜 Missing ID 오탐 제거
 * - 전체 44개 마스터 요구사항 대비 프로젝트 종합 RTM 충족률 계산 (어느 파일에도 없는 것만 '전수 누락' 지적)
 * - 산출물 간 정합성(Cross-Document Consistency) 분석 (용어 표기 불일치, 유령 ID 확산)
 */
export function aggregateBatchSSOT(batchResults = [], ssotData = {}) {
    const { requirements = [], glossary = [] } = ssotData;

    // 1) 전체 산출물 대상 요구사항 매핑 현황 집계
    const rtmMap = new Map();
    requirements.forEach(req => {
        rtmMap.set(req.id.toUpperCase(), {
            ...req,
            appearingFiles: [],
            pagesByFile: {},
            pageLabelsByFile: {},
            anyKeywordMissing: false
        });
    });

    batchResults.forEach(res => {
        const matrix = res.stage1?.rtmMatrix || [];
        matrix.forEach(m => {
            if (m.isPresent) {
                const item = rtmMap.get(m.id.toUpperCase());
                if (item) {
                    item.appearingFiles.push(res.fileName);
                    item.pagesByFile[res.fileName] = m.pages || [];
                    item.pageLabelsByFile[res.fileName] = m.pageLabels || [];
                    if (m.status && m.status.includes('누락')) {
                        item.anyKeywordMissing = true;
                    }
                }
            }
        });
    });

    const projectRtmMatrix = Array.from(rtmMap.values()).map(r => {
        const isPresentOverall = r.appearingFiles.length > 0;
        let overallStatus = '미출현 (전체 파일 누락)';
        if (isPresentOverall) {
            overallStatus = r.anyKeywordMissing ? '키워드 일부 누락' : '정상 반영';
        }
        return {
            ...r,
            isPresentOverall,
            overallStatus
        };
    });

    const coveredReqs = projectRtmMatrix.filter(r => r.isPresentOverall);
    const missingReqs = projectRtmMatrix.filter(r => !r.isPresentOverall);
    const projectRtmRate = requirements.length > 0 ? Math.round((coveredReqs.length / requirements.length) * 100) : 0;

    // 2) 프로젝트 전체 전수 누락 결함 (개별 파일마다 44개씩 발생하던 중복 오탐 제거 -> 전수 미출현 건만 1회 표기)
    const projectMissingIssues = missingReqs.map((req, idx) => ({
        id: `PROJ-MIS-${req.id}`,
        stage: '프로젝트 통합 RTM',
        category: '요구사항 누락',
        errorType: '전체 산출물 미반영 (Missing ID)',
        severity: 'HIGH',
        page: '-',
        fileName: '전체 산출물 전수 누락',
        matchedTerm: `[${req.id}] ${req.name}`,
        originalText: `(프로젝트 전체 ${batchResults.length}개 산출물 어디에서도 [${req.id}] 미발견)`,
        suggestedText: `요구사항 [${req.id}] ${req.name} 관련 산출물 본문 반영 필요`,
        reason: `RFP 마스터 요구사항 [${req.id}]이(가) 등록된 ${batchResults.length}개 산출물 중 단 한 곳에도 작성되지 않았습니다.`
    }));

    // 3) 산출물별 실제 결함 (개별 파일의 Missing ID 결함은 제외하고, 실제 용어 오표기/수치 오류/유령 ID만 수록)
    const fileActualIssues = [];
    batchResults.forEach(res => {
        const issues = res.allIssues || [];
        issues.forEach(iss => {
            if (iss.errorType === '요구사항 미출현(Missing ID)' || iss.id?.startsWith('S1-RTM-MIS-')) {
                return; // 개별 파일 레벨의 Missing ID 오탐은 제외!
            }
            fileActualIssues.push({
                ...iss,
                fileName: res.fileName
            });
        });
    });

    // 4) 산출물 간 용어 정합성 분석 (어떤 파일은 표준어, 어떤 파일은 금지어를 교차 사용하여 비정합이 발생했는지)
    const termConsistency = [];
    glossary.forEach(g => {
        const approved = (g.approved || '').trim();
        const forbiddens = (g.forbidden || '').split(',').map(f => f.trim()).filter(Boolean);
        if (!approved) return;

        const filesUsingForbidden = [];
        const filesUsingApproved = [];

        batchResults.forEach(res => {
            const text = res.stage1?.fullText || '';
            const issues = res.stage1?.glossaryIssues || [];

            const hasForbidden = issues.some(iss => iss.suggestedText === approved || (iss.matchedTerm && forbiddens.includes(iss.matchedTerm)));
            if (hasForbidden) {
                filesUsingForbidden.push(res.fileName);
            }
            if (text.includes(approved)) {
                filesUsingApproved.push(res.fileName);
            }
        });

        if (filesUsingForbidden.length > 0) {
            termConsistency.push({
                category: g.category || '행정용어',
                approved,
                forbidden: g.forbidden,
                filesUsingForbidden,
                filesUsingApproved,
                isInconsistent: filesUsingForbidden.length > 0 && filesUsingApproved.length > 0
            });
        }
    });

    // 5) 산출물 간 비표준 유령 ID 확산도 분석
    const ghostMap = new Map();
    batchResults.forEach(res => {
        const issues = res.stage1?.ghostIdIssues || [];
        issues.forEach(iss => {
            const gid = iss.detectedId || iss.matchedTerm;
            if (!gid) return;
            if (!ghostMap.has(gid)) {
                ghostMap.set(gid, {
                    ghostId: gid,
                    files: new Set(),
                    totalCount: 0
                });
            }
            const gItem = ghostMap.get(gid);
            gItem.files.add(res.fileName);
            gItem.totalCount++;
        });
    });

    const ghostIdSpread = Array.from(ghostMap.values()).map(g => ({
        ghostId: g.ghostId,
        fileCount: g.files.size,
        files: Array.from(g.files),
        totalCount: g.totalCount
    })).sort((a, b) => b.fileCount - a.fileCount);

    return {
        projectRtmMatrix,
        coveredReqCount: coveredReqs.length,
        missingReqCount: missingReqs.length,
        totalMasterReqCount: requirements.length,
        projectRtmRate,
        projectMissingIssues,
        fileActualIssues,
        termConsistency,
        ghostIdSpread
    };
}

export async function exportBatchSSOTReportToExcel({
    batchResults = [],
    ssotData = {}
}) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = '건강한 프로젝트 AI 산출물 검수기 (SSOT Batch)';
    workbook.created = new Date();

    const { metadata = {}, requirements = [] } = ssotData;

    // ── 프로젝트 레벨 통합 집계 (개별 파일 가짜 Missing ID 오탐 제거 및 정합성 분석) ──
    const agg = aggregateBatchSSOT(batchResults, ssotData);

    const headerStyle = {
        font: { name: '맑은 고딕', size: 11, bold: true, color: { argb: 'FFFFFFFF' } },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } }, // Slate 800
        alignment: { vertical: 'middle', horizontal: 'center' }
    };

    const subHeaderStyle = {
        font: { name: '맑은 고딕', size: 10.5, bold: true, color: { argb: 'FFFFFFFF' } },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF334155' } }, // Slate 700
        alignment: { vertical: 'middle', horizontal: 'center' }
    };

    const borderStyle = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
    };

    // ── 시트 1: 다중 산출물 검수 종합 요약 ──────────────────────────
    const summarySheet = workbook.addWorksheet('다중산출물 검수 종합 요약', {
        views: [{ showGridLines: true }]
    });

    summarySheet.columns = [
        { width: 6 },   // No.
        { width: 42 },  // 파일명
        { width: 14 },  // 검수 상태
        { width: 18 },  // RTM 출현 건수
        { width: 14 },  // 실제 결함 수
        { width: 14 },  // 치명(HIGH)
        { width: 14 },  // 주의(MEDIUM)
        { width: 14 }   // 권고(LOW)
    ];

    // 타이틀
    summarySheet.mergeCells('A1:H1');
    const titleCell = summarySheet.getCell('A1');
    titleCell.value = `[SSOT] 다중 산출물 정합성·용어 일괄 검수 종합 리포트`;
    titleCell.font = { name: '맑은 고딕', size: 16, bold: true, color: { argb: 'FF0F172A' } };
    titleCell.alignment = { vertical: 'middle' };
    summarySheet.getRow(1).height = 36;

    // 프로젝트 메타정보 블록
    summarySheet.mergeCells('A3:D3');
    summarySheet.getCell('A3').value = `■ 사업명: ${metadata.projectName || '-'}`;
    summarySheet.getCell('A3').font = { name: '맑은 고딕', size: 11, bold: true };

    summarySheet.mergeCells('E3:H3');
    summarySheet.getCell('E3').value = `■ 발주기관: ${metadata.clientAgency || '-'}`;
    summarySheet.getCell('E3').font = { name: '맑은 고딕', size: 11, bold: true };

    summarySheet.mergeCells('A4:D4');
    summarySheet.getCell('A4').value = `■ 총사업비: ${metadata.budget || '-'} / 수행기간: ${metadata.period || '-'}`;
    summarySheet.getCell('A4').font = { name: '맑은 고딕', size: 10, color: { argb: 'FF475569' } };

    summarySheet.mergeCells('E4:H4');
    const validCount = batchResults.filter(r => r.status === 'DONE').length;
    summarySheet.getCell('E4').value = `■ 검수 일시: ${new Date().toLocaleString()} (총 ${batchResults.length}건 중 완료 ${validCount}건)`;
    summarySheet.getCell('E4').font = { name: '맑은 고딕', size: 10, color: { argb: 'FF475569' } };

    // 프로젝트 전체 RTM 및 산출물 간 정합성 핵심 지표 요약 배너 (Row 6)
    summarySheet.mergeCells('A6:H6');
    const kpiCell = summarySheet.getCell('A6');
    kpiCell.value = `🎯 [프로젝트 전체 정합성 종합 진단] 총 요구사항 ${agg.totalMasterReqCount}건 중 충족 ${agg.coveredReqCount}건 (${agg.projectRtmRate}%) | 전수 누락: ${agg.missingReqCount}건 | 산출물 간 비정합 용어: ${agg.termConsistency.filter(t => t.isInconsistent).length}건 | 비표준 유령ID: ${agg.ghostIdSpread.length}건`;
    kpiCell.font = { name: '맑은 고딕', size: 10.5, bold: true, color: { argb: 'FF1E3A8A' } };
    kpiCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFF6FF' } };
    kpiCell.alignment = { vertical: 'middle', horizontal: 'center' };
    kpiCell.border = borderStyle;
    summarySheet.getRow(6).height = 30;

    // 파일별 요약 테이블 헤더 (Row 8)
    const tableHeaderRow = summarySheet.getRow(8);
    tableHeaderRow.height = 28;
    const headers = ['No.', '산출물 파일명', '검수 상태', '해당파일 RTM 출현', '실제 결함 수', '치명 (HIGH)', '주의 (MEDIUM)', '경미 (LOW)'];
    headers.forEach((h, idx) => {
        const cell = tableHeaderRow.getCell(idx + 1);
        cell.value = h;
        cell.style = headerStyle;
    });

    let totalActualIssuesAll = 0;
    let totalHighAll = 0;
    let totalMediumAll = 0;
    let totalLowAll = 0;

    batchResults.forEach((res, idx) => {
        const rowNum = 9 + idx;
        const row = summarySheet.getRow(rowNum);
        row.height = 26;

        // 개별 파일에서는 가짜 Missing ID 결함 제외한 실제 결함만 카운트!
        const fileActualIssues = (res.allIssues || []).filter(i => i.errorType !== '요구사항 미출현(Missing ID)' && !i.id?.startsWith('S1-RTM-MIS-'));
        const highCount = fileActualIssues.filter(i => i.severity === 'HIGH').length;
        const medCount = fileActualIssues.filter(i => i.severity === 'MEDIUM').length;
        const lowCount = fileActualIssues.filter(i => i.severity === 'LOW').length;

        totalActualIssuesAll += fileActualIssues.length;
        totalHighAll += highCount;
        totalMediumAll += medCount;
        totalLowAll += lowCount;

        const coveredInThisFile = (res.stage1?.rtmMatrix || []).filter(m => m.isPresent).length;
        const rtmStr = `${coveredInThisFile}건 반영 (${res.stage1?.coveredReqCount || coveredInThisFile}/${agg.totalMasterReqCount})`;

        row.getCell(1).value = idx + 1;
        row.getCell(2).value = res.fileName || `파일 #${idx + 1}`;
        row.getCell(3).value = res.status === 'DONE' ? '검수 완료' : (res.status === 'ERROR' ? `실패: ${res.errorMsg || '오류'}` : '대기');
        row.getCell(4).value = rtmStr;
        row.getCell(5).value = fileActualIssues.length;
        row.getCell(6).value = highCount;
        row.getCell(7).value = medCount;
        row.getCell(8).value = lowCount;

        for (let c = 1; c <= 8; c++) {
            row.getCell(c).border = borderStyle;
            if (c === 1 || c === 3 || c === 4) {
                row.getCell(c).alignment = { vertical: 'middle', horizontal: 'center' };
            } else if (c >= 5) {
                row.getCell(c).alignment = { vertical: 'middle', horizontal: 'right' };
            } else {
                row.getCell(c).alignment = { vertical: 'middle', horizontal: 'left' };
            }
        }

        if (highCount > 0) {
            row.getCell(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
            row.getCell(6).font = { name: '맑은 고딕', color: { argb: 'FFDC2626' }, bold: true };
        }
    });

    // 합계 Row
    const sumRowNum = 9 + batchResults.length;
    const sumRow = summarySheet.getRow(sumRowNum);
    sumRow.height = 28;
    sumRow.getCell(1).value = '';
    sumRow.getCell(2).value = '총 합 계 (개별 산출물 실제 결함 누적)';
    sumRow.getCell(2).font = { name: '맑은 고딕', bold: true };
    sumRow.getCell(3).value = `${batchResults.length}개 파일`;
    sumRow.getCell(4).value = `${agg.coveredReqCount}/${agg.totalMasterReqCount}건 충족`;
    sumRow.getCell(5).value = totalActualIssuesAll;
    sumRow.getCell(6).value = totalHighAll;
    sumRow.getCell(7).value = totalMediumAll;
    sumRow.getCell(8).value = totalLowAll;

    for (let c = 1; c <= 8; c++) {
        const cell = sumRow.getCell(c);
        cell.border = borderStyle;
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
        cell.font = { name: '맑은 고딕', bold: true };
        if (c === 2 || c === 3 || c === 4) {
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
        } else if (c >= 5) {
            cell.alignment = { vertical: 'middle', horizontal: 'right' };
        }
    }

    // ── 시트 2: 통합 상세 결함 목록 (오탐 완전 제거) ───────────────────
    const issueSheet = workbook.addWorksheet('통합 상세 결함 목록', {
        views: [{ showGridLines: true }]
    });

    issueSheet.columns = [
        { header: 'No.', width: 8 },
        { header: '대상 산출물', width: 35 },
        { header: '검증 단계', width: 20 },
        { header: '분류', width: 16 },
        { header: '오류 유형', width: 24 },
        { header: '심각도', width: 12 },
        { header: '위치', width: 12 },
        { header: '원문 발췌', width: 45 },
        { header: '수정 권고안', width: 35 },
        { header: '판단 근거 및 상세 사유', width: 50 }
    ];

    const issueHead = issueSheet.getRow(1);
    issueHead.height = 28;
    for (let c = 1; c <= 10; c++) {
        issueHead.getCell(c).style = headerStyle;
    }

    let globalIssueIdx = 1;

    // 1단계: 프로젝트 전체 전수 누락 결함 (어느 파일에도 전혀 없는 요구사항만 1회 표기)
    agg.projectMissingIssues.forEach(issue => {
        const row = issueSheet.addRow([
            globalIssueIdx++,
            issue.fileName,
            issue.stage,
            issue.category,
            issue.errorType,
            issue.severity,
            issue.page,
            issue.originalText,
            issue.suggestedText,
            issue.reason
        ]);
        row.height = 28;
        row.alignment = { vertical: 'middle', wrapText: true };

        const sevCell = row.getCell(6);
        sevCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
        sevCell.font = { name: '맑은 고딕', color: { argb: 'FFB91C1C' }, bold: true };

        const targetCell = row.getCell(2);
        targetCell.font = { name: '맑은 고딕', color: { argb: 'FFDC2626' }, bold: true };

        for (let col = 1; col <= 10; col++) {
            row.getCell(col).border = borderStyle;
            if (col === 1 || col === 3 || col === 4 || col === 6 || col === 7) {
                row.getCell(col).alignment = { vertical: 'middle', horizontal: 'center' };
            }
        }
    });

    // 2단계: 각 산출물별 실제 결함 (Missing ID 제외한 용어, 수치, 유령ID 등 실제 결함)
    agg.fileActualIssues.forEach(issue => {
        const row = issueSheet.addRow([
            globalIssueIdx++,
            issue.fileName,
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

        const sevCell = row.getCell(6);
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

        for (let col = 1; col <= 10; col++) {
            row.getCell(col).border = borderStyle;
            if (col === 1 || col === 3 || col === 4 || col === 6 || col === 7) {
                row.getCell(col).alignment = { vertical: 'middle', horizontal: 'center' };
            }
        }
    });

    // ── 시트 3: 통합 요구사항 추적 매트릭스(RTM) ───────────────────
    const rtmSheet = workbook.addWorksheet('통합 요구사항 추적(RTM)', {
        views: [{ showGridLines: true }]
    });

    rtmSheet.columns = [
        { header: '순번', width: 8 },
        { header: '요구사항 ID', width: 16 },
        { header: '요구사항명', width: 42 },
        { header: '필수 키워드', width: 32 },
        { header: '출현 파일 수', width: 14 },
        { header: '출현 파일 및 위치 목록', width: 55 },
        { header: '종합 정합성 상태', width: 22 }
    ];

    const rtmHead = rtmSheet.getRow(1);
    rtmHead.height = 28;
    for (let c = 1; c <= 7; c++) {
        rtmHead.getCell(c).style = headerStyle;
    }

    agg.projectRtmMatrix.forEach((r, idx) => {
        const fileLocs = r.appearingFiles.map(fn => {
            const pages = r.pagesByFile[fn] || [];
            const labels = (r.pageLabelsByFile || {})[fn] || [];
            return pages.length > 0 ? `${fn}(${formatPageLabelRanges(labels, pages)})` : fn;
        });

        const row = rtmSheet.addRow([
            idx + 1,
            r.id,
            r.name,
            r.keywords,
            r.appearingFiles.length,
            fileLocs.length > 0 ? fileLocs.join('; ') : '-',
            r.overallStatus
        ]);
        row.height = 24;
        row.alignment = { vertical: 'middle' };

        const statusCell = row.getCell(7);
        const countCell = row.getCell(5);

        if (!r.isPresentOverall) {
            statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
            statusCell.font = { name: '맑은 고딕', color: { argb: 'FFDC2626' }, bold: true };
            countCell.font = { name: '맑은 고딕', color: { argb: 'FFDC2626' }, bold: true };
        } else if (r.overallStatus.includes('누락')) {
            statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF3C7' } };
            statusCell.font = { name: '맑은 고딕', color: { argb: 'FFD97706' }, bold: true };
        } else {
            statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } };
            statusCell.font = { name: '맑은 고딕', color: { argb: 'FF16A34A' }, bold: true };
        }

        for (let col = 1; col <= 7; col++) {
            row.getCell(col).border = borderStyle;
            if (col === 1 || col === 2 || col === 5 || col === 7) {
                row.getCell(col).alignment = { vertical: 'middle', horizontal: 'center' };
            }
        }
    });

    // ── 시트 4 (신설): 산출물 간 정합성 분석 (Cross-Document Consistency) ──
    const crossSheet = workbook.addWorksheet('산출물간 정합성 분석', {
        views: [{ showGridLines: true }]
    });

    crossSheet.columns = [
        { width: 8 },   // A: No.
        { width: 14 },  // B: 구분
        { width: 22 },  // C: 표준 공식 표기 (Approved)
        { width: 24 },  // D: 비표준/금지/유령 표기
        { width: 16 },  // E: 정합성 상태
        { width: 38 },  // F: 표준어 사용 산출물 목록
        { width: 38 },  // G: 비표준어 사용 산출물 목록 (수정필요)
        { width: 30 }   // H: 비고 및 개선 권고
    ];

    crossSheet.mergeCells('A1:H1');
    const crossTitle = crossSheet.getCell('A1');
    crossTitle.value = `[SSOT] 산출물 간 교차 정합성(Cross-Document Consistency) 종합 분석`;
    crossTitle.font = { name: '맑은 고딕', size: 15, bold: true, color: { argb: 'FF0F172A' } };
    crossTitle.alignment = { vertical: 'middle' };
    crossSheet.getRow(1).height = 34;

    // 섹션 1: 산출물 간 표준용어 불일치 현황
    crossSheet.mergeCells('A3:H3');
    const termSec = crossSheet.getCell('A3');
    termSec.value = `1. 산출물 간 용어 표기 불일치 분석 (동일 프로젝트 내 표준어 vs 금지어 교차 사용 현황)`;
    termSec.font = { name: '맑은 고딕', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
    termSec.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF334155' } };
    crossSheet.getRow(3).height = 26;

    const termHead = crossSheet.getRow(4);
    termHead.height = 26;
    const termHeaders = ['No.', '분류', '표준 공식 표기', '금지동의어/오표기', '정합성 상태', '표준어 사용 산출물', '금지어 사용 산출물 (수정 권고)', '개선 조치안'];
    termHeaders.forEach((th, idx) => {
        const c = termHead.getCell(idx + 1);
        c.value = th;
        c.style = subHeaderStyle;
    });

    let crossRowIdx = 5;
    if (agg.termConsistency.length === 0) {
        const row = crossSheet.getRow(crossRowIdx++);
        row.getCell(1).value = 1;
        row.getCell(2).value = '용어';
        row.getCell(3).value = '전체 표준용어 준수';
        row.getCell(4).value = '-';
        row.getCell(5).value = '정합 (일치)';
        row.getCell(6).value = '전체 산출물 일관 표기';
        row.getCell(7).value = '-';
        row.getCell(8).value = '산출물 간 용어 불일치 없음';
        for (let c = 1; c <= 8; c++) row.getCell(c).border = borderStyle;
    } else {
        agg.termConsistency.forEach((tc, idx) => {
            const row = crossSheet.getRow(crossRowIdx++);
            row.height = 24;
            row.getCell(1).value = idx + 1;
            row.getCell(2).value = tc.category;
            row.getCell(3).value = tc.approved;
            row.getCell(4).value = tc.forbidden;
            row.getCell(5).value = tc.isInconsistent ? '불일치 (비정합)' : '금지어 단독사용';
            row.getCell(6).value = tc.filesUsingApproved.join(', ') || '-';
            row.getCell(7).value = tc.filesUsingForbidden.join(', ');
            row.getCell(8).value = `금지어 표기를 공식 표준어 '${tc.approved}'(으)로 통일 필요`;

            for (let c = 1; c <= 8; c++) {
                row.getCell(c).border = borderStyle;
                if (c === 1 || c === 2 || c === 5) row.getCell(c).alignment = { vertical: 'middle', horizontal: 'center' };
                else if (c === 6 || c === 7) row.getCell(c).alignment = { vertical: 'middle', wrapText: true };
            }

            if (tc.isInconsistent) {
                row.getCell(5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
                row.getCell(5).font = { name: '맑은 고딕', color: { argb: 'FFB91C1C' }, bold: true };
                row.getCell(7).font = { name: '맑은 고딕', color: { argb: 'FFB91C1C' } };
            }
        });
    }

    // 섹션 2: 산출물 간 비표준 유령 ID(Ghost ID) 확산 현황
    crossRowIdx += 2;
    crossSheet.mergeCells(`A${crossRowIdx}:H${crossRowIdx}`);
    const ghostSec = crossSheet.getCell(`A${crossRowIdx}`);
    ghostSec.value = `2. 산출물 간 비표준 유령 ID(Ghost ID) 확산 분석 (RFP 마스터 미등록 ID의 다중 산출물 침투 현황)`;
    ghostSec.font = { name: '맑은 고딕', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
    ghostSec.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF334155' } };
    crossSheet.getRow(crossRowIdx).height = 26;

    crossRowIdx++;
    const ghostHead = crossSheet.getRow(crossRowIdx);
    ghostHead.height = 26;
    const ghostHeaders = ['No.', '유형', '검출된 비표준 ID', '확산 파일 수', '확산 심각도', '출현 산출물 목록', '총 출현 횟수', '개선 권고안'];
    ghostHeaders.forEach((gh, idx) => {
        const c = ghostHead.getCell(idx + 1);
        c.value = gh;
        c.style = subHeaderStyle;
    });

    crossRowIdx++;
    if (agg.ghostIdSpread.length === 0) {
        const row = crossSheet.getRow(crossRowIdx++);
        row.getCell(1).value = 1;
        row.getCell(2).value = '유령ID';
        row.getCell(3).value = '비표준 ID 없음';
        row.getCell(4).value = '0개 파일';
        row.getCell(5).value = '정상';
        row.getCell(6).value = '-';
        row.getCell(7).value = 0;
        row.getCell(8).value = '모든 산출물이 RFP 마스터 ID 체계를 준수함';
        for (let c = 1; c <= 8; c++) row.getCell(c).border = borderStyle;
    } else {
        agg.ghostIdSpread.forEach((gs, idx) => {
            const row = crossSheet.getRow(crossRowIdx++);
            row.height = 24;
            row.getCell(1).value = idx + 1;
            row.getCell(2).value = '유령ID';
            row.getCell(3).value = gs.ghostId;
            row.getCell(4).value = `${gs.fileCount}개 파일`;
            row.getCell(5).value = gs.fileCount >= 3 ? '광역 확산 (치명)' : gs.fileCount >= 2 ? '다중 확산 (주의)' : '단일 파일';
            row.getCell(6).value = gs.files.join(', ');
            row.getCell(7).value = `${gs.totalCount}회`;
            row.getCell(8).value = `마스터 요구사항 ID 체계와 대조하여 공식 ID로 정정하거나 삭제`;

            for (let c = 1; c <= 8; c++) {
                row.getCell(c).border = borderStyle;
                if (c === 1 || c === 2 || c === 3 || c === 4 || c === 5 || c === 7) row.getCell(c).alignment = { vertical: 'middle', horizontal: 'center' };
                else if (c === 6) row.getCell(c).alignment = { vertical: 'middle', wrapText: true };
            }

            if (gs.fileCount >= 2) {
                row.getCell(5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
                row.getCell(5).font = { name: '맑은 고딕', color: { argb: 'FFB91C1C' }, bold: true };
            }
        });
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const downloadName = `다중_산출물_통합_정합성_검수_리포트_${new Date().toISOString().slice(0, 10)}.xlsx`;
    saveAs(blob, downloadName);
}


// ── 6-3. 공공 ISP 공식 표준 요구사항추적표 및 산출물현황표 엑셀 엔진 ──
export const OFFICIAL_RTM_TEMPLATE_ITEMS = [
    { id: 'CSR-001', name: '내·외부 환경분석', status: '충족', file: '2025_K-BDS_ISP_K-BDS_ISP_Ⅱ.환경분석.pptx', pages: '7~58' },
    { id: 'CSR-002', name: '법·제도 분석', status: '충족', file: '2025_K-BDS_ISP_K-BDS_ISP_Ⅱ.환경분석.pptx', pages: '123~195' },
    { id: 'CSR-003', name: '정기술(IT) 환경분석', status: '충족', file: '2025_K-BDS_ISP_K-BDS_ISP_Ⅱ.환경분석.pptx', pages: '59~122' },
    { id: 'CSR-004', name: '업무 현황 분석', status: '충족', file: '2025_K-BDS_ISP_K-BDS_ISP_Ⅲ.현황분석.pptx', pages: '5~75' },
    { id: 'CSR-005', name: '정보기술(IT) 현황 분석', status: '충족', file: '2025_K-BDS_ISP_K-BDS_ISP_Ⅲ.현황분석.pptx', pages: '76~347' },
    { id: 'CSR-006', name: '벤치마킹', status: '충족', file: '2025_K-BDS_ISP_K-BDS_ISP_Ⅲ.현황분석.pptx', pages: '348~484' },
    { id: 'CSR-007', name: '차이(Gap) 분석', status: '충족', file: '2025_K-BDS_ISP_K-BDS_ISP_Ⅲ.현황분석.pptx', pages: '485~487' },
    { id: 'CSR-008', name: '이슈통합 및 개선과제 도출', status: '충족', file: '2025_K-BDS_ISP_K-BDS_ISP_Ⅲ.현황분석.pptx', pages: '488~503' },
    { id: 'CSR-009', name: '정보화 비전 및 전략 수립', status: '충족', file: '2025_K-BDS_ISP_K-BDS_ISP_Ⅳ.목표모델수립.pptx', pages: '6~19' },
    { id: 'CSR-010', name: 'To-Be 개선과제 상세화', status: '충족', file: '2025_K-BDS_ISP_K-BDS_ISP_Ⅳ.목표모델수립.pptx', pages: '23~112' },
    { id: 'CSR-011', name: 'To-Be 업무 프로세스 설계', status: '충족', file: '2025_K-BDS_ISP_K-BDS_ISP_Ⅳ.목표모델수립.pptx', pages: '113~414' },
    { id: 'CSR-012', name: 'To-Be 정보시스템 구조 설계', status: '충족', file: '2025_K-BDS_ISP_K-BDS_ISP_Ⅳ.목표모델수립.pptx', pages: '415~731' },
    { id: 'CSR-013', name: 'To-Be 데이터 구조 설계', status: '충족', file: '2025_K-BDS_ISP_K-BDS_ISP_Ⅳ.목표모델수립.pptx', pages: '732~1105' },
    { id: 'CSR-014', name: 'To-Be 기술 및 보안 구조 설계', status: '충족', file: '2025_K-BDS_ISP_K-BDS_ISP_Ⅳ.목표모델수립.pptx', pages: '1106~1219' },
    { id: 'CSR-015', name: 'To-Be 정보관리 구조 설계', status: '충족', file: '2025_K-BDS_ISP_K-BDS_ISP_Ⅳ.목표모델수립.pptx', pages: '1220~1233' },
    { id: 'CSR-016', name: '통합이행계획 수립', status: '충족', file: '2025_K-BDS_ISP_K-BDS_ISP_Ⅴ.이행계획수립.pptx', pages: '1~66' },
    { id: 'CSR-017', name: '총 구축비 산출', status: '충족', file: '2025_K-BDS_ISP_K-BDS_ISP_Ⅴ.이행계획수립.pptx', pages: '67~87' },
    { id: 'CSR-018', name: '기대효과 분석', status: '충족', file: '2025_K-BDS_ISP_K-BDS_ISP_Ⅴ.이행계획수립.pptx', pages: '87~91' },
    { id: 'CSR-019', name: '제안요청서(RFP) 작성', status: '충족', file: "2025_K-BDS_ISP_제안요청서_차세대 K-BDS 시스템 구축 용역 _27년.hwp\n2025_K-BDS_ISP_제안요청서_차세대 K-BDS 시스템 구축 용역 _28년.hwp", pages: '전 페이지' },
    { id: 'CSR-020', name: 'ISP 최종산출물 검토신청서 작성', status: '충족', file: '2025_K-BDS_ISP_검토신청서.hwpx', pages: '전 페이지' },
    { id: 'DAR-001', name: '데이터 표준화 방안 제시', status: '충족', file: '2025_K-BDS_ISP_K-BDS_ISP_Ⅳ.목표모델수립.pptx', pages: '758~761' },
    { id: 'DAR-002', name: '데이터 표준관리 방안 제시', status: '충족', file: '2025_K-BDS_ISP_K-BDS_ISP_Ⅳ.목표모델수립.pptx', pages: '783~791' },
    { id: 'DAR-003', name: '데이터 구조 설계 방안 제시', status: '충족', file: '2025_K-BDS_ISP_K-BDS_ISP_Ⅳ.목표모델수립.pptx', pages: '775~782' },
    { id: 'DAR-004', name: '데이터 구조 관리 방안 제시', status: '충족', file: '2025_K-BDS_ISP_K-BDS_ISP_Ⅳ.목표모델수립.pptx', pages: '1106' },
    { id: 'DAR-005', name: '데이터 관리체계 현황분석 및 목표모델 제시', status: '충족', file: '2025_K-BDS_ISP_K-BDS_ISP_Ⅳ.목표모델수립.pptx', pages: '1109~1110' },
    { id: 'DAR-006', name: '데이터 개방 및 메타데이터 관리체계 제시', status: '충족', file: '2025_K-BDS_ISP_K-BDS_ISP_Ⅳ.목표모델수립.pptx', pages: '1107~1108' },
    { id: 'COR-001', name: '기술적용계획표 준수', status: '충족', file: '2025_K-BDS_ISP_[별지 1] 기술적용계획표 및 기술적용결과표(행정기관 및 공공기관 정보시스템 구축·운영 지침 ).hwpx', pages: '전 페이지' },
    { id: 'COR-002', name: '정보화사업 관련 법ㆍ규정 준수', status: '충족', file: '2025_K-BDS_ISP_[별지 1] 기술적용계획표 및 기술적용결과표(행정기관 및 공공기관 정보시스템 구축·운영 지침 ).hwpx', pages: '전 페이지' },
    { id: 'SER-001', name: '보안준수 일반사항', status: '충족', file: '착수계(사업수행계획서, 외부자 보안서약서, 개인정보 취급자보안서약서 등)', pages: '전 페이지' },
    { id: 'SER-002', name: '자료보안관리', status: '충족', file: '착수계(사업수행계획서, 외부자 보안서약서, 개인정보 취급자보안서약서 등)', pages: '전 페이지' },
    { id: 'SER-003', name: '인력보안관리', status: '충족', file: '착수계(사업수행계획서, 외부자 보안서약서, 개인정보 취급자보안서약서 등)', pages: '전 페이지' },
    { id: 'SER-004', name: '작업장소 보안관리', status: '충족', file: '착수계(사업수행계획서, 외부자 보안서약서, 개인정보 취급자보안서약서 등)', pages: '전 페이지' }
];

export const OFFICIAL_DOCUMENT_CHECKLIST_ITEMS = [
    { seq: 1, docName: '환경 분석서', status: '미작성', fileName: '', note: 'Ⅱ.1 ~ Ⅱ.6', keywords: ['환경분석', '개요', 'Ⅱ.1'] },
    { seq: 2, docName: '현황 분석서', status: '미작성', fileName: '', note: 'Ⅲ.1 ~ Ⅲ.8', keywords: ['현황분석', '개요', 'Ⅲ.1'] },
    { seq: 3, docName: '법·제도 분석', status: '미작성', fileName: '', note: 'Ⅱ.3', keywords: ['법제도', '법·제도', 'Ⅱ.3'] },
    { seq: 4, docName: '정보기술 동향 분석서', status: '미작성', fileName: '', note: 'Ⅱ.4, Ⅱ.5', keywords: ['정보기술', '동향', '활용동향', 'Ⅱ.4', 'Ⅱ.5'] },
    { seq: 5, docName: '업무 현황 분석서(인터뷰 결과 포함)', status: '미작성', fileName: '', note: 'Ⅲ.2', keywords: ['업무현황', '업무 현황', '인터뷰', 'Ⅲ.2'] },
    { seq: 6, docName: '정보기술(IT) 현황 분석서', status: '미작성', fileName: '', note: 'Ⅲ.3, Ⅲ.4', keywords: ['정보기술', '응용시스템', '인프라', 'Ⅲ.3', 'Ⅲ.4'] },
    { seq: 7, docName: '선진사례 분석서', status: '미작성', fileName: '', note: 'Ⅲ.7', keywords: ['선진사례', '벤치마킹', 'Ⅲ.7'] },
    { seq: 8, docName: '차이(Gap) 분석서', status: '미작성', fileName: '', note: 'Ⅲ.7, Ⅲ.8', keywords: ['차이', 'Gap', '개선과제', 'Ⅲ.8'] },
    { seq: 9, docName: '요구사항 및 개선과제 분석서', status: '미작성', fileName: '', note: 'Ⅲ.6, Ⅲ.8', keywords: ['요구사항분석', '개선과제', 'Ⅲ.6', 'Ⅲ.8'] },
    { seq: 10, docName: '정보화 전략 정의서', status: '미작성', fileName: '', note: 'Ⅳ.2', keywords: ['비전', '전략수립', '정보화전략', 'Ⅳ.2'] },
    { seq: 11, docName: 'To-Be 과제 상세 정의서', status: '미작성', fileName: '', note: 'Ⅳ.3', keywords: ['개선과제상세화', '과제상세', 'Ⅳ.3'] },
    { seq: 12, docName: 'To-Be 업무 프로세스 설계서', status: '미작성', fileName: '', note: 'Ⅳ.4', keywords: ['목표업무', '업무아키텍처', 'Ⅳ.4'] },
    { seq: 13, docName: 'To-Be 정보관리 구조 설계서', status: '미작성', fileName: '', note: 'Ⅳ.12', keywords: ['정보화통합거버넌스', '거버넌스', 'Ⅳ.12'] },
    { seq: 14, docName: 'To-Be 정보시스템 구조 설계서', status: '미작성', fileName: '', note: 'Ⅳ.5', keywords: ['목표응용시스템', '응용시스템아키텍처', 'Ⅳ.5'] },
    { seq: 15, docName: 'To-Be 데이터 구조 설계서', status: '미작성', fileName: '', note: 'Ⅳ.6, Ⅳ.10', keywords: ['목표데이터', '데이터아키텍처', 'DW', 'CDW', 'Ⅳ.6', 'Ⅳ.10'] },
    { seq: 16, docName: 'To-Be 기술 및 보안 구조 설계서', status: '미작성', fileName: '', note: 'Ⅳ.7, Ⅳ.8', keywords: ['목표인프라', '정보보호아키텍처', 'AI아키텍처', 'Ⅳ.7', 'Ⅳ.8'] },
    { seq: 17, docName: 'To-Be AI 서비스 확산 전략서', status: '미작성', fileName: '', note: 'Ⅳ.9', keywords: ['AI서비스확산', '스마트병원', '활성화', 'Ⅳ.9'] },
    { seq: 18, docName: '통합이행계획 수립서', status: '미작성', fileName: '', note: 'Ⅴ.1 ~ Ⅴ.5', keywords: ['이행계획', '로드맵', '이행과제', 'Ⅴ.1', 'Ⅴ.2', 'Ⅴ.3', 'Ⅴ.5'] },
    { seq: 19, docName: '소요예산 산정내역서(산출근거 포함)', status: '미작성', fileName: '', note: 'Ⅴ.6', keywords: ['소요예산', '예산', '산정', 'Ⅴ.6'] },
    { seq: 20, docName: '기대효과 및 투자효과 분석서', status: '미작성', fileName: '', note: 'Ⅴ.7, Ⅴ.8', keywords: ['기대효과', '투자효과', 'Ⅴ.7', 'Ⅴ.8'] },
    { seq: 21, docName: '공통기반 시스템 연계·활용 방안', status: '미작성', fileName: '', note: 'Ⅳ.11', keywords: ['공통기반', '연계활용', '연계', 'Ⅳ.11'] },
    { seq: 22, docName: '제안요청서(RFP) 산출물', status: '미작성', fileName: '', note: 'RFP/과업지시서', keywords: ['제안요청서', 'RFP', '과업지시서'] },
    { seq: 23, docName: 'ISP 최종산출물 검토신청서', status: '미작성', fileName: '', note: '검토신청서', keywords: ['검토신청서', '신청서'] },
    { seq: 24, docName: '최종산출물 점검표', status: '미작성', fileName: '', note: '점검표', keywords: ['점검표', '산출물점검'] },
    { seq: 25, docName: '목표시스템 데이터 표준화 방안', status: '미작성', fileName: '', note: '데이터 표준화', keywords: ['데이터표준화', '표준화'] },
    { seq: 26, docName: '데이터 표준관리 현황 분석서', status: '미작성', fileName: '', note: '데이터 표준관리', keywords: ['표준관리', '데이터현황', 'Ⅲ.5'] },
    { seq: 27, docName: '데이터 구조현황 분석서', status: '미작성', fileName: '', note: '데이터 구조분석', keywords: ['데이터구조', '구조현황'] },
    { seq: 28, docName: '데이터 구조 설계 방안', status: '미작성', fileName: '', note: '데이터 구조설계', keywords: ['구조설계', '데이터아키텍처'] },
    { seq: 29, docName: '데이터 구조 관리현황 분석서', status: '미작성', fileName: '', note: '데이터 관리현황', keywords: ['구조관리', '데이터관리'] },
    { seq: 30, docName: '데이터 구조 관리방안', status: '미작성', fileName: '', note: '데이터 관리방안', keywords: ['구조관리방안'] },
    { seq: 31, docName: '목표시스템 데이터(표준, 구조, 값, 연계)관리방안', status: '미작성', fileName: '', note: '데이터 종합관리', keywords: ['데이터종합', '데이터관리방안'] },
    { seq: 32, docName: '메타데이터 등록·관리현황', status: '미작성', fileName: '', note: '메타데이터', keywords: ['메타데이터'] },
    { seq: 33, docName: '기술적용계획표 및 결과표', status: '미작성', fileName: '', note: '기술적용', keywords: ['기술적용', '기술적용계획표'] },
    { seq: 34, docName: '보안계획서 및 관련 문서', status: '미작성', fileName: '', note: '보안계획', keywords: ['보안계획', '정보보호', '보안서약서'] }
];

// 정렬된 숫자 배열을 ['7~9', '12'] 형태의 구간 목록으로 축약
function compressNumericRanges(nums) {
    if (nums.length === 0) return [];
    if (nums.length >= 10 && (nums[nums.length - 1] - nums[0] + 1) === nums.length) {
        return [`${nums[0]}~${nums[nums.length - 1]}`];
    }
    const ranges = [];
    let start = nums[0];
    let prev = nums[0];
    for (let i = 1; i < nums.length; i++) {
        const cur = nums[i];
        if (cur === prev + 1) {
            prev = cur;
        } else {
            ranges.push(start === prev ? `${start}` : `${start}~${prev}`);
            start = cur;
            prev = cur;
        }
    }
    ranges.push(start === prev ? `${start}` : `${start}~${prev}`);
    return ranges;
}

/**
 * 페이지 배열을 '7~58' 또는 '12, 15~18' 형식으로 스마트 축약
 */
export function formatPageRanges(pages = []) {
    if (!pages || pages.length === 0) return '-';
    const nums = Array.from(new Set(pages.map(p => parseInt(p, 10)).filter(n => !isNaN(n)))).sort((a, b) => a - b);
    if (nums.length === 0) return '전 페이지';
    return compressNumericRanges(nums).join(', ');
}

/**
 * 슬라이드 하단 표시 라벨("장번호-슬라이드번호", 예: Ⅱ-15)을 장별로 묶어 'Ⅱ-7~9, Ⅱ-12, Ⅲ-5' 형식으로 축약.
 * 라벨이 하나도 없으면(PPTX 외 문서 등) 숫자 페이지로 폴백한다.
 */
export function formatPageLabelRanges(labels = [], plainPages = []) {
    const labelList = (labels || []).filter(Boolean);
    if (labelList.length === 0) return formatPageRanges(plainPages);

    const groups = new Map(); // 장번호 -> Set<슬라이드번호> (등장 순서 유지)
    const unlabeledNums = new Set();
    labelList.forEach(label => {
        const m = String(label).match(/^(.+?)-(\d+)$/);
        if (m) {
            if (!groups.has(m[1])) groups.set(m[1], new Set());
            groups.get(m[1]).add(parseInt(m[2], 10));
        } else {
            const n = parseInt(label, 10);
            if (!isNaN(n)) unlabeledNums.add(n);
        }
    });

    const parts = [];
    groups.forEach((numSet, prefix) => {
        const nums = Array.from(numSet).sort((a, b) => a - b);
        compressNumericRanges(nums).forEach(range => parts.push(`${prefix}-${range}`));
    });
    if (unlabeledNums.size > 0) {
        parts.push(...compressNumericRanges(Array.from(unlabeledNums).sort((a, b) => a - b)));
    }
    return parts.length > 0 ? parts.join(', ') : formatPageRanges(plainPages);
}

const ROMAN_NUMERAL_VALUES = { 'Ⅰ': 1, 'Ⅱ': 2, 'Ⅲ': 3, 'Ⅳ': 4, 'Ⅴ': 5, 'Ⅵ': 6, 'Ⅶ': 7, 'Ⅷ': 8, 'Ⅸ': 9, 'Ⅹ': 10 };
const ASCII_ROMAN_VALUES = { I: 1, II: 2, III: 3, IV: 4, V: 5, VI: 6, VII: 7, VIII: 8, IX: 9, X: 10 };
const NUM_TO_ROMAN = Object.fromEntries(Object.entries(ROMAN_NUMERAL_VALUES).map(([k, v]) => [v, k]));

// 파일명에서 장 번호 추출 (예: "..._Ⅱ.환경분석.pptx" → 2, "..._III.현황분석.pptx" → 3)
function detectChapterFromFileName(fileName = '') {
    const uni = fileName.match(/([ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ])\s*[.．]/);
    if (uni) return ROMAN_NUMERAL_VALUES[uni[1]];
    const ascii = fileName.match(/(?:^|[_\s\-(])(VIII|VII|VI|IV|IX|III|II|I|V|X)\s*[.．]/);
    if (ascii) return ASCII_ROMAN_VALUES[ascii[1]];
    return null;
}

// 비고(작성근거)의 목차 코드 파싱 (예: "Ⅱ.1 ~ Ⅱ.6" → Ⅱ장 1~6, "Ⅳ.6, Ⅳ.10" → Ⅳ장 6, 10)
function parseTocRefs(note = '') {
    const re = /([ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ])\.(\d+)/g;
    const tokens = [];
    let m;
    while ((m = re.exec(note)) !== null) {
        tokens.push({ chapter: ROMAN_NUMERAL_VALUES[m[1]], num: parseInt(m[2], 10), start: m.index, end: re.lastIndex });
    }
    const refs = [];
    tokens.forEach((t, i) => {
        const prev = tokens[i - 1];
        if (prev && prev.chapter === t.chapter && note.slice(prev.end, t.start).includes('~')) {
            for (let n = prev.num + 1; n <= t.num; n++) refs.push({ chapter: t.chapter, num: n });
        } else {
            refs.push({ chapter: t.chapter, num: t.num });
        }
    });
    return refs;
}

// 슬라이드 하단 라벨(Ⅱ-15, 3-15)에서 가장 많이 쓰인 장 번호 추출
function detectChapterFromPageLabels(pageLabelMap = {}) {
    const counts = new Map();
    Object.values(pageLabelMap).forEach(label => {
        const prefix = String(label).split('-')[0];
        const value = ROMAN_NUMERAL_VALUES[prefix] || (/^\d{1,2}$/.test(prefix) ? parseInt(prefix, 10) : null);
        if (value) counts.set(value, (counts.get(value) || 0) + 1);
    });
    let best = null;
    counts.forEach((cnt, value) => { if (!best || cnt > best.cnt) best = { value, cnt }; });
    return best ? best.value : null;
}

// 앞쪽 슬라이드(목차/CONTENTS 슬라이드 우선)에서 목차 항목 추출.
// 숫자 목차("3. 법·제도 분석", "Ⅱ.3 법·제도 분석")와 장 목차("Ⅱ. 환경분석")를 분리 보관해,
// 매 슬라이드 상단의 장 구분 헤더가 같은 번호의 숫자 목차를 가리지 않게 한다.
function extractTocEntries(pageMap = {}) {
    const slideNums = Object.keys(pageMap).map(Number).filter(n => !isNaN(n)).sort((a, b) => a - b).slice(0, 12);
    const isTocMarker = p => /목\s*차|CONTENTS|INDEX/i.test(p);
    const tocSlides = slideNums.filter(n => (pageMap[n] || []).some(isTocMarker));
    const scanSlides = tocSlides.length > 0 ? tocSlides : slideNums;

    const numeric = new Map();
    const roman = new Map();
    const addNumeric = (num, rawTitle) => {
        const title = rawTitle.trim();
        if (title.length >= 2 && title.length <= 40 && !numeric.has(num)) numeric.set(num, title);
    };

    scanSlides.forEach(n => {
        const ps = pageMap[n] || [];
        ps.forEach((p, i) => {
            let m = p.match(/^\s*0*([1-9][0-9]?)\.\s+(\D.*)$/);
            if (m) { addNumeric(parseInt(m[1], 10), m[2]); return; }
            m = p.match(/^\s*[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]\s*\.\s*([1-9][0-9]?)\.?\s+(\D.*)$/);
            if (m) { addNumeric(parseInt(m[1], 10), m[2]); return; }
            m = p.match(/^\s*([ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ])\s*[.．]\s*(\D.*)$/);
            if (m) {
                const num = ROMAN_NUMERAL_VALUES[m[1]];
                const title = m[2].trim();
                if (title.length >= 2 && title.length <= 40 && !roman.has(num)) roman.set(num, `${m[1]}. ${title}`);
                return;
            }
            // 번호 토큰과 제목이 별도 문단으로 나뉜 목차 (예: "1" / "환경분석 개요")
            if (/^0?[1-9]$/.test(p.trim())) {
                const next = (ps[i + 1] || '').trim();
                if (next && !/^\d+$/.test(next) && !isTocMarker(next) && !/CHAPTER/i.test(next)) {
                    addNumeric(parseInt(p, 10), next);
                }
            }
        });
    });

    const toList = map => Array.from(map.entries()).sort((a, b) => a[0] - b[0]).map(([num, title]) => ({ num, title }));
    return { numericToc: toList(numeric), romanToc: toList(roman) };
}

// 검수 대상 PPTX별 장 번호와 목차 색인
async function buildDeliverableTocIndex(batchResults = []) {
    const index = [];
    for (const res of batchResults) {
        if (!res.file || getFileExtension(res.fileName || '') !== 'pptx') continue;
        let content;
        try {
            content = await extractStructuredFileContent(res.file);
        } catch (e) {
            console.warn(`[산출물현황] '${res.fileName}' 목차 분석 실패:`, e);
            continue;
        }
        const { numericToc, romanToc } = extractTocEntries(content.pageMap);
        const chapter = detectChapterFromFileName(res.fileName)
            ?? detectChapterFromPageLabels(content.pageLabelMap)
            ?? (romanToc.length === 1 ? romanToc[0].num : null);
        index.push({ fileName: res.fileName, chapter, numericToc, romanToc });
    }
    return index;
}

// 목차 코드로 산출물이 담긴 파일과 해당 목차 제목을 찾는다
function resolveDocByTocRefs(refs, tocIndex) {
    const fileNames = [];
    const tocLines = [];
    const byChapter = new Map();
    refs.forEach(r => {
        if (!byChapter.has(r.chapter)) byChapter.set(r.chapter, []);
        byChapter.get(r.chapter).push(r.num);
    });

    byChapter.forEach((nums, chapter) => {
        const roman = NUM_TO_ROMAN[chapter] || String(chapter);
        const chapterFile = tocIndex.find(x => x.chapter === chapter);
        if (chapterFile) {
            if (!fileNames.includes(chapterFile.fileName)) fileNames.push(chapterFile.fileName);
            nums.forEach(n => {
                const t = chapterFile.numericToc.find(item => item.num === n);
                if (t) tocLines.push(`${roman}.${n} ${t.title}`);
            });
            return;
        }
        // 전 장이 한 파일에 담긴 종합보고서: 장 단위 목차 제목으로 표기
        const combined = tocIndex.find(x => x.romanToc.some(t => t.num === chapter));
        if (combined) {
            if (!fileNames.includes(combined.fileName)) fileNames.push(combined.fileName);
            tocLines.push(combined.romanToc.find(t => t.num === chapter).title);
        }
    });

    return { fileNames, tocLines };
}

/**
 * 공공기관 ISP/ISMP 공식 표준 요구사항추적표 및 산출물점검표 엑셀 생성
 */
export async function exportOfficialRtmToExcel({
    ssotData = {},
    batchResults = [],
    companyName = '㈜한국아이티컨설팅',
    issueDate = '2025.12.17.',
    mode = 'auto', // 'auto' (검수 결과 자동 연동) | 'template' (표준 템플릿 원본)
    includeDocChecklist = false // 산출물현황 시트 포함 여부 (기본 false: 요구사항추적표 단일 시트만 생성)
}) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = `${companyName} (건강한 프로젝트 RTM)`;
    workbook.created = new Date();

    const borderThin = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        left: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'thin', color: { argb: 'FF000000' } },
        right: { style: 'thin', color: { argb: 'FF000000' } }
    };

    const headerFill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFF2F2F2' } // 연회색
    };

    // ── 시트 1: 요구사항추적표 ────────────────────────────────────
    const rtmSheet = workbook.addWorksheet('요구사항추적표', {
        views: [{ showGridLines: true }]
    });

    // 열 너비 설정 (A, B는 여백, C~G는 데이터)
    rtmSheet.columns = [
        { width: 4 },   // A (여백)
        { width: 4 },   // B (여백)
        { width: 16 },  // C: 요구사항번호
        { width: 36 },  // D: 요구사항명칭
        { width: 12 },  // E: 충족여부
        { width: 55 },  // F: 관련 파일
        { width: 18 }   // G: 해당 페이지
    ];

    // 상단 여백 및 작성자/일자 정보 (G5 위치)
    rtmSheet.mergeCells('F5:G5');
    const authorCell = rtmSheet.getCell('F5');
    authorCell.value = `${companyName}\n${issueDate}`;
    authorCell.font = { name: '맑은 고딕', size: 10, bold: true };
    authorCell.alignment = { vertical: 'middle', horizontal: 'right', wrapText: true };
    rtmSheet.getRow(5).height = 30;

    // 헤더 행 (Row 7)
    const rtmHeaderRow = rtmSheet.getRow(7);
    rtmHeaderRow.height = 28;

    const rtmHeaders = [
        { col: 3, title: '요구사항번호' },
        { col: 4, title: '요구사항명칭' },
        { col: 5, title: '충족여부' },
        { col: 6, title: '관련 파일' },
        { col: 7, title: '해당 페이지' }
    ];

    rtmHeaders.forEach(h => {
        const cell = rtmHeaderRow.getCell(h.col);
        cell.value = h.title;
        cell.font = { name: '맑은 고딕', size: 10, bold: true };
        cell.fill = headerFill;
        cell.border = borderThin;
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    // 데이터 행 생성
    let itemsToRender = [];

    if (mode === 'template') {
        itemsToRender = OFFICIAL_RTM_TEMPLATE_ITEMS;
    } else {
        // mode === 'auto': 현재 ssotData.requirements와 검수결과(batchResults) 결합
        const masterReqs = ssotData.requirements || [];

        if (masterReqs.length > 0) {
            itemsToRender = masterReqs.map(req => {
                const reqId = (req.id || '').trim().toUpperCase();
                const matchedFiles = [];
                const pageCollector = [];
                const labelCollector = [];

                batchResults.forEach(res => {
                    const rtmMatrix = res.stage1?.rtmMatrix || [];
                    const found = rtmMatrix.find(m => (m.id || '').toUpperCase() === reqId);
                    if (found && found.isPresent) {
                        matchedFiles.push(res.fileName);
                        if (found.pages && found.pages.length > 0) {
                            pageCollector.push(...found.pages);
                        }
                        if (found.pageLabels && found.pageLabels.length > 0) {
                            labelCollector.push(...found.pageLabels);
                        }
                    }
                });

                const isPresent = matchedFiles.length > 0;
                let pageStr = '-';
                if (isPresent) {
                    pageStr = pageCollector.length > 0 ? formatPageLabelRanges(labelCollector, pageCollector) : '전 페이지';
                }

                return {
                    id: req.id,
                    name: req.name,
                    status: isPresent ? '충족' : (batchResults.length > 0 ? '미충족' : '충족'),
                    file: matchedFiles.length > 0 ? matchedFiles.join('\n') : (batchResults.length > 0 ? '미확인 산출물' : '산출물 확인 필요'),
                    pages: pageStr
                };
            });
        } else {
            itemsToRender = OFFICIAL_RTM_TEMPLATE_ITEMS;
        }
    }

    let startRow = 8;
    itemsToRender.forEach((item, idx) => {
        const row = rtmSheet.getRow(startRow + idx);
        row.height = item.file && item.file.includes('\n') ? 36 : 24;

        // C열: 요구사항번호
        const cCell = row.getCell(3);
        cCell.value = item.id;
        cCell.font = { name: '맑은 고딕', size: 9.5 };
        cCell.alignment = { vertical: 'middle', horizontal: 'center' };
        cCell.border = borderThin;

        // D열: 요구사항명칭
        const dCell = row.getCell(4);
        dCell.value = item.name;
        dCell.font = { name: '맑은 고딕', size: 9.5 };
        dCell.alignment = { vertical: 'middle', horizontal: 'left' };
        dCell.border = borderThin;

        // E열: 충족여부
        const eCell = row.getCell(5);
        eCell.value = item.status || '충족';
        eCell.font = { name: '맑은 고딕', size: 9.5, bold: true, color: { argb: item.status === '미충족' ? 'FFDC2626' : 'FF000000' } };
        eCell.alignment = { vertical: 'middle', horizontal: 'center' };
        eCell.border = borderThin;

        // F열: 관련 파일
        const fCell = row.getCell(6);
        fCell.value = item.file || '-';
        fCell.font = { name: '맑은 고딕', size: 9.5 };
        fCell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        fCell.border = borderThin;

        // G열: 해당 페이지
        const gCell = row.getCell(7);
        gCell.value = item.pages || '-';
        gCell.font = { name: '맑은 고딕', size: 9.5 };
        gCell.alignment = { vertical: 'middle', horizontal: 'center' };
        gCell.border = borderThin;
    });

    // ── 시트 2: RFP상 산출물 (선택 시에만 포함, 기본값 false) ─────────────
    if (includeDocChecklist) {
        const docSheet = workbook.addWorksheet('RFP상 산출물', {
            views: [{ showGridLines: true }]
        });

        docSheet.columns = [
            { width: 4 },   // A (여백)
            { width: 8 },   // B: 구분
            { width: 38 },  // C: 서류
            { width: 14 },  // D: 현황
            { width: 50 },  // E: 파일명
            { width: 45 },  // F: 주요 목차
            { width: 20 }   // G: 비고(작성근거)
        ];

        // 상단 작성자/일자 정보 (G5)
        docSheet.mergeCells('F5:G5');
        const docAuthorCell = docSheet.getCell('F5');
        docAuthorCell.value = `${companyName}\n${issueDate}`;
        docAuthorCell.font = { name: '맑은 고딕', size: 10, bold: true };
        docAuthorCell.alignment = { vertical: 'middle', horizontal: 'right', wrapText: true };
        docSheet.getRow(5).height = 30;

        // 헤더 행 (Row 7)
        const docHeaderRow = docSheet.getRow(7);
        docHeaderRow.height = 28;

        const docHeaders = [
            { col: 2, title: '구분' },
            { col: 3, title: '서류' },
            { col: 4, title: '현황' },
            { col: 5, title: '파일명' },
            { col: 6, title: '주요 목차' },
            { col: 7, title: '비고(작성근거)' }
        ];

        docHeaders.forEach(h => {
            const cell = docHeaderRow.getCell(h.col);
            cell.value = h.title;
            cell.font = { name: '맑은 고딕', size: 10, bold: true };
            cell.fill = headerFill;
            cell.border = borderThin;
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
        });

        // 실제 업로드된 검수 산출물과 RFP 요구서류 매칭
        // 1순위: 비고의 목차 코드(Ⅱ.3 등) ↔ 파일의 장 번호 + 목차 슬라이드 항목
        // 2순위: 파일명 키워드 매칭 (목차 코드가 없거나 장 파일을 찾지 못한 경우)
        const tocIndex = batchResults.length > 0 ? await buildDeliverableTocIndex(batchResults) : [];

        const docItemsToRender = OFFICIAL_DOCUMENT_CHECKLIST_ITEMS.map(item => {
            let fileNames = [];
            let tocLines = [];

            if (batchResults.length > 0) {
                const refs = parseTocRefs(item.note);
                if (refs.length > 0) {
                    ({ fileNames, tocLines } = resolveDocByTocRefs(refs, tocIndex));
                }

                if (fileNames.length === 0) {
                    const targetKeywords = (item.keywords || []).map(k => k.toLowerCase().replace(/[\s·\-_]/g, ''));
                    const docNameKey = item.docName.toLowerCase().replace(/[\s·\-_]/g, '');
                    const matchedFile = batchResults.find(b => {
                        const fName = (b.fileName || '').toLowerCase().replace(/[\s·\-_]/g, '');
                        if (fName.includes(docNameKey)) return true;
                        return targetKeywords.some(k => k.length >= 2 && fName.includes(k));
                    });
                    if (matchedFile) fileNames = [matchedFile.fileName];
                }
            }

            const isMatched = fileNames.length > 0;
            return {
                ...item,
                status: isMatched ? '작성완료' : '미작성',
                fileName: isMatched ? fileNames.join('\n') : '-',
                tocSummary: tocLines.length > 0 ? tocLines.join('\n') : '-'
            };
        });

        docItemsToRender.forEach((item, idx) => {
            const row = docSheet.getRow(8 + idx);
            const lineCount = Math.max(item.fileName.split('\n').length, item.tocSummary.split('\n').length);
            row.height = Math.max(24, lineCount * 15);

            // B열: 구분
            const bCell = row.getCell(2);
            bCell.value = item.seq;
            bCell.font = { name: '맑은 고딕', size: 9.5 };
            bCell.alignment = { vertical: 'middle', horizontal: 'center' };
            bCell.border = borderThin;

            // C열: 서류
            const cCell = row.getCell(3);
            cCell.value = item.docName;
            cCell.font = { name: '맑은 고딕', size: 9.5 };
            cCell.alignment = { vertical: 'middle', horizontal: 'left' };
            cCell.border = borderThin;

            // D열: 현황
            const dCell = row.getCell(4);
            dCell.value = item.status || '작성완료';
            dCell.font = { name: '맑은 고딕', size: 9.5 };
            dCell.alignment = { vertical: 'middle', horizontal: 'center' };
            dCell.border = borderThin;

            // E열: 파일명
            const eCell = row.getCell(5);
            eCell.value = item.fileName || '';
            eCell.font = { name: '맑은 고딕', size: 9.5 };
            eCell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
            eCell.border = borderThin;

            // F열: 주요 목차
            const fCell = row.getCell(6);
            fCell.value = item.tocSummary;
            fCell.font = { name: '맑은 고딕', size: 9.5 };
            fCell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
            fCell.border = borderThin;

            // G열: 비고
            const gCell = row.getCell(7);
            gCell.value = item.note || '';
            gCell.font = { name: '맑은 고딕', size: 9.5 };
            gCell.alignment = { vertical: 'middle', horizontal: 'left' };
            gCell.border = borderThin;
        });
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const filePrefix = includeDocChecklist ? '요구사항추적표_및_산출물현황표' : '요구사항추적표(조견표)';
    const downloadName = `${filePrefix}_${(ssotData.metadata?.projectName || '공공ISP').replace(/\s+/g, '_')}_${new Date().toISOString().slice(0, 10)}.xlsx`;
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
    // ID -> { id, name, keywords } 맵으로 관리하여 유효한 진짜 이름이 발견되면 즉시 업데이트!
    const reqDict = new Map();

    const isPlaceholderName = (name, id) => {
        if (!name) return true;
        const clean = String(name).trim();
        if (clean.length < 3) return true;
        if (clean === `요구사항 ${id}` || clean === `요구사항${id}` || clean === id) return true;
        if (/^요구사항\s*[A-Z]{2,6}[-_]\d{2,4}$/i.test(clean)) return true;
        if (/^[0-9\s.,\-_/]+$/.test(clean)) return true;
        return false;
    };

    const addOrUpdateReq = (id, rawName, rawDesc) => {
        if (!id) return;
        const cleanId = String(id).trim().toUpperCase();
        if (!/^[A-Z]{2,6}[-_]\d{2,4}$/.test(cleanId)) return;

        let name = String(rawName || '').trim();
        // 앞뒤 불필요한 번호, 특수기호, 줄바꿈 제거
        name = name
            .replace(/[\r\n\t]+/g, ' ')
            .replace(/^(?:[0-9]+[.)]|[-•*□■▶◆○●])\s*/, '')
            .replace(/\s+/g, ' ')
            .trim();

        let keywords = '';
        if (rawDesc) {
            const words = String(rawDesc)
                .replace(/[\r\n\t]+/g, ' ')
                .replace(/[()[\]{}.,~;:_!?"'`/\\|]/g, ' ')
                .split(/\s+/)
                .map(w => w.trim())
                .filter(w => w.length >= 2 && !['체계', '구축', '방안', '수립', '및', '관한', '위한', '지원', '제공', '등', '에', '의', '수행', '대상', '내용', '정의', '요구사항'].includes(w));
            keywords = Array.from(new Set(words)).slice(0, 5).join(', ');
        }

        // 표준 카탈로그 조회 (플레이스홀더인 경우 공공의료 AI ISP 표준 명칭 자동 복원)
        const stdCatalog = PUBLIC_HEALTHCARE_AI_ISP_REQUIREMENTS.find(r => r.id === cleanId);
        let resolvedName = name;
        let resolvedKeywords = keywords;
        if (isPlaceholderName(resolvedName, cleanId) && stdCatalog) {
            resolvedName = stdCatalog.name;
            if (!resolvedKeywords) resolvedKeywords = stdCatalog.keywords;
        }

        const existing = reqDict.get(cleanId);
        if (!existing) {
            reqDict.set(cleanId, {
                id: cleanId,
                name: !isPlaceholderName(resolvedName, cleanId) ? resolvedName : `요구사항 ${cleanId}`,
                keywords: resolvedKeywords || (resolvedName && !isPlaceholderName(resolvedName, cleanId) ? resolvedName.split(/\s+/).filter(w => w.length >= 2).slice(0, 4).join(', ') : '')
            });
        } else {
            // 기존 이름이 플레이스홀더인데 새 이름이 유효한 경우 교체!
            if (isPlaceholderName(existing.name, cleanId) && !isPlaceholderName(resolvedName, cleanId)) {
                existing.name = resolvedName;
                if (!existing.keywords && resolvedKeywords) existing.keywords = resolvedKeywords;
            } else if (!isPlaceholderName(resolvedName, cleanId) && resolvedName.length > existing.name.length && !existing.name.includes(resolvedName)) {
                // 더 구체적인 이름이 발견되면 업데이트
                existing.name = resolvedName;
                if (resolvedKeywords) existing.keywords = resolvedKeywords;
            } else if (!existing.keywords && resolvedKeywords) {
                existing.keywords = resolvedKeywords;
            }
        }
    };

    if (Array.isArray(tables)) {
        for (const tbl of tables) {
            const rows = tbl.rows || [];
            if (rows.length < 1) continue;

            // ── [패턴 1: 양식형 Key-Value 표 (공공 RFP 세부 요구사항 정의서 표)] ──
            // 한 표 안에 '요구사항 고유번호' 라벨 셀과 '요구사항명' 라벨 셀이 존재하는 경우
            let kvId = '';
            let kvName = '';
            let kvDesc = '';

            for (let r = 0; r < rows.length; r++) {
                const row = rows[r];
                for (let c = 0; c < row.length; c++) {
                    const cell = (row[c] || '').replace(/\s+/g, '');

                    // ID 라벨 탐색
                    if (cell.includes('요구사항고유번호') || cell.includes('요구사항번호') || cell.includes('요구사항id') || cell.includes('요구사항식별자')) {
                        // 같은 행 다음 셀 탐색
                        const nextCell = (row[c + 1] || '').trim();
                        const idMatch = nextCell.match(/[A-Z]{2,6}[-_]\d{2,4}/i);
                        if (idMatch) {
                            kvId = idMatch[0].toUpperCase();
                        } else if (rows[r + 1] && rows[r + 1][c]) {
                            // 아래 행 셀 탐색
                            const belowMatch = (rows[r + 1][c] || '').match(/[A-Z]{2,6}[-_]\d{2,4}/i);
                            if (belowMatch) kvId = belowMatch[0].toUpperCase();
                        }
                    }

                    // 요구사항명 라벨 탐색
                    if ((cell.includes('요구사항명') || cell.includes('과업명') || cell.includes('기능명')) && !cell.includes('내용') && !cell.includes('설명')) {
                        const nextCell = (row[c + 1] || '').trim();
                        if (nextCell && nextCell.length >= 3 && !nextCell.includes('고유번호')) {
                            kvName = nextCell;
                        } else if (rows[r + 1] && rows[r + 1][c]) {
                            const belowCell = (rows[r + 1][c] || '').trim();
                            if (belowCell && belowCell.length >= 3) kvName = belowCell;
                        }
                    }

                    // 세부내용 라벨 탐색
                    if (cell.includes('세부내용') || cell.includes('요구사항세부') || cell.includes('과업내용')) {
                        const nextCell = (row[c + 1] || '').trim();
                        if (nextCell && nextCell.length >= 5) {
                            kvDesc = nextCell;
                        } else if (rows[r + 1] && rows[r + 1][c]) {
                            const belowCell = (rows[r + 1][c] || '').trim();
                            if (belowCell && belowCell.length >= 5) kvDesc = belowCell;
                        }
                    }
                }
            }

            if (kvId && kvName) {
                addOrUpdateReq(kvId, kvName, kvDesc);
            }

            // ── [패턴 2: 그리드형 요구사항 총괄/목록 테이블] ──
            // 첫 1~3개 행 중에서 열 헤더 감지
            let bestHeaderRow = 0;
            let idCol = -1;
            let nameCol = -1;
            let descCol = -1;

            for (let r = 0; r < Math.min(4, rows.length); r++) {
                const headerRow = rows[r].map(c => (c || '').replace(/\s+/g, '').toLowerCase());
                const curId = headerRow.findIndex(h => h.includes('고유번호') || h.includes('요구사항id') || h.includes('요구사항번호') || h === 'id' || h.includes('식별자') || (h.includes('번호') && !h.includes('전화') && !h.includes('순번')));
                const curName = headerRow.findIndex(h => (h.includes('요구사항명') || h.includes('기능명') || h.includes('과업명') || h.includes('항목명') || h === '명칭' || h === '제목' || h === '과제명') && !h.includes('설명') && !h.includes('구분'));
                const curDesc = headerRow.findIndex(h => h.includes('내용') || h.includes('설명') || h.includes('정의') || h.includes('키워드') || h.includes('개요'));

                if (curId >= 0 || curName >= 0) {
                    bestHeaderRow = r;
                    idCol = curId;
                    nameCol = curName;
                    descCol = curDesc;
                    break;
                }
            }

            for (let r = bestHeaderRow + 1; r < rows.length; r++) {
                const row = rows[r];
                let rowReqId = '';
                let rowReqName = '';
                let rowReqDesc = '';

                // 1) idCol에서 ID 추출 시도
                if (idCol >= 0 && row[idCol]) {
                    const m = row[idCol].match(/[A-Z]{2,6}[-_]\d{2,4}/i);
                    if (m) rowReqId = m[0].toUpperCase();
                }

                // 2) 전체 셀에서 ID 탐색 (Fallback)
                let idCellIdx = -1;
                if (!rowReqId) {
                    for (let c = 0; c < row.length; c++) {
                        const m = (row[c] || '').match(/\b([A-Z]{2,6}[-_]\d{2,4})\b/i);
                        if (m) {
                            rowReqId = m[1].toUpperCase();
                            idCellIdx = c;
                            break;
                        }
                    }
                } else if (idCol >= 0) {
                    idCellIdx = idCol;
                }

                if (rowReqId) {
                    // nameCol에서 이름 추출 시도
                    if (nameCol >= 0 && row[nameCol]) {
                        const candidate = row[nameCol].trim();
                        if (candidate.length >= 3 && !candidate.match(/^[A-Z]{2,6}[-_]\d{2,4}$/i)) {
                            rowReqName = candidate;
                        }
                    }

                    // nameCol로 못 찾았거나 없으면, ID 셀 다음 셀(c+1) 또는 2번째 다음 셀(c+2) 탐색
                    if (!rowReqName && idCellIdx >= 0) {
                        for (let offset = 1; offset <= 3; offset++) {
                            const nextC = idCellIdx + offset;
                            if (nextC < row.length && row[nextC]) {
                                const candidate = row[nextC].trim();
                                // 길이가 3 이상이고, 단순 숫자/특수문자/구분자가 아닌 한글 포함 문자열
                                if (candidate.length >= 3 && /[가-힣]/.test(candidate) && !candidate.match(/^[A-Z]{2,6}[-_]\d{2,4}$/i) && !['공통', '기능', '보안', '품질', '운영', '필수'].includes(candidate)) {
                                    rowReqName = candidate;
                                    break;
                                }
                            }
                        }
                    }

                    // 설명 추출
                    if (descCol >= 0 && row[descCol]) {
                        rowReqDesc = row[descCol].trim();
                    }

                    addOrUpdateReq(rowReqId, rowReqName, rowReqDesc);
                }
            }
        }
    }

    // ── 3) 텍스트 본문 정규식 전수 탐색 및 보완 ──
    // 공공기관 RFP의 다양한 표기 패턴 전수 수용
    const regexPatterns = [
        // 1. [CNR-001] 요구사항명 또는 [CNR-001] : 요구사항명
        /\[([A-Z]{2,6}[-_]\d{2,4})\][ \t]*[:：\-–—./)\]]?[ \t]*([가-힣A-Za-z0-9][^\n\r]{2,80})/g,
        // 2. 글머리표 + CNR-001 요구사항명
        /(?:^|\n)[ \t]*(?:[□■▶◆○●*·-]|가\.|나\.|다\.|라\.|[0-9]+[.)])?[ \t]*([A-Z]{2,6}[-_]\d{2,4})[ \t]*[:：\-–—./)\]][ \t]*([가-힣A-Za-z0-9][^\n\r]{2,80})/g,
        // 3. 콜론 없는 공백형: CNR-001 공공의료 정보화 환경 및 법·제도 분석
        /(?:^|\n)[ \t]*(?:[□■▶◆○●*·-]|가\.|나\.|다\.|라\.|[0-9]+[.)])?[ \t]*([A-Z]{2,6}[-_]\d{2,4})[ \t]{1,6}([가-힣A-Za-z0-9][^\n\r]{2,80})/g,
        // 4. 요구사항ID : 요구사항명
        /(?:요구사항\s*(?:ID|번호|고유번호)?[ \t]*[:：|][ \t]*)([A-Z]{2,6}[-_]\d{2,4})[ \t\n\r]*(?:요구사항명[ \t]*[:：|][ \t]*)([가-힣A-Za-z0-9][^\n\r]{2,80})/g,
        // 5. 표 텍스트 변환 시 나타나는 줄바꿈 분리형: CNR-001\n요구사항명
        /(?:^|\n)[ \t]*([A-Z]{2,6}[-_]\d{2,4})[ \t]*\n[ \t]*([가-힣][^\n\r]{2,80})/g
    ];

    for (const pat of regexPatterns) {
        let m;
        while ((m = pat.exec(cleanText)) !== null) {
            const id = m[1].trim().toUpperCase();
            let rawName = m[2].trim();
            // 괄호 안의 부가설명이나 [필수] 태그 분리
            let kw = '';
            const bMatch = rawName.match(/\[([^\]]+)\]/);
            if (bMatch) {
                kw = bMatch[1].trim();
                rawName = rawName.replace(/\[[^\]]+\]/, '').trim();
            }
            addOrUpdateReq(id, rawName, kw);
        }
    }

    const requirements = Array.from(reqDict.values());

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

    const fullText = text || '';
    // 앞부분 8,000자 (사업개요, 발주처, 예산, 기간 집중)
    const introSnippet = fullText.slice(0, 8000);

    // 요구사항 ID(예: CNR-001) 또는 '요구사항' 키워드가 밀집된 본문 단락 지능적 수집
    const paragraphs = fullText.split(/\n{2,}/);
    const reqParagraphs = [];
    let reqChars = 0;

    for (const p of paragraphs) {
        if (/[A-Z]{2,6}[-_]\d{2,4}/.test(p) || (p.includes('요구사항') && (p.includes('기능') || p.includes('분석') || p.includes('체계') || p.includes('구축')))) {
            reqParagraphs.push(p);
            reqChars += p.length;
            if (reqChars > 40000) break; // 최대 40,000자
        }
    }

    const textSnippet = introSnippet + '\n\n... [중략: 요구사항 핵심 본문 발췌] ...\n\n' + reqParagraphs.join('\n\n');

    // 표 데이터도 요구사항 ID 패턴을 포함한 표들을 최우선 선별 (최대 20개)
    const relevantTables = (tables || []).filter(t => {
        const str = JSON.stringify(t.rows || []);
        return /[A-Z]{2,6}[-_]\d{2,4}/.test(str) || str.includes('요구사항') || str.includes('과업명') || str.includes('고유번호');
    });

    const tablesToUse = (relevantTables.length > 0 ? relevantTables : (tables || [])).slice(0, 20);

    const tableSnippets = tablesToUse.map((t, idx) => {
        const rowsStr = (t.rows || []).slice(0, 35).map(r => r.join(' | ')).join('\n');
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
   - RFP 내에 명시된 모든 요구사항 고유번호(ID: 예 CNR-001, SFR-001, PMR-001 등), 실제 요구사항명(과업명), 필수 핵심키워드(쉼표 구분)
   - ★ 주의: 요구사항명을 단순히 '요구사항 CNR-002'처럼 ID를 기재하지 마십시오. 반드시 제안요청서에 기재된 실제 기능명/과업명(예: '공공의료 정보화 환경 및 법·제도 분석')을 정확히 추출하십시오.
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
    { "id": "CNR-001", "name": "공공의료 정보화 환경 및 법·제도 분석", "keywords": "공공의료, 정보화, 법제도, 환경분석" }
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

        // 요구사항 병합: 플레이스홀더 여부 검사 후 실제 유효한 이름 우선 채택
        const isPlaceholder = (name, id) => {
            if (!name) return true;
            const str = String(name).trim();
            return str.length < 3 || str === `요구사항 ${id}` || str === id || /^요구사항\s*[A-Z]{2,6}[-_]\d{2,4}$/i.test(str);
        };

        const reqMap = new Map();
        // 1) 룰 기반 요구사항 먼저 세팅
        (ruleResult.requirements || []).forEach(r => {
            if (r.id) reqMap.set(r.id.toUpperCase(), r);
        });

        // 2) AI 요구사항으로 보강 (AI 이름이 유효하면 우선, AI가 플레이스홀더면 룰 결과 유지)
        (aiResult.requirements || []).forEach(aiReq => {
            if (!aiReq.id) return;
            const key = aiReq.id.toUpperCase();
            const existing = reqMap.get(key);

            if (!existing) {
                reqMap.set(key, {
                    id: key,
                    name: !isPlaceholder(aiReq.name, key) ? aiReq.name : `요구사항 ${key}`,
                    keywords: aiReq.keywords || ''
                });
            } else {
                // 기존(룰)이 플레이스홀더인데 AI가 유효한 이름을 가진 경우 AI로 교체
                if (isPlaceholder(existing.name, key) && !isPlaceholder(aiReq.name, key)) {
                    existing.name = aiReq.name;
                }
                // AI에 키워드가 더 많으면 키워드 보강
                if (!existing.keywords && aiReq.keywords) {
                    existing.keywords = aiReq.keywords;
                }
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
            ruleReqCount: ruleResult.requirements.length,
            ruleGlossaryCount: ruleResult.glossary.length
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
