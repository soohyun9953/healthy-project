import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { extractStructuredFileContent, getFileExtension } from './fileExtractor.js';
import {
    formatPageRanges, formatPageLabelRanges, normalizeForMatch, titleMatchesText,
    CHAPTER_TOPIC_STEMS, chapterTopicOf,
    buildDeckIndex, evaluateRequirements, mergeMatchingConfig, buildDeliverableLinks, gradeOfLink,
    evidenceLocationText, EVIDENCE_TYPE_LABELS, RTM_STATUS,
    parseCsvRows, parseSynonymCsv, parseManualEntries
} from './rtmEvidenceEngine.js';
import DEFAULT_RTM_MATCHING_CONFIG from '../config/rtmMatchingConfig.json';

export { DEFAULT_RTM_MATCHING_CONFIG };

export { formatPageRanges, formatPageLabelRanges };

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
        synonyms: '복지부, MOHW',
        forbidden: '보건부',
        note: '공식 명칭 및 통용 약칭 (복지부 통과 허용)'
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
        synonyms: 'NMC',
        forbidden: '중앙의료원, 국립의료원, 국립중앙병원',
        note: '발주기관 공식 명칭 (NMC 영문약어 통과 허용)'
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
        "id": "CNR-001",
        "category": "컨설팅 요구사항",
        "name": "공공의료 정보화 환경 및 현황 분석",
        "description": "공공의료 정보시스템* 관련 법령·정책, 지침, 제도 현황 검토 및 분석, 내외부 정보기술 환경을 종합 분석하고, 문제점 및 시사점 도출",
        "subItems": [
            "(법제도 분석) 지역·필수공공보건의료와 AI 관련 법, 정책, 지침 현황 검토 (국정과제, 공공보건의료계획 등 보건의료 분야 추진정책 환경 분석, 보건의료·데이터·정보기술·보안·개인정보보호 관련 법·제도 변화, 공공의료 AI정보화 추진시 고려할 제약사항 및 기회요인 도출)",
            "(환경분석) 현행 정책, 기술 환경과의 적합성 분석, 적용가능 범위 도출 (보건의료분야 국내외 공공 데이터 및 기술 생태계 동향 분석, 최신 IT 메가트렌드 및 핵심IT 기술의 공공의료 정보시스템 적용 가능성 분석, 데이터표준 기술·클라우드·데이터 파이프라인 자동화 등 인프라 기술 동향 및 적용 가능성 분석)",
            "(사례분석) 국내외 정부, 지자체, 공공기관 유사 ISP 및 구축 사례 분석 (국내외 AI기반 정보시스템 구축 사례를 수집하고 각 시스템별 장단점 분석, 국내외 AI기반 정보시스템 모범 사례 발굴 및 사업간 공통된 데이터 특징·구성요소 및 차별점 분석)"
        ],
        "deliverables": "Ⅱ.1. 환경분석 개요, Ⅱ.3. 법제도 분석",
        "keywords": "공공의료, 정보화, 환경, 현황, 법제도, 사례분석, 환경분석"
    },
    {
        "id": "CNR-002",
        "category": "컨설팅 요구사항",
        "name": "보건의료 AI 활용 동향 및 적용 가능성 분석",
        "description": "글로벌 및 국내 보건의료 분야 최신 AI/빅데이터 기술 활용 동향 및 공공의료 현장 적용 타당성 분석",
        "subItems": [
            "글로벌 보건의료 AI 적용 우수사례 및 최신 기술 트렌드(생성형 AI, CDSS 등) 조사",
            "공공병원 진료지원, 판독, 응급의료, 감염병 등 영역별 AI 적용 가능성 및 실효성 분석",
            "AI 알고리즘 도입에 따른 의료 윤리, 보안성, 설명가능성(XAI) 확보 방안 검토"
        ],
        "deliverables": "Ⅱ.5. 보건의료 AI 활용 동향 및 적용 가능성 분석",
        "keywords": "보건의료, AI, 활용, 동향, CDSS, 진료지원"
    },
    {
        "id": "CNR-003",
        "category": "컨설팅 요구사항",
        "name": "공공의료 정보시스템 업무/정보기술(IT) 현황 분석",
        "description": "전국 공공병원 및 지방의료원의 업무 프로세스, EMR/HIS 응용시스템, IT 인프라 현황 종합 실사 및 분석",
        "subItems": [
            "공공의료기관별 진료, 원무, 행정 업무 프로세스 및 운영 체계 실태 분석",
            "공공병원별 상이한 EMR/OCS/PACS 등 응용시스템 구조, 연계 규격 및 노후도 진단",
            "서버, 스토리지, 네트워크, 보안 인프라 현황 및 클라우드 전환 적합성 평가"
        ],
        "deliverables": "Ⅲ.2. 업무 현황분석, Ⅲ.3. 응용시스템 현황분석, Ⅲ.4. 인프라 현황분석",
        "keywords": "시스템, 구성, 기술, 환경, 분석, EMR, 인프라"
    },
    {
        "id": "CNR-004",
        "category": "컨설팅 요구사항",
        "name": "선진사례 벤치마킹 및 차이 분석",
        "description": "국내외 스마트병원 및 선진 공공보건의료 정보화 사례 분석을 통한 Gap 분석 및 핵심 성공 요인 도출",
        "subItems": [
            "국내외 대형 종합병원 및 선진 공공의료 네트워크 지능화 사례 벤치마킹",
            "현재 공공병원 정보화 수준과 선진 모델 간의 기능적·기술적·거버넌스 차이(Gap) 분석",
            "공공의료 지능형 플랫폼 구축을 위한 시사점 및 핵심 성공 요인(CSF) 도출"
        ],
        "deliverables": "Ⅲ.7. 선진사례 벤치마킹 및 차이 분석",
        "keywords": "선진사례, 벤치마킹, 차이, 분석, Gap, 시사점"
    },
    {
        "id": "CNR-005",
        "category": "컨설팅 요구사항",
        "name": "과업 분야별 업무 개선방향 및 개선과제 도출",
        "description": "현황 분석 및 Gap 분석 결과를 토대로 과업 분야별 병목/이슈 식별 및 구체적 업무 개선과제 Pool 정립",
        "subItems": [
            "진료정보교류, 전원 환자 관리, 만성질환 모니터링 등 핵심 업무 병목/이슈 분석",
            "현업 실무진 인터뷰 및 설문조사를 통한 개선 니즈 수렴 및 개선방향 수립",
            "단기·중장기 분야별 To-Be 개선과제(Task) 목록 도출 및 우선순위 평가"
        ],
        "deliverables": "Ⅲ.8. 분야별 이슈 및 개선과제 도출",
        "keywords": "과업, 분야별, 업무, 개선방향, 개선과제, 이슈"
    },
    {
        "id": "CNR-006",
        "category": "컨설팅 요구사항",
        "name": "공공의료 AI 정보화 비전 및 전략 수립",
        "description": "국가 공공의료 체계 강화를 위한 지능형 공공보건의료 미래 비전, 전략방향 및 추진목표 체계 수립",
        "subItems": [
            "공공의료 디지털 혁신을 위한 미래 비전 및 핵심 슬로건 정립",
            "비전 달성을 위한 3~4대 핵심 전략방향 및 세부 전략과제 도출",
            "단계별 정량적/정성적 성과지표(KPI) 및 전략적 지향점 설정"
        ],
        "deliverables": "Ⅳ.2. 정보화 비전 및 전략수립",
        "keywords": "공공의료, AI, 정보화, 비전, 전략, 추진목표"
    },
    {
        "id": "CNR-007",
        "category": "컨설팅 요구사항",
        "name": "AI 기반 공공의료 정보시스템 목표모델(To-Be) 설계",
        "description": "공공병원 및 책임의료기관 간 협진과 지능형 진료지원을 아우르는 통합 목표 시스템 모델(To-Be) 설계",
        "subItems": [
            "지능형 환자 관리 및 협진 지원 To-Be 목표 업무 프로세스(BPR) 정의",
            "공공의료 AI 플랫폼 목표 응용 아키텍처(Application Architecture) 설계",
            "외부 유관기관(복지부, 심평원, 질병청 등) 연계 및 확장 모델 수립"
        ],
        "deliverables": "Ⅳ.1. 목표모델 개요, Ⅳ.4. 목표 업무 아키텍처, Ⅳ.5. 목표 응용시스템 아키텍처",
        "keywords": "AI, 기반, 공공의료, 정보시스템, 목표모델, To-Be"
    },
    {
        "id": "CNR-008",
        "category": "컨설팅 요구사항",
        "name": "차세대 정보시스템/데이터 플랫폼 구축 전략 및 기술 아키텍쳐 설계",
        "description": "대용량 의료 빅데이터 처리와 AI 서빙을 지원하는 차세대 클라우드 기반 데이터 플랫폼 및 기술 아키텍처 설계",
        "subItems": [
            "민간·공공 하이브리드 클라우드 인프라 아키텍처(TA) 설계",
            "대용량 의료데이터 수집·적재를 위한 데이터 레이크 및 DW/CDW 아키텍처 수립",
            "마이크로서비스(MSA), 컨테이너(K8s) 기반 시스템 유연성 및 무중단 운영 체계 설계"
        ],
        "deliverables": "Ⅳ.6. 목표 데이터 아키텍처 설계, Ⅳ.7. 목표 인프라 및 정보보호 아키텍처 설계",
        "keywords": "차세대, 정보시스템, 데이터플랫폼, 구축, 기술아키텍쳐"
    },
    {
        "id": "CNR-009",
        "category": "컨설팅 요구사항",
        "name": "공공의료 AI 거버넌스 및 데이터 관리체계 설계",
        "description": "신뢰할 수 있는 공공의료 AI 서비스 운영을 위한 통합 거버넌스, 윤리 가이드라인 및 데이터 품질 관리 체계 설계",
        "subItems": [
            "공공보건의료 AI 추진 조직, 의사결정 체계 및 역할·책임(R&R) 정립",
            "AI 모델 생애주기 관리(MLOps) 및 데이터 정기 검증 체계 수립",
            "의료데이터 품질 관리, 메타데이터 표준화 및 비식별화 거버넌스 정의"
        ],
        "deliverables": "Ⅳ.12. 공공의료 정보화 통합 거버넌스",
        "keywords": "공공의료, AI, 거버넌스, 데이터, 관리체계, 품질"
    },
    {
        "id": "CNR-010",
        "category": "컨설팅 요구사항",
        "name": "공공병원 AI 서비스 확산 및 스마트병원 활성화 추진 전략",
        "description": "지역 공공병원 현장의 AI 도입 장벽을 해소하고 스마트병원 혁신 서비스를 확산하기 위한 단계별 추진 전략 수립",
        "subItems": [
            "공공병원 규모 및 특성(지방의료원, 적십자병원 등)별 맞춤형 AI 서비스 패키지 정의",
            "현장 의료진 및 행정인력 대상 AI 활용 교육 및 변화관리(Change Management) 프로그램 수립",
            "스마트병원 선도모델 실증 및 성과 확산 방안 수립"
        ],
        "deliverables": "Ⅳ.9. AI 서비스 확산 및 스마트병원 활성화 전략",
        "keywords": "공공병원, AI, 서비스, 확산, 스마트병원, 활성화"
    },
    {
        "id": "CNR-011",
        "category": "컨설팅 요구사항",
        "name": "재난·재해 대응, 개인정보보호 및 정보보안 강화 전략 수립",
        "description": "민감한 환자 의료데이터 보호 및 비상 상황 시 연속 진료 보장을 위한 재난복구(DR) 및 정보보안 체계 수립",
        "subItems": [
            "ISMS-P, 의료법 등 공공 정보보호 규정 준수를 위한 보안 아키텍처 설계",
            "랜섬웨어, 침해사고 대응 및 24/365 보안관제 체계 수립",
            "데이터 센터 이중화 및 재해복구(DR) 센터 구축 전략 수립"
        ],
        "deliverables": "Ⅳ.7. 목표 인프라 및 정보보호 아키텍처 설계",
        "keywords": "재난·재해, 대응, 개인정보보호, 정보보안, DR, ISMS-P"
    },
    {
        "id": "CNR-012",
        "category": "컨설팅 요구사항",
        "name": "공공병원 AX 단계별 추진 로드맵 수립",
        "description": "디지털 전환(AX) 목표 달성을 위한 단기(1단계), 중기(2단계), 장기(3단계) 종합 추진 로드맵 및 이행 일정 수립",
        "subItems": [
            "개선과제 간 선후행 관계 및 기술적 종속성을 고려한 단계별 로드맵 편성",
            "연차별 시범사업, 본구축, 전국 확산 일정계획 수립",
            "예산 투입 규모 및 조직 준비도를 반영한 마일스톤 정의"
        ],
        "deliverables": "Ⅴ.1. 이행계획 수립 개요, Ⅴ.5. 이행과제 로드맵",
        "keywords": "공공병원, AX, 단계별, 추진, 로드맵, 일정"
    },
    {
        "id": "CNR-013",
        "category": "컨설팅 요구사항",
        "name": "AI 기반 차세대 시스템 구축에 따른 기대효과 및 성과지표 도출",
        "description": "차세대 플랫폼 구축 후 공공의료 서비스 질 향상, 재정 절감, 국민 편익을 측정할 정량적·정성적 기대효과 및 성과지표 도출",
        "subItems": [
            "진료 대기시간 단축, 오진율 감소, 의료 질 향상 등 정량적 기대효과 산출",
            "지역 간 의료 격차 해소, 공공의료 신뢰도 제고 등 정성적 파급효과 분석",
            "사업 완료 후 사후 성과평가를 위한 핵심 성과지표(KPI) 및 측정 방법론 제시"
        ],
        "deliverables": "Ⅴ.7. 기대효과분석, Ⅴ.8. 투자효과분석",
        "keywords": "AI, 차세대, 시스템, 기대효과, 성과지표, ROI"
    },
    {
        "id": "CNR-014",
        "category": "컨설팅 요구사항",
        "name": "사업계획 수립 및 소요예산 산출, 개선과제 이행방안 수립",
        "description": "본 구축 사업 추진을 위한 세부 사업계획 수립 및 SW 개발비, H/W·S/W 도입비, 인건비 등 객관적 소요예산 산출",
        "subItems": [
            "SW 기능점수(FP) 분석 및 엔지니어링 표준품셈 기반 개발비 산정",
            "클라우드 인프라 이용료, 상용 SW 라이선스, AI 솔루션 도입 비용 산출",
            "연차별/비목별 재정 투자계획 및 비용 절감 방안 수립"
        ],
        "deliverables": "Ⅴ.2. 이행과제 도출, Ⅴ.6. 소요예산 정의",
        "keywords": "사업계획, 소요예산, 산출, 개선과제, 이행방안, 기능점수"
    },
    {
        "id": "CNR-015",
        "category": "컨설팅 요구사항",
        "name": "세부 사업계획서 및 제안요청서 작성",
        "description": "후속 본 구축 사업자 선정을 위한 공식 제안요청서(RFP) 및 세부 과업내용서 작성 지원",
        "subItems": [
            "정부 공식 서식에 맞춘 세부 사업계획서 및 제안요청서(RFP) 초안 작성",
            "기능적·기술적·관리적 세부 요구사항 정의서 및 RTM 작성",
            "사업자 선정 평가 기준표 및 기술제안서 평가 배점표 마련"
        ],
        "deliverables": "제안요청서(RFP) 산출물, 최종산출물 점검표",
        "keywords": "세부, 사업계획서, 제안요청서, RFP, 과업내용서"
    },
    {
        "id": "CNR-016",
        "category": "컨설팅 요구사항",
        "name": "공통기반 시스템 연계·활용 방안 검토",
        "description": "공통기반 시스템 연계·활용 방안 검토 마련 (챗봇, 국민비서, 범정부 생성형 AI 공통기반, AI 허브, 음성인식/합성, 데이터 분석, 모바일 신분증, Any-ID 연계)",
        "subItems": [
            "챗봇 구축에 최신 정보기술 도출 (민간 보안인증(CSAP)을 획득한 클라우드(SaaS) 연계·활용 방안 검토)",
            "국민비서 알림 서비스(행정안전부) 연계·활용 방안 검토",
            "국민비서 상담형서비스 챗봇(행정안전부) 연계·활용 방안 검토 (챗봇 서비스 중 공통 기능 및 지식베이스 기반을 제공하는 서비스)",
            "(LLM·AI 학습데이터 구축 시) 행정안전부는 민간의 여러 언어모델을 도입하여 각 기관이 다양한 AI서비스를 구현할 수 있는 시스템을 구축하였으므로 행정안전부 '범정부 생성형 AI 공통기반'을 활용하는 방안 검토",
            "(AI 학습데이터 구축 시) AI 기반 생성되는 학습데이터는 과학기술정보통신부 'AI 허브'에 연계하여 제공하여야 함.",
            "음성인식(STT) 및 음성합성(TTS) 구축 시 한국전자통신연구원 '공공 인공지능 오픈 API·DATA 서비스(aiopen.etri.re.kr)'에서 제공하는 음성지능 및 언어처리 관련 자원(기능 및 데이터)을 활용하는 방안을 검토",
            "음성인식 학습방안 마련 시 국립국어원 '말뭉치 통합 시스템(corpus.korean.go.kr)'에서 제공하는 기계어로 개발된 말뭉치를 활용하는 방안을 검토",
            "범정부 데이터 분석시스템 연계·활용 방안 검토 (정부부처, 지방자치단체를 대상으로 공공데이터와 민간데이터에 대한 분석서비스를 제공하는 빅데이터 플랫폼)",
            "본인확인(로그인, 비회원로그인, 계정찾기, 비밀번호 찾기 등)시 모바일 신분증 활용 방안 검토",
            "로그인 기능 및 사용자 인증이 필요한 경우, 정부 통합인증(Any-ID) 연계·활용 방안 검토"
        ],
        "deliverables": "Ⅱ.4. 정보기술 환경분석, Ⅲ.6. 목표모델 아키텍처 설계, Ⅳ.11. 공통기반 시스템 연계·활용 방안",
        "keywords": "공통기반, 챗봇, 국민비서, 생성형AI, AI허브, 음성인식, Any-ID, 모바일신분증, 빅데이터, CSAP"
    },
    {
        "id": "CNR-017",
        "category": "컨설팅 요구사항",
        "name": "OpenAPI 기능의 활용성 확장 및 타 시스템 간 연계",
        "description": "외부 의료기관 및 민간 헬스케어 서비스와의 원활한 데이터 교환을 위한 OpenAPI 표준 게이트웨이 및 연계 체계 수립",
        "subItems": [
            "RESTful API, HL7 FHIR 기반 표준 OpenAPI 규격 설계",
            "API 보안 인증(OAuth 2.0, API Key), 접근제어 및 트래픽 스로틀링 정책 수립",
            "외부 개발자 포털 및 개발 가이드라인 마련"
        ],
        "deliverables": "Ⅳ.11. 공통기반 시스템 연계·활용 방안",
        "keywords": "OpenAPI, 활용성, 확장, 시스템, 연계, FHIR"
    },
    {
        "id": "DAR-001",
        "category": "데이터 요구사항",
        "name": "데이터 표준화 방안 제시",
        "description": "공공의료 데이터의 일관성 및 상호운용성 확보를 위한 전사 표준단어, 표준용어, 도메인, 코드 체계 수립",
        "subItems": [
            "보건의료 전사 표준 단어사전, 표준 용어사전, 코드 정의서 정립",
            "국제 표준 의학용어(SNOMED-CT, LOINC, EDI) 매핑 방안 마련",
            "데이터 명명 규칙 및 도메인 분류 표준 정의"
        ],
        "deliverables": "Ⅳ.6. 목표 데이터 아키텍처 설계, 데이터 표준화 방안",
        "keywords": "데이터, 표준화, 단어사전, 도메인, 코드, 용어"
    },
    {
        "id": "DAR-002",
        "category": "데이터 요구사항",
        "name": "데이터 표준관리 방안 제시",
        "description": "수립된 데이터 표준의 지속적인 유지관리 및 신규 등록·변경 승인을 위한 데이터 표준관리 프로세스 수립",
        "subItems": [
            "데이터 표준 관리 조직 및 변경 심의 승인 절차 수립",
            "표준관리 도구(Meta Tool) 연계 및 정기 표준 준수율 점검 방안 제시",
            "표준 미준수 데이터 조기 식별 및 교정 프로세스 정의"
        ],
        "deliverables": "데이터 표준관리 현황 분석서, 목표 데이터 아키텍처",
        "keywords": "데이터, 표준관리, 메타데이터, 승인절차, 준수율"
    },
    {
        "id": "DAR-003",
        "category": "데이터 요구사항",
        "name": "데이터 구조 설계 방안 제시",
        "description": "의료 빅데이터 분석 및 AI 학습 파이프라인에 최적화된 개념/논리/물리 데이터 모델(ERD) 설계 방안 수립",
        "subItems": [
            "공공의료 지능형 플랫폼 목표 개념 및 논리 데이터 모델(ERD) 설계",
            "시계열 생체신호, 비정형 영상(DICOM) 데이터 수집 및 모델링 방안 마련",
            "데이터 정규화 및 분석 성능 최적화를 위한 반정규화 기준 수립"
        ],
        "deliverables": "Ⅳ.6. 목표 데이터 아키텍처 설계, 데이터 구조 설계 방안",
        "keywords": "데이터, 구조, 설계, ERD, 논리모델, 물리모델"
    },
    {
        "id": "DAR-004",
        "category": "데이터 요구사항",
        "name": "데이터 구조 검증 및 개선방안 제시",
        "description": "현행 공공병원 EMR 데이터 구조의 정합성, 중복성, 무결성 제약조건 검증 및 개선 방안 수립",
        "subItems": [
            "현행 DB 구조 역공학 분석 및 무결성 제약 누락 요소 식별",
            "테이블/컬럼 단위 중복 및 데이터 모델 불일치 개선안 도출",
            "대용량 파티셔닝 및 인덱스 최적화 방안 제시"
        ],
        "deliverables": "데이터 구조현황 분석서, 목표 데이터 아키텍처",
        "keywords": "데이터, 구조, 검증, 개선, 무결성, 역공학"
    },
    {
        "id": "DAR-005",
        "category": "데이터 요구사항",
        "name": "데이터 구조 관리 방안 제시",
        "description": "데이터 모델 형상관리 및 스키마 변경 이력 추적을 위한 구조 관리 절차 수립",
        "subItems": [
            "DB 스키마 변경 시 영향도 분석 프로세스 수립",
            "데이터 모델링 툴 형상관리 및 버전 동기화 체계 정의",
            "데이터 딕셔너리 정기 최신화 절차 수립"
        ],
        "deliverables": "데이터 구조 관리방안, 정보화 통합 거버넌스",
        "keywords": "데이터, 구조, 관리, 스키마, 형상관리, 변경이력"
    },
    {
        "id": "DAR-006",
        "category": "데이터 요구사항",
        "name": "데이터 값 검증 방안 제시",
        "description": "수집 및 적재되는 공공의료 데이터의 유효성, 완전성, 정확성을 보장하기 위한 데이터 값 품질 진단 방안 수립",
        "subItems": [
            "필수값 누락(Null), 유효 범위 초과, 포맷 불일치 등 데이터 프로파일링 기법 정의",
            "데이터 정제(Cleansing) 룰셋 및 이상치 탐지 알고리즘 적용 방안",
            "주기적 데이터 품질 지수(DQ Index) 측정 체계 수립"
        ],
        "deliverables": "데이터(표준, 구조, 값, 연계)관리방안",
        "keywords": "데이터, 값, 검증, 품질진단, 프로파일링, 정제"
    },
    {
        "id": "DAR-007",
        "category": "데이터 요구사항",
        "name": "이관 데이터 값 검증 방안 제시",
        "description": "레거시 EMR 시스템에서 신규 플랫폼으로의 의료데이터 이관(Migration) 시 데이터 정합성 검증 방안 수립",
        "subItems": [
            "원천 DB와 목표 DB 간 레코드 건수 및 핵심 수치 대조 검증 룰 수립",
            "환자 고유식별자 매핑 및 이력 데이터 정합성 검증 절차 수립",
            "이관 실패 데이터(Exception) 격리 및 재이관 프로세스 마련"
        ],
        "deliverables": "데이터 관리방안, 이행계획 수립서",
        "keywords": "이관, 데이터, 검증, 정합성, 마이그레이션, 무결성"
    },
    {
        "id": "DAR-008",
        "category": "데이터 요구사항",
        "name": "데이터 관리체계 현황분석 및 목표모델 제시",
        "description": "현재 공공의료 데이터 관리 체계의 성숙도를 진단하고 차세대 지능형 데이터 레이크 목표 모델 설계",
        "subItems": [
            "공공의료 데이터 생애주기(생성, 저장, 활용, 폐기) 현황 진단",
            "의료데이터 거버넌스 성숙도 평가 및 개선 목표 수립",
            "데이터 카탈로그 및 셀프서비스 분석 환경 목표 모델 제시"
        ],
        "deliverables": "Ⅲ.5. 데이터 현황분석, Ⅳ.6. 목표 데이터 아키텍처",
        "keywords": "데이터, 관리체계, 현황분석, 목표모델, 성숙도"
    },
    {
        "id": "DAR-009",
        "category": "데이터 요구사항",
        "name": "데이터 개방 및 메타데이터 관리체계",
        "description": "공공데이터 개방 규정에 따른 가명정보 결합, 비식별 조치 및 공공포털 메타데이터 연동 체계 구축",
        "subItems": [
            "보건의료 가명정보 처리 가이드라인에 따른 비식별화 파이프라인 설계",
            "공공데이터포털(data.go.kr) 메타데이터 자동 연계 규격 수립",
            "연구자용 안심존(Safe Haven) 데이터 제공 절차 정의"
        ],
        "deliverables": "데이터 개방 및 메타데이터 관리방안",
        "keywords": "데이터, 개방, 메타데이터, 가명정보, 비식별화"
    },
    {
        "id": "DAR-010",
        "category": "데이터 요구사항",
        "name": "보건의료데이터 표준화",
        "description": "전국 의료기관 간 진료정보의 원활한 상호연계를 위한 보건의료 핵심 용어 및 FHIR 표준화 규격 수립",
        "subItems": [
            "보건의료 용어 표준(KOSTOM) 및 국제 진단코드(KCD-8) 표준 매핑",
            "HL7 FHIR 기반 임상 데이터 리소스(Patient, Encounter, Condition 등) 표준 프로파일 정의",
            "지방의료원 레거시 용어의 중앙 표준 용어 자동 변환 매핑 테이블 구축"
        ],
        "deliverables": "목표시스템 데이터 표준화 방안, Ⅳ.6. 목표 데이터 아키텍처",
        "keywords": "보건의료데이터, 표준화, KOSTOM, FHIR, 용어표준"
    },
    {
        "id": "SER-001",
        "category": "보안 요구사항",
        "name": "보안관리 준수사항",
        "description": "국가정보보안기본지침, 전자정부 정보보호 가이드라인 및 의료법상 기술적·관리적 보안 대책 수립",
        "subItems": [
            "행정안전부 및 국가정보원 정보보안 기본지침 준수 체계 수립",
            "서버, 네트워크, DB 접근통제 및 비밀번호 정책, 세션 관리 지침 수립",
            "시스템 취약점 분석·평가 및 침투 테스트 수행 방안 마련"
        ],
        "deliverables": "보안계획서 및 관련 문서, Ⅳ.7. 목표 인프라 및 정보보호 아키텍처",
        "keywords": "보안관리, 준수사항, 보안가이드라인, 접근통제"
    },
    {
        "id": "SER-002",
        "category": "보안 요구사항",
        "name": "누출금지 대상정보",
        "description": "국가계약법상 누출금지 대상정보 정의 및 사업 수행 중 정보 유출 방지를 위한 물리적·기술적 통제 방안 수립",
        "subItems": [
            "공공병원 네트워크 구성도, 시스템 IP, 계정 정보 등 누출금지 대상 정보 정의",
            "환자 개인정보 및 전자의무기록(EMR) 데이터 무단 반출 차단 체계 수립",
            "개발 사업장 출입 통제, PC 매체 제어 및 보안서약서 징구 방안 마련"
        ],
        "deliverables": "보안계획서(외부자 보안서약서 등)",
        "keywords": "누출금지, 대상정보, 정보누출, 개인정보유출, 보안서약서"
    },
    {
        "id": "SER-003",
        "category": "보안 요구사항",
        "name": "사업자 보안위규 처리기준",
        "description": "사업 수행 인력의 보안 위규 행위 적발 시 경중별 제재 조치 및 처리 기준 정립",
        "subItems": [
            "보안 위규 사항 수준별(심각, 중대, 경미) 분류 및 벌점 기준 수립",
            "위규자 즉시 교체 및 재발방지 대책 보고 절차 마련",
            "주기적 참여인력 보안 교육 및 불시 보안점검 실시 방안"
        ],
        "deliverables": "보안계획서 및 착수계 관련 서식",
        "keywords": "사업자, 보안위규, 처리기준, 보안점검, 위규자제재"
    },
    {
        "id": "SER-004",
        "category": "보안 요구사항",
        "name": "사업자 보안위약금 부과기준",
        "description": "정보 유출 및 보안 위규 행위에 따른 객관적 보안위약금 부과 기준 및 감액 기준 수립",
        "subItems": [
            "위규 수준에 따른 계약금액 대비 위약금 산정 비율표 정립",
            "보안 사고 발생 시 손해배상 책임 및 행정 제재 기준 마련",
            "기성금 지급 시 보안 위약금 공제 처리 절차 수립"
        ],
        "deliverables": "보안계획서 및 관련 문서",
        "keywords": "사업자, 보안위약금, 부과기준, 손해배상"
    },
    {
        "id": "QUR-001",
        "category": "품질 요구사항",
        "name": "품질관리 준수",
        "description": "SW공학 및 감리 기준에 부합하는 분석·설계 산출물 품질보증(QA) 계획 및 결함 관리 절차 수립",
        "subItems": [
            "산출물 작성 표준 가이드라인 및 피어 리뷰(Peer Review) 체계 수립",
            "단계별 품질 점검표 및 산출물 정합성(SSOT) 검증 절차 적용",
            "공식 감리 지적사항에 대한 신속한 조치결과 이행 관리"
        ],
        "deliverables": "ISP 최종산출물 검토신청서, 최종산출물 점검표",
        "keywords": "품질관리, 준수, QA, 감리, 결함관리, 피어리뷰"
    },
    {
        "id": "COR-001",
        "category": "제약사항",
        "name": "정보화기반 표준화 지침 및 가이드라인 준수",
        "description": "전자정부 표준프레임워크, 행정기관 공공기관 정보시스템 구축·운영 지침 및 기술적용기준 준수",
        "subItems": [
            "전자정부 표준프레임워크 최신 버전 준수 및 오픈소스 라이선스 검토",
            "클라우드 보안인증(CSAP) 획득 인프라 활용 원칙 준수",
            "행정기관 및 공공기관 정보시스템 구축·운영 기술적용계획표 작성 및 준수"
        ],
        "deliverables": "[별지 1] 기술적용계획표 및 기술적용결과표",
        "keywords": "정보화기반, 표준화, 지침, 가이드라인, 전자정부, 기술적용"
    },
    {
        "id": "COR-002",
        "category": "제약사항",
        "name": "저작권 및 지식재산권",
        "description": "사업 수행 중 생성되는 모든 분석·설계 산출물의 지식재산권 귀속 및 오픈소스 저작권 침해 방지 대책 수립",
        "subItems": [
            "기획재정부 계약예규에 따른 산출물 지식재산권 공동 귀속 원칙 정의",
            "외부 폰트, 이미지, 오픈소스 라이선스 침해 여부 전수 검증",
            "산출물의 무단 전재, 배포 방지 및 기술자료 임치 방안"
        ],
        "deliverables": "최종산출물 점검표, 계약 관련 서류",
        "keywords": "저작권, 지식재산권, 오픈소스, 라이선스, 계약예규"
    },
    {
        "id": "COR-003",
        "category": "제약사항",
        "name": "정보시스템 운영 성과측정 기능활용도 측정 방안 마련",
        "description": "행정안전부 정보시스템 운영 성과관리 지침에 따른 기능 활용도 및 사용률 측정 방안 마련",
        "subItems": [
            "목표 시스템 기능별 호출 빈도 및 사용자 접속 로그 수집 아키텍처 설계",
            "사용 빈도가 낮은 미활용 기능에 대한 지속적 모니터링 체계 수립",
            "연차별 운영 성과측정 보고서 작성 체계 정의"
        ],
        "deliverables": "Ⅳ.5. 목표 응용시스템 아키텍처, Ⅴ.5. 이행과제 로드맵",
        "keywords": "정보시스템, 운영, 성과측정, 기능활용도, 성과관리"
    },
    {
        "id": "PMR-001",
        "category": "프로젝트관리",
        "name": "업무보고",
        "description": "사업 추진 과정의 투명한 진척 관리를 위한 정기(주간/월간) 및 비정기(착수/중간/완료/수시) 업무보고 체계 수립",
        "subItems": [
            "착수보고회, 중간보고회, 최종 완료보고회 계획 및 심의 절차 수립",
            "주간/월간 정기 진척 보고서 작성 및 계획 대비 실적 갭(Gap) 관리",
            "주요 이슈 및 리스크 발생 시 즉각적인 수시보고 체계 마련"
        ],
        "deliverables": "착수/중간/최종 보고서, 주간/월간 보고서",
        "keywords": "업무보고, 주간보고, 월간보고, 착수보고, 완료보고"
    },
    {
        "id": "PMR-002",
        "category": "프로젝트관리",
        "name": "프로젝트 일정관리 방안 제시",
        "description": "WBS 기반의 상세 일정 계획 수립, 임계경로(Critical Path) 관리 및 지연 방지 대책 마련",
        "subItems": [
            "산출물 단위 세부 작업분할구조도(WBS) 작성 및 마일스톤 관리",
            "일정 지연 위험 감지 시 자원 추가 투입 및 일정 압축(Fast-tracking) 방안",
            "발주처 및 감리 일정과의 긴밀한 동기화 관리"
        ],
        "deliverables": "Ⅴ.1. 이행계획 수립 개요, 착수계(수행계획서)",
        "keywords": "프로젝트, 일정관리, WBS, 마일스톤, 공정관리"
    },
    {
        "id": "PMR-003",
        "category": "프로젝트관리",
        "name": "사업수행 산출물 작성 및 제출",
        "description": "단계별 요구되는 공식 산출물의 작성 기준, 제출 일정 및 발주처 승인 절차 수립",
        "subItems": [
            "환경분석, 현황분석, 목표모델, 이행계획 등 34대 표준 산출물 목록 정의",
            "산출물 양식 표준화 및 전자문서(PPTX/HWPX/PDF) 포맷 준수",
            "발주처 검토 의견서 피드백 반영 및 최종 승인 징구"
        ],
        "deliverables": "ISP 최종산출물 검토신청서, 34대 표준 산출물 일체",
        "keywords": "사업수행, 산출물, 작성, 제출, 산출물점검, 승인절차"
    },
    {
        "id": "PMR-004",
        "category": "프로젝트관리",
        "name": "프로젝트 조직구성(인력관리 등)",
        "description": "전문 컨설턴트 및 기술 아키텍트 중심의 전담 투입 조직 구성 및 인력 운영·교체 관리 방안 수립",
        "subItems": [
            "총괄 PM, 도메인 컨설팅팀, TA/DA 기술팀, 품질보증팀 간 R&R 정의",
            "투입 인력의 자격 요건 검증 및 보안/서약 관리",
            "불가피한 인력 교체 시 동급 이상 대체 인력 투입 및 인수인계 계획"
        ],
        "deliverables": "착수계(사업수행계획서 내 투입조직도)",
        "keywords": "프로젝트, 조직구성, 인력관리, 투입인력, R&R"
    },
    {
        "id": "PSR-001",
        "category": "지원 요구사항",
        "name": "관련 기술 이전 및 지원",
        "description": "ISP 수립 결과물 및 AI To-Be 모델에 대한 발주처 실무진 역량 강화를 위한 기술 이전 및 교육 지원",
        "subItems": [
            "공공의료 AI 플랫폼 설계 원리 및 아키텍처 기술 이전 교육 실시",
            "발주기관 시스템 담당자 대상 관리자 교육 및 기술자료 이관",
            "본 사업 발주 시 질의응답 및 기술 자문 지원"
        ],
        "deliverables": "기술지원 계획서, 기술이전 보고서",
        "keywords": "기술이전, 지원, 교육, 역량강화, 기술자문"
    },
    {
        "id": "PSR-002",
        "category": "지원 요구사항",
        "name": "전산장비 및 각종 사무기기 확보",
        "description": "사업 수행 인력의 원활한 업무 환경 조성을 위한 독립된 작업 공간, 보안 PC 및 집기류 확보 계획 수립",
        "subItems": [
            "사업 수행 장소(발주처 인접 공간) 확보 및 출입통제 체계 수립",
            "인터넷망/업무망 분리 개발 PC 및 소프트웨어 정품 라이선스 구비",
            "보안 프린터, 문서 파쇄기 등 보안 사무기기 설치 및 운영"
        ],
        "deliverables": "착수계(수행환경 구축 보고)",
        "keywords": "전산장비, 사무기기, 작업공간, 보안PC, 사무환경"
    },
    {
        "id": "PSR-003",
        "category": "지원 요구사항",
        "name": "하자보수 일반",
        "description": "ISP 완료 후 제안요청서 작성 및 본 사업 발주 과정에서 발견되는 산출물 오류에 대한 무상 하자보수 지원",
        "subItems": [
            "사업 종료 후 1년간 산출물 오탈자 및 오류에 대한 무상 정정 지원",
            "후속 본 구축 사업자 입찰 시 RFP 관련 질의에 대한 기술적 답변 지원",
            "정부 예산 심의 시 산출 근거 추가 소명 자료 작성 지원"
        ],
        "deliverables": "하자보수 계획서, 사업완료 확인서",
        "keywords": "하자보수, 일반, 무상보수, 질의대응, 사후지원"
    },
    {
        "id": "PSR-004",
        "category": "지원 요구사항",
        "name": "행사(보고회) 및 워크숍 지원",
        "description": "이해관계자 의견 수렴 및 공감대 형성을 위한 대외 보고회, 자문회의, 전문가 워크숍 기획 및 운영 지원",
        "subItems": [
            "전국 지방의료원장 및 실무진 초청 의견수렴 워크숍 개최 지원",
            "의료 AI 전문가 자문위원회 구성 및 정기 회의체 운영",
            "보고회 발표 자료(PPT) 및 홍보 리플릿 제작 지원"
        ],
        "deliverables": "워크숍 결과보고서, 전문가 자문의견서",
        "keywords": "행사, 보고회, 워크숍, 자문회의, 의견수렴"
    },
    {
        "id": "PSR-005",
        "category": "지원 요구사항",
        "name": "추가제안",
        "description": "본 제안요청서 범위 외에 공공의료 AI 플랫폼의 성공적 구축과 혁신을 위해 제안사가 추가 제시하는 창의적 제안",
        "subItems": [
            "민간 AI 파운데이션 모델 활용 및 공공-민간 협력 방안 추가 제시",
            "지방의료원 의료진 소진 방지를 위한 의료 행정 자동화(RPA) 도입 방안 제안",
            "추가 제안 과제의 타당성, 소요 예산 및 기대효과 명시"
        ],
        "deliverables": "추가제안 기술서, Ⅳ.9. 스마트병원 활성화 전략",
        "keywords": "추가제안, 창의제안, 부가가치, 민관협력, 혁신"
    }
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

            let loadedGlossary = Array.isArray(parsed.glossary) && parsed.glossary.length > 0 ? parsed.glossary : DEFAULT_GLOSSARY;
            // 허용 동의어(synonyms) 필드 보장 및 기존 데이터 보정 (보건복지부 -> 복지부 허용, 국립중앙의료원 -> NMC 허용)
            loadedGlossary = loadedGlossary.map(g => {
                let syn = (g.synonyms || '').trim();
                let forb = (g.forbidden || '').trim();
                const app = (g.approved || '').trim();

                if (app === '보건복지부') {
                    if (forb.includes('복지부')) {
                        forb = forb.replace(/복지부/g, '').replace(/,\s*,/g, ',').replace(/^,\s*|,\s*$/g, '').trim();
                    }
                    if (!syn.includes('복지부')) {
                        syn = syn ? `${syn}, 복지부` : '복지부, MOHW';
                    }
                } else if (app === '국립중앙의료원') {
                    if (forb.includes('NMC')) {
                        forb = forb.replace(/NMC/g, '').replace(/,\s*,/g, ',').replace(/^,\s*|,\s*$/g, '').trim();
                    }
                    if (!syn.includes('NMC')) {
                        syn = syn ? `${syn}, NMC` : 'NMC';
                    }
                }

                return {
                    ...g,
                    synonyms: syn,
                    forbidden: forb
                };
            });

            return {
                metadata: { ...DEFAULT_PROJECT_METADATA, ...(parsed.metadata || {}) },
                glossary: loadedGlossary,
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

    // A. 1) 금지 동의어 / 오표기 검출 (오탐 0건 보장 로직 - 허용 동의어는 절대 오류 제외)
    // 전역 허용 동의어/약어 화이트리스트 사전 구축
    const allAllowedSynonyms = new Set();
    glossary.forEach(item => {
        const syns = (item.synonyms || '').split(',').map(s => s.trim()).filter(Boolean);
        syns.forEach(s => allAllowedSynonyms.add(s.toUpperCase()));
    });

    glossary.forEach(item => {
        const approved = (item.approved || '').trim();
        const itemSynonyms = (item.synonyms || '')
            .split(',')
            .map(s => s.trim())
            .filter(Boolean);

        // 허용 동의어(synonyms)에 등록된 단어는 금지어 목록에 있더라도 무조건 제외
        const forbiddenList = (item.forbidden || '')
            .split(',')
            .map(s => s.trim())
            .filter(s => s.length > 0 && s !== approved && !itemSynonyms.includes(s) && !allAllowedSynonyms.has(s.toUpperCase()));

        forbiddenList.forEach(forbiddenTerm => {
            if (forbiddenTerm.length < 2) return;
            // 허용 동의어로 등록된 단어는 절대 결함으로 판정하지 않음 (스킵)
            if (itemSynonyms.includes(forbiddenTerm) || allAllowedSynonyms.has(forbiddenTerm.toUpperCase())) return;

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

    const masterIdsByPage = {}; // { [page]: Set<요구사항ID> }
    requirements.forEach(req => {
        const cleanId = (req.id || '').trim();
        if (!cleanId) return;
        const upperId = cleanId.toUpperCase();

        Object.entries(pageMap || {}).forEach(([pageNum, lines]) => {
            const pageText = (Array.isArray(lines) ? lines.join(' ') : String(lines)).toUpperCase();
            if (pageText.includes(upperId)) {
                matchedIdsInDoc.add(cleanId);
                addLocation(cleanId, pageNum);
                const p = parseInt(pageNum, 10) || 1;
                if (!masterIdsByPage[p]) masterIdsByPage[p] = new Set();
                masterIdsByPage[p].add(cleanId);
            }
        });
    });
    // 요구사항 ID를 여러 개(4개 이상) 나열한 슬라이드(요구사항 매핑표·추적표 등)는 ID가 언급만 된 것이므로
    // 요구사항추적표의 '작성 근거 페이지'로 보지 않는다
    const ID_LISTING_THRESHOLD = 4;
    const idListingPages = new Set(
        Object.entries(masterIdsByPage).filter(([, ids]) => ids.size >= ID_LISTING_THRESHOLD).map(([p]) => parseInt(p, 10))
    );

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
            evidencePages: pages.filter(p => !idListingPages.has(p)),
            evidenceLabels: pageLabels.filter((_, i) => !idListingPages.has(pages[i])),
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
    // B-1-2) 세부 요구사항 기준 RTM 매트릭스 계산 (subItems 기준 전개 평가)
    const detailedReqs = flattenDetailedRequirements(requirements);
    const subRtmMatrix = [];
    detailedReqs.forEach(subReq => {
        const cleanId = (subReq.id || '').trim();
        const parentId = (subReq.parentReqId || '').trim();
        const isIdPresent = matchedIdsInDoc.has(cleanId) || matchedIdsInDoc.has(cleanId.toUpperCase());
        const isParentPresent = matchedIdsInDoc.has(parentId) || matchedIdsInDoc.has(parentId.toUpperCase());

        const kwList = (subReq.keywords || '').split(',').map(k => k.trim()).filter(Boolean);
        const matchedKws = kwList.filter(k => text.includes(k));
        const isKwHit = matchedKws.length > 0;

        const isPresent = isIdPresent || (isParentPresent && isKwHit) || (isKwHit && matchedKws.length >= 2);
        const pages = isIdPresent
            ? (idLocations[cleanId] || idLocations[cleanId.toUpperCase()] || [])
            : (isParentPresent ? (idLocations[parentId] || idLocations[parentId.toUpperCase()] || []) : []);

        const status = isPresent
            ? (matchedKws.length >= Math.min(2, Math.max(1, kwList.length)) ? '충족' : '부분충족')
            : (inspectionScope === 'unit' ? '미출현 (단위산출물 범위 외)' : '미충족');

        subRtmMatrix.push({
            id: cleanId,
            parentReqId: parentId,
            parentName: subReq.parentName || '',
            category: subReq.category || '',
            name: subReq.name || '',
            keywords: subReq.keywords || '',
            isPresent,
            pages,
            status,
            matchedKeywords: matchedKws
        });
    });

    const rtmMappingRate = totalMasterReqCount > 0 
        ? Math.round((coveredReqCount / totalMasterReqCount) * 100) 
        : 100;

    return {
        issues,
        rtmMatrix,
        subRtmMatrix,
        rtmMappingRate,
        coveredReqCount,
        totalMasterReqCount,
        totalDetailedReqCount: detailedReqs.length,
        coveredDetailedCount: subRtmMatrix.filter(s => s.isPresent).length
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

// 요구사항 근거 판정 결과(evaluateRequirements) → 일괄 검수 프로젝트 RTM 행
function buildEvidenceRtmRows(evidenceResults = [], requirements = []) {
    const keywordsById = new Map(requirements.map(r => [String(r.id).toUpperCase(), r.keywords || '']));
    return evidenceResults.map(r => {
        const pagesByFile = {};
        const pageLabelsByFile = {};
        r.evidences.forEach(e => {
            if (!pagesByFile[e.fileName]) { pagesByFile[e.fileName] = []; pageLabelsByFile[e.fileName] = []; }
            pagesByFile[e.fileName].push(...e.pages);
            pageLabelsByFile[e.fileName].push(...e.labels);
        });
        return {
            id: r.id,
            name: r.name,
            keywords: keywordsById.get(String(r.id).toUpperCase()) || '',
            appearingFiles: r.files,
            pagesByFile,
            pageLabelsByFile,
            isPresentOverall: r.files.length > 0,
            overallStatus: r.status,
            evidenceTypes: r.typeLabels,
            locations: r.evidences.map(e => `[${EVIDENCE_TYPE_LABELS[e.type]}] ${evidenceLocationText(e)}`),
            remarks: [
                ...(r.missingOutputs.length ? [`미대응 산출정보: ${r.missingOutputs.join(', ')}`] : []),
                ...r.warnings.map(w => `⚠ ${w}`),
                ...r.notes
            ],
            missingOutputs: r.missingOutputs,
            pageText: r.pageText
        };
    });
}

/**
 * 다중 산출물 일괄 검수 결과 프로젝트 레벨 통합 집계 엔진
 * - evidenceResults(요구사항추적표와 같은 근거 판정 결과)가 주어지면 RTM·충족률·누락 결함을 그 판정으로 집계
 * - 개별 파일 단위 가짜 Missing ID 오탐 제거
 * - 전체 44개 마스터 요구사항 대비 프로젝트 종합 RTM 충족률 계산 (어느 파일에도 없는 것만 '전수 누락' 지적)
 * - 산출물 간 정합성(Cross-Document Consistency) 분석 (용어 표기 불일치, 유령 ID 확산)
 */
export function aggregateBatchSSOT(batchResults = [], ssotData = {}, evidenceResults = null) {
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

    const useEvidence = Array.isArray(evidenceResults) && evidenceResults.length > 0;
    const projectRtmMatrix = useEvidence
        ? buildEvidenceRtmRows(evidenceResults, requirements)
        : Array.from(rtmMap.values()).map(r => {
            const isPresentOverall = r.appearingFiles.length > 0;
            let overallStatus = '미출현 (전체 파일 누락)';
            if (isPresentOverall) {
                overallStatus = r.anyKeywordMissing ? '키워드 일부 누락' : '정상 반영';
            }
            return { ...r, isPresentOverall, overallStatus };
        });

    const statusCount = (st) => projectRtmMatrix.filter(r => r.overallStatus === st).length;
    const coveredReqs = useEvidence ? projectRtmMatrix.filter(r => r.overallStatus === RTM_STATUS.MET) : projectRtmMatrix.filter(r => r.isPresentOverall);
    const missingReqs = useEvidence ? projectRtmMatrix.filter(r => r.overallStatus === RTM_STATUS.UNMET) : projectRtmMatrix.filter(r => !r.isPresentOverall);
    const projectRtmRate = requirements.length > 0 ? Math.round((coveredReqs.length / requirements.length) * 100) : 0;

    // 2) 프로젝트 전체 누락 결함 (파일별 중복 오탐 없이 요구사항당 1회)
    // 근거 판정 모드: 미충족 → HIGH, 부분충족 → MEDIUM(빠진 산출정보 명시). 확인필요·사업관리 문서 확인필요는 결함이 아니라 확인 대상
    const projectMissingIssues = useEvidence
        ? [
            ...missingReqs.map(req => ({
                id: `PROJ-MIS-${req.id}`,
                stage: '프로젝트 통합 RTM',
                category: '요구사항 누락',
                errorType: '요구사항 미충족 (근거 없음)',
                severity: 'HIGH',
                page: '-',
                fileName: '전체 산출물',
                matchedTerm: `[${req.id}] ${req.name}`,
                originalText: `(대응표·ID·절 제목·동의어 의미 매칭 어디에서도 [${req.id}] 근거 미발견)`,
                suggestedText: `요구사항 [${req.id}] ${req.name} 관련 내용 작성 또는 근거 위치 수동 확정`,
                reason: `RFP 마스터 요구사항 [${req.id}]의 작성 근거를 ${batchResults.length}개 산출물에서 찾지 못했습니다.`
            })),
            ...projectRtmMatrix.filter(r => r.overallStatus === RTM_STATUS.PARTIAL).map(req => ({
                id: `PROJ-PART-${req.id}`,
                stage: '프로젝트 통합 RTM',
                category: '요구사항 불완전',
                errorType: '산출정보 일부 미대응 (부분충족)',
                severity: 'MEDIUM',
                page: req.pageText || '-',
                fileName: req.appearingFiles.join(', ') || '전체 산출물',
                matchedTerm: `[${req.id}] ${req.name}`,
                originalText: `미대응 산출정보: ${(req.missingOutputs || []).join(', ') || '-'}`,
                suggestedText: `빠진 산출정보(${(req.missingOutputs || []).join(', ')}) 작성 또는 근거 위치 수동 확정`,
                reason: `요구사항 [${req.id}]의 산출정보 중 일부만 산출물에서 확인되었습니다.`
            }))
        ]
        : missingReqs.map(req => ({
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

    // 파일별 근거 요구사항 수 (요약 시트·화면의 '해당파일 RTM 반영')
    const reqCountByFile = {};
    if (useEvidence) {
        projectRtmMatrix.forEach(r => r.appearingFiles.forEach(fn => { reqCountByFile[fn] = (reqCountByFile[fn] || 0) + 1; }));
    }

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
        useEvidence,
        projectRtmMatrix,
        reqCountByFile,
        partialReqCount: useEvidence ? statusCount(RTM_STATUS.PARTIAL) : 0,
        checkReqCount: useEvidence ? statusCount(RTM_STATUS.CHECK) : 0,
        mgmtCheckReqCount: useEvidence ? statusCount(RTM_STATUS.MGMT_CHECK) : 0,
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
    ssotData = {},
    evidenceResults = null,   // 화면에서 이미 계산한 근거 판정 결과 (없으면 여기서 계산)
    managementFiles = [],
    matchingConfig = null,
    manualEntries = []
}) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = '건강한 프로젝트 AI 산출물 검수기 (SSOT Batch)';
    workbook.created = new Date();

    const { metadata = {}, requirements = [] } = ssotData;

    // ── 프로젝트 레벨 통합 집계: 요구사항추적표와 같은 근거 판정(대응표·ID·절 제목·동의어) 사용 ──
    let evidence = evidenceResults;
    if (!evidence && requirements.length > 0 && batchResults.length > 0) {
        try {
            evidence = (await analyzeRequirementEvidence({ ssotData, batchResults, managementFiles, matchingConfig, manualEntries })).results;
        } catch (e) {
            console.warn('[일괄 리포트] 요구사항 근거 판정 실패, ID 출현 기준으로 집계:', e);
            evidence = null;
        }
    }
    const agg = aggregateBatchSSOT(batchResults, ssotData, evidence);

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
    const rtmBreakdown = agg.useEvidence
        ? `충족 ${agg.coveredReqCount}건 (${agg.projectRtmRate}%) · 부분충족 ${agg.partialReqCount} · 확인필요 ${agg.checkReqCount} · 사업관리 문서 확인필요 ${agg.mgmtCheckReqCount} · 미충족 ${agg.missingReqCount}건`
        : `충족 ${agg.coveredReqCount}건 (${agg.projectRtmRate}%) | 전수 누락: ${agg.missingReqCount}건`;
    kpiCell.value = `🎯 [프로젝트 전체 정합성 종합 진단] 총 요구사항 ${agg.totalMasterReqCount}건 중 ${rtmBreakdown} | 산출물 간 비정합 용어: ${agg.termConsistency.filter(t => t.isInconsistent).length}건 | 비표준 유령ID: ${agg.ghostIdSpread.length}건`;
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

        const coveredInThisFile = agg.useEvidence
            ? (agg.reqCountByFile[res.fileName] || 0)
            : (res.stage1?.rtmMatrix || []).filter(m => m.isPresent).length;
        const rtmStr = agg.useEvidence
            ? `${coveredInThisFile}건 근거 (${coveredInThisFile}/${agg.totalMasterReqCount})`
            : `${coveredInThisFile}건 반영 (${res.stage1?.coveredReqCount || coveredInThisFile}/${agg.totalMasterReqCount})`;

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

    if (agg.useEvidence) {
        // 요구사항추적표와 같은 근거 판정
        rtmSheet.columns = [
            { header: '순번', width: 7 },
            { header: '요구사항 ID', width: 13 },
            { header: '요구사항명', width: 38 },
            { header: '판정 근거 유형', width: 18 },
            { header: '근거 파일 수', width: 11 },
            { header: '근거 위치 (파일명 / 절 번호 / 장표 제목 / 페이지)', width: 80 },
            { header: '종합 판정', width: 16 },
            { header: '비고 (누락 산출정보·경고)', width: 46 }
        ];
        const evHead = rtmSheet.getRow(1);
        evHead.height = 28;
        for (let c = 1; c <= 8; c++) evHead.getCell(c).style = headerStyle;
        const statusStyle = {
            [RTM_STATUS.MET]: ['FFDCFCE7', 'FF16A34A'],
            [RTM_STATUS.PARTIAL]: ['FFFFEDD5', 'FFB45309'],
            [RTM_STATUS.CHECK]: ['FFFEF3C7', 'FFD97706'],
            [RTM_STATUS.MGMT_CHECK]: ['FFEDE9FE', 'FF7C3AED'],
            [RTM_STATUS.UNMET]: ['FFFEE2E2', 'FFDC2626']
        };
        agg.projectRtmMatrix.forEach((r, idx) => {
            const row = rtmSheet.addRow([
                idx + 1,
                r.id,
                r.name,
                (r.evidenceTypes || []).join(', ') || '-',
                r.appearingFiles.length,
                r.locations.length ? r.locations.slice(0, 15).join('\n') + (r.locations.length > 15 ? `\n… 외 ${r.locations.length - 15}건` : '') : '-',
                r.overallStatus,
                r.remarks.join('\n')
            ]);
            const lines = Math.max(Math.min(r.locations.length, 16), r.remarks.length, 1);
            row.height = Math.max(24, lines * 14);
            const [fill, color] = statusStyle[r.overallStatus] || ['FFFFFFFF', 'FF000000'];
            row.getCell(7).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
            row.getCell(7).font = { name: '맑은 고딕', color: { argb: color }, bold: true };
            for (let col = 1; col <= 8; col++) {
                row.getCell(col).border = borderStyle;
                row.getCell(col).alignment = {
                    vertical: 'middle',
                    horizontal: [1, 2, 5, 7].includes(col) ? 'center' : 'left',
                    wrapText: [3, 4, 6, 8].includes(col)
                };
            }
        });
    } else {
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
    }

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

// 목차 슬라이드의 번호 없는 제목 목록 (예: "환경분석 개요" / "내·외부 환경분석" / "법·제도 분석" …, 번호 도형은 별도)
function extractUnnumberedTocTitles(pageMap = {}, chapterTitle = '') {
    const slideNums = Object.keys(pageMap).map(Number).filter(n => !isNaN(n)).sort((a, b) => a - b).slice(0, 12);
    const tocSlide = slideNums.find(n => (pageMap[n] || []).some(p => /^\s*(목\s*차|CONTENTS)\s*$/i.test(p)));
    if (!tocSlide) return [];
    const skip = p => /^\s*(목\s*차|CONTENTS|INDEX|CHAPTER|내\s*용)\s*$/i.test(p)
        || /^[\dⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ\s.]+$/.test(p)
        || (chapterTitle && p.replace(/\s/g, '') === chapterTitle.replace(/\s/g, ''));
    const titles = (pageMap[tocSlide] || []).map(p => p.trim()).filter(p => p.length >= 3 && p.length <= 40 && !skip(p));
    return titles.length >= 2 && titles.length <= 15 ? titles.map((title, i) => ({ num: i + 1, title })) : [];
}

// 슬라이드별 상단 경로("1. 환경분석 개요", "1.1 분석목적 및 대상")로 절/항과 해당 슬라이드 구간을 구한다
const SECTION_HEADING_REGEX = /^0?([1-9]\d?)\.\s*([^\d\s.][^\n]{1,38})$/;
const SUBSECTION_HEADING_REGEX = /^0?([1-9]\d?)\.(\d{1,2})\.?\s+(\S[^\n]{1,38})$/;

export function buildSectionIndex(pageMap = {}, pageLabelMap = {}, tocTitles = new Map()) {
    const slideNums = Object.keys(pageMap).map(Number).filter(n => !isNaN(n)).sort((a, b) => a - b);
    const secCount = new Map();
    const subCount = new Map();
    const keySlides = new Map(); // 'num|title' → 나온 슬라이드 번호들
    const perSlide = slideNums.map(n => {
        const secs = new Set();
        const subs = new Set();
        (pageMap[n] || []).forEach(raw => {
            const p = raw.trim();
            let m = p.match(SUBSECTION_HEADING_REGEX);
            if (m) { subs.add(`${parseInt(m[1], 10)}.${parseInt(m[2], 10)}|${m[3].trim()}`); return; }
            m = p.match(SECTION_HEADING_REGEX);
            if (m) secs.add(`${parseInt(m[1], 10)}|${m[2].trim()}`);
        });
        // 슬라이드 위치는 목차 슬라이드(절 제목이 3개 이상 나열)를 빼고 기록
        const isToc = secs.size >= 3 || subs.size >= 4;
        const at = (k) => { if (isToc) return; if (!keySlides.has(k)) keySlides.set(k, []); keySlides.get(k).push(n); };
        secs.forEach(k => { secCount.set(k, (secCount.get(k) || 0) + 1); at(k); });
        subs.forEach(k => { subCount.set(k, (subCount.get(k) || 0) + 1); at(k); });
        return { n, secs, subs };
    });

    // 본문 번호 목록과 구분: 여러 슬라이드에 반복되는 경로 문구만 절/항 제목으로 채택
    const minCount = slideNums.length >= 6 ? 2 : 1;
    const bestTitle = (countMap) => {
        const best = new Map();
        countMap.forEach((cnt, key) => {
            if (cnt < minCount) return;
            const [num, title] = key.split('|');
            const cur = best.get(num);
            if (!cur || cnt > cur.cnt) best.set(num, { title, cnt, key });
        });
        return best;
    };
    const secBest = bestTitle(secCount);
    const subBest = bestTitle(subCount);
    // 한 슬라이드뿐인 절·항(예: '1. 현황분석 개요' 1장, '2.1 …' 1장)도 슬라이드 순서가 맞으면 채택
    // - 절: 인식된 절 번호 사이·앞의 빈 번호이거나 목차 슬라이드에 같은 번호가 있을 때
    // - 항: 상위 절 구간 안에 있고, 앞·뒤 번호 항 사이의 슬라이드이며, 이웃 번호(n-1 또는 n+1) 항이 있을 때
    // 같은 번호에 제목 후보가 여럿이면(본문 번호 목록 등) 목차 제목 → 슬라이드 순서로 하나를 고름
    const singlesOf = (countMap, best) => {
        const out = new Map();
        countMap.forEach((cnt, key) => {
            if (cnt >= minCount) return;
            const [num, title] = key.split('|');
            if (best.has(num)) return;
            if (!out.has(num)) out.set(num, []);
            out.get(num).push({ title, cnt, key, slide: Math.min(...(keySlides.get(key) || [Infinity])) });
        });
        return out;
    };
    const firstSlide = (v) => Math.min(...(keySlides.get(v.key) || [Infinity]));
    const secSingles = singlesOf(secCount, secBest);
    const secMax = Math.max(0, ...Array.from(secBest.keys()).map(Number));
    Array.from(secSingles.entries()).sort((x, y) => Number(x[0]) - Number(y[0])).forEach(([num, list]) => {
        const n = parseInt(num, 10);
        const toc = tocTitles.get(n);
        const tocHits = toc ? list.filter(c => normalizeForMatch(c.title) === normalizeForMatch(toc)) : [];
        // 앞·뒤 번호 절의 첫 슬라이드 사이에 있는 후보만
        const known = Array.from(secBest.entries()).map(([k, v]) => ({ n: Number(k), s: firstSlide(v) }));
        const lower = known.filter(k => k.n < n).sort((a, b) => b.n - a.n)[0];
        const upper = known.filter(k => k.n > n).sort((a, b) => a.n - b.n)[0];
        const ordered = list.filter(c => (!lower || c.slide >= lower.s) && (!upper || c.slide <= upper.s));
        const pick = tocHits.length === 1 ? tocHits[0] : (ordered.length === 1 ? ordered[0] : (list.length === 1 ? list[0] : null));
        if (!pick) return;
        const tocMatch = toc && normalizeForMatch(toc) === normalizeForMatch(pick.title);
        if (n < secMax || tocMatch) secBest.set(num, pick);
    });
    const secStart = new Map(Array.from(secBest.entries()).map(([k, v]) => [Number(k), firstSlide(v)]));
    const subSingles = singlesOf(subCount, subBest);
    const knownSub = new Map(); // parent → Map(child → firstSlide)
    subBest.forEach((v, num) => {
        const [pn, cn] = num.split('.').map(x => parseInt(x, 10));
        if (!knownSub.has(pn)) knownSub.set(pn, new Map());
        knownSub.get(pn).set(cn, firstSlide(v));
    });
    const pending = Array.from(subSingles.entries())
        .map(([num, list]) => { const [pn, cn] = num.split('.').map(x => parseInt(x, 10)); return { num, pn, cn, list }; })
        .filter(x => secBest.has(String(x.pn)))
        .sort((a, b) => a.pn - b.pn || a.cn - b.cn);
    // 이웃 번호가 생기면 그 다음 번호도 채택될 수 있으므로 더 바뀌지 않을 때까지 반복
    for (let changed = true; changed;) {
        changed = false;
        pending.forEach(x => {
            if (subBest.has(x.num)) return;
            const sibs = knownSub.get(x.pn) || new Map();
            const hasNeighbor = sibs.has(x.cn - 1) || sibs.has(x.cn + 1) || (x.cn === 1 && sibs.size > 0);
            if (!hasNeighbor) return;
            const pStart = secStart.get(x.pn) ?? -Infinity;
            const nextSec = Array.from(secStart.entries()).filter(([k]) => k > x.pn).sort((a, b) => a[0] - b[0])[0];
            const pEnd = nextSec ? nextSec[1] : Infinity;
            const lowerS = Math.max(-Infinity, ...Array.from(sibs.entries()).filter(([c]) => c < x.cn).map(([, sl]) => sl));
            const upperS = Math.min(Infinity, ...Array.from(sibs.entries()).filter(([c]) => c > x.cn).map(([, sl]) => sl));
            const ok = x.list.filter(c => c.slide >= pStart && c.slide < pEnd && c.slide >= lowerS && c.slide <= upperS)
                .sort((a, b) => a.slide - b.slide);
            if (!ok.length) return;
            subBest.set(x.num, ok[0]);
            if (!knownSub.has(x.pn)) knownSub.set(x.pn, new Map());
            knownSub.get(x.pn).set(x.cn, ok[0].slide);
            changed = true;
        });
    }
    const secKeys = new Set(Array.from(secBest.values()).map(v => v.key));
    const subKeys = new Set(Array.from(subBest.values()).map(v => v.key));

    const sections = new Map();
    const ensureSection = (num) => {
        if (!sections.has(num)) sections.set(num, { num, title: secBest.get(String(num))?.title || '', pages: [], subsections: new Map() });
        return sections.get(num);
    };

    let curSec = null;
    let curSub = null;
    perSlide.forEach(({ n, secs, subs }) => {
        const secHits = Array.from(secs).filter(k => secKeys.has(k));
        const subHits = Array.from(subs).filter(k => subKeys.has(k));
        // 여러 절이 한꺼번에 나열된 슬라이드 = 목차/간지 → 구간에 넣지 않는다
        if (secHits.length >= 3 || subHits.length >= 4) { curSec = null; curSub = null; return; }
        const byCount = (map) => (a, b) => (map.get(b) || 0) - (map.get(a) || 0);
        if (secHits.length > 0) {
            const num = parseInt(secHits.sort(byCount(secCount))[0].split('|')[0], 10);
            if (num !== curSec) curSub = null;
            curSec = num;
        }
        if (subHits.length > 0) {
            const key = subHits.sort(byCount(subCount))[0];
            const [secNum, subNum] = key.split('|')[0].split('.').map(x => parseInt(x, 10));
            curSec = secNum;
            curSub = subNum;
        }
        if (curSec == null) return;
        const sec = ensureSection(curSec);
        sec.pages.push(n);
        if (curSub != null) {
            const subTitle = subBest.get(`${curSec}.${curSub}`)?.title || '';
            if (!sec.subsections.has(curSub)) sec.subsections.set(curSub, { num: curSub, title: subTitle, pages: [] });
            sec.subsections.get(curSub).pages.push(n);
        }
    });

    const withLabels = node => ({ ...node, labels: node.pages.map(p => pageLabelMap[p] || String(p)) });
    return Array.from(sections.values())
        .sort((a, b) => a.num - b.num)
        .map(sec => ({
            ...withLabels(sec),
            subsections: Array.from(sec.subsections.values()).sort((a, b) => a.num - b.num).map(withLabels)
        }));
}

// 장 제목 ↔ 장 번호 기본값 (공공 ISP 표준 목차). 파일명이 "1장. 환경분석"처럼 내부 장 번호와 달라도 주제어로 판별
const DEFAULT_CHAPTER_TOPICS = [
    { num: 1, keys: ['사업개요', '사업의개요'] },
    { num: 2, keys: ['환경분석'] },
    { num: 3, keys: ['현황분석'] },
    { num: 4, keys: ['목표모델', '목표시스템'] },
    { num: 5, keys: ['이행계획', '통합이행'] }
];

function majorityChapter(slideChapterMap = {}) {
    const counts = new Map();
    Object.values(slideChapterMap).forEach(c => {
        const num = ROMAN_NUMERAL_VALUES[c.roman];
        if (!num) return;
        const cur = counts.get(num) || { cnt: 0, title: c.title };
        cur.cnt += 1;
        counts.set(num, cur);
    });
    let best = null;
    counts.forEach((v, num) => { if (!best || v.cnt > best.cnt) best = { num, ...v }; });
    return best;
}

function detectChapterFromFileTopic(fileName = '', knownTitles = new Map()) {
    const fName = normalizeForMatch(fileName);
    for (const [num, title] of knownTitles) {
        const t = normalizeForMatch(title);
        if (t.length >= 3 && fName.includes(t)) return num;
    }
    const hit = DEFAULT_CHAPTER_TOPICS.find(c => c.keys.some(k => fName.includes(k)));
    return hit ? hit.num : null;
}

// 검수 대상 PPTX별 장 번호, 절/항 구간, 목차 색인
export async function buildDeliverableTocIndex(batchResults = []) {
    const parsed = [];
    for (const res of batchResults) {
        if (!res.file || getFileExtension(res.fileName || '') !== 'pptx') continue;
        try {
            parsed.push({ res, content: await extractStructuredFileContent(res.file) });
        } catch (e) {
            console.warn(`[산출물현황] '${res.fileName}' 목차 분석 실패:`, e);
        }
    }

    // 모든 파일의 장 구분 헤더·표지 목차에서 장 번호 ↔ 장 제목 사전 구성 (예: 3 → '현황분석')
    const knownTitles = new Map();
    parsed.forEach(({ content }) => {
        Object.values(content.slideChapterMap || {}).forEach(c => {
            const num = ROMAN_NUMERAL_VALUES[c.roman];
            if (num && !knownTitles.has(num)) knownTitles.set(num, c.title);
        });
        extractTocEntries(content.pageMap).romanToc.forEach(t => {
            if (!knownTitles.has(t.num)) knownTitles.set(t.num, t.title.replace(/^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]\.\s*/, ''));
        });
    });

    return parsed.map(({ res, content }) => {
        const { numericToc, romanToc } = extractTocEntries(content.pageMap);
        const headerChapter = majorityChapter(content.slideChapterMap);
        const chapter = detectChapterFromPageLabels(content.pageLabelMap)
            ?? headerChapter?.num
            ?? detectChapterFromFileName(res.fileName)
            ?? detectChapterFromFileTopic(res.fileName, knownTitles)
            ?? (romanToc.length === 1 ? romanToc[0].num : null);
        const chapterTitle = (headerChapter && headerChapter.num === chapter ? headerChapter.title : null) || knownTitles.get(chapter) || '';

        // 절 제목 보완: 상단 경로 > 번호 목차 > 번호 없는 목차 슬라이드
        const sections = buildSectionIndex(content.pageMap, content.pageLabelMap, new Map(numericToc.map(t => [t.num, t.title])));
        const tocFallback = numericToc.length > 0 ? numericToc : extractUnnumberedTocTitles(content.pageMap, chapterTitle);
        tocFallback.forEach(t => {
            const sec = sections.find(s => s.num === t.num);
            if (sec && !sec.title) sec.title = t.title;
        });
        const mergedToc = new Map(tocFallback.map(t => [t.num, t.title]));
        sections.forEach(s => { if (s.title) mergedToc.set(s.num, s.title); });

        return {
            fileName: res.fileName,
            chapter,
            chapterTitle,
            sections,
            numericToc: Array.from(mergedToc.entries()).sort((a, b) => a[0] - b[0]).map(([num, title]) => ({ num, title })),
            romanToc,
            content
        };
    });
}

// 요구사항/서류명에 해당하는 절·항 목록 (절이 일치하면 하위 항은 생략)
function findSectionsForText(text, tocIndex = []) {
    const hits = [];
    const normText = normalizeForMatch(text);
    const textTopics = CHAPTER_TOPIC_STEMS.filter(t => normText.includes(t));
    tocIndex.forEach(doc => {
        const roman = NUM_TO_ROMAN[doc.chapter] || '';
        const docTopic = chapterTopicOf(doc.chapterTitle) || chapterTopicOf(doc.fileName);
        const allowPartial = textTopics.length === 0 || (docTopic && textTopics.includes(docTopic));
        doc.sections.forEach(sec => {
            if (sec.title && titleMatchesText(sec.title, text, { partial: allowPartial })) {
                hits.push({ fileName: doc.fileName, code: `${roman}.${sec.num}`, title: sec.title, pages: sec.pages, labels: sec.labels });
                return;
            }
            sec.subsections.forEach(sub => {
                if (sub.title && titleMatchesText(sub.title, text)) {
                    hits.push({ fileName: doc.fileName, code: `${roman}.${sec.num}.${sub.num}`, title: sub.title, pages: sub.pages, labels: sub.labels });
                }
            });
        });
    });
    return hits;
}

const formatTocLine = (code, title, node) => {
    const range = node && node.pages && node.pages.length > 0 ? ` (${formatPageLabelRanges(node.labels, node.pages)})` : '';
    return `${code} ${title}${range}`;
};

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
            const lines = [];
            nums.forEach(n => {
                const sec = (chapterFile.sections || []).find(item => item.num === n);
                const t = chapterFile.numericToc.find(item => item.num === n);
                const title = (sec && sec.title) || (t && t.title);
                if (title) lines.push(formatTocLine(`${roman}.${n}`, title, sec));
            });
            // 절 구조를 읽어낸 파일인데 해당 절이 하나도 없으면 이 파일에 작성되지 않은 것으로 본다
            const structureKnown = (chapterFile.sections || []).length > 0 || chapterFile.numericToc.length > 0;
            if (lines.length === 0 && structureKnown) return;
            if (!fileNames.includes(chapterFile.fileName)) fileNames.push(chapterFile.fileName);
            tocLines.push(...lines);
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

// ── 매칭 설정 / 수동 확정 목록 파일 읽기 ─────────────────────────────
/**
 * 동의어 사전·매칭 설정 파일 읽기 (JSON: 설정 전체 형식, CSV: "요구사항 용어, 산출물 용어1|용어2")
 */
export async function parseMatchingConfigFile(file) {
    const text = await file.text();
    const ext = getFileExtension(file.name);
    if (ext === 'json') {
        const json = JSON.parse(text);
        if (typeof json !== 'object' || Array.isArray(json)) throw new Error('JSON 최상위는 객체여야 합니다.');
        return json;
    }
    if (ext === 'csv' || ext === 'txt' || ext === 'tsv') {
        const parsed = parseSynonymCsv(text);
        if (parsed.synonyms.length === 0) throw new Error('CSV에서 "요구사항 용어, 산출물 용어" 행을 찾지 못했습니다.');
        return parsed;
    }
    throw new Error('JSON 또는 CSV 파일만 지원합니다.');
}

/**
 * 수동 확정 목록 파일 읽기 (CSV / JSON / XLSX - 요구사항추적표의 '근거 상세' 시트를 그대로 올려도 됨)
 * 열: 요구사항ID, 파일명, 절 번호, 장표 제목, 페이지, (산출정보), (판정), (확정(Y))
 */
export async function parseManualConfirmationFile(file) {
    const ext = getFileExtension(file.name);
    if (ext === 'json') {
        const json = JSON.parse(await file.text());
        return parseManualEntries(Array.isArray(json) ? json : (json.entries || []));
    }
    if (ext === 'csv' || ext === 'txt' || ext === 'tsv') {
        return parseManualEntries(parseCsvRows(await file.text()));
    }
    if (ext === 'xlsx') {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(await file.arrayBuffer());
        const sheet = wb.getWorksheet('수동 확정') || wb.getWorksheet('근거 상세') || wb.worksheets[0];
        if (!sheet) return [];
        const rows = [];
        sheet.eachRow({ includeEmpty: false }, row => {
            const cells = [];
            for (let c = 1; c <= row.cellCount; c++) {
                const v = row.getCell(c).value;
                cells.push(v == null ? '' : (typeof v === 'object' ? (v.richText ? v.richText.map(t => t.text).join('') : (v.text ?? v.result ?? '')) : String(v)));
            }
            rows.push(cells);
        });
        return parseManualEntries(rows);
    }
    throw new Error('CSV, JSON, XLSX 파일만 지원합니다.');
}

export const MANUAL_CONFIRMATION_TEMPLATE_CSV = '\uFEFF요구사항ID,산출정보,파일명,절 번호,장표 제목,페이지,판정\nDAR-005,방안,3장. 목표모델.pptx,Ⅳ.6 6.5.8,데이터 구조 관리방안,Ⅳ-690~692,충족\n';

// 사업관리 근거 문서(수행계획서·보안계획서·보고회 자료 등) 텍스트 추출. 지원하지 않는 형식(HWP 등)은 파일명만 사용
async function extractManagementDocs(files = []) {
    const docs = [];
    for (const file of files) {
        if (!file) continue;
        try {
            const content = await extractStructuredFileContent(file);
            docs.push({ fileName: file.name, pageMap: content.pageMap || {} });
        } catch (e) {
            console.warn(`[사업관리 문서] '${file.name}' 본문 추출 실패, 파일명만 사용:`, e?.message || e);
            docs.push({ fileName: file.name, pageMap: {} });
        }
    }
    return docs;
}

const RTM_STATUS_COLORS = {
    [RTM_STATUS.MET]: 'FF000000',
    [RTM_STATUS.PARTIAL]: 'FFB45309',
    [RTM_STATUS.CHECK]: 'FFD97706',
    [RTM_STATUS.UNMET]: 'FFDC2626',
    [RTM_STATUS.MGMT_CHECK]: 'FF7C3AED'
};

/**
 * 요구사항별 근거 수집·판정 (엑셀 생성과 화면 미리보기에서 공용)
 */
/**
 * 마스터 요구사항 목록을 1단계 하위 세부 과업(subItems) 단위로 전개하는 헬퍼
 * 예: CNR-016 (subItems 10건) -> CNR-016-01 ~ CNR-016-10 (10건의 독립 세부 요구사항 객체)
 */
export function flattenDetailedRequirements(requirements = []) {
    const flattened = [];
    (requirements || []).forEach(req => {
        const subItems = Array.isArray(req.subItems) && req.subItems.length > 0
            ? req.subItems.filter(Boolean)
            : [];

        if (subItems.length > 0) {
            subItems.forEach((sub, sIdx) => {
                const subId = `${req.id}-${String(sIdx + 1).padStart(2, '0')}`;
                // 세부 항목 내 주요 키워드 추출
                const subWords = String(sub)
                    .replace(/[\r\n\t]+/g, ' ')
                    .replace(/[()[\]{}.,~;:_!?"'`/\\|]/g, ' ')
                    .split(/\s+/)
                    .map(w => w.trim())
                    .filter(w => w.length >= 2 && !['체계', '구축', '방안', '수립', '및', '관한', '위한', '지원', '제공', '등', '에', '의', '수행', '대상', '내용', '정의', '검토', '활용', '마련'].includes(w));
                const subKeywords = Array.from(new Set(subWords)).slice(0, 5).join(', ');

                flattened.push({
                    id: subId,
                    parentReqId: req.id,
                    parentName: req.name,
                    category: req.category || '컨설팅 요구사항',
                    name: sub,
                    description: sub,
                    deliverables: req.deliverables || '',
                    keywords: req.keywords ? `${req.keywords}, ${subKeywords}` : subKeywords,
                    subIndex: sIdx + 1,
                    totalSubCount: subItems.length,
                    isSubItem: true
                });
            });
        } else {
            flattened.push({
                id: req.id,
                parentReqId: req.id,
                parentName: req.name,
                category: req.category || '일반 요구사항',
                name: req.name,
                description: req.description || req.name,
                deliverables: req.deliverables || '',
                keywords: req.keywords || '',
                subIndex: 1,
                totalSubCount: 1,
                isSubItem: false
            });
        }
    });
    return flattened;
}

export async function analyzeRequirementEvidence({
    ssotData = {},
    batchResults = [],
    managementFiles = [],
    matchingConfig = null,
    manualEntries = [],
    evaluationScope = 'subItem' // 'subItem' (세부 요구사항 기준) | 'parent' (상위 요구사항 기준)
}) {
    const config = mergeMatchingConfig(DEFAULT_RTM_MATCHING_CONFIG, matchingConfig);
    const tocIndex = batchResults.length > 0 ? await buildDeliverableTocIndex(batchResults) : [];
    const decks = tocIndex.map(doc => buildDeckIndex({ fileName: doc.fileName, content: doc.content, doc }));
    const otherFiles = batchResults
        .filter(b => getFileExtension(b.fileName || '') !== 'pptx')
        .map(b => ({ fileName: b.fileName }));
    const managementDocs = await extractManagementDocs(managementFiles);

    // 1) 세부 요구사항 전개 목록
    const detailedReqs = flattenDetailedRequirements(ssotData.requirements || []);
    // 2) 상위 요구사항 목록
    const parentReqs = ssotData.requirements || [];

    // 세부 요구사항 기준 평가
    const subItemResults = evaluateRequirements({
        requirements: detailedReqs,
        decks,
        otherFiles,
        managementDocs,
        manualEntries,
        config,
        deliverableItems: OFFICIAL_DOCUMENT_CHECKLIST_ITEMS
    });

    // 상위 요구사항 기준 평가
    const parentResults = evaluateRequirements({
        requirements: parentReqs,
        decks,
        otherFiles,
        managementDocs,
        manualEntries,
        config,
        deliverableItems: OFFICIAL_DOCUMENT_CHECKLIST_ITEMS
    });

    const activeResults = evaluationScope === 'subItem' ? subItemResults : parentResults;

    return {
        config,
        tocIndex,
        decks,
        results: activeResults,
        subItemResults,
        parentResults,
        evaluationScope
    };
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
    includeDocChecklist = false, // 산출물현황 시트 포함 여부 (기본 false)
    managementFiles = [],        // 사업관리 근거 문서 (SER·QUR·COR·PMR·PSR 우선 검색)
    matchingConfig = null,       // 동의어 사전·산출정보 설정 덮어쓰기
    manualEntries = [],          // PM 수동 확정 목록
    evaluationScope = 'subItem'  // 'subItem' (세부 요구사항 기준) | 'parent' (상위 요구사항 기준)
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
    const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
    const baseFont = { name: '맑은 고딕', size: 9.5 };
    const writeHeader = (sheet, headers) => {
        const row = sheet.getRow(7);
        row.height = 28;
        headers.forEach(h => {
            const cell = row.getCell(h.col);
            cell.value = h.title;
            cell.font = { name: '맑은 고딕', size: 10, bold: true };
            cell.fill = headerFill;
            cell.border = borderThin;
            cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        });
    };
    const writeAuthor = (sheet, range) => {
        sheet.mergeCells(range);
        const cell = sheet.getCell(range.split(':')[0]);
        cell.value = `${companyName}\n${issueDate}`;
        cell.font = { name: '맑은 고딕', size: 10, bold: true };
        cell.alignment = { vertical: 'middle', horizontal: 'right', wrapText: true };
        sheet.getRow(5).height = 30;
    };
    const setCell = (row, col, value, { align = 'left', bold = false, color = null, wrap = true } = {}) => {
        const cell = row.getCell(col);
        cell.value = value;
        cell.font = { ...baseFont, bold, ...(color ? { color: { argb: color } } : {}) };
        cell.alignment = { vertical: 'middle', horizontal: align, wrapText: wrap };
        cell.border = borderThin;
    };

    // ── 분석 ──────────────────────────────────────────────────────
    const masterReqs = ssotData.requirements || [];
    const useAuto = mode !== 'template' && masterReqs.length > 0;
    let analysis = { config: mergeMatchingConfig(DEFAULT_RTM_MATCHING_CONFIG, matchingConfig), tocIndex: [], decks: [], results: [] };
    if (useAuto) {
        analysis = await analyzeRequirementEvidence({ ssotData, batchResults, managementFiles, matchingConfig, manualEntries, evaluationScope });
    } else if (batchResults.length > 0) {
        analysis.tocIndex = await buildDeliverableTocIndex(batchResults);
    }
    const { config, tocIndex, results } = analysis;
    const resultsById = new Map(results.map(r => [String(r.id).toUpperCase(), r]));

    // ── 시트 1: 요구사항추적표 (세부 요구사항 기준 정밀 평가) ────────────────────
    const isSubScope = evaluationScope === 'subItem';
    const mainSheetName = isSubScope ? '요구사항추적표(세부과업기준)' : '요구사항추적표';
    const rtmSheet = workbook.addWorksheet(mainSheetName, { views: [{ showGridLines: true }] });

    if (isSubScope) {
        rtmSheet.columns = [
            { width: 4 }, { width: 4 },
            { width: 14 },  // C 세부요구사항ID
            { width: 28 },  // D 모 요구사항
            { width: 42 },  // E 세부 과업 내용 (하위 요건)
            { width: 13 },  // F 충족여부
            { width: 30 },  // G 관련 파일
            { width: 22 },  // H 해당 페이지
            { width: 16 },  // I 판정 근거 유형
            { width: 68 },  // J 근거 위치
            { width: 40 }   // K 비고
        ];
        writeAuthor(rtmSheet, 'J5:K5');
        writeHeader(rtmSheet, [
            { col: 3, title: '세부요구사항ID' },
            { col: 4, title: '모 요구사항' },
            { col: 5, title: '세부 과업 내용 (하위 실행 요건)' },
            { col: 6, title: '충족여부' },
            { col: 7, title: '관련 파일' },
            { col: 8, title: '해당 페이지' },
            { col: 9, title: '판정 근거 유형' },
            { col: 10, title: '근거 위치\n(파일명 / 절 번호 / 장표 제목 / 페이지)' },
            { col: 11, title: '비고\n(누락 산출정보·경고)' }
        ]);
    } else {
        rtmSheet.columns = [
            { width: 4 }, { width: 4 },
            { width: 13 },  // C 요구사항번호
            { width: 34 },  // D 요구사항명칭
            { width: 13 },  // E 충족여부
            { width: 30 },  // F 관련 파일
            { width: 22 },  // G 해당 페이지
            { width: 16 },  // H 판정 근거 유형
            { width: 72 },  // I 근거 위치
            { width: 46 }   // J 비고
        ];
        writeAuthor(rtmSheet, 'I5:J5');
        writeHeader(rtmSheet, [
            { col: 3, title: '요구사항번호' },
            { col: 4, title: '요구사항명칭' },
            { col: 5, title: '충족여부' },
            { col: 6, title: '관련 파일' },
            { col: 7, title: '해당 페이지' },
            { col: 8, title: '판정 근거 유형' },
            { col: 9, title: '근거 위치\n(파일명 / 절 번호 / 장표 제목 / 페이지)' },
            { col: 10, title: '비고\n(누락 산출정보·경고)' }
        ]);
    }

    const MAX_LOCATION_LINES = 15;
    const rtmRows = useAuto
        ? results.map(r => {
            const locLines = r.evidences.map(e => `[${EVIDENCE_TYPE_LABELS[e.type]}] ${evidenceLocationText(e)}`);
            const shown = locLines.slice(0, MAX_LOCATION_LINES);
            if (locLines.length > MAX_LOCATION_LINES) shown.push(`… 외 ${locLines.length - MAX_LOCATION_LINES}건 ('근거 상세' 시트 참고)`);
            const remarks = [
                ...(r.missingOutputs.length ? [`미대응 산출정보: ${r.missingOutputs.join(', ')}`] : []),
                ...r.warnings.map(w => `⚠ ${w}`),
                ...r.notes
            ];
            return {
                id: r.id,
                parentInfo: r.parentName ? `[${r.parentReqId}] ${r.parentName}` : (r.parentReqId || r.id),
                name: r.name,
                status: r.status,
                file: r.files.length ? r.files.join('\n') : (r.isManagement ? '사업관리 문서 미확인' : (batchResults.length ? '미확인 산출물' : '산출물 확인 필요')),
                pages: r.files.length ? r.pageText : '-',
                types: r.typeLabels.join(', ') || '-',
                locations: shown.join('\n') || '-',
                remarks: remarks.join('\n') || ''
            };
        })
        : OFFICIAL_RTM_TEMPLATE_ITEMS.map(item => ({ ...item, parentInfo: item.id, types: '', locations: '', remarks: '' }));

    rtmRows.forEach((item, idx) => {
        const row = rtmSheet.getRow(8 + idx);
        const lines = Math.max(
            String(item.file || '').split('\n').length,
            String(item.locations || '').split('\n').length,
            String(item.remarks || '').split('\n').length,
            Math.ceil(String(item.pages || '').length / 22)
        );
        row.height = Math.max(24, lines * 14);

        if (isSubScope) {
            setCell(row, 3, item.id, { align: 'center', wrap: false, bold: true, color: 'FF1D4ED8' });
            setCell(row, 4, item.parentInfo);
            setCell(row, 5, item.name);
            setCell(row, 6, item.status || '충족', { align: 'center', bold: true, color: RTM_STATUS_COLORS[item.status] || 'FF000000' });
            setCell(row, 7, item.file || '-');
            setCell(row, 8, item.pages || '-', { align: 'center' });
            setCell(row, 9, item.types || '', { align: 'center' });
            setCell(row, 10, item.locations || '');
            setCell(row, 11, item.remarks || '');
        } else {
            setCell(row, 3, item.id, { align: 'center', wrap: false });
            setCell(row, 4, item.name);
            setCell(row, 5, item.status || '충족', { align: 'center', bold: true, color: RTM_STATUS_COLORS[item.status] || 'FF000000' });
            setCell(row, 6, item.file || '-');
            setCell(row, 7, item.pages || '-', { align: 'center' });
            setCell(row, 8, item.types || '', { align: 'center' });
            setCell(row, 9, item.locations || '');
            setCell(row, 10, item.remarks || '');
        }
    });

    // ── 시트 2 (세부과업 기준일 경우): 상위 요구사항 종합 추적표 시트 추가 ──────
    if (isSubScope && useAuto && analysis.parentResults && analysis.parentResults.length > 0) {
        const parentSheet = workbook.addWorksheet('요구사항추적표(상위종합)', { views: [{ showGridLines: true }] });
        parentSheet.columns = [
            { width: 4 }, { width: 4 },
            { width: 13 },  // C 요구사항번호
            { width: 34 },  // D 요구사항명칭
            { width: 13 },  // E 충족여부
            { width: 30 },  // F 관련 파일
            { width: 22 },  // G 해당 페이지
            { width: 16 },  // H 판정 근거 유형
            { width: 72 },  // I 근거 위치
            { width: 46 }   // J 비고
        ];
        writeAuthor(parentSheet, 'I5:J5');
        writeHeader(parentSheet, [
            { col: 3, title: '요구사항번호' },
            { col: 4, title: '요구사항명칭' },
            { col: 5, title: '충족여부' },
            { col: 6, title: '관련 파일' },
            { col: 7, title: '해당 페이지' },
            { col: 8, title: '판정 근거 유형' },
            { col: 9, title: '근거 위치\n(파일명 / 절 번호 / 장표 제목 / 페이지)' },
            { col: 10, title: '비고\n(누락 산출정보·경고)' }
        ]);

        analysis.parentResults.forEach((r, idx) => {
            const row = parentSheet.getRow(8 + idx);
            const locLines = r.evidences.map(e => `[${EVIDENCE_TYPE_LABELS[e.type]}] ${evidenceLocationText(e)}`);
            const shown = locLines.slice(0, MAX_LOCATION_LINES);
            const remarks = [
                ...(r.missingOutputs.length ? [`미대응 산출정보: ${r.missingOutputs.join(', ')}`] : []),
                ...r.warnings.map(w => `⚠ ${w}`),
                ...r.notes
            ];
            row.height = 26;
            setCell(row, 3, r.id, { align: 'center', wrap: false, bold: true });
            setCell(row, 4, r.name);
            setCell(row, 5, r.status || '충족', { align: 'center', bold: true, color: RTM_STATUS_COLORS[r.status] || 'FF000000' });
            setCell(row, 6, r.files.length ? r.files.join('\n') : '-');
            setCell(row, 7, r.files.length ? r.pageText : '-', { align: 'center' });
            setCell(row, 8, r.typeLabels.join(', ') || '-', { align: 'center' });
            setCell(row, 9, shown.join('\n') || '-');
            setCell(row, 10, remarks.join('\n') || '');
        });
    }

    // ── 시트 2: RFP상 산출물 (추적표 판정에서 도출) ─────────────────────
    if (includeDocChecklist) {
        const docSheet = workbook.addWorksheet('RFP상 산출물', { views: [{ showGridLines: true }] });
        docSheet.columns = [
            { width: 4 },
            { width: 7 },   // B 구분
            { width: 34 },  // C 서류
            { width: 11 },  // D 현황
            { width: 30 },  // E 파일명
            { width: 60 },  // F 주요 목차
            { width: 18 },  // G 비고(작성근거)
            { width: 34 }   // H 추적표 연계
        ];
        writeAuthor(docSheet, 'G5:H5');
        writeHeader(docSheet, [
            { col: 2, title: '구분' },
            { col: 3, title: '서류' },
            { col: 4, title: '현황' },
            { col: 5, title: '파일명' },
            { col: 6, title: '주요 목차' },
            { col: 7, title: '비고(작성근거)' },
            { col: 8, title: '추적표 연계\n(요구사항: 판정)' }
        ]);

        // 1) 비고의 목차 코드(Ⅱ.3 등) ↔ 파일의 장·절 구간  2) 서류명 ↔ 절 제목  3) 파일명
        // 4) 요구사항추적표 근거(연계된 요구사항·산출정보) → 두 시트가 같은 근거를 쓰도록 우선 반영
        const links = buildDeliverableLinks(config);
        const docRows = OFFICIAL_DOCUMENT_CHECKLIST_ITEMS.map(item => {
            let fileNames = [];
            let tocLines = [];
            const refs = parseTocRefs(item.note);
            if (batchResults.length > 0) {
                if (refs.length > 0) ({ fileNames, tocLines } = resolveDocByTocRefs(refs, tocIndex));
                if (fileNames.length === 0 && refs.length === 0) {
                    findSectionsForText(item.docName.replace(/\(.*?\)/g, ''), tocIndex).forEach(hit => {
                        if (!fileNames.includes(hit.fileName)) fileNames.push(hit.fileName);
                        tocLines.push(formatTocLine(hit.code, hit.title, hit));
                    });
                }
                if (fileNames.length === 0 && refs.length === 0) {
                    const targetKeywords = (item.keywords || []).map(k => k.toLowerCase().replace(/[\s·\-_]/g, ''));
                    const docNameKey = item.docName.toLowerCase().replace(/[\s·\-_]/g, '');
                    const matchedFile = batchResults.find(b => {
                        const fName = (b.fileName || '').toLowerCase().replace(/[\s·\-_]/g, '');
                        if (fName.includes(docNameKey)) return true;
                        return targetKeywords.some(k => k.length >= 3 && fName.includes(k));
                    });
                    if (matchedFile) fileNames = [matchedFile.fileName];
                }
            }
            const itemRefSections = parseTocRefs(item.note).map(r => `${NUM_TO_ROMAN[r.chapter]}.${r.num}`);
            // 비고에 목차 코드가 있는 서류는 연계 요구사항의 근거 중 그 장·절 안의 것만 반영
            const linkInfos = (useAuto ? (links.get(item.seq) || []).map(l => ({ ...l, ...gradeOfLink(l, resultsById) })) : [])
                .map(l => {
                    if (itemRefSections.length === 0) return l;
                    const evidences = l.evidences.filter(e => e.type === 'manual' || (e.sections || []).some(sec => itemRefSections.includes(sec)));
                    const strongEv = evidences.some(e => !['semantic', 'mgmtFile'].includes(e.type));
                    return { ...l, evidences, grade: evidences.length === 0 ? 'none' : (strongEv && l.grade !== 'none' ? 'strong' : 'weak') };
                });
            // 추적표 근거가 있는 서류는 추적표 근거만으로 위치를 적는다 (서류명 매칭으로 찾은 넓은 절은 생략)
            if (refs.length === 0 && linkInfos.some(l => l.grade !== 'none')) {
                fileNames = [];
                tocLines = [];
            }
            const sectionFound = fileNames.length > 0;

            const remarks = [];
            let status = sectionFound ? '작성완료' : '미작성';
            if (linkInfos.length > 0) {
                const strong = linkInfos.filter(l => l.grade === 'strong');
                const weak = linkInfos.filter(l => l.grade === 'weak');
                if (strong.length) status = '작성완료';
                else if (sectionFound) {
                    status = '작성완료';
                    remarks.push('추적표에서는 미충족 - 근거 위치 확인');
                } else if (weak.length) status = '확인필요';
                else status = '미작성';

                const seen = new Set(tocLines);
                [...strong, ...weak].forEach(l => {
                    l.evidences.forEach(e => {
                        if (!fileNames.includes(e.fileName)) fileNames.push(e.fileName);
                        if (sectionFound && ['section', 'semantic'].includes(e.type)) return;
                        const pages = e.labels.length ? formatPageLabelRanges(e.labels, e.pages) : (e.pagesText || '전 페이지');
                        const loc = [e.sectionCode, e.code && e.code !== String(e.sectionCode).split('.').pop() ? e.code : ''].filter(Boolean).join(' ');
                        const line = `${loc ? `${loc} ` : ''}${e.title || e.fileName} (${pages})`.trim();
                        if (!seen.has(line)) { seen.add(line); tocLines.push(line); }
                    });
                });
            }
            const linkText = linkInfos
                .map(l => `${l.reqId}${l.outputKey ? `(${l.outputKey})` : ''}: ${l.result ? l.result.status : '-'}`)
                .join('\n');
            return {
                ...item,
                status,
                fileName: fileNames.length ? fileNames.join('\n') : '-',
                tocSummary: tocLines.length ? tocLines.join('\n') : '-',
                note: [item.note, ...remarks].filter(Boolean).join('\n'),
                linkText
            };
        });

        docRows.forEach((item, idx) => {
            const row = docSheet.getRow(8 + idx);
            const lines = Math.max(item.fileName.split('\n').length, item.tocSummary.split('\n').length, item.linkText.split('\n').length, item.note.split('\n').length);
            row.height = Math.max(24, lines * 14);
            setCell(row, 2, item.seq, { align: 'center', wrap: false });
            setCell(row, 3, item.docName);
            setCell(row, 4, item.status, { align: 'center', bold: item.status !== '작성완료', color: item.status === '미작성' ? 'FFDC2626' : (item.status === '확인필요' ? 'FFD97706' : null) });
            setCell(row, 5, item.fileName);
            setCell(row, 6, item.tocSummary);
            setCell(row, 7, item.note);
            setCell(row, 8, item.linkText);
        });
    }

    // ── 시트 3: 근거 상세 (PM 검토용. '확정(Y)'에 Y를 표시해 다시 올리면 수동 확정 목록으로 반영) ──
    if (useAuto) {
        const evSheet = workbook.addWorksheet('근거 상세', { views: [{ state: 'frozen', ySplit: 1 }] });
        evSheet.columns = [
            { header: '요구사항ID', key: 'id', width: 12 },
            { header: '요구사항명칭', key: 'name', width: 32 },
            { header: '판정', key: 'status', width: 12 },
            { header: '근거 유형', key: 'type', width: 14 },
            { header: '산출정보', key: 'output', width: 16 },
            { header: '파일명', key: 'file', width: 30 },
            { header: '절 번호', key: 'loc', width: 18 },
            { header: '장표 제목', key: 'title', width: 44 },
            { header: '페이지', key: 'pages', width: 22 },
            { header: '확정(Y)', key: 'confirm', width: 9 }
        ];
        results.forEach(r => {
            const evs = r.evidences.length ? r.evidences : [null];
            evs.forEach(e => {
                evSheet.addRow({
                    id: r.id,
                    name: r.name,
                    status: r.status,
                    type: e ? EVIDENCE_TYPE_LABELS[e.type] : '-',
                    output: e ? (e.outputs || [e.output]).filter(Boolean).join(', ') : '',
                    file: e ? e.fileName : '',
                    loc: e ? [e.sectionCode, e.code && e.code !== String(e.sectionCode).split('.').pop() ? e.code : ''].filter(Boolean).join(' ') : '',
                    title: e ? e.title : '',
                    pages: e ? (e.labels.length ? formatPageLabelRanges(e.labels, e.pages) : (e.pagesText || '전 페이지')) : '',
                    confirm: e && e.type === 'manual' ? 'Y' : ''
                });
            });
        });
        evSheet.getRow(1).font = { name: '맑은 고딕', size: 10, bold: true };
        evSheet.getRow(1).fill = headerFill;
        evSheet.eachRow((row, n) => {
            if (n > 1) row.font = baseFont;
            row.alignment = { vertical: 'top', wrapText: true };
        });
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const filePrefix = '요구사항추적표';
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
    // ID -> { id, category, name, description, subItems, deliverables, keywords } 맵
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

    // 텍스트 블록에서 번호 매김 하위 세부 항목들을 분리 추출하는 헬퍼
    const parseSubItemsFromText = (descText) => {
        if (!descText) return [];
        const lines = String(descText).split(/\r?\n/).map(l => l.trim()).filter(Boolean);
        const subList = [];
        let currentItem = null;

        for (const line of lines) {
            // 대항목 판별: '○', '●', '□', '■', '▶', '◆', 'ㅇ'(한글 이응 불릿), 번호 매김('1.', '2.', '(1)', '①'), 한글 번호('가.', '나.')
            const isMajor = /^(?:[○●□■▶◆ㅇ]|(?:\d+[.)])|(?:\(\d+\))|(?:[①-⑳])|(?:[가-힣]\.))\s*/.test(line);
            // 소항목(보조 설명) 판별: '-', '•', '*', '·'
            const isMinor = /^[-•*·]\s*/.test(line);

            if (isMajor) {
                if (currentItem) {
                    subList.push(currentItem.trim());
                }
                currentItem = line.replace(/^(?:[○●□■▶◆ㅇ]|(?:\d+[.)])|(?:\(\d+\))|(?:[①-⑳])|(?:[가-힣]\.))\s*/, '').trim();
            } else if (isMinor) {
                const minorText = line.replace(/^[-•*·]\s*/, '').trim();
                if (currentItem) {
                    currentItem += ` (${minorText})`;
                } else {
                    currentItem = minorText;
                }
            } else {
                if (currentItem) {
                    currentItem += ` ${line}`;
                } else if (line.length >= 6) {
                    currentItem = line;
                }
            }
        }
        if (currentItem) {
            subList.push(currentItem.trim());
        }

        if (subList.length === 0) {
            for (const line of lines) {
                const clean = line.replace(/^(?:[0-9]+[.)]|[-•*□■▶◆○●·ㅇ]|(?:\(\d+\))|(?:[가-힣]\.))\s*/, '').trim();
                if (clean.length >= 6) subList.push(clean);
            }
        }

        return subList.slice(0, 50); // 세부 요구사항이 10건 이상이어도 전수 수용
    };

    const addOrUpdateReq = (id, rawName, rawDesc, rawCategory = '', rawDeliverables = '', rawSubItems = null) => {
        if (!id) return;
        const cleanId = String(id).trim().toUpperCase();
        if (!/^[A-Z]{2,6}[-_]\d{2,4}$/.test(cleanId)) return;

        let name = String(rawName || '').trim();
        name = name
            .replace(/[\r\n\t]+/g, ' ')
            .replace(/^(?:[0-9]+[.)]|[-•*□■▶◆○●])\s*/, '')
            .replace(/\s+/g, ' ')
            .trim();

        let desc = String(rawDesc || '').trim();
        let category = String(rawCategory || '').trim();
        let deliverables = String(rawDeliverables || '').trim();

        // subItems 파싱
        let subItems = Array.isArray(rawSubItems) && rawSubItems.length > 0
            ? rawSubItems.map(s => String(s).trim()).filter(Boolean)
            : parseSubItemsFromText(desc);

        // 키워드 추출
        let keywords = '';
        if (desc || name) {
            const words = `${name} ${desc}`
                .replace(/[\r\n\t]+/g, ' ')
                .replace(/[()[\]{}.,~;:_!?"'`/\\|]/g, ' ')
                .split(/\s+/)
                .map(w => w.trim())
                .filter(w => w.length >= 2 && !['체계', '구축', '방안', '수립', '및', '관한', '위한', '지원', '제공', '등', '에', '의', '수행', '대상', '내용', '정의', '요구사항', '관리'].includes(w));
            keywords = Array.from(new Set(words)).slice(0, 5).join(', ');
        }

        // 표준 44종 카탈로그 조회 및 자동 보강
        const stdCatalog = PUBLIC_HEALTHCARE_AI_ISP_REQUIREMENTS.find(r => r.id === cleanId);
        let resolvedName = name;
        let resolvedCategory = category;
        let resolvedDesc = desc;
        let resolvedSubItems = subItems;
        let resolvedDeliverables = deliverables;
        let resolvedKeywords = keywords;

        if (stdCatalog) {
            if (isPlaceholderName(resolvedName, cleanId)) {
                resolvedName = stdCatalog.name;
            }
            if (!resolvedCategory && stdCatalog.category) {
                resolvedCategory = stdCatalog.category;
            }
            if ((!resolvedDesc || resolvedDesc.length < 10) && stdCatalog.description) {
                resolvedDesc = stdCatalog.description;
            }
            if ((!resolvedSubItems || resolvedSubItems.length === 0) && Array.isArray(stdCatalog.subItems)) {
                resolvedSubItems = [...stdCatalog.subItems];
            }
            if (!resolvedDeliverables && stdCatalog.deliverables) {
                resolvedDeliverables = stdCatalog.deliverables;
            }
            if (!resolvedKeywords && stdCatalog.keywords) {
                resolvedKeywords = stdCatalog.keywords;
            }
        }

        if (!resolvedCategory) {
            if (cleanId.startsWith('CNR')) resolvedCategory = '컨설팅 요구사항';
            else if (cleanId.startsWith('SFR')) resolvedCategory = '기능 요구사항';
            else if (cleanId.startsWith('DAR')) resolvedCategory = '데이터 요구사항';
            else if (cleanId.startsWith('SER')) resolvedCategory = '보안 요구사항';
            else if (cleanId.startsWith('QUR')) resolvedCategory = '품질 요구사항';
            else if (cleanId.startsWith('PMR')) resolvedCategory = '프로젝트관리 요구사항';
            else if (cleanId.startsWith('PSR')) resolvedCategory = '지원 요구사항';
            else if (cleanId.startsWith('COR')) resolvedCategory = '제약사항';
            else resolvedCategory = '일반 요구사항';
        }

        const existing = reqDict.get(cleanId);
        if (!existing) {
            reqDict.set(cleanId, {
                id: cleanId,
                category: resolvedCategory,
                name: !isPlaceholderName(resolvedName, cleanId) ? resolvedName : `요구사항 ${cleanId}`,
                description: resolvedDesc,
                subItems: resolvedSubItems,
                deliverables: resolvedDeliverables,
                keywords: resolvedKeywords || (resolvedName && !isPlaceholderName(resolvedName, cleanId) ? resolvedName.split(/\s+/).filter(w => w.length >= 2).slice(0, 4).join(', ') : '')
            });
        } else {
            // 업데이트 및 세부내용 보강
            if (isPlaceholderName(existing.name, cleanId) && !isPlaceholderName(resolvedName, cleanId)) {
                existing.name = resolvedName;
            } else if (!isPlaceholderName(resolvedName, cleanId) && resolvedName.length > existing.name.length && !existing.name.includes(resolvedName)) {
                existing.name = resolvedName;
            }
            if (!existing.category && resolvedCategory) existing.category = resolvedCategory;
            if ((!existing.description || existing.description.length < resolvedDesc.length) && resolvedDesc) {
                existing.description = resolvedDesc;
            }
            if ((!existing.subItems || existing.subItems.length === 0) && resolvedSubItems.length > 0) {
                existing.subItems = resolvedSubItems;
            }
            if (!existing.deliverables && resolvedDeliverables) existing.deliverables = resolvedDeliverables;
            if (!existing.keywords && resolvedKeywords) existing.keywords = resolvedKeywords;
        }
    };

    if (Array.isArray(tables)) {
        for (const tbl of tables) {
            const rows = tbl.rows || [];
            if (rows.length < 1) continue;

            // ── [패턴 1: 양식형 Key-Value 표 (공공 RFP 세부 요구사항 정의서 표)] ──
            let kvId = '';
            let kvName = '';
            let kvCategory = '';
            let kvDesc = '';
            let kvDeliverables = '';

            for (let r = 0; r < rows.length; r++) {
                const row = rows[r];
                for (let c = 0; c < row.length; c++) {
                    const cell = (row[c] || '').replace(/\s+/g, '');

                    // ID 라벨 탐색
                    if (cell.includes('요구사항고유번호') || cell.includes('요구사항번호') || cell.includes('요구사항id') || cell.includes('요구사항식별자')) {
                        const nextCell = (row[c + 1] || '').trim();
                        const idMatch = nextCell.match(/[A-Z]{2,6}[-_]\d{2,4}/i);
                        if (idMatch) {
                            kvId = idMatch[0].toUpperCase();
                        } else if (rows[r + 1] && rows[r + 1][c]) {
                            const belowMatch = (rows[r + 1][c] || '').match(/[A-Z]{2,6}[-_]\d{2,4}/i);
                            if (belowMatch) kvId = belowMatch[0].toUpperCase();
                        }
                    }

                    // 분류 라벨 탐색
                    if (cell.includes('요구사항구분') || cell.includes('요구사항분류') || cell === '구분' || cell === '분류' || cell.includes('과업구분')) {
                        const nextCell = (row[c + 1] || '').trim();
                        if (nextCell && nextCell.length >= 2 && !nextCell.includes('고유번호')) {
                            kvCategory = nextCell;
                        }
                    }

                    // 요구사항명 라벨 탐색
                    if ((cell.includes('요구사항명') || cell.includes('명칭') || cell.includes('과업명') || cell.includes('기능명')) && !cell.includes('내용') && !cell.includes('설명') && !cell.includes('고유번호')) {
                        const nextCell = (row[c + 1] || '').trim();
                        if (nextCell && nextCell.length >= 3 && !nextCell.includes('고유번호')) {
                            kvName = nextCell;
                        } else if (rows[r + 1] && rows[r + 1][c]) {
                            const belowCell = (rows[r + 1][c] || '').trim();
                            if (belowCell && belowCell.length >= 3) kvName = belowCell;
                        }
                    }

                    // 세부내용 / 상세설명 / 과업내용 라벨 탐색
                    if (cell.includes('세부내용') || cell.includes('요구사항세부') || cell.includes('과업내용') || cell.includes('상세설명') || cell.includes('세부요건') || cell.includes('과업상세')) {
                        const nextCell = (row[c + 1] || '').trim();
                        if (nextCell && nextCell.length >= 5) {
                            kvDesc = nextCell;
                        } else if (rows[r + 1] && rows[r + 1][c]) {
                            const belowCell = (rows[r + 1][c] || '').trim();
                            if (belowCell && belowCell.length >= 5) kvDesc = belowCell;
                        }
                    }

                    // 관련 산출물 라벨 탐색
                    if (cell.includes('산출물') || cell.includes('관련산출물') || cell.includes('제출산출물') || cell.includes('증빙')) {
                        const nextCell = (row[c + 1] || '').trim();
                        if (nextCell && nextCell.length >= 2) {
                            kvDeliverables = nextCell;
                        }
                    }
                }
            }

            if (kvId && (kvName || kvDesc)) {
                addOrUpdateReq(kvId, kvName, kvDesc, kvCategory, kvDeliverables);
            }

            // ── [패턴 2: 그리드형 요구사항 총괄/목록 테이블] ──
            let bestHeaderRow = 0;
            let idCol = -1;
            let nameCol = -1;
            let catCol = -1;
            let descCol = -1;
            let delivCol = -1;

            for (let r = 0; r < Math.min(4, rows.length); r++) {
                const headerRow = rows[r].map(c => (c || '').replace(/\s+/g, '').toLowerCase());
                const curId = headerRow.findIndex(h => h.includes('고유번호') || h.includes('요구사항id') || h.includes('요구사항번호') || h === 'id' || h.includes('식별자') || (h.includes('번호') && !h.includes('전화') && !h.includes('순번')));
                const curName = headerRow.findIndex(h => (h.includes('요구사항명') || h.includes('기능명') || h.includes('과업명') || h.includes('항목명') || h === '명칭' || h === '제목' || h === '과제명') && !h.includes('설명') && !h.includes('구분'));
                const curCat = headerRow.findIndex(h => h.includes('구분') || h.includes('분류') || h.includes('유형'));
                const curDesc = headerRow.findIndex(h => h.includes('세부') || h.includes('내용') || h.includes('설명') || h.includes('정의') || h.includes('키워드') || h.includes('개요'));
                const curDeliv = headerRow.findIndex(h => h.includes('산출물') || h.includes('증빙'));

                if (curId >= 0 || curName >= 0) {
                    bestHeaderRow = r;
                    idCol = curId;
                    nameCol = curName;
                    catCol = curCat;
                    descCol = curDesc;
                    delivCol = curDeliv;
                    break;
                }
            }

            for (let r = bestHeaderRow + 1; r < rows.length; r++) {
                const row = rows[r];
                let rowReqId = '';
                let rowReqName = '';
                let rowReqCat = '';
                let rowReqDesc = '';
                let rowReqDeliv = '';

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
                                if (candidate.length >= 3 && /[가-힣]/.test(candidate) && !candidate.match(/^[A-Z]{2,6}[-_]\d{2,4}$/i) && !['공통', '기능', '보안', '품질', '운영', '필수'].includes(candidate)) {
                                    rowReqName = candidate;
                                    break;
                                }
                            }
                        }
                    }

                    if (catCol >= 0 && row[catCol]) {
                        rowReqCat = row[catCol].trim();
                    }
                    if (descCol >= 0 && row[descCol]) {
                        rowReqDesc = row[descCol].trim();
                    }
                    if (delivCol >= 0 && row[delivCol]) {
                        rowReqDeliv = row[delivCol].trim();
                    }

                    addOrUpdateReq(rowReqId, rowReqName, rowReqDesc, rowReqCat, rowReqDeliv);
                }
            }
        }
    }

    // ── 3) 텍스트 본문 정규식 전수 탐색 및 보완 ──
    const regexPatterns = [
        /\[([A-Z]{2,6}[-_]\d{2,4})\][ \t]*[:：\-–—./)\]]?[ \t]*([가-힣A-Za-z0-9][^\n\r]{2,80})/g,
        /(?:^|\n)[ \t]*(?:[□■▶◆○●*·-]|가\.|나\.|다\.|라\.|[0-9]+[.)])?[ \t]*([A-Z]{2,6}[-_]\d{2,4})[ \t]*[:：\-–—./)\]][ \t]*([가-힣A-Za-z0-9][^\n\r]{2,80})/g,
        /(?:^|\n)[ \t]*(?:[□■▶◆○●*·-]|가\.|나\.|다\.|라\.|[0-9]+[.)])?[ \t]*([A-Z]{2,6}[-_]\d{2,4})[ \t]{1,6}([가-힣A-Za-z0-9][^\n\r]{2,80})/g,
        /(?:요구사항\s*(?:ID|번호|고유번호)?[ \t]*[:：|][ \t]*)([A-Z]{2,6}[-_]\d{2,4})[ \t\n\r]*(?:요구사항명[ \t]*[:：|][ \t]*)([가-힣A-Za-z0-9][^\n\r]{2,80})/g,
        /(?:^|\n)[ \t]*([A-Z]{2,6}[-_]\d{2,4})[ \t]*\n[ \t]*([가-힣][^\n\r]{2,80})/g
    ];

    for (const pat of regexPatterns) {
        let m;
        while ((m = pat.exec(cleanText)) !== null) {
            const id = m[1].trim().toUpperCase();
            let rawName = m[2].trim();
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
        let autoSynonyms = '';
        if (clientAgency.includes('국립중앙의료원')) {
            autoSynonyms = 'NMC';
            autoForbidden = '중앙의료원, 국립의료원, 국립중앙병원';
        } else if (clientAgency.includes('보건복지부')) {
            autoSynonyms = '복지부, MOHW';
            autoForbidden = '보건부';
        } else if (clientAgency.includes('정보원')) {
            autoForbidden = clientAgency.replace(/정보원$/, '센터');
        } else {
            autoForbidden = `${clientAgency}재단, ${clientAgency}센터`;
        }

        glossary.unshift({
            id: `glo-rfp-agency`,
            category: '기관명',
            approved: clientAgency,
            synonyms: autoSynonyms,
            forbidden: autoForbidden,
            note: 'RFP 발주기관 공식 명칭 (허용 약어/동의어는 정상 통과)'
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
                forbidden: sysMatch[1].split('').join(' '),
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
 * Gemini AI 기반 심층 RFP 분석기 (한 단계 더 세부적인 과업 및 하위 항목 추출)
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
    const introSnippet = fullText.slice(0, 8000);

    const paragraphs = fullText.split(/\n{2,}/);
    const reqParagraphs = [];
    let reqChars = 0;

    for (const p of paragraphs) {
        if (/[A-Z]{2,6}[-_]\d{2,4}/.test(p) || (p.includes('요구사항') && (p.includes('기능') || p.includes('분석') || p.includes('체계') || p.includes('구축')))) {
            reqParagraphs.push(p);
            reqChars += p.length;
            if (reqChars > 40000) break;
        }
    }

    const textSnippet = introSnippet + '\n\n... [중략: 요구사항 핵심 본문 발췌] ...\n\n' + reqParagraphs.join('\n\n');

    const relevantTables = (tables || []).filter(t => {
        const str = JSON.stringify(t.rows || []);
        return /[A-Z]{2,6}[-_]\d{2,4}/.test(str) || str.includes('요구사항') || str.includes('과업명') || str.includes('고유번호') || str.includes('세부내용');
    });

    const tablesToUse = (relevantTables.length > 0 ? relevantTables : (tables || [])).slice(0, 25);

    const tableSnippets = tablesToUse.map((t, idx) => {
        const rowsStr = (t.rows || []).slice(0, 40).map(r => r.join(' | ')).join('\n');
        return `[표 #${idx + 1} (${t.source || '표'})]:\n${rowsStr}`;
    }).join('\n\n');

    const prompt = `당신은 공공기관 정보화전략계획(ISP/ISMP) 및 IT 공공제안요청서(RFP) 분석 전문 수석 감리위원이자 수석 PMO입니다.
제공된 제안요청서(RFP) 문서 원문과 표 데이터를 정밀 분석하여, 향후 제안서 및 산출물 정합성 검증의 기준이 될 [단일 진실 원천(SSOT) 마스터 정보]를 한 단계 더 세부 내용까지 추출하십시오.

[제안요청서(RFP) 텍스트 발췌]:
${textSnippet}

${tableSnippets ? `[제안요청서 표 발췌]:\n${tableSnippets}\n` : ''}

[정밀 추출 지침]:
1. metadata:
   - projectName: 공식 사업명 (예: 2026년 공공보건의료 AI 기반 지능형 플랫폼 구축 ISP)
   - clientAgency: 발주기관/수요기관 공식 명칭 (예: 국립중앙의료원)
   - period: 사업기간 (예: 계약체결일로부터 8개월간 (2026.04 ~ 2026.11))
   - budget: 총사업비/배정예산 (숫자 및 한글 표기, 부가세 포함 여부)
2. requirements (★ 한 단계 더 세부적인 과업 내용까지 정밀 추출):
   - id: 요구사항 고유번호 (예: CNR-001, SFR-001, DAR-001 등)
   - category: 요구사항 분류 (예: 컨설팅 요구사항, 기능 요구사항, 데이터 요구사항, 보안 요구사항, 품질 요구사항, 프로젝트관리 요구사항 등)
   - name: ★ RFP 요구사항 정의서의 '요구사항 명칭' 필드에 명시된 공식 요구사항명을 그대로 사용. 절대 '요구사항 CNR-001'처럼 ID만 기재하지 말고, 반드시 RFP 원문의 '요구사항 명칭' 필드값 그대로 기재
   - description: RFP 요구사항 정의서의 '정의' 필드에 명시된 과업 상세 설명 (2~3문장으로 구체적 목표와 범위를 명확히 기술)
   - subItems: ★★ RFP '세부내용' 중 'ㅇ'(한글 이응) 또는 '○' 불릿으로 시작하는 각 문장을 하나의 독립 세부 요구사항으로 인식하여 배열로 추출. 'ㅇ' 아래의 '-' 소항목은 해당 ㅇ 항목에 괄호로 결합. 세부 항목 수 제한 없이 전수 추출
   - deliverables: 관련 산출물 (예: Ⅱ.3. 법제도 분석, Ⅲ.2. 업무 현황분석 등)
   - keywords: 필수 핵심 키워드 (쉼표 구분)
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
    {
      "id": "CNR-001",
      "category": "컨설팅 요구사항",
      "name": "공공의료 정보화 환경 및 현황 분석",
      "description": "공공의료 정보시스템 관련 법령·정책, 지침, 제도 현황 검토 및 분석, 내외부 정보기술 환경을 종합 분석하고, 문제점 및 시사점 도출",
      "subItems": [
        "(법제도 분석) 지역·필수공공보건의료와 AI 관련 법, 정책, 지침 현황 검토",
        "(환경분석) 현행 정책, 기술 환경과의 적합성 분석, 적용가능 범위 도출",
        "(사례분석) 국내외 정부, 지자체, 공공기관 유사 ISP 및 구축 사례 분석"
      ],
      "deliverables": "Ⅱ.1. 환경분석 개요, Ⅱ.3. 법제도 분석",
      "keywords": "공공의료, 정보화, 환경분석, 현황분석, 사례분석"
    }
  ],
  "glossary": [
    { "category": "기관명", "approved": "표준어", "forbidden": "오표기, 금지어", "note": "비고" }
  ]
}
`;

    onProgress({ percent: 50, message: 'Gemini AI로 마스터 메타데이터, 세부 과업 및 하위 항목 정밀 추출 중...' });

    const aiRes = await callGeminiProxyWithJson({
        apiKey,
        modelId,
        prompt,
        chunkIndex: 0
    });

    onProgress({ percent: 90, message: '추출 데이터 구조화 및 세부 항목 정합성 검증 중...' });

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
 * RFP 통합 분석 실행기 (규칙 기반 + AI 하이브리드 지원, 세부 과업 통합)
 */
export async function parseRfpContent({
    text,
    tables = [],
    apiKey = '',
    modelId = 'gemini-2.5-flash',
    useAI = false,
    onProgress = () => {}
}) {
    onProgress({ percent: 10, message: '결정론적 룰 기반 RFP 1차 파싱 및 세부항목 분석 중...' });
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

        const isPlaceholder = (name, id) => {
            if (!name) return true;
            const str = String(name).trim();
            return str.length < 3 || str === `요구사항 ${id}` || str === id || /^요구사항\s*[A-Z]{2,6}[-_]\d{2,4}$/i.test(str);
        };

        const reqMap = new Map();
        // 1) 룰 기반 요구사항 먼저 세팅
        (ruleResult.requirements || []).forEach(r => {
            if (r.id) reqMap.set(r.id.toUpperCase(), { ...r });
        });

        // 2) AI 요구사항으로 보강 (세부 설명, 하위 항목, 카테고리, 산출물 병합)
        (aiResult.requirements || []).forEach(aiReq => {
            if (!aiReq.id) return;
            const key = aiReq.id.toUpperCase();
            const existing = reqMap.get(key);

            let subItems = [];
            if (Array.isArray(aiReq.subItems)) {
                subItems = aiReq.subItems.map(s => String(s).trim()).filter(Boolean);
            } else if (typeof aiReq.subItems === 'string') {
                subItems = aiReq.subItems.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
            }

            if (!existing) {
                reqMap.set(key, {
                    id: key,
                    category: aiReq.category || '컨설팅 요구사항',
                    name: !isPlaceholder(aiReq.name, key) ? aiReq.name : `요구사항 ${key}`,
                    description: aiReq.description || '',
                    subItems: subItems,
                    deliverables: aiReq.deliverables || '',
                    keywords: aiReq.keywords || ''
                });
            } else {
                if (isPlaceholder(existing.name, key) && !isPlaceholder(aiReq.name, key)) {
                    existing.name = aiReq.name;
                }
                if (!existing.category && aiReq.category) existing.category = aiReq.category;
                if ((!existing.description || existing.description.length < (aiReq.description || '').length) && aiReq.description) {
                    existing.description = aiReq.description;
                }
                if ((!existing.subItems || existing.subItems.length === 0) && subItems.length > 0) {
                    existing.subItems = subItems;
                }
                if (!existing.deliverables && aiReq.deliverables) existing.deliverables = aiReq.deliverables;
                if (!existing.keywords && aiReq.keywords) existing.keywords = aiReq.keywords;
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

        onProgress({ percent: 100, message: 'Gemini AI 심층 분석 및 세부과업 마스터 추출 완료' });

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
                const existing = reqMap.get(key);
                reqMap.set(key, {
                    ...(existing || {}),
                    ...r,
                    subItems: (Array.isArray(r.subItems) && r.subItems.length > 0) ? r.subItems : (existing?.subItems || [])
                });
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
                    synonyms: g.synonyms || existing?.synonyms || '',
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
