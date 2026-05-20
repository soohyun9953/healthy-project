/**
 * ISMP 데이터 아키텍처 품질 검증 실제 모의 데이터셋
 * 최신 검토 보고서(IV.3.2.1-6, IV.3.2.7)의 실질적인 결함과 분석 팩트를 구조화
 */
export const ismpDaData = {
    complianceScore: 70, // 전체 컴플라이언스 초기 준수 점수
    
    // 12개 공식 주제영역 체계와 개념 구성도 정합성 갭
    subjectAreas: [
        { id: 1, name: "공통기반 데이터 영역", status: "matched", mappedName: "공통기반 (시스템관리 포함)" },
        { id: 2, name: "커뮤니케이션/협업 데이터 영역", status: "mismatched", mappedName: "커뮤니케이션" },
        { id: 3, name: "서비스/프로젝트/계약 데이터 영역", status: "omitted", mappedName: "누락" },
        { id: 4, name: "워크플로 자산/SOP 데이터 영역", status: "mismatched", mappedName: "워크플로" },
        { id: 5, name: "실험/ELN 데이터 영역", status: "omitted", mappedName: "누락 (NoSQL 구조 기획에 흡수됨)" },
        { id: 6, name: "결과제공/전달 데이터 영역", status: "omitted", mappedName: "누락" },
        { id: 7, name: "장비/스테이션 데이터 영역", status: "omitted", mappedName: "누락" },
        { id: 8, name: "바이오 객체 데이터 영역", status: "mismatched", mappedName: "바이오뱅크" },
        { id: 9, name: "자원/재고 운영 데이터 영역", status: "mismatched", mappedName: "재고관리" },
        { id: 10, name: "교육/지식 데이터 영역", status: "omitted", mappedName: "누락" },
        { id: 11, name: "AI분석 데이터 영역", status: "mismatched", mappedName: "AI 분석 가속도" },
        { id: 12, name: "시스템운영 데이터 영역", status: "omitted", mappedName: "누락" }
    ],

    // 개념 설계상 폴리글랏 DB 기획 대비 물리 설계 구현 갭
    polyglotDb: [
        { dbName: "PostgreSQL", conceptualRole: "RDBMS (재고, 기준정보, 관계형 트랜잭션)", physicalImplementation: "PostgreSQL 100% 물리 구현", isMismatch: false },
        { dbName: "MongoDB", conceptualRole: "NoSQL (데이터 레이크 원시 적재, 유연한 실험 워크플로)", physicalImplementation: "미반영 (일반 PostgreSQL 2D 테이블로 단순 매핑)", isMismatch: true },
        { dbName: "InfluxDB", conceptualRole: "시계열 DB (38종 실험 장비 IoT 로그 및 시계열 센서)", physicalImplementation: "미반영 (일반 PostgreSQL 2D 테이블로 단순 매핑)", isMismatch: true },
        { dbName: "Milvus", conceptualRole: "벡터 DB (RAG 유사도 및 대규모 인공지능 벡터 검색)", physicalImplementation: "미반영 (일반 PostgreSQL 2D 테이블로 단순 매핑)", isMismatch: true },
        { dbName: "Elasticsearch", conceptualRole: "대용량 검색 엔진 (RAG/LLM 연계 지식 텍스트 분석)", physicalImplementation: "미반영 (일반 PostgreSQL 2D 테이블로 단순 매핑)", isMismatch: true }
    ],

    // 괄호 짝 불일치 및 비표준 공백(NBSP) 검출 리스트
    typoIssues: [
        { id: 1, slide: "S006", originalText: "ISA-Tab (Investigation Study ...", correctedText: "ISA-Tab (Investigation Study ...)", errorType: "괄호 누락", isPatched: false },
        { id: 2, slide: "S013", originalText: "인 실리코2)", correctedText: "인 실리코[2]", errorType: "여는 괄호 누락 각주", isPatched: false },
        { id: 3, slide: "S013", originalText: "IUPAC1)/IUB", correctedText: "IUPAC[1]/IUB", errorType: "여는 괄호 누락 각주", isPatched: false },
        { id: 4, slide: "S015", originalText: "fasta_sequence\u00a0/\u00a0sbol_document", correctedText: "fasta_sequence / sbol_document", errorType: "NBSP 비표준 공백", isPatched: false },
        { id: 5, slide: "S016", originalText: "file_object\u00a0FK 참조 헤더", correctedText: "file_object FK 참조 헤더", errorType: "NBSP 비표준 공백", isPatched: false },
        { id: 6, slide: "S017", originalText: "감사 로그\u00a0\u00a0인덱스 설정", correctedText: "감사 로그  인덱스 설정", errorType: "NBSP 비표준 공백", isPatched: false },
        { id: 7, slide: "S050", originalText: "ACID1)", correctedText: "ACID[1]", errorType: "여는 괄호 누락 각주", isPatched: false },
        { id: 8, slide: "S054", originalText: "SBOM1) SBOM 스캔 수행", correctedText: "SBOM[1] SBOM 스캔 수행", errorType: "여는 괄호 누락 각주", isPatched: false },
        { id: 9, slide: "S055", originalText: "데이터 암호화2)", correctedText: "데이터 암호화[2]", errorType: "여는 괄호 누락 각주", isPatched: false }
    ],

    // 용어 혼용 빈도 데이터
    termMixUsage: [
        {
            termGroup: "저장소 인프라 표기",
            variants: [
                { term: "Storage (영문)", count: 15 },
                { term: "저장소 (한글 순화)", count: 17 },
                { term: "스토리지 (한글 발음)", count: 14 }
            ]
        },
        {
            termGroup: "데이터베이스 표기",
            variants: [
                { term: "DB (영문 약어)", count: 86 },
                { term: "데이터베이스 (한글 명칭)", count: 12 },
                { term: "DBMS (시스템 통칭)", count: 12 }
            ]
        },
        {
            termGroup: "고가용성 아키텍처 표기",
            variants: [
                { term: "HA (영문 약어)", count: 32 },
                { term: "고가용성 (한글 명칭)", count: 1 }
            ]
        }
    ]
};
