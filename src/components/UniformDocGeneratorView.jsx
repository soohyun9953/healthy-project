import React, { useState, useEffect, useMemo, useRef } from 'react';
import {
    Layers,
    FileSpreadsheet,
    FileText,
    Presentation,
    Upload,
    Download,
    Plus,
    Trash2,
    Sparkles,
    Palette,
    RefreshCw,
    CheckCircle2,
    Eye,
    ChevronDown,
    ChevronUp,
    ExternalLink,
    Grid,
    Sliders,
    Monitor,
    ShieldAlert,
    Copy,
    ArrowRight
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import {
    THEME_PRESETS,
    DEFAULT_DESIGN_TOKENS,
    loadDesignTokens,
    saveDesignTokens
} from '../utils/designTokens';
import {
    LAYOUT_TYPES,
    LAYOUT_LABELS,
    buildPptxDeck,
    buildDocxDocument,
    buildXlsxWorkbook
} from '../utils/documentBuilders';

// ── 공공의료/ISP 표준 8대 핵심 목차 기본 프리셋 ────────────────────
export const DEFAULT_TOC_SLIDES = [
    {
        id: 'slide-1',
        chapter: 'Ⅰ. 사업 개요',
        section: '1. 추진 배경 및 필요성',
        subsection: '1.1 대내외 환경 변화',
        title: '공공보건의료 데이터 통합의 정책적 시급성 및 당위성',
        reqId: 'PMR-001',
        owner: '김수현 PM',
        layoutType: LAYOUT_TYPES.TYPE_B_BEFORE_AFTER,
        keywords: '데이터통합, 정책시급성, 디지털헬스케어, 전원환자',
        headMessage: '지방의료원 간 데이터 고립으로 환자 전원 시 진료 차질 및 중복 검사가 빈번하게 발생하고 있습니다.\n국가 보건의료 표준(FHIR) 기반의 실시간 데이터 연계 체계를 구축하여 필수의료 안전망을 완결해야 합니다.',
        asIsBullets: [
            '의료원별 이기종 EMR 운영으로 환자 전원 시 의무기록 수기 복사 전달',
            '원내 통계 수기 집계로 공공의료 정책 수립을 위한 데이터 수집 3주 지연',
            '권역 책임의료기관과 지역의료원 간 실시간 병상 자원 정보 공유 불가'
        ],
        toBeBullets: [
            '국가 표준 진료정보교류(FHIR) 연계로 전국 공공의료원 데이터 실시간 동기화',
            '중앙 데이터 레이크 구축으로 전국 공공의료 임상·경영 지표 실시간 시각화',
            '실시간 응급의료 관제 대시보드 구축으로 골든타임 환자 즉시 이송 지원'
        ],
        deltaText: '🎯 [기대성과] 환자 재검사율 28% 감소 | 응급환자 골든타임 확보율 35% 향상 | 연간 행정비용 45억원 절감'
    },
    {
        id: 'slide-2',
        chapter: 'Ⅰ. 사업 개요',
        section: '2. 목표 및 추진 전략',
        subsection: '2.1 단계별 추진 로드맵',
        title: '지능형 공공보건의료 플랫폼 4단계 연차별 이행 로드맵',
        reqId: 'PMR-001',
        owner: '이진우 수석',
        layoutType: LAYOUT_TYPES.TYPE_C_PROCESS,
        keywords: '추진전략, 단계별로드맵, 클라우드, EMR연계',
        headMessage: '단위 시스템 중심의 점진적 전환을 통해 대규모 구축에 따른 현장 혼선을 최소화합니다.\n1단계 기반 인프라 구축부터 4단계 전국 확산까지 8개월에 걸쳐 체계적으로 추진합니다.',
        steps: [
            { num: 'Step 1', title: '기반 분석 및 진단', period: '1~2개월', desc: '• 레거시 EMR 연계성 평가\n• 개인정보보호 요건 분석\n• 표준화 대상 항목 도출' },
            { num: 'Step 2', title: '목표 아키텍처 설계', period: '3~4개월', desc: '• 클라우드 인프라 설계\n• FHIR 데이터 모델 정의\n• 보안 거버넌스 수립' },
            { num: 'Step 3', title: '플랫폼 구축 및 실증', period: '5~6개월', desc: '• AI CDSS 알고리즘 탑재\n• 응급관제 대시보드 구축\n• 5개 선도병원 시범 실증' },
            { num: 'Step 4', title: '전국 확산 및 안정화', period: '7~8개월', desc: '• 전국 35개 의료원 연계\n• 운영 자동화 및 감리\n• 최종 산출물 완성' }
        ]
    },
    {
        id: 'slide-3',
        chapter: 'Ⅱ. 현황 및 문제점 분석',
        section: '1. 업무 및 기술 현황',
        subsection: '1.1 EMR 연계 현황 분석',
        title: '전국 지방의료원 EMR 시스템 연계 현황 및 표준화 수준 진단',
        reqId: 'CNR-001',
        owner: '박서연 책임',
        layoutType: LAYOUT_TYPES.TYPE_A_TABLE,
        keywords: 'EMR, 표준화, 상호운용성, 인터페이스',
        headMessage: '전국 35개 지방의료원 중 28개소가 독자적인 코드 및 벤더 솔루션을 사용하고 있어 연계성이 극히 취약합니다.\n국가 표준 인터페이스 가이드라인을 수립하고 공통 매핑 테이블을 일괄 배포해야 합니다.',
        keyPoints: [
            '의료원별 EMR 버전 상이로 진료 서식 및 임상 용어의 일관성 결여',
            'RESTful API 기반 표준 게이트웨이 부재로 직접 DB 링크에 따른 보안 리스크 노출',
            '보건복지부 전자의무기록(EMR) 인증 획득 비율 42% 수준으로 기술 고도화 시급'
        ],
        tableData: {
            headers: ['구분', '현행(As-Is) 수준', '목표(To-Be) 모델', '표준 적용 방안', '비고'],
            rows: [
                ['데이터 표준', '의료원별 독자 코드체계 사용', '보건의료표준용어 및 FHIR 적용', '국가 표준 코드 매핑 테이블 구축', '필수준수'],
                ['연계 방식', '수기 추출 및 배치 파일 전송', '실시간 API 게이트웨이 연계', 'RESTful API 및 EAI 표준화', '무중단'],
                ['보안 거버넌스', '원내 자체 보안 기준 적용', '국가정보원 및 개인정보보호법 엄수', '가명정보 처리 및 전송구간 암호화', '규제대응'],
                ['유지관리 체계', '개별 유지보수 계약 체결', '중앙 통합 관제 및 표준 SLA 적용', '통합 헬프데스크 및 기술지원', '비용절감']
            ]
        }
    },
    {
        id: 'slide-4',
        chapter: 'Ⅱ. 현황 및 문제점 분석',
        section: '2. 환경 분석 및 시사점',
        subsection: '2.1 SWOT 분석',
        title: '공공의료 AI 플랫폼 도입을 위한 내·외부 환경(SWOT) 전략 분석',
        reqId: 'CNR-001',
        owner: '김수현 PM',
        layoutType: LAYOUT_TYPES.TYPE_E_MATRIX,
        keywords: 'SWOT, 환경분석, 강점약점, 기회위협',
        headMessage: '공공의료원 네트워크라는 강력한 내부 자산을 바탕으로 정부 디지털 헬스케어 정책 기회를 극대화합니다.\n취약한 전산 인력과 보안 리스크는 중앙 클라우드 통합 거버넌스로 극복해야 합니다.',
        matrixData: [
            { title: '강점 (Strength)', color: '1E3A8A', items: ['전국 35개 공공의료원 단일 네트워크', '오랜 기간 축적된 임상 코호트 데이터', '공공보건의료 지원단 협력 체계'] },
            { title: '약점 (Weakness)', color: 'DC2626', items: ['의료원별 IT 예산 및 전담 인력 부족', '레거시 시스템 노후화로 인한 확장 한계', '의료진의 신규 디지털 도구 수용 저항'] },
            { title: '기회 (Opportunity)', color: '16A34A', items: ['정부 AI 일상화 및 디지털 헬스케어 육성', '공공 클라우드 전환 지원 및 예산 확대', '지역 필수의료 강화에 대한 사회적 공감대'] },
            { title: '위협 (Threat)', color: 'D97706', items: ['의료법 및 민감 개인정보 유출 리스크', '민간 대형병원과의 기술 격차 심화', '유지관리 예산 지속성 확보 불확실성'] }
        ]
    },
    {
        id: 'slide-5',
        chapter: 'Ⅲ. 목표 모델 수립',
        section: '1. 목표 아키텍처',
        subsection: '1.1 인프라 구성 방안',
        title: '하이브리드 클라우드 기반 목표 인프라 및 보안 아키텍처',
        reqId: 'TER-001',
        owner: '최민호 이사',
        layoutType: LAYOUT_TYPES.TYPE_D_2COL_CARDS,
        keywords: '클라우드, 아키텍처, 보안, 프레임워크',
        headMessage: '민감 진료 데이터는 원내 가명화 존에서 보호하고, AI 연산 및 빅데이터 분석은 공공 클라우드 존에서 수행합니다.\n보안성과 확장성을 동시에 충족하는 하이브리드 멀티클라우드 체계를 가동합니다.',
        leftCard: {
            title: '인프라 및 플랫폼 아키텍처 요건',
            items: [
                '공공 클라우드(CSAP 인증) PaaS 기반 데이터 레이크 구성',
                '쿠버네티스(K8s) 기반 마이크로서비스(MSA) 무중단 배포 체계',
                '원내 레거시 EMR과 클라우드 간 전용 암호화 통신망(VPN) 구성'
            ]
        },
        rightCard: {
            title: '보안 및 데이터 거버넌스 요건',
            items: [
                '개인정보보호법 제28조의2 준수 가명정보 결합 전문기관 연계',
                '접근 제어 및 감사 로그 위변조 방지 WORM 스토리지 적용',
                '24시간 365일 통합 보안관제(SOC) 및 재해복구(DR) 이중화'
            ]
        }
    },
    {
        id: 'slide-6',
        chapter: 'Ⅲ. 목표 모델 수립',
        section: '2. 서비스 기능 정의',
        subsection: '2.1 AI 진료지원 기능',
        title: 'AI 기반 임상의사결정지원(CDSS) 핵심 알고리즘 사양 및 연계 방안',
        reqId: 'SFR-001',
        owner: '박서연 책임',
        layoutType: LAYOUT_TYPES.TYPE_A_TABLE,
        keywords: 'AI, CDSS, 진료지원, 패혈증, 흉부Xray',
        headMessage: '지방의료원의 전문의 부족을 보완하기 위해 3대 필수의료 영역에 AI 알고리즘을 우선 배포합니다.\n의료진의 진단 정확도를 18% 향상시키고 위중증 전원 골든타임을 확보합니다.',
        keyPoints: [
            '급성 패혈증 및 심정지 조기 예측 바이탈 사인 모니터링 알고리즘 내장',
            '전공의 판독 지원을 위한 인공지능 기반 흉부 X-ray 1차 선별 판독 서비스',
            'EMR 처방 오더 입력 시 약물 상호작용(DUR) 및 부작용 실시간 알림'
        ],
        tableData: {
            headers: ['알고리즘 명칭', '적용 임상 영역', '핵심 입력 데이터', '예측 성능(AUROC)', '도입 효과'],
            rows: [
                ['패혈증 조기예측(qSOFA)', '중환자실·응급실', '혈압, 맥박, 체온, 산소포화도', '0.92 이상', '패혈증 사망률 24% 감소'],
                ['AI 흉부영상 판독', '영상의학과·호흡기내과', '디지털 X-ray DICOM 이미지', '0.94 이상', '판독 대기시간 70% 단축'],
                ['심정지 위험도 예측', '일반 병동 입원환자', '연령, 기저질환, 12유도 심전도', '0.89 이상', '원내 급성 심정지 사전 예방'],
                ['약물 부작용 예측', '전 임상과 외래·입원', '환자 복용 이력, 알레르기 정보', '0.95 이상', '처방 오류율 0% 수렴']
            ]
        }
    },
    {
        id: 'slide-7',
        chapter: 'Ⅳ. 세부 이행계획',
        section: '1. 소요 예산 산정',
        subsection: '1.1 총구축비 내역',
        title: '총구축비 세부 비목별 및 3개년 연차별 소요 예산안',
        reqId: 'CNR-001',
        owner: '김수현 PM',
        layoutType: LAYOUT_TYPES.TYPE_A_TABLE,
        keywords: '예산, 총구축비, SW개발비, 연차별투자',
        headMessage: '기획재정부 ISP 공통가이드(제9판) 및 SW사업 대가산정 가이드라인을 준수하여 총 50억원을 산출했습니다.\n1차년도 플랫폼 구축 25억원, 2차년도 전국 확산 18억원, 3차년도 고도화 7억원으로 분할 투입합니다.',
        keyPoints: [
            '기능점수(FP) 간이법 기반 맞춤형 SW 개발비 22.5억원(45%) 정밀 산정',
            '클라우드 인프라 이용료 및 CSAP 보안 인증 솔루션 15.5억원(31%) 배정',
            '사업관리 PMO 및 3단계 감리비, 데이터 표준화 컨설팅 12.0억원(24%) 반영'
        ],
        tableData: {
            headers: ['비목 구분', '1차년도(구축)', '2차년도(확산)', '3차년도(고도화)', '총합계', '산출 근거'],
            rows: [
                ['SW 개발비', '1,200 백만원', '800 백만원', '250 백만원', '2,250 백만원', '기능점수(FP) 4,200 FP 산출'],
                ['클라우드/인프라', '750 백만원', '550 백만원', '250 백만원', '1,550 백만원', 'CSAP 클라우드 인프라 및 DB'],
                ['데이터 표준화', '300 백만원', '250 백만원', '100 백만원', '650 백만원', 'FHIR 코드 매핑 및 클렌징'],
                ['감리 및 PMO', '250 백만원', '200 백만원', '100 백만원', '550 백만원', '행안부 감리기준 및 PMO 대가'],
                ['총 합계', '2,500 백만원', '1,800 백만원', '700 백만원', '5,000 백만원', '총사업비 지침 100% 준수']
            ]
        }
    },
    {
        id: 'slide-8',
        chapter: 'Ⅳ. 세부 이행계획',
        section: '2. 기대 효과 분석',
        subsection: '2.1 정량·정성적 성과',
        title: '지능형 공공보건의료 플랫폼 도입에 따른 기대 성과 및 정책적 효과',
        reqId: 'SFR-002',
        owner: '이진우 수석',
        layoutType: LAYOUT_TYPES.TYPE_B_BEFORE_AFTER,
        keywords: '기대효과, 정량적성과, 정성적효과, 국민체감',
        headMessage: '지역 간 필수의료 격차를 해소하고 응급환자의 재이송(일명 응급실 뺑뺑이)을 근본적으로 차단합니다.\n국민 의료비 부담 완화 및 공공의료 신뢰도 회복의 획기적 계기를 마련합니다.',
        asIsBullets: [
            '지역별 중증환자 타 지역 유출률 38%로 지역의료 공동화 가속',
            '전원 지연으로 인한 응급의료 사망률 OECD 평균 대비 1.4배',
            '지방의료원 의료수익 적자 누적 및 진료 인프라 노후화'
        ],
        toBeBullets: [
            'AI 진료지원으로 지역 내 중증환자 완결 치료율 78%로 상향',
            '실시간 병상 관제로 응급환자 골든타임 내 수술 개시율 90% 달성',
            '공공의료 질 향상으로 연간 1,200억원 규모 사회경제적 편익 창출'
        ],
        deltaText: '🎯 [종합 효과] B/C 분석 1.34 달성 | 경제적 타당성 확보 | 지역완결형 필수의료 안전망 완성'
    }
];

export default function UniformDocGeneratorView({ apiKey, llmProvider, omniRouteModel }) {
    // ── 디자인 토큰 상태 ─────────────────────────────────────────
    const [tokens, setTokens] = useState(() => loadDesignTokens());
    const [isTokenPanelOpen, setIsTokenPanelOpen] = useState(false);

    // ── 목차 슬라이드 목록 상태 ──────────────────────────────────
    const [slides, setSlides] = useState(() => {
        try {
            const raw = localStorage.getItem('healthy_project_toc_slides_v1');
            if (raw) return JSON.parse(raw);
        } catch (_) {}
        return DEFAULT_TOC_SLIDES;
    });

    // 선택된 슬라이드 (미리보기 대상)
    const [selectedSlideId, setSelectedSlideId] = useState(slides[0]?.id || 'slide-1');
    const [isAiLoading, setIsAiLoading] = useState(false);
    const [isExporting, setIsExporting] = useState(false);

    // 파일 업로드 Ref
    const excelImportRef = useRef(null);

    // 변경 시 LocalStorage 저장
    useEffect(() => {
        saveDesignTokens(tokens);
    }, [tokens]);

    useEffect(() => {
        try {
            localStorage.setItem('healthy_project_toc_slides_v1', JSON.stringify(slides));
        } catch (_) {}
    }, [slides]);

    // 현재 선택된 슬라이드 객체
    const activeSlide = useMemo(() => {
        return slides.find(s => s.id === selectedSlideId) || slides[0] || {};
    }, [slides, selectedSlideId]);

    // ── 테마 프리셋 선택 핸들러 ───────────────────────────────────
    const handleSelectPreset = (presetId) => {
        const preset = THEME_PRESETS[presetId];
        if (!preset) return;
        setTokens(prev => ({
            ...prev,
            selectedPresetId: presetId,
            colors: { ...preset.colors }
        }));
    };

    // 토큰 커스텀 핸들러
    const handleColorChange = (key, val) => {
        setTokens(prev => ({
            ...prev,
            colors: { ...prev.colors, [key]: val.replace('#', '') }
        }));
    };

    const handleLayoutSettingChange = (field, val) => {
        setTokens(prev => ({
            ...prev,
            pageLayout: { ...prev.pageLayout, [field]: val }
        }));
    };

    // ── 목차 슬라이드 조작 핸들러 ────────────────────────────────
    const handleAddSlide = () => {
        const newSlide = {
            id: `slide-${Date.now()}`,
            chapter: 'Ⅲ. 세부 추진계획',
            section: '1. 세부 과제 정의',
            subsection: '1.1 신규 과제',
            title: `신규 추진 과제 ${slides.length + 1}`,
            reqId: 'SFR-001',
            owner: '홍길동 수석',
            layoutType: LAYOUT_TYPES.TYPE_A_TABLE,
            keywords: '신규과제, 아키텍처, 표준화',
            headMessage: '본 장표는 신규 정보화 세부 과제의 핵심 추진 내용과 표준화 사양을 정의합니다.\n현행 시스템과의 연계성을 확보하고 안정적인 서비스 전환을 도모합니다.'
        };
        setSlides(prev => [...prev, newSlide]);
        setSelectedSlideId(newSlide.id);
    };

    const handleDeleteSlide = (idToDelete) => {
        if (slides.length <= 1) {
            alert('최소 1개 이상의 장표 목차가 유지되어야 합니다.');
            return;
        }
        setSlides(prev => {
            const next = prev.filter(s => s.id !== idToDelete);
            if (selectedSlideId === idToDelete && next.length > 0) {
                setSelectedSlideId(next[0].id);
            }
            return next;
        });
    };

    const handleUpdateSlideField = (id, field, value) => {
        setSlides(prev => prev.map(s => s.id === id ? { ...s, [field]: value } : s));
    };

    const handleResetTocPreset = () => {
        if (window.confirm('공공의료/ISP 표준 8대 핵심 목차 기본 프리셋으로 초기화하시겠습니까?')) {
            setSlides(DEFAULT_TOC_SLIDES);
            setSelectedSlideId(DEFAULT_TOC_SLIDES[0].id);
        }
    };

    // ── 목차 엑셀 (.xlsx) 내보내기 / 가져오기 ────────────────────
    const handleExportTocExcel = () => {
        const rows = slides.map((s, idx) => ({
            '순번': idx + 1,
            '대장': s.chapter || '',
            '중절': s.section || '',
            '소항': s.subsection || '',
            '장표제목': s.title || '',
            '요구사항ID': s.reqId || '',
            '담당자': s.owner || '',
            '레이아웃타입': s.layoutType || LAYOUT_TYPES.TYPE_A_TABLE,
            '핵심키워드': s.keywords || '',
            '헤드메시지(거버닝)': s.headMessage || ''
        }));
        const ws = XLSX.utils.json_to_sheet(rows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, '장표목차스캐폴딩');
        XLSX.writeFile(wb, `장표_스캐폴딩_목차엑셀_${new Date().toISOString().slice(0, 10)}.xlsx`);
    };

    const handleImportTocExcel = (e) => {
        const file = e.target.files?.[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (evt) => {
            try {
                const data = new Uint8Array(evt.target.result);
                const wb = XLSX.read(data, { type: 'array' });
                const ws = wb.Sheets[wb.SheetNames[0]];
                const json = XLSX.utils.sheet_to_json(ws);

                const imported = json.map((r, i) => ({
                    id: `slide-imp-${Date.now()}-${i}`,
                    chapter: r['대장'] || 'Ⅰ. 사업 개요',
                    section: r['중절'] || '1. 추진 배경',
                    subsection: r['소항'] || '',
                    title: r['장표제목'] || `가져온 장표 #${i + 1}`,
                    reqId: r['요구사항ID'] || 'CNR-001',
                    owner: r['담당자'] || '담당자',
                    layoutType: r['레이아웃타입'] || LAYOUT_TYPES.TYPE_A_TABLE,
                    keywords: r['핵심키워드'] || '',
                    headMessage: r['헤드메시지(거버닝)'] || r['헤드메시지'] || '본 장표는 세부 이행방안을 표준 양식에 따라 명시합니다.'
                }));

                if (imported.length > 0) {
                    setSlides(imported);
                    setSelectedSlideId(imported[0].id);
                    alert(`총 ${imported.length}건의 목차 장표가 성공적으로 로드되었습니다.`);
                }
            } catch (err) {
                console.error(err);
                alert('목차 엑셀 파싱 중 오류가 발생했습니다.');
            }
        };
        reader.readAsArrayBuffer(file);
        e.target.value = '';
    };

    // ── AI 초안 자동 채우기 (/api/gemini-proxy 연동) ──────────────
    const handleRunAiAutoFill = async () => {
        if (!apiKey) {
            alert('AI 초안 생성을 위해 상단에 Gemini API Key를 먼저 입력해 주세요.');
            return;
        }

        setIsAiLoading(true);
        try {
            // 현재 활성화된 슬라이드의 제목, 키워드, 레이아웃을 기반으로 초안 생성
            const prompt = `당신은 공공기관 ISP/ISMP 전문 컨설팅 수석 감리위원입니다.
아래 장표 제목과 핵심 키워드를 바탕으로 고품질 컨설팅 장표의 세부 내용(헤드메시지, 불릿 요약 포인트, 표 데이터 또는 카드 데이터)을 작성해 주십시오.

[장표 정보]
- 프로젝트명: ${tokens.pageLayout.projectName}
- 대장/중절: ${activeSlide.chapter} > ${activeSlide.section}
- 장표 제목: ${activeSlide.title}
- 요구사항 ID: ${activeSlide.reqId}
- 레이아웃 타입: ${activeSlide.layoutType}
- 핵심 키워드: ${activeSlide.keywords}

[출력 요구사항]:
반드시 아래 JSON 스키마를 만족하는 단일 JSON 객체로만 응답하십시오. (마크다운 백틱 없이 순수 JSON만 출력)
{
  "headMessage": "핵심을 관통하는 거버닝 메시지 (2줄 문장, 굵직한 결론 제시)",
  "keyPoints": [
    "핵심 시사점/이행 포인트 1 (구체적 수치 포함)",
    "핵심 시사점/이행 포인트 2",
    "핵심 시사점/이행 포인트 3"
  ],
  "asIsBullets": ["현행 문제점 1", "현행 문제점 2", "현행 문제점 3"],
  "toBeBullets": ["목표 혁신안 1", "목표 혁신안 2", "목표 혁신안 3"],
  "deltaText": "🎯 [기대성과] 정량적 지표 개선치 및 정책적 편익",
  "tableData": {
    "headers": ["구분", "현행 수준", "목표 모델", "추진 방안", "비고"],
    "rows": [
      ["항목 1", "내용...", "내용...", "내용...", "비고..."],
      ["항목 2", "내용...", "내용...", "내용...", "비고..."],
      ["항목 3", "내용...", "내용...", "내용...", "비고..."]
    ]
  }
}`;

            const response = await fetch('/api/gemini-proxy', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    apiKey,
                    modelId: 'gemini-2.5-flash',
                    contents: [{ role: 'user', parts: [{ text: prompt }] }],
                    generationConfig: { temperature: 0.2, responseMimeType: 'application/json' }
                })
            });

            let jsonText = '';
            if (response.ok) {
                const resData = await response.json();
                jsonText = resData.candidates?.[0]?.content?.parts?.[0]?.text || '';
            } else {
                // 직접 호출 폴백
                const directUrl = `https://generativelanguage.googleapis.com/v1beta/gemini-2.5-flash:generateContent`;
                const directRes = await fetch(directUrl, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json', 'x-goog-api-key': apiKey },
                    body: JSON.stringify({
                        contents: [{ role: 'user', parts: [{ text: prompt }] }],
                        generationConfig: { temperature: 0.2, responseMimeType: 'application/json' }
                    })
                });
                const directData = await directRes.json();
                jsonText = directData.candidates?.[0]?.content?.parts?.[0]?.text || '';
            }

            const cleaned = jsonText.replace(/```json/gi, '').replace(/```/g, '').trim();
            const parsed = JSON.parse(cleaned);

            // 현재 선택된 슬라이드 업데이트
            setSlides(prev => prev.map(s => {
                if (s.id === activeSlide.id) {
                    return {
                        ...s,
                        headMessage: parsed.headMessage || s.headMessage,
                        keyPoints: parsed.keyPoints || s.keyPoints,
                        asIsBullets: parsed.asIsBullets || s.asIsBullets,
                        toBeBullets: parsed.toBeBullets || s.toBeBullets,
                        deltaText: parsed.deltaText || s.deltaText,
                        tableData: parsed.tableData || s.tableData
                    };
                }
                return s;
            }));

            alert(`'${activeSlide.title}' 장표의 AI 세부 초안이 성공적으로 생성·반영되었습니다!`);
        } catch (err) {
            console.error('AI 생성 오류:', err);
            alert(`AI 초안 생성 중 오류가 발생했습니다: ${err.message}`);
        } finally {
            setIsAiLoading(false);
        }
    };

    // ── 동일 디자인 포맷 일괄 내보내기 (PPTX, DOCX, XLSX) ─────────
    const handleExportSingleFormat = async (format) => {
        setIsExporting(true);
        const fileNameBase = `표준산출물_${tokens.pageLayout.projectName || '스캐폴딩'}_${new Date().toISOString().slice(0, 10)}`;

        try {
            if (format === 'pptx') {
                const blob = await buildPptxDeck(slides, tokens);
                saveAs(blob, `${fileNameBase}.pptx`);
            } else if (format === 'docx') {
                const blob = await buildDocxDocument(slides, tokens);
                saveAs(blob, `${fileNameBase}.docx`);
            } else if (format === 'xlsx') {
                const blob = await buildXlsxWorkbook(slides, tokens);
                saveAs(blob, `${fileNameBase}.xlsx`);
            }
        } catch (err) {
            console.error(`${format} 내보내기 실패:`, err);
            alert(`${format.toUpperCase()} 파일 생성 중 오류가 발생했습니다: ${err.message}`);
        } finally {
            setIsExporting(false);
        }
    };

    const handleExportAllFormats = async () => {
        setIsExporting(true);
        const fileNameBase = `표준산출물_${tokens.pageLayout.projectName || '통합'}_${new Date().toISOString().slice(0, 10)}`;

        try {
            // 1. PPTX
            const pptxBlob = await buildPptxDeck(slides, tokens);
            saveAs(pptxBlob, `${fileNameBase}.pptx`);

            // 2. DOCX
            const docxBlob = await buildDocxDocument(slides, tokens);
            saveAs(docxBlob, `${fileNameBase}.docx`);

            // 3. XLSX
            const xlsxBlob = await buildXlsxWorkbook(slides, tokens);
            saveAs(xlsxBlob, `${fileNameBase}.xlsx`);

            alert('PPTX, DOCX, XLSX 3종 표준 문서가 동일 디자인 토큰 규격으로 일괄 생성되어 다운로드되었습니다!');
        } catch (err) {
            console.error('일괄 내보내기 오류:', err);
            alert('일괄 내보내기 중 오류가 발생했습니다.');
        } finally {
            setIsExporting(false);
        }
    };

    return (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
            {/* 상단 타이틀 & 디자인 토큰 프리셋 퀵 바 */}
            <div className="glass-panel" style={{ padding: '20px 24px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '16px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '14px' }}>
                    <div style={{
                        width: '46px',
                        height: '46px',
                        borderRadius: '12px',
                        background: `linear-gradient(135deg, #${tokens.colors.primary}33, #${tokens.colors.secondary}33)`,
                        border: `1px solid #${tokens.colors.primary}88`,
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        color: `#${tokens.colors.primary}`
                    }}>
                        <Layers size={26} />
                    </div>
                    <div>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                            <h2 style={{ fontSize: '18px', fontWeight: 700, margin: 0, color: 'var(--text-primary)' }}>
                                통합 디자인 토큰 & 목차 기반 표준 산출물 일괄 스캐폴딩 생성기
                            </h2>
                            <span style={{
                                padding: '2px 8px',
                                borderRadius: '12px',
                                fontSize: '11px',
                                fontWeight: 700,
                                background: `rgba(59, 130, 246, 0.15)`,
                                color: 'var(--accent-blue)',
                                border: '1px solid rgba(59, 130, 246, 0.3)'
                            }}>
                                PPTX · DOCX · XLSX 일치화
                            </span>
                        </div>
                        <p style={{ margin: '4px 0 0', fontSize: '13px', color: 'var(--text-secondary)' }}>
                            단일 디자인 토큰을 공유하여 5대 표준 컨설팅 장표(테이블, As-Is/To-Be, 프로세스, 2단 카드, SWOT)를 3종 문서로 일괄 빌드합니다.
                        </p>
                    </div>
                </div>

                <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                    <button
                        onClick={handleResetTocPreset}
                        className="btn-outline"
                        style={{ padding: '8px 14px', fontSize: '12.5px', display: 'flex', alignItems: 'center', gap: '6px' }}
                        title="기본 8대 표준 목차로 초기화"
                    >
                        <RefreshCw size={14} /> 목차 프리셋 복원
                    </button>
                    <button
                        onClick={() => setIsTokenPanelOpen(!isTokenPanelOpen)}
                        className="btn-primary"
                        style={{
                            padding: '8px 16px',
                            fontSize: '13px',
                            display: 'flex',
                            alignItems: 'center',
                            gap: '6px',
                            background: isTokenPanelOpen ? 'var(--accent-purple)' : undefined
                        }}
                    >
                        <Palette size={15} />
                        {isTokenPanelOpen ? '디자인 토큰 닫기' : '디자인 토큰 테마 설정'}
                        {isTokenPanelOpen ? <ChevronUp size={15} /> : <ChevronDown size={15} />}
                    </button>
                </div>
            </div>

            {/* 통합 디자인 토큰 설정 패널 (접기/펼치기) */}
            {isTokenPanelOpen && (
                <div className="glass-panel" style={{ padding: '20px', display: 'flex', flexDirection: 'column', gap: '18px', border: `1px solid #${tokens.colors.primary}55` }}>
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: '12px' }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                            <Sliders size={18} style={{ color: 'var(--accent-blue)' }} />
                            <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 700, color: 'var(--text-primary)' }}>
                                표준 테마 프리셋 3종 선택
                            </h3>
                        </div>
                        {/* 16:9 와이드 vs 4:3 A4 라디오 */}
                        <div style={{ display: 'flex', alignItems: 'center', gap: '12px', background: 'rgba(0,0,0,0.2)', padding: '4px 12px', borderRadius: '8px', border: '1px solid var(--panel-border)' }}>
                            <span style={{ fontSize: '12px', fontWeight: 600, color: 'var(--text-muted)' }}>슬라이드 비율:</span>
                            <label style={{ display: 'flex', alignItems: 'center', gap: '4px', fontSize: '12px', cursor: 'pointer', color: 'var(--text-primary)' }}>
                                <input
                                    type="radio"
                                    name="aspectRatio"
                                    value="16:9"
                                    checked={tokens.pageLayout.aspectRatio === '16:9'}
                                    onChange={() => handleLayoutSettingChange('aspectRatio', '16:9')}
                                />
                                16:9 와이드 (표준)
                            </label>
                            <label style={{ display: 'flex', alignItems: 'center', gap: '4px', fontSize: '12px', cursor: 'pointer', color: 'var(--text-primary)' }}>
                                <input
                                    type="radio"
                                    name="aspectRatio"
                                    value="4:3"
                                    checked={tokens.pageLayout.aspectRatio === '4:3'}
                                    onChange={() => handleLayoutSettingChange('aspectRatio', '4:3')}
                                />
                                4:3 (인쇄·보고서형)
                            </label>
                        </div>
                    </div>

                    {/* 테마 프리셋 카드 3종 */}
                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: '14px' }}>
                        {Object.values(THEME_PRESETS).map(preset => {
                            const isSelected = tokens.selectedPresetId === preset.id;
                            return (
                                <div
                                    key={preset.id}
                                    onClick={() => handleSelectPreset(preset.id)}
                                    style={{
                                        padding: '16px',
                                        borderRadius: '10px',
                                        cursor: 'pointer',
                                        border: `2px solid ${isSelected ? `#${preset.colors.primary}` : 'var(--panel-border)'}`,
                                        background: isSelected ? 'rgba(59, 130, 246, 0.08)' : 'rgba(255,255,255,0.02)',
                                        transition: 'all 0.2s ease',
                                        display: 'flex',
                                        flexDirection: 'column',
                                        gap: '10px'
                                    }}
                                >
                                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                                        <span style={{ fontSize: '13.5px', fontWeight: 700, color: 'var(--text-primary)' }}>
                                            {preset.name}
                                        </span>
                                        {isSelected && (
                                            <span style={{ fontSize: '11px', fontWeight: 700, color: '#10b981', display: 'flex', alignItems: 'center', gap: '3px' }}>
                                                <CheckCircle2 size={13} /> 적용중
                                            </span>
                                        )}
                                    </div>
                                    <p style={{ margin: 0, fontSize: '12px', color: 'var(--text-secondary)' }}>
                                        {preset.description}
                                    </p>
                                    {/* 색상 팔레트 미리보기 칩 */}
                                    <div style={{ display: 'flex', gap: '6px', alignItems: 'center', marginTop: '4px' }}>
                                        <div style={{ width: '24px', height: '24px', borderRadius: '4px', background: `#${preset.colors.primary}` }} title="Primary" />
                                        <div style={{ width: '24px', height: '24px', borderRadius: '4px', background: `#${preset.colors.secondary}` }} title="Secondary" />
                                        <div style={{ width: '24px', height: '24px', borderRadius: '4px', background: `#${preset.colors.accent}` }} title="Accent" />
                                        <div style={{ width: '24px', height: '24px', borderRadius: '4px', background: `#${preset.colors.keyPointBg}`, border: '1px solid #ccc' }} title="KeyPoint Bg" />
                                        <div style={{ width: '24px', height: '24px', borderRadius: '4px', background: `#${preset.colors.tableHeaderBg}`, border: '1px solid #ccc' }} title="Table Header" />
                                    </div>
                                </div>
                            );
                        })}
                    </div>

                    {/* 세부 커스텀 입력 (프로젝트명, 폰트, 기본 색상) */}
                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '14px', paddingTop: '10px', borderTop: '1px solid var(--panel-border)' }}>
                        <div>
                            <label style={{ display: 'block', fontSize: '12px', fontWeight: 600, color: 'var(--text-muted)', marginBottom: '4px' }}>
                                상단 고정 프로젝트명
                            </label>
                            <input
                                type="text"
                                value={tokens.pageLayout.projectName}
                                onChange={(e) => handleLayoutSettingChange('projectName', e.target.value)}
                                style={{ width: '100%', padding: '7px 10px', borderRadius: '6px', border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)', color: 'var(--text-primary)', fontSize: '12.5px' }}
                            />
                        </div>
                        <div>
                            <label style={{ display: 'block', fontSize: '12px', fontWeight: 600, color: 'var(--text-muted)', marginBottom: '4px' }}>
                                공통 본문 서체 (Font)
                            </label>
                            <select
                                value={tokens.typography.fontFamily}
                                onChange={(e) => setTokens(prev => ({ ...prev, typography: { ...prev.typography, fontFamily: e.target.value } }))}
                                style={{ width: '100%', padding: '7px 10px', borderRadius: '6px', border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.3)', color: 'var(--text-primary)', fontSize: '12.5px' }}
                            >
                                <option value="맑은 고딕">맑은 고딕 (윈도우 공식 호환)</option>
                                <option value="Pretendard">Pretendard (현대적 산세리프)</option>
                                <option value="KoPub돋움체 Medium">KoPub돋움체 Medium (공공출판)</option>
                            </select>
                        </div>
                        <div>
                            <label style={{ display: 'block', fontSize: '12px', fontWeight: 600, color: 'var(--text-muted)', marginBottom: '4px' }}>
                                Primary 색상 (HEX)
                            </label>
                            <div style={{ display: 'flex', gap: '6px', alignItems: 'center' }}>
                                <input
                                    type="color"
                                    value={`#${tokens.colors.primary}`}
                                    onChange={(e) => handleColorChange('primary', e.target.value)}
                                    style={{ width: '32px', height: '32px', border: 'none', background: 'none', cursor: 'pointer' }}
                                />
                                <input
                                    type="text"
                                    value={tokens.colors.primary}
                                    onChange={(e) => handleColorChange('primary', e.target.value)}
                                    style={{ flex: 1, padding: '7px 10px', borderRadius: '6px', border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)', color: 'var(--text-primary)', fontSize: '12.5px' }}
                                />
                            </div>
                        </div>
                        <div>
                            <label style={{ display: 'block', fontSize: '12px', fontWeight: 600, color: 'var(--text-muted)', marginBottom: '4px' }}>
                                Secondary 색상 (HEX)
                            </label>
                            <div style={{ display: 'flex', gap: '6px', alignItems: 'center' }}>
                                <input
                                    type="color"
                                    value={`#${tokens.colors.secondary}`}
                                    onChange={(e) => handleColorChange('secondary', e.target.value)}
                                    style={{ width: '32px', height: '32px', border: 'none', background: 'none', cursor: 'pointer' }}
                                />
                                <input
                                    type="text"
                                    value={tokens.colors.secondary}
                                    onChange={(e) => handleColorChange('secondary', e.target.value)}
                                    style={{ flex: 1, padding: '7px 10px', borderRadius: '6px', border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)', color: 'var(--text-primary)', fontSize: '12.5px' }}
                                />
                            </div>
                        </div>
                    </div>
                </div>
            )}

            {/* 메인 2단 작업 영역: 좌측 목차 테이블 & 우측 실시간 장표 프리뷰 */}
            <div style={{ display: 'grid', gridTemplateColumns: 'minmax(420px, 1.1fr) minmax(480px, 1.2fr)', gap: '20px', alignItems: 'start' }}>
                {/* ── 좌측: 목차 스캐폴딩 관리 패널 ─────────────────────── */}
                <div className="glass-panel" style={{ padding: '20px', display: 'flex', flexDirection: 'column', gap: '14px' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '10px' }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                            <FileSpreadsheet size={18} style={{ color: 'var(--accent-blue)' }} />
                            <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 700, color: 'var(--text-primary)' }}>
                                산출물 장표 목차 구성 ({slides.length}개 슬라이드)
                            </h3>
                        </div>
                        <div style={{ display: 'flex', gap: '6px' }}>
                            <input
                                type="file"
                                accept=".xlsx,.xls"
                                ref={excelImportRef}
                                style={{ display: 'none' }}
                                onChange={handleImportTocExcel}
                            />
                            <button
                                onClick={() => excelImportRef.current?.click()}
                                className="btn-outline"
                                style={{ padding: '5px 10px', fontSize: '12px', display: 'flex', alignItems: 'center', gap: '4px' }}
                                title="목차 엑셀 가져오기"
                            >
                                <Upload size={13} /> 엑셀 로드
                            </button>
                            <button
                                onClick={handleExportTocExcel}
                                className="btn-outline"
                                style={{ padding: '5px 10px', fontSize: '12px', display: 'flex', alignItems: 'center', gap: '4px' }}
                                title="목차 엑셀 다운로드"
                            >
                                <Download size={13} /> 엑셀 저장
                            </button>
                            <button
                                onClick={handleAddSlide}
                                className="btn-primary"
                                style={{ padding: '5px 12px', fontSize: '12px', display: 'flex', alignItems: 'center', gap: '4px' }}
                            >
                                <Plus size={13} /> 장표 추가
                            </button>
                        </div>
                    </div>

                    {/* 목차 리스트 (인라인 편집 가능) */}
                    <div style={{ maxHeight: '580px', overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '8px', paddingRight: '4px' }}>
                        {slides.map((slide, idx) => {
                            const isSelected = slide.id === selectedSlideId;
                            return (
                                <div
                                    key={slide.id}
                                    onClick={() => setSelectedSlideId(slide.id)}
                                    style={{
                                        padding: '12px 14px',
                                        borderRadius: '8px',
                                        cursor: 'pointer',
                                        background: isSelected ? 'rgba(59, 130, 246, 0.12)' : 'rgba(255,255,255,0.02)',
                                        border: `1.5px solid ${isSelected ? 'var(--accent-blue)' : 'var(--panel-border)'}`,
                                        transition: 'all 0.15s ease',
                                        display: 'flex',
                                        flexDirection: 'column',
                                        gap: '8px'
                                    }}
                                >
                                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                                        <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                                            <span style={{
                                                padding: '2px 6px',
                                                borderRadius: '4px',
                                                fontSize: '10.5px',
                                                fontWeight: 800,
                                                background: isSelected ? 'var(--accent-blue)' : 'rgba(255,255,255,0.1)',
                                                color: '#fff'
                                            }}>
                                                #{idx + 1}
                                            </span>
                                            <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>
                                                {slide.chapter} &gt; {slide.section}
                                            </span>
                                        </div>
                                        <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                                            <span style={{
                                                padding: '1px 6px',
                                                borderRadius: '4px',
                                                fontSize: '10px',
                                                fontWeight: 700,
                                                background: 'rgba(16, 185, 129, 0.15)',
                                                color: '#10b981'
                                            }}>
                                                {slide.reqId || 'CNR-001'}
                                            </span>
                                            <button
                                                onClick={(e) => { e.stopPropagation(); handleDeleteSlide(slide.id); }}
                                                style={{ background: 'none', border: 'none', color: 'var(--text-muted)', cursor: 'pointer', padding: '2px' }}
                                                title="장표 삭제"
                                            >
                                                <Trash2 size={13} />
                                            </button>
                                        </div>
                                    </div>

                                    {/* 장표 제목 인풋 */}
                                    <input
                                        type="text"
                                        value={slide.title}
                                        onClick={(e) => e.stopPropagation()}
                                        onChange={(e) => handleUpdateSlideField(slide.id, 'title', e.target.value)}
                                        placeholder="장표 제목 입력"
                                        style={{
                                            width: '100%',
                                            padding: '5px 8px',
                                            borderRadius: '4px',
                                            border: '1px solid var(--panel-border)',
                                            background: 'rgba(0,0,0,0.2)',
                                            color: 'var(--text-primary)',
                                            fontSize: '13px',
                                            fontWeight: 600
                                        }}
                                    />

                                    {/* 레이아웃 선택 및 담당자 */}
                                    <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                                        <select
                                            value={slide.layoutType}
                                            onClick={(e) => e.stopPropagation()}
                                            onChange={(e) => handleUpdateSlideField(slide.id, 'layoutType', e.target.value)}
                                            style={{
                                                flex: 1,
                                                padding: '4px 6px',
                                                borderRadius: '4px',
                                                background: 'rgba(0,0,0,0.3)',
                                                border: '1px solid var(--panel-border)',
                                                color: 'var(--accent-purple)',
                                                fontSize: '11.5px',
                                                fontWeight: 600
                                            }}
                                        >
                                            {Object.entries(LAYOUT_LABELS).map(([k, label]) => (
                                                <option key={k} value={k}>{label}</option>
                                            ))}
                                        </select>
                                        <input
                                            type="text"
                                            value={slide.reqId || ''}
                                            onClick={(e) => e.stopPropagation()}
                                            onChange={(e) => handleUpdateSlideField(slide.id, 'reqId', e.target.value)}
                                            placeholder="요구사항ID"
                                            style={{ width: '85px', padding: '4px 6px', borderRadius: '4px', border: '1px solid var(--panel-border)', background: 'rgba(0,0,0,0.2)', color: 'var(--text-secondary)', fontSize: '11px' }}
                                        />
                                    </div>
                                </div>
                            );
                        })}
                    </div>
                </div>

                {/* ── 우측: 선택 장표 실시간 프리뷰 & AI 초안 채우기 ───────── */}
                <div className="glass-panel" style={{ padding: '20px', display: 'flex', flexDirection: 'column', gap: '16px' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '10px' }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                            <Presentation size={18} style={{ color: `#${tokens.colors.primary}` }} />
                            <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 700, color: 'var(--text-primary)' }}>
                                장표 레이아웃 실시간 미리보기 (Preview)
                            </h3>
                        </div>

                        {/* AI 자동 채우기 버튼 */}
                        <button
                            onClick={handleRunAiAutoFill}
                            disabled={isAiLoading}
                            className="btn-primary"
                            style={{
                                padding: '6px 14px',
                                fontSize: '12px',
                                fontWeight: 700,
                                display: 'flex',
                                alignItems: 'center',
                                gap: '5px',
                                background: 'linear-gradient(135deg, #a855f7, #6366f1)',
                                boxShadow: '0 2px 8px rgba(168, 85, 247, 0.3)'
                            }}
                            title="ISP 가이드와 연계하여 헤드메시지, 불릿, 표 데이터를 자동 생성"
                        >
                            <Sparkles size={13} className={isAiLoading ? 'animate-spin' : ''} />
                            {isAiLoading ? 'AI 초안 생성 중...' : 'AI 초안 자동 채우기'}
                        </button>
                    </div>

                    {/* 실시간 슬라이드 디자인 프리뷰 프레임 (16:9 와이드 비율) */}
                    <div
                        style={{
                            width: '100%',
                            aspectRatio: tokens.pageLayout.aspectRatio === '16:9' ? '16 / 9' : '4 / 3',
                            background: '#FFFFFF',
                            borderRadius: '10px',
                            border: `2px solid #${tokens.colors.borderColor}`,
                            boxShadow: '0 8px 24px rgba(0,0,0,0.25)',
                            padding: '16px 20px',
                            display: 'flex',
                            flexDirection: 'column',
                            justifyContent: 'space-between',
                            color: '#0F172A',
                            fontFamily: tokens.typography.fontFamily,
                            overflow: 'hidden'
                        }}
                    >
                        {/* 1. 상단 마스터 헤더 */}
                        <div>
                            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: '10px', color: `#${tokens.colors.textSecondary}`, fontWeight: 600, borderBottom: `1px solid #${tokens.colors.borderColor}`, paddingBottom: '4px' }}>
                                <span>{activeSlide.chapter} &gt; {activeSlide.section}</span>
                                <span style={{ color: `#${tokens.colors.primary}`, fontWeight: 700 }}>
                                    {tokens.pageLayout.projectName} [{activeSlide.reqId || 'CNR-001'}]
                                </span>
                            </div>

                            {/* 장표 제목 */}
                            <h4 style={{ margin: '8px 0 6px', fontSize: '15px', fontWeight: 800, color: `#${tokens.colors.primary}`, letterSpacing: '-0.3px' }}>
                                {activeSlide.title}
                            </h4>

                            {/* 거버닝 헤드메시지 박스 */}
                            <div style={{
                                background: `#${tokens.colors.keyPointBg}`,
                                border: `1px solid #${tokens.colors.secondary}`,
                                borderLeft: `4px solid #${tokens.colors.primary}`,
                                borderRadius: '4px',
                                padding: '6px 10px',
                                fontSize: '10.5px',
                                fontWeight: 700,
                                color: `#${tokens.colors.textPrimary}`,
                                lineHeight: '1.4'
                            }}>
                                💡 {activeSlide.headMessage || '본 장표는 세부 이행방안을 표준 양식에 따라 명시합니다.'}
                            </div>
                        </div>

                        {/* 2. 본문 5대 레이아웃 프리뷰 */}
                        <div style={{ flex: 1, marginTop: '8px', marginBottom: '8px', overflow: 'hidden', display: 'flex', flexDirection: 'column' }}>
                            {activeSlide.layoutType === LAYOUT_TYPES.TYPE_A_TABLE && (
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px', height: '100%' }}>
                                    <div style={{ background: `#${tokens.colors.bgLight}`, border: `1px solid #${tokens.colors.borderColor}`, borderRadius: '4px', padding: '6px 8px', fontSize: '9px', lineHeight: '1.3' }}>
                                        <div style={{ fontWeight: 700, color: `#${tokens.colors.primary}`, marginBottom: '2px' }}>📌 핵심 시사점 (Key Points)</div>
                                        <div>• 공공보건의료 데이터 통합 연계를 위한 표준 EMR 인터페이스 규격 수립 필요</div>
                                        <div>• 국가 표준 진료정보교류(FHIR) 준수율을 95% 이상으로 상향하여 데이터 고립 방지</div>
                                    </div>
                                    <div style={{ flex: 1, border: `1px solid #${tokens.colors.borderColor}`, borderRadius: '4px', overflow: 'hidden' }}>
                                        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '8.5px' }}>
                                            <thead>
                                                <tr style={{ background: `#${tokens.colors.tableHeaderBg}`, color: `#${tokens.colors.tableHeaderText}` }}>
                                                    <th style={{ padding: '3px 4px', textAlign: 'center' }}>구분</th>
                                                    <th style={{ padding: '3px 4px' }}>현행(As-Is) 수준</th>
                                                    <th style={{ padding: '3px 4px' }}>목표(To-Be) 모델</th>
                                                    <th style={{ padding: '3px 4px' }}>표준 적용 방안</th>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                <tr style={{ borderBottom: '1px solid #e2e8f0' }}>
                                                    <td style={{ padding: '3px', textAlign: 'center', fontWeight: 600 }}>데이터 표준</td>
                                                    <td style={{ padding: '3px' }}>의료원별 독자 코드체계 사용</td>
                                                    <td style={{ padding: '3px' }}>보건의료표준용어 및 FHIR 적용</td>
                                                    <td style={{ padding: '3px' }}>국가 표준 매핑 테이블 구축</td>
                                                </tr>
                                                <tr>
                                                    <td style={{ padding: '3px', textAlign: 'center', fontWeight: 600 }}>연계 방식</td>
                                                    <td style={{ padding: '3px' }}>수기 추출 및 배치 파일 전송</td>
                                                    <td style={{ padding: '3px' }}>실시간 API 게이트웨이 연계</td>
                                                    <td style={{ padding: '3px' }}>RESTful API 및 EAI 표준화</td>
                                                </tr>
                                            </tbody>
                                        </table>
                                    </div>
                                </div>
                            )}

                            {activeSlide.layoutType === LAYOUT_TYPES.TYPE_B_BEFORE_AFTER && (
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px', height: '100%' }}>
                                    <div style={{ display: 'grid', gridTemplateColumns: '1fr auto 1fr', gap: '8px', alignItems: 'center', flex: 1 }}>
                                        <div style={{ background: '#FFF5F5', border: `1px solid #${tokens.colors.danger}`, borderRadius: '6px', padding: '6px', height: '100%', fontSize: '8.5px' }}>
                                            <div style={{ fontWeight: 800, color: `#${tokens.colors.danger}`, marginBottom: '3px' }}>[As-Is] 현행 문제점</div>
                                            <div>• 환자 전원 시 진료기록 수기 복사 전달</div>
                                            <div>• 통계 수기 집계로 집계 지연(평균 3주)</div>
                                            <div>• 의료 AI 도입을 위한 표준 인프라 부재</div>
                                        </div>
                                        <div style={{ fontSize: '14px', fontWeight: 900, color: `#${tokens.colors.secondary}` }}>➔</div>
                                        <div style={{ background: '#F0FDF4', border: `1px solid #${tokens.colors.success}`, borderRadius: '6px', padding: '6px', height: '100%', fontSize: '8.5px' }}>
                                            <div style={{ fontWeight: 800, color: `#${tokens.colors.success}`, marginBottom: '3px' }}>[To-Be] 목표 혁신안</div>
                                            <div>• 국가 표준 FHIR 기반 실시간 진료정보교류</div>
                                            <div>• 중앙 데이터 레이크로 실시간 통계 시각화</div>
                                            <div>• 클라우드 기반 AI CDSS 표준 탑재·배포</div>
                                        </div>
                                    </div>
                                    <div style={{ background: `#${tokens.colors.keyPointBg}`, border: `1px solid #${tokens.colors.secondary}`, padding: '4px 8px', borderRadius: '4px', fontSize: '8.5px', fontWeight: 700, color: `#${tokens.colors.primary}` }}>
                                        {activeSlide.deltaText || '🎯 [기대성과] 환자 재검사율 28% 감소 | 응급환자 골든타임 확보율 35% 향상'}
                                    </div>
                                </div>
                            )}

                            {activeSlide.layoutType === LAYOUT_TYPES.TYPE_C_PROCESS && (
                                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '6px', height: '100%' }}>
                                    {[
                                        { s: 'Step 1', t: '기반 분석·진단', p: '1~2개월' },
                                        { s: 'Step 2', t: '목표 아키텍처', p: '3~4개월' },
                                        { s: 'Step 3', t: '플랫폼 구축·실증', p: '5~6개월' },
                                        { s: 'Step 4', t: '전국 확산·안정화', p: '7~8개월' }
                                    ].map((step, idx) => (
                                        <div key={idx} style={{ border: `1px solid #${tokens.colors.borderColor}`, borderRadius: '4px', overflow: 'hidden', display: 'flex', flexDirection: 'column' }}>
                                            <div style={{ background: idx === 0 ? `#${tokens.colors.primary}` : `#${tokens.colors.secondary}`, color: '#fff', padding: '3px 4px', fontSize: '8.5px', fontWeight: 700, textAlign: 'center' }}>
                                                {step.s}: {step.t}
                                            </div>
                                            <div style={{ flex: 1, background: `#${tokens.colors.bgLight}`, padding: '4px', fontSize: '8px', lineHeight: '1.3' }}>
                                                <div>• 세부 이행활동 수행</div>
                                                <div>• 단계별 산출물 작성</div>
                                                <div>• 감리 및 품질 점검</div>
                                            </div>
                                        </div>
                                    ))}
                                </div>
                            )}

                            {activeSlide.layoutType === LAYOUT_TYPES.TYPE_D_2COL_CARDS && (
                                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px', height: '100%' }}>
                                    <div style={{ border: `1px solid #${tokens.colors.primary}`, borderRadius: '6px', overflow: 'hidden', display: 'flex', flexDirection: 'column' }}>
                                        <div style={{ background: `#${tokens.colors.primary}`, color: '#fff', padding: '4px 8px', fontSize: '9px', fontWeight: 700 }}>
                                            전략 1 : 공공의료 데이터 거버넌스
                                        </div>
                                        <div style={{ flex: 1, background: `#${tokens.colors.bgLight}`, padding: '6px', fontSize: '8.5px', lineHeight: '1.4' }}>
                                            <div>• 전국 지방의료원 EMR 표준 인터페이스(API) 단일화</div>
                                            <div>• 보건의료 마이데이터 연계 표준화 및 가명처리 엄수</div>
                                        </div>
                                    </div>
                                    <div style={{ border: `1px solid #${tokens.colors.secondary}`, borderRadius: '6px', overflow: 'hidden', display: 'flex', flexDirection: 'column' }}>
                                        <div style={{ background: `#${tokens.colors.secondary}`, color: '#fff', padding: '4px 8px', fontSize: '9px', fontWeight: 700 }}>
                                            전략 2 : 지능형 AI 임상결정지원(CDSS)
                                        </div>
                                        <div style={{ flex: 1, background: `#${tokens.colors.bgLight}`, padding: '6px', fontSize: '8.5px', lineHeight: '1.4' }}>
                                            <div>• 급성 중증 환자 조기 예측 및 패혈증 실시간 모니터링</div>
                                            <div>• 전공의 판독 지원을 위한 AI 흉부 X-ray 자동 판독</div>
                                        </div>
                                    </div>
                                </div>
                            )}

                            {activeSlide.layoutType === LAYOUT_TYPES.TYPE_E_MATRIX && (
                                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gridTemplateRows: '1fr 1fr', gap: '6px', height: '100%' }}>
                                    {[
                                        { t: '강점 (Strength)', c: tokens.colors.primary, it: ['전국 공공의료원 네트워크', '풍부한 임상 코호트 데이터'] },
                                        { t: '약점 (Weakness)', c: tokens.colors.danger, it: ['의료원별 EMR 버전 상이', '원내 전산 인력 전문성 부족'] },
                                        { t: '기회 (Opportunity)', c: tokens.colors.success, it: ['정부 디지털 헬스케어 정책', '클라우드 도입 가이드라인 완화'] },
                                        { t: '위협 (Threat)', c: tokens.colors.warning, it: ['개인정보 유출 등 보안 리스크', '의료계 이해관계자 반발 가능성'] }
                                    ].map((q, idx) => (
                                        <div key={idx} style={{ border: `1px solid #${q.c}`, borderRadius: '4px', overflow: 'hidden', display: 'flex', flexDirection: 'column' }}>
                                            <div style={{ background: `#${q.c}`, color: '#fff', padding: '2px 6px', fontSize: '8px', fontWeight: 700 }}>
                                                {q.t}
                                            </div>
                                            <div style={{ flex: 1, background: `#${tokens.colors.bgLight}`, padding: '4px', fontSize: '7.5px' }}>
                                                {q.it.map((item, iIdx) => <div key={iIdx}>• {item}</div>)}
                                            </div>
                                        </div>
                                    ))}
                                </div>
                            )}
                        </div>

                        {/* 3. 하단 슬라이드 번호 */}
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: '9px', color: `#${tokens.colors.textSecondary}`, borderTop: `1px solid #${tokens.colors.borderColor}`, paddingTop: '4px' }}>
                            <span>국가 공공보건의료 표준 정보화전략계획(ISP)</span>
                            <span>Ⅰ - {slides.findIndex(s => s.id === activeSlide.id) + 1}</span>
                        </div>
                    </div>

                    {/* 일괄 내보내기 버튼 영역 */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', marginTop: '4px' }}>
                        <span style={{ fontSize: '12.5px', fontWeight: 700, color: 'var(--text-primary)' }}>
                            🚀 단일 디자인 토큰 기반 산출물 일괄 내보내기 (Export)
                        </span>

                        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(130px, 1fr))', gap: '8px' }}>
                            <button
                                onClick={() => handleExportSingleFormat('pptx')}
                                disabled={isExporting}
                                className="btn-outline"
                                style={{ padding: '8px 10px', fontSize: '12px', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '5px' }}
                            >
                                <Presentation size={14} style={{ color: '#f97316' }} />
                                .pptx 프레젠테이션
                            </button>
                            <button
                                onClick={() => handleExportSingleFormat('docx')}
                                disabled={isExporting}
                                className="btn-outline"
                                style={{ padding: '8px 10px', fontSize: '12px', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '5px' }}
                            >
                                <FileText size={14} style={{ color: '#3b82f6' }} />
                                .docx 정형 보고서
                            </button>
                            <button
                                onClick={() => handleExportSingleFormat('xlsx')}
                                disabled={isExporting}
                                className="btn-outline"
                                style={{ padding: '8px 10px', fontSize: '12px', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '5px' }}
                            >
                                <FileSpreadsheet size={14} style={{ color: '#10b981' }} />
                                .xlsx 목차 명세서
                            </button>
                        </div>

                        {/* 원클릭 3종 세트 일괄 빌드 */}
                        <button
                            onClick={handleExportAllFormats}
                            disabled={isExporting}
                            className="btn-primary"
                            style={{
                                padding: '12px 18px',
                                fontSize: '13.5px',
                                fontWeight: 800,
                                display: 'flex',
                                alignItems: 'center',
                                justifyContent: 'center',
                                gap: '8px',
                                background: 'linear-gradient(135deg, #2563EB, #7C3AED)',
                                boxShadow: '0 4px 14px rgba(37, 99, 235, 0.35)'
                            }}
                        >
                            <Download size={16} />
                            {isExporting ? '표준 산출물 일괄 렌더링 중...' : '동일 디자인 일괄 내보내기 (PPTX + DOCX + XLSX 동시 생성)'}
                        </button>
                    </div>
                </div>
            </div>
        </div>
    );
}
