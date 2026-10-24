// LocalStorage 영속화 키
export const DESIGN_TOKENS_STORAGE_KEY = 'healthy_project_design_tokens';

// ── 테마 프리셋 3종 정의 ───────────────────────────────────────────
export const THEME_PRESETS = {
    PUBLIC_NAVY: {
        id: 'PUBLIC_NAVY',
        name: '① 공공기관 표준 네이비 (Standard Navy)',
        description: '공공기관 ISP, 정부 감리 및 공공 보고서 공식 표준',
        colors: {
            primary: '1E3A8A',       // 딥 네이비 (Blue 900)
            secondary: '2563EB',     // 로얄 블루 (Blue 600)
            accent: '7C3AED',        // 퍼플 포인트
            tableHeaderBg: 'F1F5F9', // 연회색 (Slate 100)
            tableHeaderText: '0F172A',// 짙은 텍스트 (Slate 900)
            borderColor: 'CBD5E1',   // 테두리 회색 (Slate 300)
            keyPointBg: 'EFF6FF',    // 강조 배경 (Blue 50)
            bgLight: 'F8FAFC',       // 카드 연한 배경
            textPrimary: '0F172A',   // 본문 검정
            textSecondary: '475569', // 보조 텍스트
            white: 'FFFFFF',
            danger: 'DC2626',
            success: '16A34A',
            warning: 'D97706'
        }
    },
    KITC_BLUE_PURPLE: {
        id: 'KITC_BLUE_PURPLE',
        name: '② KITC 컨설팅 블루/퍼플 (Consulting Brand)',
        description: 'IT 컨설팅 제안서, 전략 보고서 및 고급 브리핑용',
        colors: {
            primary: '1D4ED8',       // 인텐스 블루 (Blue 700)
            secondary: '6366F1',     // 인디고 (Indigo 500)
            accent: '8B5CF6',        // 바이올렛 (Violet 500)
            tableHeaderBg: 'EEF2FF', // 인디고 50
            tableHeaderText: '1E1B4B',// 딥 인디고 950
            borderColor: 'C7D2FE',   // 인디고 200
            keyPointBg: 'F5F3FF',    // 퍼플 50
            bgLight: 'FAFAFA',
            textPrimary: '111827',
            textSecondary: '4B5563',
            white: 'FFFFFF',
            danger: 'E11D48',
            success: '059669',
            warning: 'D97706'
        }
    },
    MODERN_DARK_GRAY: {
        id: 'MODERN_DARK_GRAY',
        name: '③ 모던 슬레이트/그레이 (Modern Slate)',
        description: '차분하고 정돈된 테크니컬 아키텍처 및 데이터 명세용',
        colors: {
            primary: '334155',       // 슬레이트 700
            secondary: '64748B',     // 슬레이트 500
            accent: '0EA5E9',        // 스카이블루 (Sky 500)
            tableHeaderBg: 'F8FAFC', // 슬레이트 50
            tableHeaderText: '0F172A',
            borderColor: 'E2E8F0',   // 슬레이트 200
            keyPointBg: 'F0FDF4',    // 연녹색 (Green 50)
            bgLight: 'F8FAFC',
            textPrimary: '0F172A',
            textSecondary: '64748B',
            white: 'FFFFFF',
            danger: 'EF4444',
            success: '10B981',
            warning: 'F59E0B'
        }
    }
};

// ── 기본 통합 디자인 토큰 사양 ─────────────────────────────────────
export const DEFAULT_DESIGN_TOKENS = {
    selectedPresetId: 'PUBLIC_NAVY',
    colors: { ...THEME_PRESETS.PUBLIC_NAVY.colors },
    typography: {
        fontFamily: '맑은 고딕',      // '맑은 고딕' | 'Pretendard' | 'KoPub돋움체 Medium'
        titleSize: 18,               // pt (장표 대제목)
        subTitleSize: 13,            // pt (소제목/섹션헤더)
        headMessageSize: 11.5,       // pt (거버닝 헤드메시지)
        bodySize: 10,                // pt (일반 본문 텍스트)
        tableTextSize: 9,            // pt (표 셀 내부 텍스트)
        badgeSize: 8.5               // pt (요구사항 ID 뱃지)
    },
    pageLayout: {
        projectName: '공공보건의료 AI 지능형 플랫폼 구축 ISP',
        aspectRatio: '16:9',         // '16:9' | '4:3'
        pageNumberFormat: 'Ⅰ - {page}', // 'Ⅰ - {page}' | '{page}' | 'Page {page}'
        showBreadcrumb: true,
        showHeadMessage: true,
        showReqBadge: true,
        showFooterLogo: true
    }
};

// ── LocalStorage 로드 / 저장 ──────────────────────────────────────
export function loadDesignTokens() {
    try {
        const raw = localStorage.getItem(DESIGN_TOKENS_STORAGE_KEY);
        if (raw) {
            const parsed = JSON.parse(raw);
            const presetId = parsed.selectedPresetId || 'PUBLIC_NAVY';
            const baseColors = THEME_PRESETS[presetId]?.colors || THEME_PRESETS.PUBLIC_NAVY.colors;
            return {
                selectedPresetId: presetId,
                colors: { ...baseColors, ...(parsed.colors || {}) },
                typography: { ...DEFAULT_DESIGN_TOKENS.typography, ...(parsed.typography || {}) },
                pageLayout: { ...DEFAULT_DESIGN_TOKENS.pageLayout, ...(parsed.pageLayout || {}) }
            };
        }
    } catch (e) {
        console.warn('DesignTokens 로드 실패, 기본값 사용:', e);
    }
    return JSON.parse(JSON.stringify(DEFAULT_DESIGN_TOKENS));
}

export function saveDesignTokens(tokens) {
    try {
        localStorage.setItem(DESIGN_TOKENS_STORAGE_KEY, JSON.stringify(tokens));
        return true;
    } catch (e) {
        console.error('DesignTokens 저장 실패:', e);
        return false;
    }
}
