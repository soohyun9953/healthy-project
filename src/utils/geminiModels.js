/**
 * geminiModels.js
 * Gemini API 모델 목록 공통 상수 파일
 *
 * - 1순위(FALLBACK_MODELS[0]): gemini-2.5-flash (빠르고 성능 우수, 현재 권장)
 * - 오류 발생 시 다음 모델로 자동 전환되는 Fallback 구조에 사용됩니다.
 * - 모델 추가/변경 시 이 파일만 수정하면 전체 앱에 반영됩니다.
 *
 * 최종 수정: 2026-05-20 v2.0
 */

export const FALLBACK_MODELS = [
    "models/gemini-2.5-flash",       // 1순위: 빠르고 성능 우수 (기본 권장)
    "models/gemini-2.5-pro",         // 2순위: 최고 성능 (복잡한 작업)
    "models/gemini-2.5-flash-lite",  // 3순위: 경량 버전
    "models/gemini-2.0-flash-exp",   // 4순위: 실험적 버전
    "models/gemini-1.5-pro",         // 5순위: 구버전 Pro
    "models/gemini-1.5-flash",       // 6순위: 구버전 Flash
    "models/gemini-1.5-flash-8b",    // 7순위: 구버전 경량
];
