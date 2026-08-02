/**
 * geminiModels.js
 * Gemini API 모델 목록 공통 상수 파일
 *
 * - 1순위(FALLBACK_MODELS[0]): gemini-3.6-flash (최신 효율 모델, 2026-07-21 출시)
 * - 오류 발생 시 다음 모델로 자동 전환되는 Fallback 구조에 사용됩니다.
 * - 모델 추가/변경 시 이 파일만 수정하면 전체 앱에 반영됩니다.
 *
 * 최종 수정: 2026-08-02 v3.0
 */

export const FALLBACK_MODELS = [
    "models/gemini-3.6-flash",          // 1순위: 최신 효율 모델 (2026-07-21, 기본 권장)
    "models/gemini-3.5-flash",          // 2순위: 고성능 에이전틱 작업 (2026-05)
    "models/gemini-3.5-flash-lite",     // 3순위: 경량 고처리량 모델 (2026-07-21)
    "models/gemini-2.5-flash",          // 4순위: 구세대 Flash (폴백 유지)
    "models/gemini-2.5-pro",            // 5순위: 구세대 Pro (폴백 유지)
    "models/gemini-1.5-pro",            // 6순위: 구버전 Pro
    "models/gemini-1.5-flash",          // 7순위: 구버전 Flash
];

