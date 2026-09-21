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
    "models/gemini-2.5-flash",          // 1순위: 최신 고성능 표준 모델 (안정적 & 빠른 응답)
    "models/gemini-2.5-flash-lite",     // 2순위: 최신 경량 고속 모델
    "models/gemini-2.0-flash",          // 3순위: 2.0 세대 표준 Flash 모델
    "models/gemini-2.0-flash-lite",     // 4순위: 2.0 세대 경량 Flash Lite 모델
    "models/gemini-3.6-flash",          // 5순위: 3.6 Flash 모델
    "models/gemini-3.5-flash",          // 6순위: 3.5 Flash 모델
    "models/gemini-3.5-flash-lite",     // 7순위: 3.5 Flash Lite 모델
];

