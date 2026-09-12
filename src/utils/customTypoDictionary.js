// customTypoDictionary.js
// 사용자가 직접 등록하는 맞춤형 교정 사전(오류 표현 -> 올바른 표현). 브라우저 localStorage에 영속 저장되며,
// typoDictionary.js의 extract_dictionary_typos(text, customDict)에 병합되어 1단계 사전 스캔에 즉시 반영된다.

const STORAGE_KEY = 'custom_typo_dictionary_v1';

export function getCustomDictionary() {
    try {
        const raw = localStorage.getItem(STORAGE_KEY);
        return raw ? JSON.parse(raw) : {};
    } catch (e) {
        console.error('커스텀 사전 로드 실패:', e);
        return {};
    }
}

function saveCustomDictionary(dict) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(dict));
}

export function addCustomTerm(wrong, correction, desc = '') {
    const wrongTrim = String(wrong || '').trim();
    const correctionTrim = String(correction || '').trim();
    if (!wrongTrim || !correctionTrim) {
        throw new Error('오류 표현과 올바른 표현을 모두 입력해 주세요.');
    }
    if (wrongTrim === correctionTrim) {
        throw new Error('오류 표현과 올바른 표현이 서로 달라야 합니다.');
    }

    const dict = getCustomDictionary();
    dict[wrongTrim] = {
        correction: correctionTrim,
        desc: desc.trim() || `사용자 등록 용어: '${wrongTrim}' → '${correctionTrim}'`,
        type: '사용자 정의'
    };
    saveCustomDictionary(dict);
    return dict;
}

export function deleteCustomTerm(wrong) {
    const dict = getCustomDictionary();
    delete dict[wrong];
    saveCustomDictionary(dict);
    return dict;
}

export function clearCustomDictionary() {
    saveCustomDictionary({});
}
