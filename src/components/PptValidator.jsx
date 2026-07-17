import React, { useState, useRef, useEffect } from 'react';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import { 
  FileUp, 
  ShieldAlert, 
  CheckCircle2, 
  AlertTriangle, 
  Download, 
  RefreshCw, 
  FileSpreadsheet,
  AlertCircle,
  ChevronRight,
  HelpCircle,
  Info,
  Plus,
  Trash2,
  Sparkles
} from 'lucide-react';

// 오탈자 내장 사전 정의
const TYPO_DICTIONARY = {
  // 맞춤법 오류
  '역활': { correction: '역할', desc: '역할(役割)의 잘못된 표기입니다.', type: '맞춤법' },
  '되서': { correction: '돼서', desc: '되어서의 축약형인 돼서로 써야 합니다.', type: '맞춤법' },
  '바램': { correction: '바람', desc: '바라다에서 파생된 명사는 바람이 맞습니다.', type: '맞춤법' },
  '않하고': { correction: '안 하고', desc: '부정을 뜻하는 부사 안과 하다의 합성어인 안 하고가 맞습니다.', type: '맞춤법' },
  '몇일': { correction: '며칠', desc: '몇 일은 며칠의 잘못된 표기입니다.', type: '맞춤법' },
  '일일히': { correction: '일일이', desc: '일일이(하나씩 하나씩)가 표준어입니다.', type: '맞춤법' },
  '문안히': { correction: '무난히', desc: '무난히(별일 없이/어렵지 않게)가 문맥상 올바릅니다.', type: '맞춤법' },
  '어의없다': { correction: '어이없다', desc: '어처구니없다 또는 어이없다가 올바른 표기입니다.', type: '맞춤법' },
  '않돼': { correction: '안 돼', desc: '안 되어의 축약형인 안 돼가 표준 표기입니다.', type: '맞춤법' },
  '안돼다': { correction: '안되다', desc: '일이나 현상이 좋지 않게 흘러갈 때는 안되다가 표준어입니다.', type: '맞춤법' },
  '갯수': { correction: '개수', desc: '개수(個數)가 맞춤법에 부합하는 올바른 표기입니다.', type: '맞춤법' },
  '댓가': { correction: '대가', desc: '대가(代價)가 맞춤법에 부합하는 올바른 표기입니다. 사이시옷이 들어가지 않습니다.', type: '맞춤법' },
  '바꼈다': { correction: '바뀌었다', desc: '바뀌었다가 줄어든 말은 바뀌었다로 표기해야 합니다.', type: '맞춤법' },
  '가게부': { correction: '가계부', desc: '가계부(家計簿)가 올바른 표기입니다.', type: '맞춤법' },
  '금새': { correction: '금세', desc: '금세(금시에의 준말)가 올바른 맞춤법 표기입니다.', type: '맞춤법' },
  '희안하다': { correction: '희한하다', desc: '희한하다(稀罕-)가 올바른 표기입니다.', type: '맞춤법' },
  '설레임': { correction: '설렘', desc: '설레다의 명사형은 설렘이 올바른 맞춤법입니다.', type: '맞춤법' },
  '와부': { correction: '외부', desc: '외부(外部)의 오타 표기입니다.', type: '맞춤법' },
  '제사하다': { correction: '제시하다', desc: '비즈니스 제안 맥락상 제시하다의 오타 표기입니다.', type: '맞춤법' },
  '계발하다': { correction: '개발하다', desc: '소프트웨어 개발 맥락상 개발하다의 오타 표기입니다.', type: '맞춤법' },
  
  // 외래어 표기법 오류
  '아키텍쳐': { correction: '아키텍처', desc: '외래어 표기법에 의하면 아키텍처(Architecture)가 표준입니다.', type: '외래어 표기' },
  '컨텐츠': { correction: '콘텐츠', desc: '콘텐츠(Contents)가 표준 외래어 표기법입니다.', type: '외래어 표기' },
  '컴퍼넌트': { correction: '컴포넌트', desc: '컴포넌트(Component)가 올바른 외래어 표기입니다.', type: '외래어 표기' },
  '데이타베이스': { correction: '데이터베이스', desc: '데이터베이스(Database)가 올바른 표준 외래어 표기입니다.', type: '외래어 표기' },
  '라이센스': { correction: '라이선스', desc: '라이선스(License)가 올바른 표준 외래어 표기입니다.', type: '외래어 표기' },
  '스케쥴': { correction: '스케줄', desc: '스케줄(Schedule)이 올바른 표준 외래어 표기입니다.', type: '외래어 표기' },
  '레포트': { correction: '리포트', desc: '리포트(Report)가 외래어 표기법에 부합합니다.', type: '외래어 표기' },
  '플렛폼': { correction: '플랫폼', desc: '플랫폼(Platform)이 표준 외래어 표기입니다.', type: '외래어 표기' },
  '디렉토리': { correction: '디렉터리', desc: '디렉터리(Directory)가 표준 외래어 표기입니다.', type: '외래어 표기' },
  '가테고리': { correction: '카테고리', desc: '카테고리(Category)가 올바른 표기입니다.', type: '외래어 표기' },
  '포퍼먼스': { correction: '퍼포먼스', desc: '퍼포먼스(Performance)가 올바른 외래어 표기입니다.', type: '외래어 표기' },
  '프로세씽': { correction: '프로세싱', desc: '프로세싱(Processing)이 올바른 외래어 표기입니다.', type: '외래어 표기' },
  '인터페이서': { correction: '인터페이스', desc: '인터페이스(Interface)가 표준 표기입니다.', type: '외래어 표기' },
  '코뮤니케이션': { correction: '커뮤니케이션', desc: '커뮤니케이션(Communication)이 올바른 외래어 표기입니다.', type: '외래어 표기' },
  '넷트웍': { correction: '네트워크', desc: '네트워크(Network)가 올바른 표준 외래어 표기입니다.', type: '외래어 표기' },
  '메세지': { correction: '메시지', desc: '메시지(Message)가 표준 외래어 표기법에 부합합니다.', type: '외래어 표기' },
  '네비게이션': { correction: '내비게이션', desc: '내비게이션(Navigation)이 표준 외래어 표기입니다.', type: '외래어 표기' },
  '카다로그': { correction: '카탈로그', desc: '카탈로그(Catalogue)가 올바른 표준 외래어 표기입니다.', type: '외래어 표기' },
  '심포지움': { correction: '심포지엄', desc: '심포지엄(Symposium)이 올바른 외래어 표기입니다.', type: '외래어 표기' },
  '컨퍼런스': { correction: '콘퍼런스', desc: '콘퍼런스(Conference)가 올바른 표준 외래어 표기입니다.', type: '외래어 표기' },
  
  // 비즈니스/용어 혼동 및 순화어
  '임계치': { correction: '임계값', desc: '순화어 권고 사항에 의하면 임계값을 사용하는 것을 권장합니다.', type: '순화어/비즈니스' },
  '가이도라인': { correction: '가이드라인', desc: '가이드라인(Guideline)의 오타 표기입니다.', type: '순화어/비즈니스' },
  '임프라': { correction: '인프라', desc: '인프라(Infrastructure)의 오타 표기입니다.', type: '순화어/비즈니스' },
  '구비서류': { correction: '제출서류', desc: '구비서류는 제출서류 또는 갖춤서류로 순화하여 사용하는 것이 권장됩니다.', type: '순화어/비즈니스' },
  '시말서': { correction: '경위서', desc: '시말서는 경위서로 순화하여 쓰는 것이 비즈니스 표준에 가깝습니다.', type: '순화어/비즈니스' },
  '도모하다': { correction: '꾀하다', desc: '가급적 쉬운 한글 표현인 꾀하다 혹은 마련하다 등으로 순화해 쓰는 것이 좋습니다.', type: '순화어/비즈니스' },
  '노하우': { correction: '비법/비결', desc: '노하우(Know-how) 대신 비법, 전문 비결 등으로 순화하여 사용하는 것이 권장됩니다.', type: '순화어/비즈니스' },
  '바이어': { correction: '구매자/거래처', desc: '바이어(Buyer) 대신 구매자 또는 거래처로 명시하는 것이 권장됩니다.', type: '순화어/비즈니스' },
  '옵티마이징': { correction: '최적화', desc: '외래어 용어 대신 순화어인 최적화를 권장합니다.', type: '순화어/비즈니스' },
  
  // 띄어쓰기
  '개선방안': { correction: '개선 방안', desc: '가독성을 위해 띄어쓰기를 적용하는 것이 좋습니다.', type: '띄어쓰기' },
  '일정계획': { correction: '일정 계획', desc: '가독성을 위해 띄어쓰기를 적용하는 것이 좋습니다.', type: '띄어쓰기' },
  '상세설계': { correction: '상세 설계', desc: '가독성을 위해 띄어쓰기를 적용하는 것이 좋습니다.', type: '띄어쓰기' },
  '수행과제': { correction: '수행 과제', desc: '가독성을 위해 띄어쓰기를 적용하는 것이 좋습니다.', type: '띄어쓰기' },
  '요구사항': { correction: '요구 사항', desc: '가독성을 위해 띄어쓰기를 적용하는 것이 좋습니다.', type: '띄어쓰기' }
};

// 오탈자 매칭 시 한국어 단어 경계 및 오탐 분석 헬퍼 함수 (스네이크 케이스 준수)
const validate_typo_match = (text, typo) => {
  let index = text.indexOf(typo);
  while (index !== -1) {
    const prev_char = index > 0 ? text[index - 1] : '';
    const next_char = index + typo.length < text.length ? text[index + typo.length] : '';

    const is_korean_or_alphanumeric = (char) => {
      if (!char) return false;
      return /[a-zA-Z0-9가-힣ㄱ-ㅎㅏ-ㅣ]/.test(char);
    };

    // 1. 앞 문자 검증: 매칭된 단어 바로 앞에 한글이나 알파벳/숫자가 연속해서 오면 다른 단어의 일부로 판단하여 제외
    // 예: '영역활성화'에서 '역활' 앞의 '영'
    let is_prev_valid = true;
    if (is_korean_or_alphanumeric(prev_char)) {
      is_prev_valid = false;
    }

    // 2. 뒤 문자 검증: 매칭된 단어 바로 뒤에 한글이나 알파벳/숫자가 오되, 조사인 경우는 매칭 허용
    // 한국어 조사의 대표적인 예: '이', '가', '을', '를', '은', '는', '에', '의', '도', '만', '과', '와', '로', '으로' 등
    // 단, 조사 뒤에 추가로 한글이 더 오면 조사 결합이 아닌 다른 단어의 일부일 확률이 높음 (예: '역활성화' -> '성', '화'가 추가로 붙음)
    let is_next_valid = true;
    if (is_korean_or_alphanumeric(next_char)) {
      const particle_regex = /^(이|가|을|를|은|는|에|의|도|만|과|와|로|로써|로서|으로|까지|부터|요|만)\b|^[.,;:!?()\s]/;
      const remaining_text = text.substring(index + typo.length);
      // 조사가 결합되었거나 바로 문장 부호/공백이 오는 경우라면 유효한 오탈자 매칭으로 봄
      if (!particle_regex.test(remaining_text)) {
        is_next_valid = false;
      }
    }

    if (is_prev_valid && is_next_valid) {
      return true;
    }

    index = text.indexOf(typo, index + 1);
  }
  return false;
};

// 한글 동사/형용사 어미 변화형을 자동으로 생성해 사전에 추가해주는 알고리즘 헬퍼 함수
const generate_conjugation_rules = (dict) => {
  const extended_dict = { ...dict };
  
  // 국어 어미 변화 정의 (하다 계열)
  const HADA_ENDINGS = [
    { suffix: '합니다', corr_suffix: '합니다' },
    { suffix: '하고', corr_suffix: '하고' },
    { suffix: '한', corr_suffix: '한' },
    { suffix: '할', corr_suffix: '할' },
    { suffix: '하며', corr_suffix: '하며' },
    { suffix: '하면', corr_suffix: '하면' },
    { suffix: '하므로', corr_suffix: '하므로' },
    { suffix: '하여', corr_suffix: '하여' },
    { suffix: '했다', corr_suffix: '했다' },
    { suffix: '했음', corr_suffix: '했음' },
    { suffix: '함', corr_suffix: '함' },
    { suffix: '하되', corr_suffix: '하되' },
    { suffix: '하는', corr_suffix: '하는' },
    { suffix: '하신', corr_suffix: '하신' }
  ];

  // 되다 계열 어미 변화 정의
  const DOEDA_ENDINGS = [
    { suffix: '됩니다', corr_suffix: '됩니다' },
    { suffix: '되고', corr_suffix: '되고' },
    { suffix: '된', corr_suffix: '된' },
    { suffix: '될', corr_suffix: '될' },
    { suffix: '되며', corr_suffix: '되며' },
    { suffix: '되면', corr_suffix: '되면' },
    { suffix: '되므로', corr_suffix: '되므로' },
    { suffix: '되어', corr_suffix: '되어' },
    { suffix: '됐다', corr_suffix: '됐다' },
    { suffix: '됐음', corr_suffix: '됐음' },
    { suffix: '됨', corr_suffix: '됨' },
    { suffix: '되되', corr_suffix: '되되' },
    { suffix: '되는', corr_suffix: '되는' }
  ];

  Object.keys(dict).forEach(key => {
    const item = dict[key];
    
    // 1. ~하다 계열 파생
    if (key.endsWith('하다') && item.correction.endsWith('하다')) {
      const typo_stem = key.substring(0, key.length - 2); // '제사'
      const corr_stem = item.correction.substring(0, item.correction.length - 2); // '제시'
      
      HADA_ENDINGS.forEach(ending => {
        const derived_typo = typo_stem + ending.suffix;
        const derived_corr = corr_stem + ending.corr_suffix;
        if (!extended_dict[derived_typo]) {
          extended_dict[derived_typo] = {
            correction: derived_corr,
            desc: `어미 변화 규칙: [${key} ➜ ${item.correction}]의 파생형 오류입니다.`,
            type: item.type
          };
        }
      });
    }

    // 2. ~되다 계열 파생
    if (key.endsWith('되다') && item.correction.endsWith('되다')) {
      const typo_stem = key.substring(0, key.length - 2);
      const corr_stem = item.correction.substring(0, item.correction.length - 2);
      
      DOEDA_ENDINGS.forEach(ending => {
        const derived_typo = typo_stem + ending.suffix;
        const derived_corr = corr_stem + ending.corr_suffix;
        if (!extended_dict[derived_typo]) {
          extended_dict[derived_typo] = {
            correction: derived_corr,
            desc: `어미 변화 규칙: [${key} ➜ ${item.correction}]의 파생형 오류입니다.`,
            type: item.type
          };
        }
      });
    }
  });

  return extended_dict;
};

// 동일 단어 중복 검출 시 단어 경계(Boundary) 유효성 검사 헬퍼 함수 (스네이크 케이스 준수)
const check_word_boundary = (text, match_index, matched_length, duplicated_word) => {
  const start1 = match_index;
  const end1 = start1 + duplicated_word.length;
  const start2 = match_index + matched_length - duplicated_word.length;
  const end2 = match_index + matched_length;

  const is_alphanumeric = (char) => {
    if (!char) return false;
    return /[a-zA-Z0-9가-힣ㄱ-ㅎㅏ-ㅣ]/.test(char);
  };

  const word_start_char = duplicated_word[0];
  const word_end_char = duplicated_word[duplicated_word.length - 1];

  // 첫 번째 유닛의 앞 경계 검사
  if (is_alphanumeric(word_start_char)) {
    const char_before1 = text[start1 - 1];
    if (is_alphanumeric(char_before1)) return false;
  }

  // 첫 번째 유닛의 뒤 경계 검사
  if (is_alphanumeric(word_end_char)) {
    const char_after1 = text[end1];
    // 두 번째 유닛과 바로 붙어 있지 않은 경우에만 뒤 문자 검사
    if (end1 !== start2 && is_alphanumeric(char_after1)) return false;
  }

  // 두 번째 유닛의 앞 경계 검사
  if (is_alphanumeric(word_start_char)) {
    const char_before2 = text[start2 - 1];
    // 첫 번째 유닛과 바로 붙어 있지 않은 경우에만 앞 문자 검사
    if (end1 !== start2 && is_alphanumeric(char_before2)) return false;
  }

  // 두 번째 유닛의 뒤 경계 검사
  if (is_alphanumeric(word_end_char)) {
    const char_after2 = text[end2];
    if (is_alphanumeric(char_after2)) return false;
  }

  return true;
};

// ── 로마자 변환 헬퍼 함수 ─────────────────────────
const roman_to_int = (roman) => {
  const map = {
    'I': 1, 'II': 2, 'III': 3, 'IV': 4, 'V': 5,
    'VI': 6, 'VII': 7, 'VIII': 8, 'IX': 9, 'X': 10
  };
  return map[roman.toUpperCase()] || null;
};

// Gemini API를 직접 호출하여 문맥적인 오탈자 및 맞춤법을 진단받는 비동기 함수
const call_gemini_ai_typos = async (file_name, slides_text_map, api_key) => {
  const keys = String(api_key).split(',').map(k => k.trim()).filter(k => k.match(/^(AIza|AQ\.)/));
  if (keys.length === 0) {
    throw new Error('유효한 Gemini API 키가 없습니다. 설정 탭에서 API 키를 등록해 주세요.');
  }
  const key = keys[0];

  let document_content = '';
  Object.keys(slides_text_map).forEach(slide_num => {
    document_content += `[Slide ${slide_num}]\n${slides_text_map[slide_num].join('\n')}\n\n`;
  });

  const prompt = `당신은 공공 IT 사업 제안서·수행계획서·설계서 분야에 특화된 한국어 맞춤법 및 문서 교정 전문가입니다.
아래 슬라이드 텍스트를 **극도로 엄격(Strict Mode)**하게 분석하여, 다음 6가지 유형의 오류를 빠짐없이 찾아 주세요.

[검사 유형 및 기준]
1. **오타·오자**: 자모 오입력으로 인한 명백한 오타. 예) 제사합니다→제시합니다, 와부→외부, 아키텍쳐→아키텍처
2. **혼동 어휘**: 의미가 달라지는 동음이의어/유사어 오용. 예) 결재(決裁)↔결제(決濟), 대표(代表)↔데표, 계발(啓發)↔개발(開發)
3. **외래어 표기**: 국립국어원 외래어 표기법 위반. 예) 컨텐츠→콘텐츠, 레포트→리포트, 스케쥴→스케줄, 플렛폼→플랫폼
4. **띄어쓰기 오류**: 조사·어미·의존명사 등 표준 띄어쓰기 위반. 예) 확인바랍니다→확인 바랍니다, 될수있는→될 수 있는
5. **이중 피동·이중 부정**: 비문법적 표현. 예) ~되어지다(→~되다), ~지 않을 수가 없다
6. **비격식·구어체 표현**: 공문서에 어울리지 않는 표현. 예) 이것보다→이보다, 좀→다소

[절대 지적 금지 항목]
- 복합명사 붙여쓰기 (예: 데이터전송, 클라우드인프라)
- 숫자+의존명사 (예: 3개, 5단계, 2개월)
- 영문 약어 및 전문 기술 용어 (예: AI, API, ISP, ISMP, UI, UX)
- 고유명사(기관명, 시스템명)
- 슬라이드 번호·날짜 등 형식 데이터

[응답 규칙]
- 반드시 아래 JSON 배열 형식으로만 응답하십시오. 다른 텍스트·마크다운·설명을 절대 포함하지 마십시오.
- 오류가 없을 경우 빈 배열 []을 반환하십시오.
- 동일 슬라이드에서 여러 오류가 있을 경우 각각 별도 항목으로 반환하십시오.

[응답 JSON 형식]
[
  {
    "slideNum": <슬라이드번호 (숫자)>,
    "originalText": "<오류가 포함된 문장 전체>",
    "typo": "<오류 단어 또는 구절>",
    "correction": "<교정 단어 또는 구절>",
    "desc": "<오류 유형 및 간략 설명>"
  }
]

[검토할 슬라이드 텍스트]
${document_content}`;

  const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${key}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: {
        responseMimeType: 'application/json'
      }
    })
  });

  if (!response.ok) {
    const err_text = await response.text();
    throw new Error(`Gemini API 호출 실패: ${err_text}`);
  }

  const res_json = await response.json();
  const text_response = res_json.candidates[0].content.parts[0].text;
  
  try {
    const parsed = JSON.parse(text_response.trim());
    return parsed.map(item => ({
      fileName: file_name,
      slideNum: parseInt(item.slideNum || '1', 10),
      sentence: item.originalText,
      typo: item.typo,
      correction: item.correction,
      type: 'AI 맞춤법 (Gemini)',
      desc: item.desc,
      isAi: true // AI 검출 항목 식별용
    }));
  } catch (err) {
    console.error('Gemini JSON 파싱 오류:', err, text_response);
    return [];
  }
};

// 외부 전문 맞춤법 검사 API를 호출하는 비동기 함수 (선택 옵션)
const call_external_spell_check = async (file_name, slides_text_map, api_url) => {
  if (!api_url || !api_url.trim()) {
    throw new Error('외부 맞춤법 검사 API URL이 설정되지 않았습니다. 설정에서 API 주소를 입력해 주세요.');
  }

  const results = [];

  for (const slide_num of Object.keys(slides_text_map)) {
    const text = slides_text_map[slide_num].join('\n');
    if (!text.trim()) continue;

    try {
      const response = await fetch(api_url.trim(), {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ text, slide: parseInt(slide_num, 10) })
      });

      if (!response.ok) {
        console.warn(`외부 API 슬라이드 ${slide_num} 호출 실패: ${response.status}`);
        continue;
      }

      const data = await response.json();

      // 외부 API 응답 규격: { errors: [{ typo, correction, desc, sentence }] }
      if (Array.isArray(data.errors)) {
        data.errors.forEach(item => {
          if (item.typo && item.correction) {
            results.push({
              fileName: file_name,
              slideNum: parseInt(slide_num, 10),
              sentence: item.sentence || text,
              typo: item.typo,
              correction: item.correction,
              type: '외부 맞춤법 API',
              desc: item.desc || '외부 맞춤법 검사 API 검출 항목',
              isAi: true
            });
          }
        });
      }
    } catch (err) {
      console.warn(`외부 API 슬라이드 ${slide_num} 오류:`, err.message);
    }
  }

  return results;
};

export default function PptValidator({ apiKey }) {
  const [pptFiles, setPptFiles] = useState([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isValidated, setIsValidated] = useState(false);
  
  // 검증 옵션 분리 상태
  const [checkTypos, setCheckTypos] = useState(true);
  const [check_ai_typos, set_check_ai_typos] = useState(false);
  const [check_external_api, set_check_external_api] = useState(false);
  const [external_api_url, set_external_api_url] = useState('');
  const [external_server_status, set_external_server_status] = useState('unknown'); // 'unknown' | 'checking' | 'running' | 'stopped'
  const [checkNumbering, setCheckNumbering] = useState(true);
  const [checkAltText, setCheckAltText] = useState(true);
  const [checkForbiddenWords, setCheckForbiddenWords] = useState(true);
  const [checkEngKoMixed, setCheckEngKoMixed] = useState(true);
  // 추가 옵션: 동일 단어 중복 검증 (스네이크 케이스 규칙 적용)
  const [check_duplicate_words, set_check_duplicate_words] = useState(true);
  const [checkPageRange, setCheckPageRange] = useState(true);
  const [checkMacImages, setCheckMacImages] = useState(true);
  
  // 결과 데이터 저장
  const [typoResults, setTypoResults] = useState([]);
  const [macImageResults, setMacImageResults] = useState([]);
  const [numberingResults, setNumberingResults] = useState([]);
  const [altTextResults, setAltTextResults] = useState([]);
  const [forbiddenResults, setForbiddenResults] = useState([]);
  const [engKoMixedResults, setEngKoMixedResults] = useState([]);
  // 동일 단어 중복 검증 결과 저장 (스네이크 케이스 규칙 적용)
  const [duplicate_results, set_duplicate_results] = useState([]);
  const [pageRangeResults, setPageRangeResults] = useState([]);
  const [fileStats, setFileStats] = useState([]); // [{ name: '', typos: 0, numberingErrors: 0, altTextErrors: 0, forbiddenErrors: 0, engKoMixedErrors: 0, duplicateErrors: 0, startPage: 1, endPage: 1, totalSlides: 1 }]
  
  const [activeResultTab, setActiveResultTab] = useState('summary'); // summary, typo, numbering, altText, forbidden
  const [userDictText, setUserDictText] = useState(() => {
    return localStorage.getItem('ppt_validator_user_dict') || '';
  }); // 사용자 정의 사전 입력란
  const [userDictTab, setUserDictTab] = useState('table'); // 'table' 또는 'text'
  const dictFileInputRef = useRef(null);
  const [showDictModal, setShowDictModal] = useState(false);
  const [dictFilterCategory, setDictFilterCategory] = useState('전체');

  // 현재 사용자 사전을 [{typo, correction}] 리스트로 파싱하는 헬퍼
  const get_user_dict_list = () => {
    const list = [];
    if (!userDictText.trim()) return list;
    const lines = userDictText.split('\n');
    lines.forEach(line => {
      const parts = line.split(/[=➜➔>:\-]/);
      if (parts.length >= 2) {
        const typo = parts[0].trim();
        const correction = parts[1].trim();
        if (typo && correction) {
          list.push({ typo, correction });
        }
      }
    });
    return list;
  };

  // 리스트를 다시 userDictText로 동기화
  const save_user_dict_list = (list) => {
    const text = list.map(item => `${item.typo} ➜ ${item.correction}`).join('\n');
    setUserDictText(text);
  };

  // 행 추가
  const handle_add_dict_row = (typo, correction) => {
    if (!typo || !correction) return;
    const list = get_user_dict_list();
    if (list.some(item => item.typo === typo)) {
      alert('이미 등록된 오탈자 단어입니다.');
      return;
    }
    list.push({ typo, correction });
    save_user_dict_list(list);
  };

  // 행 삭제
  const handle_delete_dict_row = (typo) => {
    const list = get_user_dict_list();
    const filtered = list.filter(item => item.typo !== typo);
    save_user_dict_list(filtered);
  };

  // 엑셀 내보내기 (다운로드)
  const handle_export_dict_excel = () => {
    const list = get_user_dict_list();
    if (list.length === 0) {
      alert('내보낼 오탈자 규칙이 없습니다. 먼저 규칙을 등록해 주세요.');
      return;
    }
    const data = list.map(item => ({
      '오탈자 단어': item.typo,
      '교정 단어': item.correction
    }));
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, '사용자 정의 사전');
    XLSX.writeFile(workbook, '오탈자_사용자정의사전.xlsx');
  };

  // 템플릿 다운로드
  const handle_download_dict_template = () => {
    const sample_data = [
      { '오탈자 단어': '아키택처', '교정 단어': '아키텍처' },
      { '오탈자 단어': '스프링부터', '교정 단어': '스프링부트' }
    ];
    const worksheet = XLSX.utils.json_to_sheet(sample_data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, '템플릿');
    XLSX.writeFile(workbook, '오탈자_사전_템플릿.xlsx');
  };

  // 엑셀 가져오기 (업로드)
  const handle_import_dict_excel = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet_name = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheet_name];
        const json = XLSX.utils.sheet_to_json(worksheet);

        const list = get_user_dict_list();
        let added_count = 0;

        json.forEach(row => {
          const typo = (row['오탈자 단어'] || row['오탈자'] || row['typo'] || '').toString().trim();
          const correction = (row['교정 단어'] || row['교정'] || row['correction'] || '').toString().trim();

          if (typo && correction) {
            if (!list.some(item => item.typo === typo)) {
              list.push({ typo, correction });
              added_count++;
            }
          }
        });

        if (added_count > 0) {
          save_user_dict_list(list);
          alert(`엑셀 파일에서 ${added_count}개의 규칙을 성공적으로 가져왔습니다.`);
        } else {
          alert('가져올 새로운 규칙이 없거나 엑셀 포맷이 잘못되었습니다. [오탈자 단어], [교정 단어] 헤더가 포함되어 있는지 확인해 주세요.');
        }
      } catch (err) {
        console.error('엑셀 파싱 실패:', err);
        alert('엑셀 파일을 읽는 도중 오류가 발생했습니다.');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const [forbiddenWordsText, setForbiddenWordsText] = useState(() => {
    const saved = localStorage.getItem('ppt_validator_forbidden_words');
    return saved !== null ? saved : "미정\n임시\nTBD\n검토필요\n작성중";
  }); // 특정 점검 단어 입력란
  const fileInputRef = useRef(null);

  useEffect(() => {
    localStorage.setItem('ppt_validator_user_dict', userDictText);
  }, [userDictText]);

  useEffect(() => {
    localStorage.setItem('ppt_validator_forbidden_words', forbiddenWordsText);
  }, [forbiddenWordsText]);

  // 파일 업로드 핸들러
  const handleFileChange = (e) => {
    const files = Array.from(e.target.files).filter(f => f.name.endsWith('.pptx') || f.name.endsWith('.hwpx'));
    if (files.length > 0) {
      setPptFiles(prev => [...prev, ...files]);
      setIsValidated(false);
    }
  };

  const handleDragOver = (e) => {
    e.preventDefault();
  };

  const handleDrop = (e) => {
    e.preventDefault();
    const files = Array.from(e.dataTransfer.files).filter(f => f.name.endsWith('.pptx') || f.name.endsWith('.hwpx'));
    if (files.length > 0) {
      setPptFiles(prev => [...prev, ...files]);
      setIsValidated(false);
    }
  };

  const removeFile = (index) => {
    setPptFiles(prev => prev.filter((_, i) => i !== index));
    setIsValidated(false);
  };

  const clearAllFiles = () => {
    setPptFiles([]);
    setIsValidated(false);
    setTypoResults([]);
    setNumberingResults([]);
    setAltTextResults([]);
    setForbiddenResults([]);
    setEngKoMixedResults([]);
    set_duplicate_results([]);
    setPageRangeResults([]);
    setMacImageResults([]);
    setFileStats([]);
  };

  // 사용자 정의 사전 파싱
  const parseUserDictionary = () => {
    const dict = {};
    if (!userDictText.trim()) return dict;

    const lines = userDictText.split('\n');
    lines.forEach(line => {
      const parts = line.split(/[=➜➔>:\-]/);
      if (parts.length >= 2) {
        const typo = parts[0].trim();
        const correction = parts[1].trim();
        if (typo && correction) {
          dict[typo] = {
            correction,
            desc: `사용자가 정의한 오탈자 교정 규칙입니다. (${typo} ➜ ${correction})`,
            type: '사용자 정의'
          };
        }
      }
    });
    return dict;
  };

  // PPTX 검증 핵심 프로세스
  const handleValidate = async () => {
    if (pptFiles.length === 0) return;
    if (!checkTypos && !checkNumbering && !checkAltText && !checkForbiddenWords && !checkEngKoMixed && !check_duplicate_words && !checkPageRange) {
      alert('오탈자, 넘버링, 대체텍스트, 특정 단어, 영어/한글 혼용 단어, 동일 단어 중복, 페이지 범위 분석 중 최소 하나 이상의 검증 옵션을 선택해야 합니다.');
      return;
    }
    setIsProcessing(true);
    
    const allTypos = [];
    const allNumberings = [];
    const allAltTexts = [];
    const allForbiddens = [];
    const allEngKoMixed = [];
    const all_duplicates = [];
    const allPageRanges = [];
    const allMacImageErrors = [];
    const stats = [];
    const userDict = parseUserDictionary();
    const mergedDict = generate_conjugation_rules({ ...TYPO_DICTIONARY, ...userDict });

    try {
      for (const file of pptFiles) {
        let fileTyposCount = 0;
        let fileNumErrorsCount = 0;
        let fileAltErrorsCount = 0;
        let fileForbiddenCount = 0;
        let fileEngKoMixedCount = 0;
        let file_duplicate_count = 0;
        let fileMacImageCount = 0;

        const slides_text_map = {};

        const arrayBuffer = await file.arrayBuffer();
        const zip = await JSZip.loadAsync(arrayBuffer);
        
        const is_hwpx = file.name.endsWith('.hwpx');
        
        if (is_hwpx) {
          const section_files = Object.keys(zip.files).filter(p => 
            p.startsWith('Contents/section') && p.endsWith('.xml')
          ).sort((a, b) => {
            const numA = parseInt(a.replace(/[^\d]/g, ''), 10);
            const numB = parseInt(b.replace(/[^\d]/g, ''), 10);
            return numA - numB;
          });
          
          for (let sIdx = 0; sIdx < section_files.length; sIdx++) {
            const section_path = section_files[sIdx];
            const section_xml_str = await zip.file(section_path).async('text');
            
            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(section_xml_str, 'application/xml');
            if (xmlDoc.getElementsByTagName('parsererror').length > 0) continue;
            
            const allNodes = xmlDoc.getElementsByTagName('*');
            for (let i = 0; i < allNodes.length; i++) {
              const node = allNodes[i];
              const localName = node.localName || node.tagName.split(':').pop();
              if (localName === 'p') {
                const tNodes = node.getElementsByTagName('*');
                let p_text_list = [];
                for (let j = 0; j < tNodes.length; j++) {
                  const tNode = tNodes[j];
                  const tLocalName = tNode.localName || tNode.tagName.split(':').pop();
                  if (tLocalName === 't') {
                    p_text_list.push(tNode.textContent || '');
                  }
                }
                const paragraph_text = p_text_list.join('').trim();
                if (paragraph_text) {
                  const slideNum = sIdx + 1;
                  if (!slides_text_map[slideNum]) {
                    slides_text_map[slideNum] = [];
                  }
                  slides_text_map[slideNum].push(paragraph_text);
                  
                  // 1) 오탈자 점검
                  if (checkTypos) {
                    Object.keys(mergedDict).forEach(typo => {
                      if (paragraph_text.includes(typo) && validate_typo_match(paragraph_text, typo)) {
                        const exists = allTypos.some(t => 
                          t.fileName === file.name && 
                          t.slideNum === slideNum && 
                          t.sentence === paragraph_text && 
                          t.typo === typo
                        );
                        if (!exists) {
                          const info = mergedDict[typo];
                          allTypos.push({
                            fileName: file.name,
                            slideNum,
                            sentence: paragraph_text,
                            typo,
                            correction: info.correction,
                            type: info.type,
                            desc: info.desc
                          });
                          fileTyposCount++;
                        }
                      }
                    });
                  }
                  
                  // 2) 특정 단어 검증
                  if (checkForbiddenWords && forbiddenWordsText.trim()) {
                    const targetWords = forbiddenWordsText.split('\n')
                      .map(w => w.trim())
                      .filter(w => w.length > 0);
                    targetWords.forEach(word => {
                      if (paragraph_text.includes(word)) {
                        const exists = allForbiddens.some(e => 
                          e.fileName === file.name && 
                          e.slideNum === slideNum && 
                          e.sentence === paragraph_text && 
                          e.word === word
                        );
                        if (!exists) {
                          allForbiddens.push({
                            fileName: file.name,
                            slideNum,
                            sentence: paragraph_text,
                            word,
                            error: `지정 단어 검출 (${word})`,
                            guide: `문서 본문 내에 점검 지정 단어인 "${word}"가 포함되어 있습니다. 최종 제출 시 적절한 표현인지 확인해 주세요.`
                          });
                          fileForbiddenCount++;
                        }
                      }
                    });
                  }
                  
                  // 3) 영어/한글 혼용 단어 검증
                  if (checkEngKoMixed) {
                    const words = paragraph_text.split(/\s+/);
                    words.forEach(word => {
                      const cleanWord = word.replace(/^[.,;:!?()\[\]"']+|[.,;:!?()\[\]"']+$/g, '');
                      const mixRegex = /[a-zA-Z][가-힣ㄱ-ㅎㅏ-ㅣ]|[가-힣ㄱ-ㅎㅏ-ㅣ][a-zA-Z]/;
                      if (mixRegex.test(cleanWord)) {
                        const exists = allEngKoMixed.some(e => 
                          e.fileName === file.name && 
                          e.slideNum === slideNum && 
                          e.word === cleanWord &&
                          e.sentence === paragraph_text
                        );
                        if (!exists) {
                          allEngKoMixed.push({
                            fileName: file.name,
                            slideNum,
                            sentence: paragraph_text,
                            word: cleanWord,
                            error: '영어/한글 혼용 단어 검출',
                            guide: `단어 "${cleanWord}" 내에 영어와 한글이 공백 없이 혼용되어 표시되어 있습니다. 오타가 아닌지 혹은 의도된 결합 표기인지 확인해 주세요.`
                          });
                          fileEngKoMixedCount++;
                        }
                      }
                    });
                  }
                  
                  // 4) 동일 단어 중복 검증
                  if (check_duplicate_words) {
                    const matches = [...paragraph_text.matchAll(/(\S{2,})\s*\1/g)];
                    matches.forEach(match => {
                      const duplicated_word = match[1];
                      const matched_text = match[0];
                      
                      // 숫자로만 구성되었거나 숫자와 결합된 단순 기호 시퀀스인 경우 중복 검사 제외
                      const is_numeric = /^[0-9.,()\-/[\]\s]+$/.test(duplicated_word);
                      if (is_numeric) return;

                      // 단어 경계가 제대로 지켜지지 않은 경우 제외 (예: xxxse sexxx 등)
                      if (!check_word_boundary(paragraph_text, match.index, match[0].length, duplicated_word)) return;

                      const exists = all_duplicates.some(e => 
                        e.fileName === file.name && 
                        e.slideNum === slideNum && 
                        e.sentence === paragraph_text && 
                        e.word === matched_text
                      );
                      if (!exists) {
                        all_duplicates.push({
                          fileName: file.name,
                          slideNum,
                          sentence: paragraph_text,
                          word: matched_text,
                          duplicateWord: duplicated_word,
                          error: '동일 단어 중복 검출',
                          guide: `문서 본문 내에 동일한 단어 "${duplicated_word}"가 연속으로 중복 기재되어 있습니다. ("${matched_text}")`
                        });
                        file_duplicate_count++;
                      }
                    });
                  }
                }
              }
            }
          }
          allPageRanges.push({
            fileName: file.name,
            startPage: 1,
            endPage: section_files.length,
            totalSlides: section_files.length
          });
        } else {
        // 1. 모든 슬라이드 파일 추출 및 정렬
        const slideFiles = Object.keys(zip.files).filter(p => 
          p.startsWith('ppt/slides/slide') && p.endsWith('.xml')
        ).sort((a, b) => {
          const numA = parseInt(a.replace(/[^\d]/g, ''), 10);
          const numB = parseInt(b.replace(/[^\d]/g, ''), 10);
          return numA - numB;
        });

        // 미디어 리소스 목록 추출 (맥 이미지 누락 대조용)
        const mediaFiles = Object.keys(zip.files).filter(p => p.startsWith('ppt/media/'));
        const mediaFilesLower = mediaFiles.map(m => m.toLowerCase());
 
        const slideList = []; // 각 슬라이드의 타이틀 정보 및 텍스트 데이터 수집용
        const detectedPages = []; // 파일별 감지된 페이지 번호 수집용

        // 2. 각 슬라이드 XML 해석
        for (let sIdx = 0; sIdx < slideFiles.length; sIdx++) {
          const slidePath = slideFiles[sIdx];
          const slideNum = sIdx + 1;
          const slideXmlStr = await zip.file(slidePath).async('text');
          
          const parser = new DOMParser();
          const xmlDoc = parser.parseFromString(slideXmlStr, 'application/xml');
          
          if (xmlDoc.getElementsByTagName('parsererror').length > 0) continue;

          // 슬라이드 내 모든 shape(<p:sp>) 노드 추출
          const allNodes = xmlDoc.getElementsByTagName('*');
          const shapes = [];

          for (let i = 0; i < allNodes.length; i++) {
            const node = allNodes[i];
            const localName = node.localName || node.tagName.split(':').pop();
            if (localName === 'sp') {
              // 텍스트 및 좌표값 추출
              let x = null, y = null, cx = null, cy = null;
              let textContent = '';

              // 2-1. 좌표 값(off 및 ext) 찾기
              const offNodes = node.getElementsByTagName('a:off');
              const extNodes = node.getElementsByTagName('a:ext');
              
              // 네임스페이스 매칭 안될 경우 대비 폴백 순회
              const childs = node.getElementsByTagName('*');
              let offNode = offNodes.length > 0 ? offNodes[0] : null;
              let extNode = extNodes.length > 0 ? extNodes[0] : null;
              
              if (!offNode || !extNode) {
                for (let k = 0; k < childs.length; k++) {
                  const cNode = childs[k];
                  const cName = cNode.localName || cNode.tagName.split(':').pop();
                  if (cName === 'off') offNode = cNode;
                  if (cName === 'ext') extNode = cNode;
                }
              }

              if (offNode) {
                x = parseInt(offNode.getAttribute('x') || '0', 10);
                y = parseInt(offNode.getAttribute('y') || '0', 10);
              }
              if (extNode) {
                cx = parseInt(extNode.getAttribute('cx') || '0', 10);
                cy = parseInt(extNode.getAttribute('cy') || '0', 10);
              }

              // 2-2. 텍스트 추출
              const tNodes = node.getElementsByTagName('a:t');
              let tList = [];
              if (tNodes.length > 0) {
                for (let k = 0; k < tNodes.length; k++) {
                  tList.push(tNodes[k].textContent || '');
                }
              } else {
                for (let k = 0; k < childs.length; k++) {
                  const cNode = childs[k];
                  const cName = cNode.localName || cNode.tagName.split(':').pop();
                  if (cName === 't') tList.push(cNode.textContent || '');
                }
              }
              textContent = tList.join(' ').trim();

              if (textContent) {
                shapes.push({ x, y, cx, cy, text: textContent });
                if (!slides_text_map[slideNum]) {
                  slides_text_map[slideNum] = [];
                }
                slides_text_map[slideNum].push(textContent);
              }
            }
          }

          // 3. 오탈자 점검 수행 (모든 텍스트 대상)
          if (checkTypos) {
            shapes.forEach(shape => {
              const text = shape.text;
              
              // 등록된 사전 키워드들 검사
              Object.keys(mergedDict).forEach(typo => {
                // 한글 조사 등 경계 고려 정규식 생성
                // 예: '역활이', '역활을'도 잡히도록 단어 포함 여부 검증
                if (text.includes(typo) && validate_typo_match(text, typo)) {
                  // 문장 내 중복 검출 방지
                  const exists = allTypos.some(t => 
                    t.fileName === file.name && 
                    t.slideNum === slideNum && 
                    t.sentence === text && 
                    t.typo === typo
                  );

                  if (!exists) {
                    const info = mergedDict[typo];
                    allTypos.push({
                      fileName: file.name,
                      slideNum,
                      sentence: text,
                      typo,
                      correction: info.correction,
                      type: info.type,
                      desc: info.desc
                    });
                    fileTyposCount++;
                  }
                }
              });
            });
          }

          // 3-2. 대체텍스트 검색 및 추출 수행
          if (checkAltText) {
            const cNvPrs = xmlDoc.getElementsByTagName('*');
            for (let i = 0; i < cNvPrs.length; i++) {
              const node = cNvPrs[i];
              const localName = node.localName || node.tagName.split(':').pop();
              if (localName === 'cNvPr') {
                const name = node.getAttribute('name') || '';
                const descr = node.getAttribute('descr') || '';
                const parentNode = node.parentNode;
                const parentName = parentNode ? (parentNode.localName || parentNode.tagName.split(':').pop()) : '';
                
                // 그림(pic), 비디오, 그룹, 스마트아트 등 대체텍스트가 필요한 성격의 개체들 체크
                const isVisualElement = parentName === 'pic' || parentName === 'graphicFrame' || parentName === 'grpSp' || 
                  name.toLowerCase().includes('picture') || name.toLowerCase().includes('그림') || 
                  name.toLowerCase().includes('image') || name.toLowerCase().includes('이미지') ||
                  name.toLowerCase().includes('chart') || name.toLowerCase().includes('차트') ||
                  name.toLowerCase().includes('diagram') || name.toLowerCase().includes('다이어그램');
                
                if (isVisualElement && descr.trim()) {
                  const exists = allAltTexts.some(e => 
                    e.fileName === file.name && 
                    e.slideNum === slideNum && 
                    e.objName === name &&
                    e.descr === descr
                  );
                  if (!exists) {
                    allAltTexts.push({
                      fileName: file.name,
                      slideNum,
                      objName: name,
                      descr: descr,
                      guide: '개체에 등록되어 있는 대체텍스트 내용입니다.'
                    });
                    fileAltErrorsCount++;
                  }
                }
              }
            }
          }

          // 3-3. 특정 단어 검증 수행
          if (checkForbiddenWords && forbiddenWordsText.trim()) {
            const targetWords = forbiddenWordsText.split('\n')
              .map(w => w.trim())
              .filter(w => w.length > 0);

            shapes.forEach(shape => {
              const text = shape.text;
              targetWords.forEach(word => {
                if (text.includes(word)) {
                  const exists = allForbiddens.some(e => 
                    e.fileName === file.name && 
                    e.slideNum === slideNum && 
                    e.sentence === text && 
                    e.word === word
                  );
                  if (!exists) {
                    allForbiddens.push({
                      fileName: file.name,
                      slideNum,
                      sentence: text,
                      word,
                      error: `지정 단어 검출 (${word})`,
                      guide: `문서 본문 내에 점검 지정 단어인 "${word}"가 포함되어 있습니다. 최종 제출 시 적절한 표현인지 확인해 주세요.`
                    });
                    fileForbiddenCount++;
                  }
                }
              });
            });
          }

          // 3-4. 영어/한글 혼용 단어 검증 수행
          if (checkEngKoMixed) {
            shapes.forEach(shape => {
              const text = shape.text;
              const words = text.split(/\s+/);
              words.forEach(word => {
                const cleanWord = word.replace(/^[.,;:!?()\[\]"']+|[.,;:!?()\[\]"']+$/g, '');
                const mixRegex = /[a-zA-Z][가-힣ㄱ-ㅎㅏ-ㅣ]|[가-힣ㄱ-ㅎㅏ-ㅣ][a-zA-Z]/;
                if (mixRegex.test(cleanWord)) {
                  const exists = allEngKoMixed.some(e => 
                    e.fileName === file.name && 
                    e.slideNum === slideNum && 
                    e.word === cleanWord &&
                    e.sentence === text
                  );
                  if (!exists) {
                    allEngKoMixed.push({
                      fileName: file.name,
                      slideNum,
                      sentence: text,
                      word: cleanWord,
                      error: '영어/한글 혼용 단어 검출',
                      guide: `단어 "${cleanWord}" 내에 영어와 한글이 공백 없이 혼용되어 표시되어 있습니다. 오타가 아닌지(예: re에이전트 등) 혹은 의도된 결합 표기인지 확인해 주세요.`
                    });
                    fileEngKoMixedCount++;
                  }
                }
              });
            });
          }

          // 3-5. 동일 단어 중복 검증 수행
          if (check_duplicate_words) {
            shapes.forEach(shape => {
              const text = shape.text;
              const matches = [...text.matchAll(/(\S{2,})\s*\1/g)];
              matches.forEach(match => {
                const duplicated_word = match[1];
                const matched_text = match[0];
                
                // 숫자로만 구성되었거나 숫자와 결합된 단순 기호 시퀀스인 경우 중복 검사 제외
                const is_numeric = /^[0-9.,()\-/[\]\s]+$/.test(duplicated_word);
                if (is_numeric) return;

                // 단어 경계가 제대로 지켜지지 않은 경우 제외 (예: xxxse sexxx 등)
                if (!check_word_boundary(text, match.index, match[0].length, duplicated_word)) return;

                const exists = all_duplicates.some(e => 
                  e.fileName === file.name && 
                  e.slideNum === slideNum && 
                  e.sentence === text && 
                  e.word === matched_text
                );
                if (!exists) {
                  all_duplicates.push({
                    fileName: file.name,
                    slideNum,
                    sentence: text,
                    word: matched_text,
                    duplicateWord: duplicated_word,
                    error: '동일 단어 중복 검출',
                    guide: `문서 본문 내에 동일한 단어 "${duplicated_word}"가 연속으로 중복 기재되어 있습니다. ("${matched_text}")`
                  });
                  file_duplicate_count++;
                }
              });
            });
          }

          // 4. 슬라이드 상단 타이틀 수집 및 넘버링 분류 준비
          // 보통 상단 타이틀은 Y 좌표가 약 1,300,000 EMU 이하인 개체에 해당
          // 전체 슬라이드 X 중앙선은 대략 6,000,000 EMU
          // 3-6. 맥 이미지 누락 검증 수행
          if (checkMacImages) {
            const slideRelsPath = slidePath.replace('ppt/slides/slide', 'ppt/slides/_rels/slide') + '.rels';
            let relsMap = {};
            if (zip.files[slideRelsPath]) {
              try {
                const relsXmlStr = await zip.file(slideRelsPath).async('text');
                const relsDoc = parser.parseFromString(relsXmlStr, 'application/xml');
                const relationships = relsDoc.getElementsByTagName('Relationship');
                for (let rIdx = 0; rIdx < relationships.length; rIdx++) {
                  const rel = relationships[rIdx];
                  const rId = rel.getAttribute('Id');
                  const target = rel.getAttribute('Target');
                  const type = rel.getAttribute('Type');
                  relsMap[rId] = { target, type };
                }
              } catch (relsErr) {
                console.warn(`${slideRelsPath} 관계 파일 파싱 실패:`, relsErr);
              }
            }

            const blipNodes = xmlDoc.getElementsByTagName('*');
            const blips = [];
            for (let bIdx = 0; bIdx < blipNodes.length; bIdx++) {
              const bNode = blipNodes[bIdx];
              const bName = bNode.localName || bNode.tagName.split(':').pop();
              if (bName === 'blip') {
                let rEmbed = bNode.getAttribute('r:embed') || bNode.getAttribute('embed') || bNode.getAttribute('r:link') || bNode.getAttribute('link');
                if (!rEmbed) {
                  const embedAttr = bNode.getAttributeNodeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed') || 
                                    bNode.getAttributeNodeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'link');
                  if (embedAttr) rEmbed = embedAttr.value;
                }
                if (rEmbed) {
                  blips.push(rEmbed);
                }
              }
            }

            blips.forEach(rId => {
              const relInfo = relsMap[rId];
              if (!relInfo) {
                const exists = allMacImageErrors.some(e => 
                  e.fileName === file.name && 
                  e.slideNum === slideNum && 
                  e.rId === rId
                );
                if (!exists) {
                  allMacImageErrors.push({
                    fileName: file.name,
                    slideNum,
                    errorType: '관계 정의 누락 (Relationship Missing)',
                    rId,
                    targetPath: '없음',
                    desc: `슬라이드 내 이미지가 관계 ID "${rId}"를 참조하고 있으나, 관계 정의(.rels) 파일에 해당 ID가 누락되어 이미지가 표시되지 않습니다.`
                  });
                  fileMacImageCount++;
                }
              } else {
                const target = relInfo.target;
                let cleanTarget = target;
                if (target.startsWith('../')) {
                  cleanTarget = 'ppt/' + target.replace('../', '');
                } else if (target.startsWith('media/')) {
                  cleanTarget = 'ppt/' + target;
                } else if (!target.startsWith('ppt/')) {
                  cleanTarget = 'ppt/media/' + target;
                }

                const cleanTargetLower = cleanTarget.toLowerCase();
                const fileExists = mediaFilesLower.includes(cleanTargetLower);
                const isExternal = target.includes('://') || target.startsWith('/') || /^[a-zA-Z]:/.test(target);

                if (isExternal) {
                  const exists = allMacImageErrors.some(e => 
                    e.fileName === file.name && 
                    e.slideNum === slideNum && 
                    e.rId === rId
                  );
                  if (!exists) {
                    allMacImageErrors.push({
                      fileName: file.name,
                      slideNum,
                      errorType: '외부 절대 경로 참조 (External Link)',
                      rId,
                      targetPath: target,
                      desc: `이미지가 외부 절대 경로 "${target}"로 연결되어 있어 다른 기기에서 열었을 때 깨질 수 있습니다.`
                    });
                    fileMacImageCount++;
                  }
                } else if (!fileExists) {
                  const exists = allMacImageErrors.some(e => 
                    e.fileName === file.name && 
                    e.slideNum === slideNum && 
                    e.rId === rId
                  );
                  if (!exists) {
                    allMacImageErrors.push({
                      fileName: file.name,
                      slideNum,
                      errorType: '이미지 파일 누락 (Media File Missing)',
                      rId,
                      targetPath: cleanTarget,
                      desc: `관계 정의 상 이미지 경로는 "${cleanTarget}"이나, PPTX 파일 내부에 리소스 파일이 실제로 누락되어 있습니다.`
                    });
                    fileMacImageCount++;
                  }
                }
              }
            });
          }

          if (checkNumbering) {
            // 헤더 수집: y좌표가 있으면 상단 영역(y<1500000), y=null인 레이아웃 상속 도형은 넘버링 패턴으로 판단
            const num_title_regex = /^\s*([0-9]+(\.[0-9]+)*)[\s\.]|^\s*(I{1,3}|IV|V|VI{0,3}|IX|X)\b/i;
            const headerShapes = shapes.filter(s => {
              if (s.y !== null) return s.y < 1500000;
              return num_title_regex.test(s.text);
            });
            
            let leftTitles = [];
            let rightTitle = '';
            
            headerShapes.forEach(hs => {
              // x=null(레이아웃 상속)이거나 x<6000000이면 좌측 타이틀
              if (hs.x === null || hs.x < 6000000) {
                const is_dup = leftTitles.some(t => t.text.trim() === hs.text.trim());
                if (!is_dup) {
                  leftTitles.push(hs);
                }
              } else if (hs.x !== null && hs.x >= 6000000) {
                if (!rightTitle || hs.y < (rightTitle.y || 99999999)) {
                  rightTitle = hs;
                }
              }
            });

            leftTitles.sort((a, b) => a.y - b.y);

            let majorTitleText = '';
            const roman_regex = /^\s*(I|II|III|IV|V|VI|VII|VIII|IX|X)\b/i;

            leftTitles.forEach(t => {
              const txt = t.text.trim();
              if (roman_regex.test(txt)) {
                majorTitleText = txt;
              }
            });

            slideList.push({
              slideNum,
              majorTitle: majorTitleText,
              titles: leftTitles.map(t => t.text.trim()).filter(Boolean),
              rightTitle: rightTitle ? rightTitle.text : ''
            });

            if (rightTitle && rightTitle.text) {
              const pageMatch = rightTitle.text.match(/(?:Page|P\.|-)?\s*(\d+)\s*(?:-)?$/i);
              if (pageMatch) {
                detectedPages.push(parseInt(pageMatch[1], 10));
              }
            }
          }
        }

        let startPage = 1;
        try {
          const presXmlStr = await zip.file('ppt/presentation.xml').async('text');
          const presParser = new DOMParser();
          const presDoc = presParser.parseFromString(presXmlStr, 'application/xml');
          let presEl = presDoc.getElementsByTagName('p:presentation')[0] || presDoc.getElementsByTagName('presentation')[0];
          if (!presEl) {
            const allEls = presDoc.getElementsByTagName('*');
            for (let i = 0; i < allEls.length; i++) {
              const el = allEls[i];
              const localName = el.localName || el.tagName.split(':').pop();
              if (localName === 'presentation') {
                presEl = el;
                break;
              }
            }
          }
          if (presEl) {
            const attrVal = presEl.getAttribute('firstSlideNum');
            if (attrVal !== null && attrVal !== undefined) {
              startPage = parseInt(attrVal, 10);
            }
          }
        } catch (presErr) {
          console.warn('ppt/presentation.xml 파싱 실패. 기본값 1 적용:', presErr);
        }
        const endPage = startPage + slideFiles.length - 1;
        allPageRanges.push({
          fileName: file.name,
          startPage,
          endPage,
          totalSlides: slideFiles.length
        });

        // 5. 상단 넘버링 규칙 검증 분석
        // 5-1. 좌측 넘버링 시퀀스 검증
        if (checkNumbering) {
          // 대주제 추적 및 검증용 상태 변수
          let current_major_roman = '';
          let current_major_val = null;
          let last_major_val = null;
          let last_major_roman = '';

          // 넘버링별 타이틀 불일치 검증을 위한 해시 맵
          const numbering_title_map = {};

          // 넘버링 순차성 검증을 위한 부모 경로별 마지막 자식 값 추적 맵
          const sibling_tracker = {};

          // 연속 페이지((순번/전체수)) 트래커 맵 (슬라이드 내 다른 헤더로 인한 상태 덮어쓰기 간섭 방지용)
          const alt_page_tracker = {};

          const roman_regex = /^\s*(I|II|III|IV|V|VI|VII|VIII|IX|X)\b/i;
          const num_regex = /^\s*([0-9]+(\.[0-9]+)*)[\s\.]/;

          for (let i = 0; i < slideList.length; i++) {
            const slide = slideList[i];
            
            // 1) 대주제(로마자) 정보 업데이트 및 대주제 자체의 순차성 검증
            if (slide.majorTitle) {
              const major_match = slide.majorTitle.match(roman_regex);
              if (major_match) {
                current_major_roman = major_match[1];
                current_major_val = roman_to_int(current_major_roman);

                if (last_major_val !== null && current_major_val !== last_major_val) {
                  if (current_major_val !== last_major_val + 1) {
                    allNumberings.push({
                      fileName: file.name,
                      slideNum: slide.slideNum,
                      area: '좌측 타이틀',
                      text: slide.majorTitle,
                      error: '대주제 로마자 순차성 단절',
                      guide: `대주제 장 번호가 순차적으로 증가하지 않고 단절되었습니다. (이전: ${last_major_roman} ➜ 현재: ${current_major_roman})`
                    });
                    fileNumErrorsCount++;
                  }
                }
                last_major_val = current_major_val;
                last_major_roman = current_major_roman;
              }
            }

            // 1-2) 2단/3단 타이틀 간의 구조적 계층 관계 정합성 검증 (원칙 2, 3, 4)
            const subTitles = (slide.titles || []).filter(text => !roman_regex.test(text));
            if (subTitles.length === 2) {
              const level2Title = subTitles[0];
              const level3Title = subTitles[1];
              
              const level2Parts = level2Title.split('>').map(s => s.trim());
              const level3Parts = level3Title.split('>').map(s => s.trim());
              
              const getNumDepth = (txt) => {
                const match = txt.match(num_regex);
                if (!match) return 0;
                return match[1].split('.').length;
              };
              
              const l2Depths = level2Parts.map(getNumDepth);
              const l3Depths = level3Parts.map(getNumDepth);
              const l3MainDepth = l3Depths[0];
              
              if (l3MainDepth === 2) {
                // 원칙 2: 중분류(1.1)가 3단 타이틀로 오는 케이스 -> 2단은 대분류(1.) 단독 구조
                const isL2Valid = level2Parts.length === 1 && l2Depths[0] === 1;
                if (!isL2Valid) {
                  allNumberings.push({
                    fileName: file.name,
                    slideNum: slide.slideNum,
                    area: '좌측 타이틀',
                    text: `${level2Title} / ${level3Title}`,
                    error: '타이틀 계층 규칙 위반 (원칙 2)',
                    guide: `중분류(${level3Parts[0]})가 3단 타이틀로 올 경우, 2단 타이틀은 '대분류' 단독 구조(예: "1. 제목")여야 합니다. (현재 2단: "${level2Title}")`
                  });
                  fileNumErrorsCount++;
                }
              } else if (l3MainDepth === 3) {
                if (level3Parts.length === 1) {
                  // 원칙 3: 소분류(1.1.1) 단독이 3단 타이틀로 오는 케이스 -> 2단은 대분류 > 중분류 구조
                  const isL2Valid = level2Parts.length === 2 && l2Depths[0] === 1 && l2Depths[1] === 2;
                  if (!isL2Valid) {
                    allNumberings.push({
                      fileName: file.name,
                      slideNum: slide.slideNum,
                      area: '좌측 타이틀',
                      text: `${level2Title} / ${level3Title}`,
                      error: '타이틀 계층 규칙 위반 (원칙 3)',
                      guide: `소분류(${level3Parts[0]})가 3단 타이틀로 올 경우, 2단 타이틀은 '대분류 > 중분류' 구조(예: "1. 대분류 > 1.1 중분류")여야 합니다. (현재 2단: "${level2Title}")`
                    });
                    fileNumErrorsCount++;
                  }
                } else if (level3Parts.length > 1) {
                  // 원칙 4: 소분류 하위가 3단 타이틀로 오는 케이스 -> 2단은 대분류 > 중분류 구조
                  const isL2Valid = level2Parts.length === 2 && l2Depths[0] === 1 && l2Depths[1] === 2;
                  if (!isL2Valid) {
                    allNumberings.push({
                      fileName: file.name,
                      slideNum: slide.slideNum,
                      area: '좌측 타이틀',
                      text: `${level2Title} / ${level3Title}`,
                      error: '타이틀 계층 규칙 위반 (원칙 4)',
                      guide: `소분류 하위가 3단 타이틀로 올 경우, 2단 타이틀은 '대분류 > 중분류' 구조(예: "1. 대분류 > 1.1 중분류")여야 합니다. (현재 2단: "${level2Title}")`
                    });
                    fileNumErrorsCount++;
                  }
                }
              }
            }

            // 슬라이드 내 모든 수집된 타이틀 순차 검증 (1단, 2단, 3단 통합 순회)
            const slide_titles = slide.titles || [];

            for (let tIdx = 0; tIdx < slide_titles.length; tIdx++) {
              const text = slide_titles[tIdx];
              
              // 대주제 텍스트 자체는 아래의 세부 넘버링(숫자) 검증에서는 제외
              if (roman_regex.test(text)) continue;

              // "목차 명칭 (1/4)" 정규식 파싱
              const altPageMatch = text.match(/\((\d+)\/(\d+)\)\s*$/);
              let hasAltPages = false;
              let currentAltPage = null;
              let totalAltPage = null;
              let pureTitle = text.trim();

              if (altPageMatch) {
                hasAltPages = true;
                currentAltPage = parseInt(altPageMatch[1], 10);
                totalAltPage = parseInt(altPageMatch[2], 10);
                pureTitle = text.replace(/\s*\(\d+\/\d+\)\s*$/, '').trim();
              }

              // 순수 제목 추출: 넘버링 숫자 접두사(1., 1.1 등) 제거 후 실제 제목 텍스트만 추출
              // 예: "1.1 재무구조 및 대외신용평가 현황" → "재무구조 및 대외신용평가 현황"
              const pureTitleWithoutNum = pureTitle.replace(/^\s*[0-9]+(\.[0-9]+)*[\s\.]+/, '').trim();

              // 넘버링 형식 추출 정규식: "1. ", "1.1 ", "1.1.1 " 등
              const numMatch = text.match(num_regex);
              
              if (numMatch) {
                const rawNumStr = numMatch[1];
                const parts = rawNumStr.split('.').map(n => parseInt(n, 10));
                
                if (parts.length > 0) {
                  const current_val = parts[parts.length - 1];

                  // 1. 동일 넘버링이 다른 슬라이드에서 쓰였는데, 제목 텍스트가 다르면 오류 검출
                  // pureTitleWithoutNum: 넘버링 숫자를 제거한 순수 제목만으로 비교
                  const existingTitle = numbering_title_map[rawNumStr];
                  if (existingTitle !== undefined) {
                    if (existingTitle !== pureTitleWithoutNum) {
                      allNumberings.push({
                        fileName: file.name,
                        slideNum: slide.slideNum,
                        area: '좌측 타이틀',
                        text,
                        error: `동일 넘버링 내 타이틀 텍스트 불일치 (${rawNumStr})`,
                        guide: `넘버링 "${rawNumStr}"에 대해 이전 슬라이드("${rawNumStr} ${existingTitle}")와 현재 슬라이드("${pureTitle}")의 타이틀 텍스트가 서로 일치하지 않습니다.`
                      });
                      fileNumErrorsCount++;
                    }
                  } else {
                    numbering_title_map[rawNumStr] = pureTitleWithoutNum;
                  }

                  // 2. 형제 노드 간의 순차적(sequential) 검증 진행
                  const parent_path = parts.slice(0, -1).join('.');
                  const last_sibling_val = sibling_tracker[parent_path];

                  if (last_sibling_val !== undefined) {
                    if (current_val === last_sibling_val) {
                      // 동일 번호 반복: 1단 넘버링(중분류 등, 예: "1", "2")이 아니고, 연속 페이지(1/2 등) 표기가 없으면 중복 오류
                      if (!hasAltPages && parts.length > 1) {
                        allNumberings.push({
                          fileName: file.name,
                          slideNum: slide.slideNum,
                          area: '좌측 타이틀',
                          text,
                          error: `넘버링 중복 사용 오류 (${rawNumStr})`,
                          guide: `이전 슬라이드에서 이미 "${rawNumStr}" 번호가 사용되었습니다. 넘버링이 동일한데 제목이 다르다면 오류입니다. 순차적으로 다음 번호로 증가시켜야 합니다.`
                        });
                        fileNumErrorsCount++;
                      }
                    } else if (current_val !== last_sibling_val + 1) {
                      // 순서 건너뜀 (시퀀스 단절)
                      allNumberings.push({
                        fileName: file.name,
                        slideNum: slide.slideNum,
                        area: '좌측 타이틀',
                        text,
                        error: `넘버링 시퀀스 단절 (${rawNumStr})`,
                        guide: `동일 계층 수준 내에서 일련번호는 순차적으로 1씩 증가해야 합니다. (이전 형제 값: ${last_sibling_val} ➜ 현재 표기: ${current_val})`
                      });
                      fileNumErrorsCount++;
                    }
                  } else {
                    // 부모 계층 아래 최초의 자식 노드 진입 시에는 반드시 1이어야 함 (로마자 대주제 제외)
                    const isRoman = roman_regex.test(text);
                    if (!isRoman && current_val !== 1) {
                      allNumberings.push({
                        fileName: file.name,
                        slideNum: slide.slideNum,
                        area: '좌측 타이틀',
                        text,
                        error: `하위 넘버링 시작값 오류 (${rawNumStr})`,
                        guide: `새로운 하위 계층으로 진입할 때 일련번호는 항상 1부터 순차 시작해야 합니다. (${parent_path ? parent_path + '.1' : '1.'} 권장)`
                      });
                      fileNumErrorsCount++;
                    }
                  }
                  // 트래커 기록 갱신 (연속 페이지인 경우에도 번호를 유지하여 다음 연속 장에서 비교 가능하도록)
                  if (!hasAltPages || current_val !== last_sibling_val) {
                    sibling_tracker[parent_path] = current_val;
                  }
                }
              }

              // 3. 연속 페이지 (1/4) 형식의 연속성 세부 검사 (alt_page_tracker 기반 독립 검사)
              if (hasAltPages) {
                const prevPageInfo = alt_page_tracker[pureTitle];
                if (prevPageInfo && prevPageInfo.slideNum === slide.slideNum - 1) {
                  if (totalAltPage !== prevPageInfo.totalAltPage) {
                    allNumberings.push({
                      fileName: file.name,
                      slideNum: slide.slideNum,
                      area: '좌측 타이틀',
                      text,
                      error: '목차 전체 페이지 수 불일치',
                      guide: `연속 목차의 전체 페이지 수(분모)가 이전 슬라이드(${prevPageInfo.totalAltPage}장)와 현재 슬라이드(${totalAltPage}장)가 서로 다릅니다.`
                    });
                    fileNumErrorsCount++;
                  }
                  
                  if (currentAltPage !== prevPageInfo.currentAltPage + 1) {
                    allNumberings.push({
                      fileName: file.name,
                      slideNum: slide.slideNum,
                      area: '좌측 타이틀',
                      text,
                      error: '목차 연속 페이지 번호 단절',
                      guide: `연속된 목차 페이지 번호가 순차적으로 증가하지 않았습니다. (이전: ${prevPageInfo.currentAltPage}/${prevPageInfo.totalAltPage} ➜ 현재: ${currentAltPage}/${totalAltPage})`
                    });
                    fileNumErrorsCount++;
                  }

                  if (currentAltPage > totalAltPage) {
                    allNumberings.push({
                      fileName: file.name,
                      slideNum: slide.slideNum,
                      area: '좌측 타이틀',
                      text,
                      error: '목차 페이지 범위 초과',
                      guide: `목차 페이지 번호(${currentAltPage})가 전체 페이지 수(${totalAltPage})를 초과하였습니다.`
                    });
                    fileNumErrorsCount++;
                  }
                } else {
                  // 직전 슬라이드에 연속되는 목차가 없었던 경우 시작 번호는 무조건 1이어야 함
                  if (currentAltPage !== 1) {
                    allNumberings.push({
                      fileName: file.name,
                      slideNum: slide.slideNum,
                      area: '좌측 타이틀',
                      text,
                      error: '목차 페이지 시작 번호 오류',
                      guide: `연속 목차가 시작될 때는 1페이지부터 시작해야 합니다. (현재: ${currentAltPage}/${totalAltPage})`
                    });
                    fileNumErrorsCount++;
                  }
                }

                // 현재 슬라이드의 연속 목차 상태 등록
                alt_page_tracker[pureTitle] = {
                  slideNum: slide.slideNum,
                  currentAltPage,
                  totalAltPage
                };
              }
            }
          }

          // 5-2. 우측 넘버링(페이지 번호 및 서브텍스트) 일관성 및 시퀀스 검증
          let lastPageNum = null;
          for (let i = 0; i < slideList.length; i++) {
            const slide = slideList[i];
            const text = slide.rightTitle;
            if (!text) continue;

            // 페이지 번호 형태 검출 (예: "01", "Page 1", "P. 3", "- 4 -")
            const pageMatch = text.match(/(?:Page|P\.|-)?\s*(\d+)\s*(?:-)?$/i);
            if (pageMatch) {
              const pageNum = parseInt(pageMatch[1], 10);
              if (lastPageNum !== null) {
                if (pageNum !== lastPageNum + 1) {
                  allNumberings.push({
                    fileName: file.name,
                    slideNum: slide.slideNum,
                    area: '우측 페이지',
                    text,
                    error: `우측 페이지 번호 시퀀스 단절 (이전 페이지: ${lastPageNum} ➜ 현재 표기: ${pageNum})`,
                    guide: `페이지 번호가 순차적으로 증가하지 않았습니다. 페이지수 흐름을 점검해 주세요.`
                  });
                  fileNumErrorsCount++;
                }
              }
              lastPageNum = pageNum;
            }
          }
        }
        }

        if (check_ai_typos && apiKey) {
          try {
            const aiTypos = await call_gemini_ai_typos(file.name, slides_text_map, apiKey);
            aiTypos.forEach(t => {
              allTypos.push(t);
              fileTyposCount++;
            });
          } catch (err) {
            console.error("AI 오탈자 검증 중 오류 발생:", err);
          }
        }

        if (check_external_api && external_api_url) {
          try {
            const ext_typos = await call_external_spell_check(file.name, slides_text_map, external_api_url);
            ext_typos.forEach(t => {
              allTypos.push(t);
              fileTyposCount++;
            });
          } catch (err) {
            console.error("외부 맞춤법 API 검증 중 오류 발생:", err);
          }
        }

        stats.push({
          name: file.name,
          typos: fileTyposCount,
          numberingErrors: fileNumErrorsCount,
          altTextErrors: fileAltErrorsCount,
          forbiddenErrors: fileForbiddenCount,
          engKoMixedErrors: fileEngKoMixedCount,
          duplicateErrors: file_duplicate_count,
          macImageErrors: fileMacImageCount,
          startPage: allPageRanges[allPageRanges.length - 1]?.startPage ?? 1,
          endPage: allPageRanges[allPageRanges.length - 1]?.endPage ?? 1,
          totalSlides: allPageRanges[allPageRanges.length - 1]?.totalSlides ?? 1
        });
      }

      // 각 파일별 시작 페이지 맵 구성
      const startPageMap = {};
      stats.forEach(s => {
        startPageMap[s.name] = s.startPage;
      });

      // 모든 검출 결과 객체들에 displayPageNum 추가
      const addDisplayPageNum = (list) => {
        return list.map(item => {
          const start = startPageMap[item.fileName] || 1;
          const displayPageNum = start + item.slideNum - 1;
          return { ...item, displayPageNum };
        });
      };

      const finalTypos = addDisplayPageNum(allTypos);
      const finalNumberings = addDisplayPageNum(allNumberings);
      const finalAltTexts = addDisplayPageNum(allAltTexts);
      const finalForbiddens = addDisplayPageNum(allForbiddens);
      const finalEngKoMixed = addDisplayPageNum(allEngKoMixed);
      const finalDuplicates = addDisplayPageNum(all_duplicates);
      const finalMacImages = addDisplayPageNum(allMacImageErrors);

      setTypoResults(finalTypos);
      setNumberingResults(finalNumberings);
      setAltTextResults(finalAltTexts);
      setForbiddenResults(finalForbiddens);
      setEngKoMixedResults(finalEngKoMixed);
      set_duplicate_results(finalDuplicates);
      setPageRangeResults(allPageRanges);
      setMacImageResults(finalMacImages);
      setFileStats(stats);
      setIsValidated(true);
      setActiveResultTab('summary');
    } catch (err) {
      console.error('검증 중 오류 발생:', err);
      alert(`검증 오류 발생: ${err.message}`);
    } finally {
      setIsProcessing(false);
    }
  };

  // 점검 결과를 엑셀 파일로 추출
  const handleExportToExcel = () => {
    if (
      typoResults.length === 0 && 
      numberingResults.length === 0 && 
      altTextResults.length === 0 && 
      forbiddenResults.length === 0 &&
      engKoMixedResults.length === 0 &&
      duplicate_results.length === 0 &&
      pageRangeResults.length === 0 &&
      macImageResults.length === 0
    ) {
      alert('출력할 검증 결과 데이터가 존재하지 않습니다.');
      return;
    }

    const workbook = XLSX.utils.book_new();

    const startPageMap = {};
    fileStats.forEach(s => {
      startPageMap[s.name] = s.startPage;
    });

    const getSafeDisplayPage = (item) => {
      if (item.displayPageNum !== undefined && item.displayPageNum !== null && !isNaN(item.displayPageNum)) {
        return item.displayPageNum;
      }
      const start = startPageMap[item.fileName] || 1;
      return start + item.slideNum - 1;
    };

    // 1. 오탈자 시트 데이터 구성
    if (checkTypos || check_ai_typos) {
      const typoRows = typoResults.map((t, idx) => ({
        '순번': idx + 1,
        '대상 파일명': t.fileName,
        '페이지수': `${t.slideNum} 페이지`,
        '표시 페이지수': `${getSafeDisplayPage(t)} 페이지`,
        '의심 단어': t.typo,
        '추천 교정안': t.correction,
        '검증 구분': t.type,
        '교정 가이드': t.desc,
        '검출 문장(전체)': t.sentence
      }));
      const typoSheet = XLSX.utils.json_to_sheet(typoRows);
      XLSX.utils.book_append_sheet(workbook, typoSheet, '오탈자_점검결과');
    }

    // 2. 넘버링 시트 데이터 구성
    if (checkNumbering) {
      const numberingRows = numberingResults.map((n, idx) => ({
        '순번': idx + 1,
        '대상 파일명': n.fileName,
        '페이지수': `${n.slideNum} 페이지`,
        '표시 페이지수': `${getSafeDisplayPage(n)} 페이지`,
        '검증 영역': n.area,
        '표기 텍스트': n.text,
        '검출 오류': n.error,
        '올바른 규칙 가이드': n.guide
      }));
      const numberingSheet = XLSX.utils.json_to_sheet(numberingRows);
      XLSX.utils.book_append_sheet(workbook, numberingSheet, '넘버링_점검결과');
    }

    // 3. 대체텍스트 시트 데이터 구성
    if (checkAltText) {
      const altRows = altTextResults.map((a, idx) => ({
        '순번': idx + 1,
        '대상 파일명': a.fileName,
        '페이지수': `${a.slideNum} 페이지`,
        '표시 페이지수': `${getSafeDisplayPage(a)} 페이지`,
        '대상 개체명': a.objName,
        '대체텍스트 내용': a.descr,
        '참고': a.guide
      }));
      const altSheet = XLSX.utils.json_to_sheet(altRows);
      XLSX.utils.book_append_sheet(workbook, altSheet, '대체텍스트_검색결과');
    }

    // 4. 지정 단어 시트 데이터 구성
    if (checkForbiddenWords) {
      const forbiddenRows = forbiddenResults.map((f, idx) => ({
        '순번': idx + 1,
        '대상 파일명': f.fileName,
        '페이지수': `${f.slideNum} 페이지`,
        '표시 페이지수': `${getSafeDisplayPage(f)} 페이지`,
        '검출 지정 단어': f.word,
        '검출 오류': f.error,
        '올바른 규칙 가이드': f.guide,
        '검출 문장(전체)': f.sentence
      }));
      const forbiddenSheet = XLSX.utils.json_to_sheet(forbiddenRows);
      XLSX.utils.book_append_sheet(workbook, forbiddenSheet, '지정단어_점검결과');
    }

    // 5. 영한 혼용 단어 시트 데이터 구성
    if (checkEngKoMixed) {
      const engKoMixedRows = engKoMixedResults.map((e, idx) => ({
        '순번': idx + 1,
        '대상 파일명': e.fileName,
        '페이지수': `${e.slideNum} 페이지`,
        '표시 페이지수': `${getSafeDisplayPage(e)} 페이지`,
        '검출 단어': e.word,
        '검출 오류': e.error,
        '올바른 규칙 가이드': e.guide,
        '검출 문장(전체)': e.sentence
      }));
      const engKoMixedSheet = XLSX.utils.json_to_sheet(engKoMixedRows);
      XLSX.utils.book_append_sheet(workbook, engKoMixedSheet, '영한혼용_점검결과');
    }

    // 6. 동일 단어 중복 시트 데이터 구성
    if (check_duplicate_words) {
      const duplicateRows = duplicate_results.map((d, idx) => ({
        '순번': idx + 1,
        '대상 파일명': d.fileName,
        '페이지수': `${d.slideNum} 페이지`,
        '표시 페이지수': `${getSafeDisplayPage(d)} 페이지`,
        '중복 표현': d.word,
        '검출 단어': d.duplicateWord,
        '검출 오류': d.error,
        '올바른 규칙 가이드': d.guide,
        '검출 문장(전체)': d.sentence
      }));
      const duplicateSheet = XLSX.utils.json_to_sheet(duplicateRows);
      XLSX.utils.book_append_sheet(workbook, duplicateSheet, '중복단어_점검결과');
    }

    // 7. 페이지 범위 및 수량 분석 시트 데이터 구성
    if (checkPageRange) {
      const pageRangeRows = pageRangeResults.map((p, idx) => ({
        '순번': idx + 1,
        '파일명': p.fileName,
        '시작페이지': p.startPage,
        '최종 페이지': p.endPage,
        '총 페이지수': p.totalSlides
      }));
      const pageRangeSheet = XLSX.utils.json_to_sheet(pageRangeRows);
      XLSX.utils.book_append_sheet(workbook, pageRangeSheet, '페이지범위_분석결과');
    }

    // 8. 맥 이미지 누락 시트 데이터 구성
    if (checkMacImages) {
      const macImageRows = macImageResults.map((m, idx) => ({
        '순번': idx + 1,
        '대상 파일명': m.fileName,
        '페이지수': `${m.slideNum} 페이지`,
        '표시 페이지수': `${getSafeDisplayPage(m)} 페이지`,
        '오류 유형': m.errorType,
        '관계 ID': m.rId,
        '대상 경로': m.targetPath,
        '상세 설명': m.desc
      }));
      const macImageSheet = XLSX.utils.json_to_sheet(macImageRows);
      XLSX.utils.book_append_sheet(workbook, macImageSheet, '맥이미지_누락결과');
    }

    // 엑셀 파일 다운로드 실행
    const dateStr = new Date().toISOString().split('T')[0].replace(/-/g, '');
    XLSX.writeFile(workbook, `PPT_산출물_검증결과_${dateStr}.xlsx`);
  };

  return (
    <div style={{ padding: '24px', display: 'flex', flexDirection: 'column', gap: '24px', minHeight: 'calc(100vh - 120px)', color: 'var(--text-primary)' }}>
      
      {/* 타이틀 및 가이드 배너 */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: 'linear-gradient(135deg, rgba(225, 29, 72, 0.15), rgba(168, 85, 247, 0.05))', border: '1px solid rgba(225, 29, 72, 0.25)', padding: '24px', borderRadius: '16px', backdropFilter: 'blur(10px)' }}>
        <div>
          <h2 style={{ display: 'flex', alignItems: 'center', gap: '10px', fontSize: '22px', fontWeight: 800, margin: 0, color: 'var(--text-primary)' }}>
            <ShieldAlert size={28} color="#e11d48" /> PPT / HWPX 검증(표준산출물)
          </h2>
          <p style={{ margin: '8px 0 0 0', fontSize: '14px', color: 'var(--text-secondary)', lineHeight: '1.6' }}>
            PPTX 제안서 또는 HWPX 한글 보고서를 업로드하면 텍스트 내의 **오탈자(비즈니스/외래어/맞춤법)**와 **중복 단어, 영한 혼용 단어** 등을 정교하게 분석합니다.<br />
            (PPTX의 경우 상단 헤더 넘버링 규칙성 및 대체텍스트도 점검하며, HWPX는 본문 텍스트 기반 검증이 수행됩니다.)
          </p>
        </div>
        <div style={{ display: 'flex', gap: '12px' }}>
          <button 
            onClick={clearAllFiles}
            disabled={pptFiles.length === 0}
            style={{ padding: '10px 16px', background: 'rgba(255,255,255,0.05)', color: pptFiles.length > 0 ? 'var(--text-primary)' : 'var(--text-muted)', border: '1px solid var(--panel-border)', borderRadius: '8px', cursor: pptFiles.length > 0 ? 'pointer' : 'not-allowed', display: 'flex', alignItems: 'center', gap: '8px', fontSize: '13px', fontWeight: 600 }}
          >
            <RefreshCw size={15} /> 초기화
          </button>
        </div>
      </div>

      {/* 메인 콘텐츠 영역 (2열 레이아웃) */}
      <div style={{ display: 'grid', gridTemplateColumns: '1.2fr 1fr', gap: '24px' }}>
        
        {/* 좌측: 파일 드롭존 및 업로드 파일 목록 */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
          
          {/* 드롭존 카드 */}
          <div 
            onDragOver={handleDragOver}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current?.click()}
            style={{ 
              background: 'var(--panel-bg)', 
              border: '2px dashed var(--panel-border)', 
              borderRadius: '16px', 
              padding: '40px 20px', 
              textAlign: 'center', 
              cursor: 'pointer',
              transition: 'all 0.3s ease',
              display: 'flex',
              flexDirection: 'column',
              alignItems: 'center',
              justifyContent: 'center',
              gap: '14px'
            }}
            className="interactive-card"
          >
            <div style={{ width: '64px', height: '64px', borderRadius: '50%', background: 'rgba(225, 29, 72, 0.08)', display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#e11d48' }}>
              <FileUp size={32} />
            </div>
            <div>
              <p style={{ margin: 0, fontSize: '16px', fontWeight: 700, color: 'var(--text-primary)' }}>검증할 PPTX 또는 HWPX 파일을 끌어다 놓으세요</p>
              <p style={{ margin: '6px 0 0 0', fontSize: '13px', color: 'var(--text-muted)' }}>또는 컴퓨터에서 파일 찾아보기 (다중 선택 가능)</p>
            </div>
            <input 
              ref={fileInputRef}
              type="file" 
              accept=".pptx, .hwpx" 
              multiple 
              onChange={handleFileChange}
              style={{ display: 'none' }}
            />
          </div>

          {/* 업로드 파일 리스트 */}
          <div style={{ background: 'var(--panel-bg)', border: '1px solid var(--panel-border)', borderRadius: '16px', padding: '20px', display: 'flex', flexDirection: 'column', gap: '14px' }}>
            <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '8px' }}>
              대기 중인 파일 <span style={{ background: 'rgba(225, 29, 72, 0.1)', color: '#e11d48', padding: '2px 8px', borderRadius: '10px', fontSize: '12px', fontWeight: 800 }}>{pptFiles.length}</span>
            </h3>
            
            {pptFiles.length === 0 ? (
              <div style={{ padding: '30px 0', textAlign: 'center', color: 'var(--text-muted)', fontSize: '13px' }}>
                검증 대기 중인 파일이 없습니다.
              </div>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', maxHeight: '220px', overflowY: 'auto', paddingRight: '4px' }}>
                {pptFiles.map((file, idx) => (
                  <div key={idx} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: 'rgba(255,255,255,0.02)', border: '1px solid var(--panel-border)', padding: '10px 14px', borderRadius: '10px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '10px', overflow: 'hidden' }}>
                      <span style={{ fontSize: '12px', background: 'rgba(255,255,255,0.08)', color: 'var(--text-secondary)', width: '22px', height: '22px', borderRadius: '50%', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                        {idx + 1}
                      </span>
                      <span style={{ fontSize: '13.5px', fontWeight: 600, color: 'var(--text-primary)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                        {file.name}
                      </span>
                      <span style={{ fontSize: '11px', color: 'var(--text-muted)', flexShrink: 0 }}>
                        ({(file.size / 1024 / 1024).toFixed(2)} MB)
                      </span>
                    </div>
                    <button 
                      onClick={() => removeFile(idx)}
                      style={{ background: 'transparent', border: 'none', color: '#ef4444', cursor: 'pointer', fontSize: '12px', fontWeight: 600 }}
                    >
                      삭제
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* 검증 옵션 설정 */}
          <div style={{ 
            background: 'var(--panel-bg)', 
            border: '1px solid var(--panel-border)', 
            borderRadius: '16px', 
            padding: '20px', 
            display: 'flex', 
            flexDirection: 'column', 
            gap: '14px' 
          }}>
            <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '8px', color: 'var(--text-primary)' }}>
              <ShieldAlert size={17} color="#e11d48" /> 검증 옵션 설정
            </h3>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              {/* 1. 오탈자 검증 */}
              <div 
                onClick={() => setCheckTypos(prev => !prev)}
                style={{ 
                  background: checkTypos ? 'rgba(225, 29, 72, 0.05)' : 'rgba(255, 255, 255, 0.01)', 
                  border: checkTypos ? '1px solid rgba(225, 29, 72, 0.4)' : '1px solid var(--panel-border)',
                  borderRadius: '10px', 
                  padding: '12px 16px', 
                  cursor: 'pointer', 
                  display: 'flex', 
                  alignItems: 'center', 
                  gap: '12px',
                  transition: 'all 0.2s ease'
                }}
              >
                <input 
                  type="checkbox" 
                  checked={checkTypos} 
                  onChange={(e) => {
                    e.stopPropagation();
                    setCheckTypos(e.target.checked);
                  }}
                  style={{ width: '16px', height: '16px', cursor: 'pointer', accentColor: '#e11d48' }}
                />
                <div style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                  <span style={{ fontSize: '13.5px', fontWeight: 700, color: 'var(--text-primary)' }}>오탈자 검증</span>
                  <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>맞춤법 및 외래어 오류 검출</span>
                </div>
              </div>

              {/* 1-2. AI 오탈자 정밀 검증 */}
              <div 
                onClick={() => set_check_ai_typos(prev => !prev)}
                style={{ 
                  background: check_ai_typos ? 'rgba(99, 102, 241, 0.05)' : 'rgba(255, 255, 255, 0.01)', 
                  border: check_ai_typos ? '1px solid rgba(99, 102, 241, 0.4)' : '1px solid var(--panel-border)',
                  borderRadius: '10px', 
                  padding: '12px 16px', 
                  cursor: 'pointer', 
                  display: 'flex', 
                  alignItems: 'center', 
                  gap: '12px',
                  transition: 'all 0.2s ease'
                }}
              >
                <input 
                  type="checkbox" 
                  checked={check_ai_typos} 
                  onChange={(e) => {
                    e.stopPropagation();
                    set_check_ai_typos(e.target.checked);
                  }}
                  style={{ width: '16px', height: '16px', cursor: 'pointer', accentColor: '#6366f1' }}
                />
                <div style={{ display: 'flex', flex: 1, flexDirection: 'column', gap: '2px' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                    <span style={{ fontSize: '13.5px', fontWeight: 700, color: 'var(--text-primary)' }}>AI 오탈자 정밀 검증</span>
                    <Sparkles size={14} style={{ color: '#6366f1' }} />
                  </div>
                  <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>Gemini AI를 이용한 실시간 문맥 맞춤법 및 문맥 오류 진단</span>
                </div>
              </div>

              {/* 외부 전문 맞춤법 검사 API 연동 옵션 */}
              <div
                onClick={() => {
                  const next = !check_external_api;
                  set_check_external_api(next);
                  // 체크박스 ON 시 로컬 서버 URL 자동 입력
                  if (next && !external_api_url) {
                    set_external_api_url('http://localhost:8000/check');
                    set_external_server_status('unknown');
                  }
                }}
                style={{
                  background: check_external_api ? 'rgba(234, 179, 8, 0.05)' : 'rgba(255, 255, 255, 0.01)',
                  border: check_external_api ? '1px solid rgba(234, 179, 8, 0.4)' : '1px solid var(--panel-border)',
                  borderRadius: '10px',
                  padding: '12px 16px',
                  cursor: 'pointer',
                  display: 'flex',
                  alignItems: 'flex-start',
                  gap: '12px',
                  transition: 'all 0.2s ease'
                }}
              >
                <input
                  type="checkbox"
                  checked={check_external_api}
                  onChange={(e) => {
                    e.stopPropagation();
                    const next = e.target.checked;
                    set_check_external_api(next);
                    if (next && !external_api_url) {
                      set_external_api_url('http://localhost:8000/check');
                      set_external_server_status('unknown');
                    }
                  }}
                  style={{ width: '16px', height: '16px', cursor: 'pointer', accentColor: '#eab308', marginTop: '2px' }}
                />
                <div style={{ display: 'flex', flex: 1, flexDirection: 'column', gap: '8px' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                    <span style={{ fontSize: '13.5px', fontWeight: 700, color: 'var(--text-primary)' }}>외부 전문 맞춤법 API 연동</span>
                    <span style={{ fontSize: '10px', background: 'rgba(234, 179, 8, 0.15)', color: '#eab308', padding: '1px 6px', borderRadius: '4px', fontWeight: 700 }}>옵션</span>
                  </div>
                  <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>로컬 맞춤법 서버 또는 사내 API와 실시간 연동 (체크 시 자동 입력)</span>

                  {check_external_api && (
                    <div onClick={(e) => e.stopPropagation()} style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>

                      {/* URL 입력창 */}
                      <div style={{ display: 'flex', gap: '6px', alignItems: 'center' }}>
                        <input
                          type="text"
                          placeholder="http://localhost:8000/check"
                          value={external_api_url}
                          onChange={(e) => { set_external_api_url(e.target.value); set_external_server_status('unknown'); }}
                          style={{
                            flex: 1,
                            padding: '7px 10px',
                            borderRadius: '6px',
                            border: '1px solid rgba(234, 179, 8, 0.5)',
                            background: 'rgba(234, 179, 8, 0.05)',
                            color: 'var(--text-primary)',
                            fontSize: '12px',
                            fontFamily: 'monospace',
                            outline: 'none',
                            boxSizing: 'border-box'
                          }}
                        />
                        {/* 서버 상태 확인 버튼 */}
                        <button
                          onClick={async () => {
                            set_external_server_status('checking');
                            try {
                              const health_url = external_api_url.replace('/check', '/health').replace('/check/', '/health');
                              const res = await fetch(health_url, { signal: AbortSignal.timeout(3000) });
                              set_external_server_status(res.ok ? 'running' : 'stopped');
                            } catch {
                              set_external_server_status('stopped');
                            }
                          }}
                          style={{
                            padding: '7px 12px',
                            borderRadius: '6px',
                            border: '1px solid rgba(234, 179, 8, 0.5)',
                            background: 'rgba(234, 179, 8, 0.1)',
                            color: '#eab308',
                            fontSize: '11px',
                            fontWeight: 700,
                            cursor: 'pointer',
                            whiteSpace: 'nowrap'
                          }}
                        >
                          {external_server_status === 'checking' ? '확인 중...' : '상태 확인'}
                        </button>
                      </div>

                      {/* 서버 상태 표시 */}
                      {external_server_status !== 'unknown' && (
                        <div style={{
                          display: 'flex', alignItems: 'flex-start', gap: '8px',
                          padding: '10px 12px', borderRadius: '8px',
                          background: external_server_status === 'running'
                            ? 'rgba(34, 197, 94, 0.08)'
                            : external_server_status === 'checking'
                            ? 'rgba(234, 179, 8, 0.08)'
                            : 'rgba(239, 68, 68, 0.08)',
                          border: `1px solid ${
                            external_server_status === 'running' ? 'rgba(34,197,94,0.3)'
                            : external_server_status === 'checking' ? 'rgba(234,179,8,0.3)'
                            : 'rgba(239,68,68,0.3)'
                          }`
                        }}>
                          <span style={{ fontSize: '16px', lineHeight: 1 }}>
                            {external_server_status === 'running' ? '🟢' : external_server_status === 'checking' ? '🟡' : '🔴'}
                          </span>
                          <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', flex: 1 }}>
                            <span style={{ fontSize: '12px', fontWeight: 700,
                              color: external_server_status === 'running' ? '#22c55e'
                                : external_server_status === 'checking' ? '#eab308' : '#ef4444'
                            }}>
                              {external_server_status === 'running' ? '서버 정상 실행 중' : external_server_status === 'checking' ? '서버 상태 확인 중...' : '서버가 실행되지 않음'}
                            </span>
                            {external_server_status === 'stopped' && (
                              <div style={{ fontSize: '11px', color: 'var(--text-muted)', lineHeight: 1.6 }}>
                                <div>📁 서버 실행 방법:</div>
                                <div style={{ marginTop: '4px', display: 'flex', flexDirection: 'column', gap: '3px' }}>
                                  <div>① 파일 탐색기에서 아래 경로를 열어주세요</div>
                                  <code style={{
                                    display: 'block', padding: '4px 8px',
                                    background: 'rgba(0,0,0,0.2)', borderRadius: '4px',
                                    fontSize: '10px', userSelect: 'all',
                                    color: '#fbbf24', wordBreak: 'break-all'
                                  }}>
                                    C:\Users\JEJU_MEC_3\Desktop\AI Coding\spell-check-server
                                  </code>
                                  <div>② <strong>서버실행.bat</strong> 파일을 더블클릭하세요</div>
                                  <div>③ 서버 시작 후 [상태 확인] 버튼을 다시 눌러주세요</div>
                                </div>
                              </div>
                            )}
                            {external_server_status === 'running' && (
                              <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>
                                ✅ 검증 시 맞춤법 서버가 자동으로 연동됩니다
                              </span>
                            )}
                          </div>
                        </div>
                      )}

                    </div>
                  )}
                </div>
              </div>

              {/* 2. 넘버링 및 목차 검증 */}
              <div 
                onClick={() => setCheckNumbering(prev => !prev)}
                style={{ 
                  background: checkNumbering ? 'rgba(168, 85, 247, 0.05)' : 'rgba(255, 255, 255, 0.01)', 
                  border: checkNumbering ? '1px solid rgba(168, 85, 247, 0.4)' : '1px solid var(--panel-border)',
                  borderRadius: '10px', 
                  padding: '12px 16px', 
                  cursor: 'pointer', 
                  display: 'flex', 
                  alignItems: 'center', 
                  gap: '12px',
                  transition: 'all 0.2s ease'
                }}
              >
                <input 
                  type="checkbox" 
                  checked={checkNumbering} 
                  onChange={(e) => {
                    e.stopPropagation();
                    setCheckNumbering(e.target.checked);
                  }}
                  style={{ width: '16px', height: '16px', cursor: 'pointer', accentColor: '#a855f7' }}
                />
                <div style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                  <span style={{ fontSize: '13.5px', fontWeight: 700, color: 'var(--text-primary)' }}>넘버링 및 목차 검증</span>
                  <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>페이지 순서 및 타이틀 중복 점검</span>
                </div>
              </div>

              {/* 3. 대체텍스트 검색 및 추출 */}
              <div 
                onClick={() => setCheckAltText(prev => !prev)}
                style={{ 
                  background: checkAltText ? 'rgba(6, 182, 212, 0.05)' : 'rgba(255, 255, 255, 0.01)', 
                  border: checkAltText ? '1px solid rgba(6, 182, 212, 0.4)' : '1px solid var(--panel-border)',
                  borderRadius: '10px', 
                  padding: '12px 16px', 
                  cursor: 'pointer', 
                  display: 'flex', 
                  alignItems: 'center', 
                  gap: '12px',
                  transition: 'all 0.2s ease'
                }}
              >
                <input 
                  type="checkbox" 
                  checked={checkAltText} 
                  onChange={(e) => {
                    e.stopPropagation();
                    setCheckAltText(e.target.checked);
                  }}
                  style={{ width: '16px', height: '16px', cursor: 'pointer', accentColor: '#06b6d4' }}
                />
                <div style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                  <span style={{ fontSize: '13.5px', fontWeight: 700, color: 'var(--text-primary)' }}>대체텍스트 검색 및 추출</span>
                  <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>시각 개체(그림/차트)에 작성된 대체텍스트(descr) 내용 검색 및 수집</span>
                </div>
              </div>

              {/* 4. 특정 단어(금지어) 포함 검증 */}
              <div 
                onClick={() => setCheckForbiddenWords(prev => !prev)}
                style={{ 
                  background: checkForbiddenWords ? 'rgba(236, 72, 153, 0.05)' : 'rgba(255, 255, 255, 0.01)', 
                  border: checkForbiddenWords ? '1px solid rgba(236, 72, 153, 0.4)' : '1px solid var(--panel-border)',
                  borderRadius: '10px', 
                  padding: '12px 16px', 
                  cursor: 'pointer', 
                  display: 'flex', 
                  alignItems: 'center', 
                  gap: '12px',
                  transition: 'all 0.2s ease'
                }}
              >
                <input 
                  type="checkbox" 
                  checked={checkForbiddenWords} 
                  onChange={(e) => {
                    e.stopPropagation();
                    setCheckForbiddenWords(e.target.checked);
                  }}
                  style={{ width: '16px', height: '16px', cursor: 'pointer', accentColor: '#ec4899' }}
                />
                <div style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                  <span style={{ fontSize: '13.5px', fontWeight: 700, color: 'var(--text-primary)' }}>특정 단어(금지어) 포함 검증</span>
                  <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>지정한 검토 필요 단어가 문서 본문에 포함되어 있는지 점검</span>
                </div>
              </div>

              {/* 5. 영어/한글 혼용 단어 검증 */}
              <div 
                onClick={() => setCheckEngKoMixed(prev => !prev)}
                style={{ 
                  background: checkEngKoMixed ? 'rgba(99, 102, 241, 0.05)' : 'rgba(255, 255, 255, 0.01)', 
                  border: checkEngKoMixed ? '1px solid rgba(99, 102, 241, 0.4)' : '1px solid var(--panel-border)',
                  borderRadius: '10px', 
                  padding: '12px 16px', 
                  cursor: 'pointer', 
                  display: 'flex', 
                  alignItems: 'center', 
                  gap: '12px',
                  transition: 'all 0.2s ease'
                }}
              >
                <input 
                  type="checkbox" 
                  checked={checkEngKoMixed} 
                  onChange={(e) => {
                    e.stopPropagation();
                    setCheckEngKoMixed(e.target.checked);
                  }}
                  style={{ width: '16px', height: '16px', cursor: 'pointer', accentColor: '#6366f1' }}
                />
                <div style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                  <span style={{ fontSize: '13.5px', fontWeight: 700, color: 'var(--text-primary)' }}>영어/한글 혼용 단어 검증</span>
                  <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>영어와 한글이 공백 없이 한 단어로 혼용되어 표시된 단어(예: re에이전트) 검출</span>
                </div>
              </div>

              {/* 6. 동일 단어 중복 검증 */}
              <div 
                onClick={() => set_check_duplicate_words(prev => !prev)}
                style={{ 
                  background: check_duplicate_words ? 'rgba(249, 115, 22, 0.05)' : 'rgba(255, 255, 255, 0.01)', 
                  border: check_duplicate_words ? '1px solid rgba(249, 115, 22, 0.4)' : '1px solid var(--panel-border)',
                  borderRadius: '10px', 
                  padding: '12px 16px', 
                  cursor: 'pointer', 
                  display: 'flex', 
                  alignItems: 'center', 
                  gap: '12px',
                  transition: 'all 0.2s ease'
                }}
              >
                <input 
                  type="checkbox" 
                  checked={check_duplicate_words} 
                  onChange={(e) => {
                    e.stopPropagation();
                    set_check_duplicate_words(e.target.checked);
                  }}
                  style={{ width: '16px', height: '16px', cursor: 'pointer', accentColor: '#f97316' }}
                />
                <div style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                  <span style={{ fontSize: '13.5px', fontWeight: 700, color: 'var(--text-primary)' }}>동일 단어 중복 검증</span>
                  <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>동일 단어/기호가 연속으로 두 번 기재된 오탈자(예: (ISP)(ISP), 데이터 데이터) 검출</span>
                </div>
              </div>

              {/* 7. 페이지 범위 및 수량 분석 */}
              <div 
                onClick={() => setCheckPageRange(prev => !prev)}
                style={{ 
                  background: checkPageRange ? 'rgba(20, 184, 166, 0.05)' : 'rgba(255, 255, 255, 0.01)', 
                  border: checkPageRange ? '1px solid rgba(20, 184, 166, 0.4)' : '1px solid var(--panel-border)',
                  borderRadius: '10px', 
                  padding: '12px 16px', 
                  cursor: 'pointer', 
                  display: 'flex', 
                  alignItems: 'center', 
                  gap: '12px',
                  transition: 'all 0.2s ease'
                }}
              >
                <input 
                  type="checkbox" 
                  checked={checkPageRange} 
                  onChange={(e) => {
                    e.stopPropagation();
                    setCheckPageRange(e.target.checked);
                  }}
                  style={{ width: '16px', height: '16px', cursor: 'pointer', accentColor: '#14b8a6' }}
                />
                <div style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                  <span style={{ fontSize: '13.5px', fontWeight: 700, color: 'var(--text-primary)' }}>페이지 범위 및 수량 분석</span>
                  <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>각 파일의 시작 페이지, 최종 페이지 번호 및 총 페이지수(슬라이드 수) 분석</span>
                </div>
              </div>

              {/* 8. 맥(Mac) 이미지 누락 검증 */}
              <div 
                onClick={() => setCheckMacImages(prev => !prev)}
                style={{ 
                  background: checkMacImages ? 'rgba(16, 185, 129, 0.05)' : 'rgba(255, 255, 255, 0.01)', 
                  border: checkMacImages ? '1px solid rgba(16, 185, 129, 0.4)' : '1px solid var(--panel-border)',
                  borderRadius: '10px', 
                  padding: '12px 16px', 
                  cursor: 'pointer', 
                  display: 'flex', 
                  alignItems: 'center', 
                  gap: '12px',
                  transition: 'all 0.2s ease'
                }}
              >
                <input 
                  type="checkbox" 
                  checked={checkMacImages} 
                  onChange={(e) => {
                    e.stopPropagation();
                    setCheckMacImages(e.target.checked);
                  }}
                  style={{ width: '16px', height: '16px', cursor: 'pointer', accentColor: '#10b981' }}
                />
                <div style={{ display: 'flex', flexDirection: 'column', gap: '2px' }}>
                  <span style={{ fontSize: '13.5px', fontWeight: 700, color: 'var(--text-primary)' }}>맥(Mac) 이미지 누락 검증</span>
                  <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>맥 OS 파워포인트 저장 시 발생하는 일부 이미지(Media) 리소스 누락/깨짐 여부 검출</span>
                </div>
              </div>
            </div>

            <button
              onClick={handleValidate}
              disabled={pptFiles.length === 0 || isProcessing}
              style={{
                width: '100%',
                padding: '14px',
                background: (pptFiles.length === 0 || isProcessing) ? 'rgba(255,255,255,0.05)' : 'linear-gradient(135deg, #e11d48, #a855f7)',
                color: (pptFiles.length === 0 || isProcessing) ? 'var(--text-muted)' : 'white',
                border: 'none',
                borderRadius: '10px',
                fontSize: '15px',
                fontWeight: 700,
                cursor: (pptFiles.length === 0 || isProcessing) ? 'not-allowed' : 'pointer',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                gap: '10px',
                transition: 'all 0.3s ease',
                marginTop: '6px'
              }}
            >
              {isProcessing ? (
                <><RefreshCw size={18} className="animate-spin" /> PPTX 구조 해독 및 규칙 검증 중...</>
              ) : (
                <><ShieldAlert size={18} /> 검증 프로세스 실행</>
              )}
            </button>
          </div>
        </div>

        {/* 우측: 검사 사전 및 금지어/검색어 설정 */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
          
          {/* 사용자 정의 검증 사전 등록 */}
          <div style={{ background: 'var(--panel-bg)', border: '1px solid var(--panel-border)', borderRadius: '16px', padding: '20px', display: 'flex', flexDirection: 'column', gap: '14px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '8px' }}>
                <HelpCircle size={17} color="var(--accent-purple)" /> 사용자 정의 검사 규칙
              </h3>
              {/* 탭 컨트롤러 */}
              <div style={{ display: 'flex', background: 'rgba(255,255,255,0.05)', borderRadius: '8px', padding: '2px' }}>
                <button
                  onClick={() => setUserDictTab('table')}
                  style={{
                    padding: '4px 10px',
                    fontSize: '11px',
                    fontWeight: 700,
                    border: 'none',
                    borderRadius: '6px',
                    background: userDictTab === 'table' ? 'var(--panel-border)' : 'transparent',
                    color: userDictTab === 'table' ? 'var(--text-primary)' : 'var(--text-muted)',
                    cursor: 'pointer',
                    transition: 'all 0.2s ease'
                  }}
                >
                  표 편집
                </button>
                <button
                  onClick={() => setUserDictTab('text')}
                  style={{
                    padding: '4px 10px',
                    fontSize: '11px',
                    fontWeight: 700,
                    border: 'none',
                    borderRadius: '6px',
                    background: userDictTab === 'text' ? 'var(--panel-border)' : 'transparent',
                    color: userDictTab === 'text' ? 'var(--text-primary)' : 'var(--text-muted)',
                    cursor: 'pointer',
                    transition: 'all 0.2s ease'
                  }}
                >
                  텍스트 편집
                </button>
              </div>
            </div>

            <p style={{ margin: 0, fontSize: '12px', color: 'var(--text-secondary)', lineHeight: '1.4' }}>
              프로젝트 고유 명사나 내장 사전에 없는 오탈자 규칙을 설정해 검출 정확도를 향상시킬 수 있습니다.
            </p>

            {userDictTab === 'text' ? (
              <textarea
                value={userDictText}
                onChange={(e) => setUserDictText(e.target.value)}
                placeholder={`예시 입력:
프로젝트명 ➜ 건강한 프로젝트
아키택처 ➜ 아키텍처
스프링부터 ➜ 스프링부트`}
                style={{
                  width: '100%',
                  height: '180px',
                  padding: '12px',
                  background: 'rgba(0,0,0,0.2)',
                  border: '1px solid var(--panel-border)',
                  borderRadius: '10px',
                  color: 'var(--text-primary)',
                  fontSize: '13px',
                  fontFamily: 'monospace',
                  lineHeight: '1.6',
                  resize: 'none',
                  outline: 'none'
                }}
              />
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                {/* 엑셀 관리 버튼바 */}
                <div style={{ display: 'flex', gap: '6px', flexWrap: 'wrap' }}>
                  <button
                    onClick={() => dictFileInputRef.current?.click()}
                    style={{
                      flex: 1,
                      padding: '8px 10px',
                      background: 'rgba(99, 102, 241, 0.1)',
                      border: '1px solid rgba(99, 102, 241, 0.2)',
                      borderRadius: '8px',
                      color: '#818cf8',
                      fontSize: '11.5px',
                      fontWeight: 700,
                      cursor: 'pointer',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      gap: '4px',
                      transition: 'all 0.2s ease'
                    }}
                  >
                    <FileSpreadsheet size={13} />
                    가져오기
                  </button>
                  <button
                    onClick={handle_export_dict_excel}
                    style={{
                      flex: 1,
                      padding: '8px 10px',
                      background: 'rgba(255, 255, 255, 0.05)',
                      border: '1px solid var(--panel-border)',
                      borderRadius: '8px',
                      color: 'var(--text-primary)',
                      fontSize: '11.5px',
                      fontWeight: 700,
                      cursor: 'pointer',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      gap: '4px',
                      transition: 'all 0.2s ease'
                    }}
                  >
                    <Download size={13} />
                    내보내기
                  </button>
                  <button
                    onClick={handle_download_dict_template}
                    style={{
                      padding: '8px 10px',
                      background: 'rgba(255, 255, 255, 0.02)',
                      border: '1px dashed var(--panel-border)',
                      borderRadius: '8px',
                      color: 'var(--text-secondary)',
                      fontSize: '11px',
                      cursor: 'pointer',
                      transition: 'all 0.2s ease'
                    }}
                  >
                    템플릿 받기
                  </button>
                  <input
                    type="file"
                    ref={dictFileInputRef}
                    onChange={handle_import_dict_excel}
                    accept=".xlsx, .xls"
                    style={{ display: 'none' }}
                  />
                </div>

                {/* 행 추가 영역 */}
                <div style={{ display: 'flex', gap: '6px' }}>
                  <input
                    id="new-typo-input"
                    type="text"
                    placeholder="오타 단어"
                    style={{
                      flex: 1,
                      padding: '8px 10px',
                      background: 'rgba(0,0,0,0.2)',
                      border: '1px solid var(--panel-border)',
                      borderRadius: '8px',
                      color: 'var(--text-primary)',
                      fontSize: '12px',
                      outline: 'none'
                    }}
                  />
                  <input
                    id="new-corr-input"
                    type="text"
                    placeholder="올바른 단어"
                    style={{
                      flex: 1,
                      padding: '8px 10px',
                      background: 'rgba(0,0,0,0.2)',
                      border: '1px solid var(--panel-border)',
                      borderRadius: '8px',
                      color: 'var(--text-primary)',
                      fontSize: '12px',
                      outline: 'none'
                    }}
                  />
                  <button
                    onClick={() => {
                      const typoInput = document.getElementById('new-typo-input');
                      const corrInput = document.getElementById('new-corr-input');
                      if (typoInput && corrInput) {
                        const tVal = typoInput.value.trim();
                        const cVal = corrInput.value.trim();
                        if (tVal && cVal) {
                          handle_add_dict_row(tVal, cVal);
                          typoInput.value = '';
                          corrInput.value = '';
                        } else {
                          alert('단어를 입력해 주세요.');
                        }
                      }
                    }}
                    style={{
                      padding: '8px 12px',
                      background: 'linear-gradient(135deg, #a855f7, #6366f1)',
                      border: 'none',
                      borderRadius: '8px',
                      color: 'white',
                      fontSize: '12px',
                      fontWeight: 700,
                      cursor: 'pointer',
                      display: 'flex',
                      alignItems: 'center',
                      gap: '4px'
                    }}
                  >
                    <Plus size={14} />
                    추가
                  </button>
                </div>

                {/* 그리드형 표 */}
                <div style={{ maxHeight: '200px', overflowY: 'auto', border: '1px solid var(--panel-border)', borderRadius: '8px', background: 'rgba(0,0,0,0.1)' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px' }}>
                    <thead>
                      <tr style={{ background: 'rgba(255,255,255,0.03)', borderBottom: '1px solid var(--panel-border)', color: 'var(--text-secondary)' }}>
                        <th style={{ padding: '8px', textAlign: 'left', fontWeight: 600 }}>오탈자 단어</th>
                        <th style={{ padding: '8px', textAlign: 'left', fontWeight: 600 }}>교정 단어</th>
                        <th style={{ padding: '8px', textAlign: 'center', width: '50px', fontWeight: 600 }}>삭제</th>
                      </tr>
                    </thead>
                    <tbody>
                      {get_user_dict_list().length === 0 ? (
                        <tr>
                          <td colSpan={3} style={{ padding: '20px', textAlign: 'center', color: 'var(--text-muted)' }}>
                            등록된 사용자 정의 규칙이 없습니다.
                          </td>
                        </tr>
                      ) : (
                        get_user_dict_list().map((item, idx) => (
                          <tr key={idx} style={{ borderBottom: '1px solid rgba(255,255,255,0.02)' }}>
                            <td style={{ padding: '8px', color: '#f43f5e', fontWeight: 600 }}>{item.typo}</td>
                            <td style={{ padding: '8px', color: 'var(--success-color)', fontWeight: 600 }}>{item.correction}</td>
                            <td style={{ padding: '8px', textAlign: 'center' }}>
                              <button
                                onClick={() => handle_delete_dict_row(item.typo)}
                                style={{
                                  background: 'transparent',
                                  border: 'none',
                                  color: 'var(--text-muted)',
                                  cursor: 'pointer',
                                  padding: '4px',
                                  borderRadius: '4px',
                                  display: 'inline-flex',
                                  alignItems: 'center',
                                  justifyContent: 'center',
                                  transition: 'all 0.2s'
                                }}
                                onMouseEnter={(e) => e.currentTarget.style.color = '#ef4444'}
                                onMouseLeave={(e) => e.currentTarget.style.color = 'var(--text-muted)'}
                              >
                                <Trash2 size={13} />
                              </button>
                            </td>
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', fontSize: '11.5px', color: 'var(--text-secondary)', marginTop: '4px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '11px', color: 'var(--text-muted)' }}>
                <Info size={12} />
                <span>내장 사전에 있는 규칙은 기본적으로 적용됩니다.</span>
              </div>
              <button
                onClick={() => setShowDictModal(true)}
                style={{
                  background: 'rgba(99, 102, 241, 0.1)',
                  color: '#818cf8',
                  border: '1px solid rgba(99, 102, 241, 0.2)',
                  borderRadius: '6px',
                  padding: '3px 8px',
                  fontSize: '11px',
                  fontWeight: 600,
                  cursor: 'pointer',
                  transition: 'all 0.2s'
                }}
                onMouseEnter={(e) => {
                  e.currentTarget.style.background = 'rgba(99, 102, 241, 0.2)';
                  e.currentTarget.style.borderColor = 'rgba(99, 102, 241, 0.4)';
                }}
                onMouseLeave={(e) => {
                  e.currentTarget.style.background = 'rgba(99, 102, 241, 0.1)';
                  e.currentTarget.style.borderColor = 'rgba(99, 102, 241, 0.2)';
                }}
              >
                표준 사전 보기
              </button>
            </div>
          </div>

          {/* 특정 검색/금지 단어 지정 */}
          <div style={{ background: 'var(--panel-bg)', border: '1px solid var(--panel-border)', borderRadius: '16px', padding: '20px', display: 'flex', flexDirection: 'column', gap: '14px' }}>
            <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '8px' }}>
              <HelpCircle size={17} color="#ec4899" /> 특정 점검 단어 지정 (선택)
            </h3>
            <p style={{ margin: 0, fontSize: '12.5px', color: 'var(--text-secondary)', lineHeight: '1.5' }}>
              문서 본문에 포함되면 안 되는 금지어나 검토가 필요한 단어들을 <strong>한 줄에 하나씩</strong> 입력해 주세요.
            </p>
            <textarea
              value={forbiddenWordsText}
              onChange={(e) => setForbiddenWordsText(e.target.value)}
              placeholder={`예시 입력:
미정
임시
TBD
검토필요
작성중`}
              style={{
                width: '100%',
                height: '140px',
                padding: '12px',
                background: 'rgba(0,0,0,0.2)',
                border: '1px solid var(--panel-border)',
                borderRadius: '10px',
                color: 'var(--text-primary)',
                fontSize: '13px',
                fontFamily: 'monospace',
                lineHeight: '1.6',
                resize: 'none',
                outline: 'none'
              }}
            />
            <div style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '11px', color: 'var(--text-muted)' }}>
              <Info size={12} />
              <span>'특정 단어(금지어) 포함 검증' 옵션이 켜져 있어야 작동합니다.</span>
            </div>
          </div>

        </div>

      </div>

      {/* 하단: 검증 결과 보고 대시보드 */}
      {isValidated && (
        <div style={{ background: 'var(--panel-bg)', border: '1px solid var(--panel-border)', borderRadius: '16px', padding: '24px', display: 'flex', flexDirection: 'column', gap: '20px', animation: 'fadeIn 0.4s ease' }}>
          
          {/* 검증 요약 요약바 */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid var(--panel-border)', paddingBottom: '16px' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
              <CheckCircle2 size={24} color="var(--success-color)" />
              <h3 style={{ margin: 0, fontSize: '18px', fontWeight: 800 }}>산출물 검증 완료 리포트</h3>
            </div>
            
            <button
              onClick={handleExportToExcel}
              style={{
                padding: '10px 18px',
                background: 'linear-gradient(135deg, #10b981, #059669)',
                color: 'white',
                border: 'none',
                borderRadius: '8px',
                cursor: 'pointer',
                fontWeight: 700,
                fontSize: '13px',
                display: 'flex',
                alignItems: 'center',
                gap: '8px',
                transition: 'opacity 0.3s'
              }}
              onMouseEnter={(e) => e.target.style.opacity = '0.9'}
              onMouseLeave={(e) => e.target.style.opacity = '1'}
            >
              <FileSpreadsheet size={16} /> 검증 결과 엑셀로 내려받기
            </button>
          </div>

          {/* 종합 요약 카드 */}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(8, 1fr)', gap: '16px' }}>
            <div style={{ background: 'rgba(255,255,255,0.02)', border: '1px solid var(--panel-border)', padding: '16px 20px', borderRadius: '12px', display: 'flex', flexDirection: 'column', gap: '6px' }}>
              <span style={{ fontSize: '13px', color: 'var(--text-secondary)', fontWeight: 600 }}>검증 파일 수</span>
              <span style={{ fontSize: '24px', fontWeight: 900, color: 'var(--text-primary)' }}>{pptFiles.length}개 파일</span>
            </div>
            <div style={{ background: checkTypos ? 'rgba(239, 68, 68, 0.05)' : 'rgba(255, 255, 255, 0.01)', border: checkTypos ? '1px solid rgba(239, 68, 68, 0.15)' : '1px solid var(--panel-border)', padding: '16px 20px', borderRadius: '12px', display: 'flex', flexDirection: 'column', gap: '6px' }}>
              <span style={{ fontSize: '13px', color: checkTypos ? '#f87171' : 'var(--text-muted)', fontWeight: 600 }}>검출된 오탈자 건수</span>
              <span style={{ fontSize: '24px', fontWeight: 900, color: checkTypos ? '#ef4444' : 'var(--text-muted)' }}>{checkTypos ? `${typoResults.length}건` : '비활성'}</span>
            </div>
            <div style={{ background: checkNumbering ? 'rgba(245, 158, 11, 0.05)' : 'rgba(255, 255, 255, 0.01)', border: checkNumbering ? '1px solid rgba(245, 158, 11, 0.15)' : '1px solid var(--panel-border)', padding: '16px 20px', borderRadius: '12px', display: 'flex', flexDirection: 'column', gap: '6px' }}>
              <span style={{ fontSize: '13px', color: checkNumbering ? '#fbbf24' : 'var(--text-muted)', fontWeight: 600 }}>넘버링 위반 건수</span>
              <span style={{ fontSize: '24px', fontWeight: 900, color: checkNumbering ? '#f59e0b' : 'var(--text-muted)' }}>{checkNumbering ? `${numberingResults.length}건` : '비활성'}</span>
            </div>
            <div style={{ background: checkAltText ? 'rgba(6, 182, 212, 0.05)' : 'rgba(255, 255, 255, 0.01)', border: checkAltText ? '1px solid rgba(6, 182, 212, 0.15)' : '1px solid var(--panel-border)', padding: '16px 20px', borderRadius: '12px', display: 'flex', flexDirection: 'column', gap: '6px' }}>
              <span style={{ fontSize: '13px', color: checkAltText ? '#22d3ee' : 'var(--text-muted)', fontWeight: 600 }}>대체텍스트 검출 건수</span>
              <span style={{ fontSize: '24px', fontWeight: 900, color: checkAltText ? '#06b6d4' : 'var(--text-muted)' }}>{checkAltText ? `${altTextResults.length}건` : '비활성'}</span>
            </div>
            <div style={{ background: checkForbiddenWords ? 'rgba(236, 72, 153, 0.05)' : 'rgba(255, 255, 255, 0.01)', border: checkForbiddenWords ? '1px solid rgba(236, 72, 153, 0.15)' : '1px solid var(--panel-border)', padding: '16px 20px', borderRadius: '12px', display: 'flex', flexDirection: 'column', gap: '6px' }}>
              <span style={{ fontSize: '13px', color: checkForbiddenWords ? '#f472b6' : 'var(--text-muted)', fontWeight: 600 }}>특정 단어 검출 건수</span>
              <span style={{ fontSize: '24px', fontWeight: 900, color: checkForbiddenWords ? '#ec4899' : 'var(--text-muted)' }}>{checkForbiddenWords ? `${forbiddenResults.length}건` : '비활성'}</span>
            </div>
            <div style={{ background: checkEngKoMixed ? 'rgba(99, 102, 241, 0.05)' : 'rgba(255, 255, 255, 0.01)', border: checkEngKoMixed ? '1px solid rgba(99, 102, 241, 0.15)' : '1px solid var(--panel-border)', padding: '16px 20px', borderRadius: '12px', display: 'flex', flexDirection: 'column', gap: '6px' }}>
              <span style={{ fontSize: '13px', color: checkEngKoMixed ? '#818cf8' : 'var(--text-muted)', fontWeight: 600 }}>영한 혼용 검출 건수</span>
              <span style={{ fontSize: '24px', fontWeight: 900, color: checkEngKoMixed ? '#6366f1' : 'var(--text-muted)' }}>{checkEngKoMixed ? `${engKoMixedResults.length}건` : '비활성'}</span>
            </div>
            <div style={{ background: check_duplicate_words ? 'rgba(249, 115, 22, 0.05)' : 'rgba(255, 255, 255, 0.01)', border: check_duplicate_words ? '1px solid rgba(249, 115, 22, 0.15)' : '1px solid var(--panel-border)', padding: '16px 20px', borderRadius: '12px', display: 'flex', flexDirection: 'column', gap: '6px' }}>
              <span style={{ fontSize: '13px', color: check_duplicate_words ? '#fb923c' : 'var(--text-muted)', fontWeight: 600 }}>동일 단어 중복 건수</span>
              <span style={{ fontSize: '24px', fontWeight: 900, color: check_duplicate_words ? '#f97316' : 'var(--text-muted)' }}>{check_duplicate_words ? `${duplicate_results.length}건` : '비활성'}</span>
            </div>
            <div style={{ background: checkMacImages ? 'rgba(16, 185, 129, 0.05)' : 'rgba(255, 255, 255, 0.01)', border: checkMacImages ? '1px solid rgba(16, 185, 129, 0.15)' : '1px solid var(--panel-border)', padding: '16px 20px', borderRadius: '12px', display: 'flex', flexDirection: 'column', gap: '6px' }}>
              <span style={{ fontSize: '13px', color: checkMacImages ? '#34d399' : 'var(--text-muted)', fontWeight: 600 }}>맥 이미지 누락 건수</span>
              <span style={{ fontSize: '24px', fontWeight: 900, color: checkMacImages ? '#10b981' : 'var(--text-muted)' }}>{checkMacImages ? `${macImageResults.length}건` : '비활성'}</span>
            </div>
          </div>

          {/* 결과 상세 확인 테이블 탭 */}
          <div style={{ marginTop: '10px' }}>
            <div style={{ display: 'flex', gap: '8px', borderBottom: '1px solid var(--panel-border)', paddingBottom: '10px' }}>
              <button 
                onClick={() => setActiveResultTab('summary')}
                style={{
                  padding: '8px 16px',
                  background: activeResultTab === 'summary' ? 'rgba(255,255,255,0.06)' : 'transparent',
                  border: 'none',
                  borderRadius: '6px',
                  color: activeResultTab === 'summary' ? 'var(--text-primary)' : 'var(--text-muted)',
                  cursor: 'pointer',
                  fontWeight: 700,
                  fontSize: '13.5px'
                }}
              >
                📊 파일별 점검 요약
              </button>
              {checkTypos && (
                <button 
                  onClick={() => setActiveResultTab('typo')}
                  style={{
                    padding: '8px 16px',
                    background: activeResultTab === 'typo' ? 'rgba(239, 68, 68, 0.1)' : 'transparent',
                    border: 'none',
                    borderRadius: '6px',
                    color: activeResultTab === 'typo' ? '#ef4444' : 'var(--text-muted)',
                    cursor: 'pointer',
                    fontWeight: 700,
                    fontSize: '13.5px',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '6px'
                  }}
                >
                  ❌ 오탈자 ({typoResults.length})
                </button>
              )}
              {checkNumbering && (
                <button 
                  onClick={() => setActiveResultTab('numbering')}
                  style={{
                    padding: '8px 16px',
                    background: activeResultTab === 'numbering' ? 'rgba(245, 158, 11, 0.1)' : 'transparent',
                    border: 'none',
                    borderRadius: '6px',
                    color: activeResultTab === 'numbering' ? '#f59e0b' : 'var(--text-muted)',
                    cursor: 'pointer',
                    fontWeight: 700,
                    fontSize: '13.5px',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '6px'
                  }}
                >
                  ⚠️ 넘버링 규칙성 ({numberingResults.length})
                </button>
              )}
              {checkAltText && (
                <button 
                  onClick={() => setActiveResultTab('altText')}
                  style={{
                    padding: '8px 16px',
                    background: activeResultTab === 'altText' ? 'rgba(6, 182, 212, 0.1)' : 'transparent',
                    border: 'none',
                    borderRadius: '6px',
                    color: activeResultTab === 'altText' ? '#06b6d4' : 'var(--text-muted)',
                    cursor: 'pointer',
                    fontWeight: 700,
                    fontSize: '13.5px',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '6px'
                  }}
                >
                  🖼️ 대체텍스트 검색 ({altTextResults.length})
                </button>
              )}
              {checkForbiddenWords && (
                <button 
                  onClick={() => setActiveResultTab('forbidden')}
                  style={{
                    padding: '8px 16px',
                    background: activeResultTab === 'forbidden' ? 'rgba(236, 72, 153, 0.1)' : 'transparent',
                    border: 'none',
                    borderRadius: '6px',
                    color: activeResultTab === 'forbidden' ? '#ec4899' : 'var(--text-muted)',
                    cursor: 'pointer',
                    fontWeight: 700,
                    fontSize: '13.5px',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '6px'
                  }}
                >
                  🔍 특정 단어 검출 ({forbiddenResults.length})
                </button>
              )}
              {checkEngKoMixed && (
                <button 
                  onClick={() => setActiveResultTab('engKoMixed')}
                  style={{
                    padding: '8px 16px',
                    background: activeResultTab === 'engKoMixed' ? 'rgba(99, 102, 241, 0.1)' : 'transparent',
                    border: 'none',
                    borderRadius: '6px',
                    color: activeResultTab === 'engKoMixed' ? '#6366f1' : 'var(--text-muted)',
                    cursor: 'pointer',
                    fontWeight: 700,
                    fontSize: '13.5px',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '6px'
                  }}
                >
                  🔤 영한 혼용 단어 ({engKoMixedResults.length})
                </button>
              )}
              {check_duplicate_words && (
                <button 
                  onClick={() => setActiveResultTab('duplicate')}
                  style={{
                    padding: '8px 16px',
                    background: activeResultTab === 'duplicate' ? 'rgba(249, 115, 22, 0.1)' : 'transparent',
                    border: 'none',
                    borderRadius: '6px',
                    color: activeResultTab === 'duplicate' ? '#f97316' : 'var(--text-muted)',
                    cursor: 'pointer',
                    fontWeight: 700,
                    fontSize: '13.5px',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '6px'
                  }}
                >
                  🔁 동일 단어 중복 ({duplicate_results.length})
                </button>
              )}
              {checkPageRange && (
                <button 
                  onClick={() => setActiveResultTab('pageRange')}
                  style={{
                    padding: '8px 16px',
                    background: activeResultTab === 'pageRange' ? 'rgba(20, 184, 166, 0.1)' : 'transparent',
                    border: 'none',
                    borderRadius: '6px',
                    color: activeResultTab === 'pageRange' ? '#14b8a6' : 'var(--text-muted)',
                    cursor: 'pointer',
                    fontWeight: 700,
                    fontSize: '13.5px',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '6px'
                  }}
                >
                  📄 페이지 범위 분석 ({pageRangeResults.length})
                </button>
              )}
              {checkMacImages && (
                <button 
                  onClick={() => setActiveResultTab('macImages')}
                  style={{
                    padding: '8px 16px',
                    background: activeResultTab === 'macImages' ? 'rgba(16, 185, 129, 0.1)' : 'transparent',
                    border: 'none',
                    borderRadius: '6px',
                    color: activeResultTab === 'macImages' ? '#10b981' : 'var(--text-muted)',
                    cursor: 'pointer',
                    fontWeight: 700,
                    fontSize: '13.5px',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '6px'
                  }}
                >
                  🖼️ 맥 이미지 누락 ({macImageResults.length})
                </button>
              )}
            </div>

            {/* 탭 1: 파일별 점검 요약 */}
            {activeResultTab === 'summary' && (
              <div style={{ marginTop: '16px', overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '13.5px' }}>
                  <thead>
                    <tr style={{ borderBottom: '2px solid var(--panel-border)', color: 'var(--text-secondary)' }}>
                      <th style={{ padding: '12px 8px', fontWeight: 700 }}>파일명</th>
                      <th style={{ padding: '12px 8px', fontWeight: 700, width: '130px' }}>오탈자 의심</th>
                      <th style={{ padding: '12px 8px', fontWeight: 700, width: '130px' }}>넘버링 오류</th>
                      <th style={{ padding: '12px 8px', fontWeight: 700, width: '140px' }}>대체텍스트 검출</th>
                      <th style={{ padding: '12px 8px', fontWeight: 700, width: '140px' }}>특정 단어 검출</th>
                      <th style={{ padding: '12px 8px', fontWeight: 700, width: '140px' }}>영한 혼용 검출</th>
                      <th style={{ padding: '12px 8px', fontWeight: 700, width: '140px' }}>중복 단어 검출</th>
                      <th style={{ padding: '12px 8px', fontWeight: 700, width: '140px' }}>맥 이미지 누락</th>
                      <th style={{ padding: '12px 8px', fontWeight: 700, width: '100px' }}>시작페이지</th>
                      <th style={{ padding: '12px 8px', fontWeight: 700, width: '100px' }}>최종 페이지</th>
                      <th style={{ padding: '12px 8px', fontWeight: 700, width: '100px' }}>총 페이지수</th>
                      <th style={{ padding: '12px 8px', fontWeight: 700, width: '100px' }}>종합 상태</th>
                    </tr>
                  </thead>
                  <tbody>
                    {fileStats.map((stat, idx) => {
                      const totalErrors = 
                        (checkTypos ? stat.typos : 0) + 
                        (checkNumbering ? stat.numberingErrors : 0) +
                        (checkAltText ? stat.altTextErrors : 0) +
                        (checkForbiddenWords ? stat.forbiddenErrors : 0) +
                        (checkEngKoMixed ? stat.engKoMixedErrors : 0) +
                        (check_duplicate_words ? stat.duplicateErrors : 0) +
                        (checkMacImages ? stat.macImageErrors : 0);
                      return (
                        <tr key={idx} style={{ borderBottom: '1px solid var(--panel-border)' }}>
                          <td style={{ padding: '14px 8px', fontWeight: 600 }}>{stat.name}</td>
                          <td style={{ padding: '14px 8px', color: !checkTypos ? 'var(--text-muted)' : stat.typos > 0 ? '#ef4444' : 'var(--text-muted)', fontWeight: 700 }}>
                            {checkTypos ? (stat.typos > 0 ? `${stat.typos}건` : '없음') : '비활성'}
                          </td>
                          <td style={{ padding: '14px 8px', color: !checkNumbering ? 'var(--text-muted)' : stat.numberingErrors > 0 ? '#f59e0b' : 'var(--text-muted)', fontWeight: 700 }}>
                            {checkNumbering ? (stat.numberingErrors > 0 ? `${stat.numberingErrors}건` : '없음') : '비활성'}
                          </td>
                          <td style={{ padding: '14px 8px', color: !checkAltText ? 'var(--text-muted)' : stat.altTextErrors > 0 ? '#06b6d4' : 'var(--text-muted)', fontWeight: 700 }}>
                            {checkAltText ? (stat.altTextErrors > 0 ? `${stat.altTextErrors}건` : '없음') : '비활성'}
                          </td>
                          <td style={{ padding: '14px 8px', color: !checkForbiddenWords ? 'var(--text-muted)' : stat.forbiddenErrors > 0 ? '#ec4899' : 'var(--text-muted)', fontWeight: 700 }}>
                            {checkForbiddenWords ? (stat.forbiddenErrors > 0 ? `${stat.forbiddenErrors}건` : '없음') : '비활성'}
                          </td>
                          <td style={{ padding: '14px 8px', color: !checkEngKoMixed ? 'var(--text-muted)' : stat.engKoMixedErrors > 0 ? '#6366f1' : 'var(--text-muted)', fontWeight: 700 }}>
                            {checkEngKoMixed ? (stat.engKoMixedErrors > 0 ? `${stat.engKoMixedErrors}건` : '없음') : '비활성'}
                          </td>
                          <td style={{ padding: '14px 8px', color: !check_duplicate_words ? 'var(--text-muted)' : stat.duplicateErrors > 0 ? '#f97316' : 'var(--text-muted)', fontWeight: 700 }}>
                            {check_duplicate_words ? (stat.duplicateErrors > 0 ? `${stat.duplicateErrors}건` : '없음') : '비활성'}
                          </td>
                          <td style={{ padding: '14px 8px', color: !checkMacImages ? 'var(--text-muted)' : stat.macImageErrors > 0 ? '#10b981' : 'var(--text-muted)', fontWeight: 700 }}>
                            {checkMacImages ? (stat.macImageErrors > 0 ? `${stat.macImageErrors}건` : '없음') : '비활성'}
                          </td>
                          <td style={{ padding: '14px 8px', color: !checkPageRange ? 'var(--text-muted)' : 'var(--text-secondary)', fontWeight: 600 }}>
                            {checkPageRange ? `${stat.startPage}p` : '비활성'}
                          </td>
                          <td style={{ padding: '14px 8px', color: !checkPageRange ? 'var(--text-muted)' : 'var(--text-secondary)', fontWeight: 600 }}>
                            {checkPageRange ? `${stat.endPage}p` : '비활성'}
                          </td>
                          <td style={{ padding: '14px 8px', color: !checkPageRange ? 'var(--text-muted)' : 'var(--text-secondary)', fontWeight: 600 }}>
                            {checkPageRange ? `${stat.totalSlides}장` : '비활성'}
                          </td>
                          <td style={{ padding: '14px 8px' }}>
                            {totalErrors === 0 ? (
                              <span style={{ background: 'rgba(16, 185, 129, 0.1)', color: '#10b981', padding: '4px 8px', borderRadius: '6px', fontSize: '11.5px', fontWeight: 800 }}>이상 없음</span>
                            ) : (
                              <span style={{ background: 'rgba(239, 68, 68, 0.1)', color: '#ef4444', padding: '4px 8px', borderRadius: '6px', fontSize: '11.5px', fontWeight: 800 }}>검토 필요</span>
                            )}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}

            {/* 탭 2: 오탈자 상세 목록 */}
            {activeResultTab === 'typo' && (
              <div style={{ marginTop: '16px' }}>
                {typoResults.length === 0 ? (
                  <div style={{ padding: '40px 0', textAlign: 'center', color: 'var(--text-muted)', fontSize: '13.5px' }}>
                    🎉 검출된 오탈자 의심 단어가 없습니다!
                  </div>
                ) : (
                  <div style={{ overflowX: 'auto' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '13px' }}>
                      <thead>
                        <tr style={{ borderBottom: '2px solid var(--panel-border)', color: 'var(--text-secondary)' }}>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '180px' }}>파일명</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '90px' }}>슬라이드(물리)</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '90px' }}>표시 페이지</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '110px' }}>의심 단어</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '110px' }}>교정 추천</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '80px' }}>유형</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700 }}>검출 문장</th>
                        </tr>
                      </thead>
                      <tbody>
                        {typoResults.map((t, idx) => (
                          <tr key={idx} style={{ borderBottom: '1px solid var(--panel-border)' }} className="table-row-hover">
                            <td style={{ padding: '12px 8px', fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: '180px' }} title={t.fileName}>
                              {t.fileName}
                            </td>
                            <td style={{ padding: '12px 8px', color: 'var(--text-secondary)', fontWeight: 600 }}>
                              {t.slideNum} 순서
                            </td>
                            <td style={{ padding: '12px 8px', color: 'var(--text-primary)', fontWeight: 700 }}>
                              {t.displayPageNum} 페이지
                            </td>
                            <td style={{ padding: '12px 8px', color: '#ef4444', fontWeight: 700 }}>{t.typo}</td>
                            <td style={{ padding: '12px 8px', color: 'var(--success-color)', fontWeight: 700 }}>
                              {t.correction}
                            </td>
                            <td style={{ padding: '12px 8px' }}>
                              {t.isAi ? (
                                <span style={{ 
                                  background: 'rgba(99, 102, 241, 0.15)', 
                                  color: '#818cf8',
                                  padding: '3px 8px', 
                                  borderRadius: '6px', 
                                  fontSize: '11px',
                                  fontWeight: 700,
                                  display: 'inline-flex',
                                  alignItems: 'center',
                                  gap: '4px'
                                }}>
                                  <Sparkles size={11} />
                                  AI 맞춤법
                                </span>
                              ) : (
                                <span style={{ background: 'rgba(255,255,255,0.06)', padding: '2px 6px', borderRadius: '4px', fontSize: '11px' }}>{t.type}</span>
                              )}
                            </td>
                            <td style={{ padding: '12px 8px', color: 'var(--text-secondary)', lineBreak: 'anywhere' }}>
                              {/* 매칭 단어 하이라이트 효과 */}
                              {t.sentence.split(t.typo).map((chunk, cIdx, arr) => (
                                <span key={cIdx}>
                                  {chunk}
                                  {cIdx < arr.length - 1 && <span style={{ background: 'rgba(239, 68, 68, 0.25)', color: '#ff8a8a', padding: '0 2px', borderRadius: '3px', fontWeight: 700 }}>{t.typo}</span>}
                                </span>
                              ))}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            )}

            {/* 탭 3: 넘버링 규칙성 오류 목록 */}
            {activeResultTab === 'numbering' && (
              <div style={{ marginTop: '16px' }}>
                {numberingResults.length === 0 ? (
                  <div style={{ padding: '40px 0', textAlign: 'center', color: 'var(--text-muted)', fontSize: '13.5px' }}>
                    🎉 검출된 상단 헤더 넘버링 오류가 없습니다!
                  </div>
                ) : (
                  <div style={{ overflowX: 'auto' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '13px' }}>
                      <thead>
                        <tr style={{ borderBottom: '2px solid var(--panel-border)', color: 'var(--text-secondary)' }}>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '180px' }}>파일명</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '90px' }}>슬라이드(물리)</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '90px' }}>표시 페이지</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '90px' }}>영역</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '120px' }}>표기 텍스트</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '180px' }}>검출된 오류</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700 }}>권장 교정 가이드</th>
                        </tr>
                      </thead>
                      <tbody>
                        {numberingResults.map((n, idx) => (
                          <tr key={idx} style={{ borderBottom: '1px solid var(--panel-border)' }} className="table-row-hover">
                            <td style={{ padding: '12px 8px', fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: '180px' }} title={n.fileName}>
                              {n.fileName}
                            </td>
                            <td style={{ padding: '12px 8px', color: 'var(--text-secondary)', fontWeight: 600 }}>
                              {n.slideNum} 순서
                            </td>
                            <td style={{ padding: '12px 8px', color: 'var(--text-primary)', fontWeight: 700 }}>
                              {n.displayPageNum} 페이지
                            </td>
                            <td style={{ padding: '12px 8px' }}>
                              <span style={{ 
                                background: n.area === '좌측 타이틀' ? 'rgba(99, 102, 241, 0.1)' : 'rgba(236, 72, 153, 0.1)', 
                                color: n.area === '좌측 타이틀' ? '#818cf8' : '#f472b6', 
                                padding: '2px 6px', 
                                borderRadius: '4px', 
                                fontSize: '11px',
                                fontWeight: 700
                              }}>
                                {n.area}
                              </span>
                            </td>
                            <td style={{ padding: '12px 8px', color: 'var(--text-primary)', fontWeight: 700, fontStyle: 'italic' }}>
                              {n.text}
                            </td>
                            <td style={{ padding: '12px 8px', color: '#f59e0b', fontWeight: 700 }}>
                              {n.error}
                            </td>
                            <td style={{ padding: '12px 8px', color: 'var(--text-secondary)', lineHeight: '1.5' }}>
                              {n.guide}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            )}

            {/* 탭 4: 대체텍스트 검색 목록 */}
            {activeResultTab === 'altText' && (
              <div style={{ marginTop: '16px' }}>
                {altTextResults.length === 0 ? (
                  <div style={{ padding: '40px 0', textAlign: 'center', color: 'var(--text-muted)', fontSize: '13.5px' }}>
                    🎉 검출된 대체텍스트 정보가 없습니다!
                  </div>
                ) : (
                  <div style={{ overflowX: 'auto' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '13px' }}>
                      <thead>
                        <tr style={{ borderBottom: '2px solid var(--panel-border)', color: 'var(--text-secondary)' }}>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '180px' }}>파일명</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '90px' }}>슬라이드(물리)</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '90px' }}>표시 페이지</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '150px' }}>대상 개체명</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '250px' }}>대체텍스트 내용</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700 }}>참고 사항</th>
                        </tr>
                      </thead>
                      <tbody>
                        {altTextResults.map((alt, idx) => (
                          <tr key={idx} style={{ borderBottom: '1px solid var(--panel-border)' }} className="table-row-hover">
                            <td style={{ padding: '12px 8px', fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: '180px' }} title={alt.fileName}>
                              {alt.fileName}
                            </td>
                            <td style={{ padding: '12px 8px', color: 'var(--text-secondary)', fontWeight: 600 }}>
                              {alt.slideNum} 순서
                            </td>
                            <td style={{ padding: '12px 8px', color: 'var(--text-primary)', fontWeight: 700 }}>
                              {alt.displayPageNum} 페이지
                            </td>
                            <td style={{ padding: '12px 8px', color: '#06b6d4', fontWeight: 700 }}>
                              {alt.objName}
                            </td>
                            <td style={{ padding: '12px 8px', color: 'var(--text-primary)', fontWeight: 700 }}>
                              {alt.descr}
                            </td>
                            <td style={{ padding: '12px 8px', color: 'var(--text-secondary)', lineHeight: '1.5' }}>
                              {alt.guide}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            )}

            {/* 탭 5: 특정 단어 검출 목록 */}
            {activeResultTab === 'forbidden' && (
              <div style={{ marginTop: '16px' }}>
                {forbiddenResults.length === 0 ? (
                  <div style={{ padding: '40px 0', textAlign: 'center', color: 'var(--text-muted)', fontSize: '13.5px' }}>
                    🎉 본문 내에 지정한 특정 단어가 검출되지 않았습니다!
                  </div>
                ) : (
                  <div style={{ overflowX: 'auto' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '13px' }}>
                      <thead>
                        <tr style={{ borderBottom: '2px solid var(--panel-border)', color: 'var(--text-secondary)' }}>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '180px' }}>파일명</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '90px' }}>슬라이드(물리)</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '90px' }}>표시 페이지</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '120px' }}>검출된 단어</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '120px' }}>상태</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700 }}>검출 문장(하이라이트)</th>
                        </tr>
                      </thead>
                      <tbody>
                        {forbiddenResults.map((forb, idx) => (
                          <tr key={idx} style={{ borderBottom: '1px solid var(--panel-border)' }} className="table-row-hover">
                            <td style={{ padding: '12px 8px', fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: '180px' }} title={forb.fileName}>
                              {forb.fileName}
                            </td>
                            <td style={{ padding: '12px 8px', color: 'var(--text-secondary)', fontWeight: 600 }}>
                              {forb.slideNum} 순서
                            </td>
                            <td style={{ padding: '12px 8px', color: 'var(--text-primary)', fontWeight: 700 }}>
                              {forb.displayPageNum} 페이지
                            </td>
                            <td style={{ padding: '12px 8px', color: '#ec4899', fontWeight: 700 }}>
                              {forb.word}
                            </td>
                            <td style={{ padding: '12px 8px', color: '#f59e0b', fontWeight: 700 }}>
                              {forb.error}
                            </td>
                            <td style={{ padding: '12px 8px', color: 'var(--text-secondary)', lineBreak: 'anywhere' }}>
                              {forb.sentence.split(forb.word).map((chunk, cIdx, arr) => (
                                <span key={cIdx}>
                                  {chunk}
                                  {cIdx < arr.length - 1 && (
                                    <span style={{ background: 'rgba(236, 72, 153, 0.25)', color: '#f472b6', padding: '0 2px', borderRadius: '3px', fontWeight: 700 }}>
                                      {forb.word}
                                    </span>
                                  )}
                                </span>
                              ))}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            )}

            {/* 탭 6: 영한 혼용 단어 검출 목록 */}
            {activeResultTab === 'engKoMixed' && (
              <div style={{ marginTop: '16px' }}>
                {engKoMixedResults.length === 0 ? (
                  <div style={{ padding: '40px 0', textAlign: 'center', color: 'var(--text-muted)', fontSize: '13.5px' }}>
                    🎉 영어와 한글이 공백 없이 혼용된 단어가 검출되지 않았습니다!
                  </div>
                ) : (
                  <div style={{ overflowX: 'auto' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '13px' }}>
                      <thead>
                        <tr style={{ borderBottom: '2px solid var(--panel-border)', color: 'var(--text-secondary)' }}>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '180px' }}>파일명</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '90px' }}>슬라이드(물리)</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '90px' }}>표시 페이지</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '150px' }}>검출 단어</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '120px' }}>상태</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700 }}>검출 문장(하이라이트)</th>
                        </tr>
                      </thead>
                      <tbody>
                        {engKoMixedResults.map((e, idx) => (
                          <tr key={idx} style={{ borderBottom: '1px solid var(--panel-border)' }} className="table-row-hover">
                            <td style={{ padding: '12px 8px', fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: '180px' }} title={e.fileName}>
                              {e.fileName}
                            </td>
                            <td style={{ padding: '12px 8px', color: 'var(--text-secondary)', fontWeight: 600 }}>
                              {e.slideNum} 순서
                            </td>
                            <td style={{ padding: '12px 8px', color: 'var(--text-primary)', fontWeight: 700 }}>
                              {e.displayPageNum} 페이지
                            </td>
                            <td style={{ padding: '12px 8px', color: '#6366f1', fontWeight: 700 }}>
                              {e.word}
                            </td>
                            <td style={{ padding: '12px 8px', color: '#f59e0b', fontWeight: 700 }}>
                              {e.error}
                            </td>
                            <td style={{ padding: '12px 8px', color: 'var(--text-secondary)', lineBreak: 'anywhere' }}>
                              {e.sentence.split(e.word).map((chunk, cIdx, arr) => (
                                <span key={cIdx}>
                                  {chunk}
                                  {cIdx < arr.length - 1 && (
                                    <span style={{ background: 'rgba(99, 102, 241, 0.25)', color: '#818cf8', padding: '0 2px', borderRadius: '3px', fontWeight: 700 }}>
                                      {e.word}
                                    </span>
                                  )}
                                </span>
                              ))}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            )}

            {/* 탭 7: 동일 단어 중복 검출 목록 */}
            {activeResultTab === 'duplicate' && (
              <div style={{ marginTop: '16px' }}>
                {duplicate_results.length === 0 ? (
                  <div style={{ padding: '40px 0', textAlign: 'center', color: 'var(--text-muted)', fontSize: '13.5px' }}>
                    🎉 본문 내에 연속으로 기재된 중복 단어가 검출되지 않았습니다!
                  </div>
                ) : (
                  <div style={{ overflowX: 'auto' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '13px' }}>
                      <thead>
                        <tr style={{ borderBottom: '2px solid var(--panel-border)', color: 'var(--text-secondary)' }}>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '180px' }}>파일명</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '90px' }}>슬라이드(물리)</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '90px' }}>표시 페이지</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '150px' }}>중복 표현</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '120px' }}>상태</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700 }}>검출 문장(하이라이트)</th>
                        </tr>
                      </thead>
                      <tbody>
                        {duplicate_results.map((dup, idx) => (
                          <tr key={idx} style={{ borderBottom: '1px solid var(--panel-border)' }} className="table-row-hover">
                            <td style={{ padding: '12px 8px', fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: '180px' }} title={dup.fileName}>
                              {dup.fileName}
                            </td>
                            <td style={{ padding: '12px 8px', color: 'var(--text-secondary)', fontWeight: 600 }}>
                              {dup.slideNum} 순서
                            </td>
                            <td style={{ padding: '12px 8px', color: 'var(--text-primary)', fontWeight: 700 }}>
                              {dup.displayPageNum} 페이지
                            </td>
                            <td style={{ padding: '12px 8px', color: '#f97316', fontWeight: 700 }}>
                              {dup.word}
                            </td>
                            <td style={{ padding: '12px 8px', color: '#f59e0b', fontWeight: 700 }}>
                              {dup.error}
                            </td>
                            <td style={{ padding: '12px 8px', color: 'var(--text-secondary)', lineBreak: 'anywhere' }}>
                              {dup.sentence.split(dup.word).map((chunk, cIdx, arr) => (
                                <span key={cIdx}>
                                  {chunk}
                                  {cIdx < arr.length - 1 && (
                                    <span style={{ background: 'rgba(249, 115, 22, 0.25)', color: '#fb923c', padding: '0 2px', borderRadius: '3px', fontWeight: 700 }}>
                                      {dup.word}
                                    </span>
                                  )}
                                </span>
                              ))}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            )}

            {/* 탭 8: 페이지 범위 분석 결과 목록 */}
            {activeResultTab === 'pageRange' && (
              <div style={{ marginTop: '16px' }}>
                {pageRangeResults.length === 0 ? (
                  <div style={{ padding: '40px 0', textAlign: 'center', color: 'var(--text-muted)', fontSize: '13.5px' }}>
                    🎉 페이지 범위 분석 데이터가 없습니다. 먼저 파일 검증을 진행해 주세요.
                  </div>
                ) : (
                  <div style={{ overflowX: 'auto' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '13px' }}>
                      <thead>
                        <tr style={{ borderBottom: '2px solid var(--panel-border)', color: 'var(--text-secondary)' }}>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '220px' }}>파일명</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '120px' }}>시작페이지</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '120px' }}>최종 페이지</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '120px' }}>총 페이지수</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700 }}>참고 사항</th>
                        </tr>
                      </thead>
                      <tbody>
                        {pageRangeResults.map((p, idx) => (
                          <tr key={idx} style={{ borderBottom: '1px solid var(--panel-border)' }} className="table-row-hover">
                            <td style={{ padding: '12px 8px', fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: '220px' }} title={p.fileName}>
                              {p.fileName}
                            </td>
                            <td style={{ padding: '12px 8px', color: '#14b8a6', fontWeight: 700 }}>
                              {p.startPage} 페이지
                            </td>
                            <td style={{ padding: '12px 8px', color: '#14b8a6', fontWeight: 700 }}>
                              {p.endPage} 페이지
                            </td>
                            <td style={{ padding: '12px 8px', color: 'var(--text-primary)', fontWeight: 700 }}>
                              {p.totalSlides} 장 (슬라이드 수)
                            </td>
                            <td style={{ padding: '12px 8px', color: 'var(--text-secondary)', lineHeight: '1.5' }}>
                              시작 번호 {p.startPage}p부터 최종 번호 {p.endPage}p까지 스캔되었으며, 물리적인 총 수량은 {p.totalSlides}장으로 감지되었습니다.
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            )}

            {/* 탭 9: 맥 이미지 누락 검출 목록 */}
            {activeResultTab === 'macImages' && (
              <div style={{ marginTop: '16px' }}>
                {macImageResults.length === 0 ? (
                  <div style={{ padding: '40px 0', textAlign: 'center', color: 'var(--text-muted)', fontSize: '13.5px' }}>
                    🎉 맥 저장으로 인한 이미지 누락/깨짐 오류가 검출되지 않았습니다!
                  </div>
                ) : (
                  <div style={{ overflowX: 'auto' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '13px' }}>
                      <thead>
                        <tr style={{ borderBottom: '2px solid var(--panel-border)', color: 'var(--text-secondary)' }}>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '180px' }}>파일명</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '90px' }}>슬라이드(물리)</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '90px' }}>표시 페이지</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '150px' }}>오류 유형</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '100px' }}>관계 ID</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700, width: '150px' }}>대상 경로</th>
                          <th style={{ padding: '12px 8px', fontWeight: 700 }}>상세 가이드 설명</th>
                        </tr>
                      </thead>
                      <tbody>
                        {macImageResults.map((m, idx) => (
                          <tr key={idx} style={{ borderBottom: '1px solid var(--panel-border)' }} className="table-row-hover">
                            <td style={{ padding: '12px 8px', fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: '180px' }} title={m.fileName}>
                              {m.fileName}
                            </td>
                            <td style={{ padding: '12px 8px', color: 'var(--text-secondary)', fontWeight: 600 }}>
                              {m.slideNum} 순서
                            </td>
                            <td style={{ padding: '12px 8px', color: 'var(--text-primary)', fontWeight: 700 }}>
                              {m.displayPageNum} 페이지
                            </td>
                            <td style={{ padding: '12px 8px', color: '#10b981', fontWeight: 700 }}>
                              {m.errorType}
                            </td>
                            <td style={{ padding: '12px 8px', color: '#f59e0b', fontWeight: 700 }}>
                              {m.rId}
                            </td>
                            <td style={{ padding: '12px 8px', color: 'var(--text-secondary)', fontFamily: 'monospace', fontSize: '11px', lineBreak: 'anywhere' }}>
                              {m.targetPath}
                            </td>
                            <td style={{ padding: '12px 8px', color: 'var(--text-secondary)', lineHeight: '1.5' }}>
                              {m.desc}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            )}

          </div>

        </div>
      )}

      {/* 내장 오탈자 표준 사전 보기 모달 */}
      {showDictModal && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          width: '100%',
          height: '100%',
          background: 'rgba(0, 0, 0, 0.65)',
          backdropFilter: 'blur(8px)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 9999,
          padding: '20px'
        }}>
          <div style={{
            background: 'var(--panel-bg)',
            border: '1px solid var(--panel-border)',
            borderRadius: '16px',
            width: '100%',
            maxWidth: '650px',
            maxHeight: '85vh',
            display: 'flex',
            flexDirection: 'column',
            boxShadow: '0 20px 25px -5px rgba(0,0,0,0.3)',
            animation: 'fadeIn 0.2s ease-out'
          }}>
            {/* 헤더 */}
            <div style={{ padding: '18px 20px', borderBottom: '1px solid var(--panel-border)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 700, color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                <Sparkles size={16} color="#818cf8" />
                내장 표준 오탈자 사전
              </h3>
              <button
                onClick={() => setShowDictModal(false)}
                style={{ background: 'transparent', border: 'none', color: 'var(--text-secondary)', cursor: 'pointer', display: 'flex', alignItems: 'center' }}
              >
                <X size={18} />
              </button>
            </div>

            {/* 필터 탭 */}
            <div style={{ padding: '10px 20px', display: 'flex', gap: '6px', borderBottom: '1px solid rgba(255,255,255,0.03)', overflowX: 'auto' }}>
              {['전체', '맞춤법', '외래어 표기', '순화어/비즈니스', '띄어쓰기'].map(cat => (
                <button
                  key={cat}
                  onClick={() => setDictFilterCategory(cat)}
                  style={{
                    padding: '6px 12px',
                    borderRadius: '8px',
                    border: 'none',
                    fontSize: '12px',
                    fontWeight: 700,
                    cursor: 'pointer',
                    background: dictFilterCategory === cat ? 'linear-gradient(135deg, #a855f7, #6366f1)' : 'rgba(255,255,255,0.05)',
                    color: 'white',
                    transition: 'all 0.2s'
                  }}
                >
                  {cat}
                </button>
              ))}
            </div>

            {/* 본문 테이블 */}
            <div style={{ flex: 1, overflowY: 'auto', padding: '20px' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px', textAlign: 'left' }}>
                <thead>
                  <tr style={{ borderBottom: '2px solid var(--panel-border)', color: 'var(--text-secondary)' }}>
                    <th style={{ padding: '10px 8px', fontWeight: 600, width: '120px' }}>오타 단어</th>
                    <th style={{ padding: '10px 8px', fontWeight: 600, width: '120px' }}>교정 단어</th>
                    <th style={{ padding: '10px 8px', fontWeight: 600 }}>설명</th>
                    <th style={{ padding: '10px 8px', fontWeight: 600, width: '110px' }}>유형</th>
                  </tr>
                </thead>
                <tbody>
                  {Object.keys(TYPO_DICTIONARY)
                    .filter(key => dictFilterCategory === '전체' || TYPO_DICTIONARY[key].type === dictFilterCategory)
                    .map((typo, idx) => {
                      const item = TYPO_DICTIONARY[typo];
                      return (
                        <tr key={idx} style={{ borderBottom: '1px solid rgba(255,255,255,0.02)' }}>
                          <td style={{ padding: '10px 8px', color: '#f43f5e', fontWeight: 600 }}>{typo}</td>
                          <td style={{ padding: '10px 8px', color: 'var(--success-color)', fontWeight: 600 }}>{item.correction}</td>
                          <td style={{ padding: '10px 8px', color: 'var(--text-secondary)', fontSize: '12px' }}>{item.desc}</td>
                          <td style={{ padding: '10px 8px' }}>
                            <span style={{
                              background: item.type === '맞춤법' ? 'rgba(239, 68, 68, 0.1)' : 
                                          item.type === '외래어 표기' ? 'rgba(59, 130, 246, 0.1)' : 
                                          item.type === '띄어쓰기' ? 'rgba(245, 158, 11, 0.1)' : 'rgba(16, 185, 129, 0.1)',
                              color: item.type === '맞춤법' ? '#f87171' : 
                                     item.type === '외래어 표기' ? '#60a5fa' : 
                                     item.type === '띄어쓰기' ? '#fbbf24' : '#34d399',
                              padding: '2px 6px',
                              borderRadius: '4px',
                              fontSize: '11px',
                              fontWeight: 700
                            }}>
                              {item.type}
                            </span>
                          </td>
                        </tr>
                      );
                    })}
                </tbody>
              </table>
            </div>

            {/* 푸터 */}
            <div style={{ padding: '12px 20px', borderTop: '1px solid var(--panel-border)', textAlign: 'right' }}>
              <button
                onClick={() => setShowDictModal(false)}
                style={{
                  padding: '8px 16px',
                  background: 'rgba(255,255,255,0.08)',
                  border: 'none',
                  borderRadius: '8px',
                  color: 'var(--text-primary)',
                  fontSize: '12px',
                  fontWeight: 600,
                  cursor: 'pointer'
                }}
              >
                닫기
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}
