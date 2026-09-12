// 오탈자 내장 사전 정의
export const TYPO_DICTIONARY = {
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
  '새호': { correction: '새로', desc: '새로의 오타 표기입니다.', type: '맞춤법' },
  '맞쳐': { correction: '맞춰', desc: '맞추어의 준말은 맞춰가 표준어입니다.', type: '맞춤법' },
  '어떻해': { correction: '어떡해', desc: '어떻게 해의 줄임말은 어떡해입니다.', type: '맞춤법' },
  
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
  '프로잭트': { correction: '프로젝트', desc: '프로젝트(Project)의 오타 표기입니다.', type: '순화어/비즈니스' },
  '프로젝터': { correction: '프로젝트', desc: '문맥상 프로젝트(Project)의 오타 표기일 수 있습니다.', type: '순화어/비즈니스' },
  '익일': { correction: '다음 날', desc: '공공언어 바로쓰기에 따라 익일보다는 다음 날(이튿날)로 순화하여 작성할 것을 권장합니다.', type: '순화어/비즈니스' },
  '금회': { correction: '이번', desc: '공공언어 바로쓰기에 따라 금회보다는 이번으로 순화하여 작성할 것을 권장합니다.', type: '순화어/비즈니스' },
  '납기': { correction: '마감일', desc: '납기보다는 마감일 또는 납품기한으로 순화할 것을 권장합니다.', type: '순화어/비즈니스' }
};

// 동사/형용사 어미 변화 자동 확장 규칙 생성기
export function generate_conjugation_rules(dict) {
  const extended_dict = { ...dict };

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
    if (key.endsWith('하다') && item.correction.endsWith('하다')) {
      const typo_stem = key.substring(0, key.length - 2);
      const corr_stem = item.correction.substring(0, item.correction.length - 2);
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
}

// 오탈자 문맥 매칭 검증기 (부분 문자열 오탐 방지)
export function validate_typo_match(fullText, typo) {
  if (!fullText || !typo) return false;
  
  if (typo === '와부') {
    const regex = /(?:^|[^\w가-힣])와부(?:[^\w가-힣]|$)/;
    return regex.test(fullText);
  }
  
  return true;
}

// "용어 사전" 탭에 업로드/입력된 지침·용어집 텍스트에서 화살표(→, ->, ⇒, ➜, =>) 형식으로
// 명시된 "오류 → 올바른 표현" 쌍만 안전하게 추출한다. 애매한 서술형 문장은 무시하여 오탐을 방지하고,
// 명확한 지침 용어만 1단계 사전 스캔(API 키 없이도 동작)에 즉시 반영되도록 한다.
export function extract_dict_pairs_from_glossary(glossaryText) {
  const result = {};
  if (!glossaryText) return result;

  const lines = String(glossaryText).split('\n');
  const ARROW_PATTERN = /^(.{1,40}?)\s*(?:→|->|⇒|➜|=>)\s*(.{1,40}?)$/;

  lines.forEach((line) => {
    const trimmed = line.trim();
    if (!trimmed) return;
    const match = trimmed.match(ARROW_PATTERN);
    if (!match) return;

    const wrong = match[1].trim().replace(/^[-•·*\d.)\s]+/, '');
    const correct = match[2].trim();
    if (!wrong || !correct || wrong === correct) return;
    // 문장 전체가 아닌 짧은 용어 쌍만 신뢰 (오탐 방지)
    if (wrong.length > 30 || correct.length > 30) return;

    result[wrong] = {
      correction: correct,
      desc: `지침(용어 사전) 등록 용어: '${wrong}' → '${correct}'`,
      type: '지침 용어'
    };
  });

  return result;
}

// 텍스트 전체에서 사전 기반 오탈자를 100% 전수 검출하는 함수
export function extract_dictionary_typos(text, customDict = {}) {
  if (!text) return [];
  const mergedDict = generate_conjugation_rules({ ...TYPO_DICTIONARY, ...customDict });
  const lines = text.split('\n');
  const detected = [];
  const seen = new Set();
  let currentLoc = '1페이지';

  lines.forEach((line) => {
    const trimmed = line.trim();
    if (!trimmed) return;

    if (trimmed.match(/^\[?(?:슬라이드|페이지|장|섹션|Section)\s*\d+\]?/i)) {
      currentLoc = trimmed;
    }

    Object.keys(mergedDict).forEach(typo => {
      if (trimmed.includes(typo) && validate_typo_match(trimmed, typo)) {
        const info = mergedDict[typo];
        const signature = `${currentLoc}_${typo}_${trimmed}`;
        if (!seen.has(signature)) {
          seen.add(signature);
          detected.push({
            page: currentLoc,
            originalText: trimmed,
            correction: trimmed.replace(new RegExp(typo, 'g'), info.correction),
            errorType: `[${info.type || '표현 품질'}] ${info.desc || `${typo} ➜ ${info.correction}`}`
          });
        }
      }
    });
  });

  return detected;
}
