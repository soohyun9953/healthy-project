// styleConsistency.js
// 문서 내 문장 종결어미의 격식체 혼용(합니다체 vs 해요체)을 감지한다.
// 자동으로 문장을 재작성하면 어색하거나 문법적으로 틀린 결과가 나올 위험이 있으므로,
// 여기서는 "자동 교정"이 아니라 사람이 검토할 수 있도록 "검토 권고" 항목만 도출한다.

const HAMNIDA_PATTERN = /(습니다|합니다|됩니다|입니다|습니까|하십시오|바랍니다)\s*[.!?]?\s*$/;
const HAEYO_PATTERN = /(해요|이에요|예요|돼요|었어요|았어요|네요|거예요|아요|어요)\s*[.!?]?\s*$/;

// 한 줄(문단)에 여러 문장이 이어져 있는 경우(PDF/DOCX 등에서 흔함)를 대비해
// 문장 종결부호(./!/?) 기준으로 다시 나눈다. 부호가 없는 개조식 항목은 줄 전체를 하나의 문장으로 본다.
function splitIntoSentences(line) {
  const matches = line.match(/[^.!?]+[.!?]+(?=\s|$)|[^.!?]+$/g) || [];
  return matches.map((s) => s.trim()).filter(Boolean);
}

export function detect_style_inconsistency(text) {
  if (!text) return [];

  const lines = text.split('\n');
  let currentLoc = '1페이지';
  const classified = [];

  lines.forEach((line) => {
    const trimmed = line.trim();
    if (!trimmed) return;

    if (trimmed.match(/^\[?(?:슬라이드|페이지|장|섹션|Section)\s*\d+\]?/i)) {
      currentLoc = trimmed;
    }

    splitIntoSentences(trimmed).forEach((sentence) => {
      if (sentence.length < 6) return;
      if (HAMNIDA_PATTERN.test(sentence)) {
        classified.push({ loc: currentLoc, line: sentence, style: '합니다체' });
      } else if (HAEYO_PATTERN.test(sentence)) {
        classified.push({ loc: currentLoc, line: sentence, style: '해요체' });
      }
    });
  });

  const hamnidaCount = classified.filter((c) => c.style === '합니다체').length;
  const haeyoCount = classified.filter((c) => c.style === '해요체').length;

  // 두 문체가 모두 일정 수 이상 등장해야 "혼용"으로 판단 (한쪽이 0건이면 혼용이 아님)
  if (hamnidaCount === 0 || haeyoCount === 0) return [];

  const dominant = hamnidaCount >= haeyoCount ? '합니다체' : '해요체';
  const minorStyle = dominant === '합니다체' ? '해요체' : '합니다체';

  return classified
    .filter((c) => c.style === minorStyle)
    .map((c) => ({
      page: c.loc,
      sentence: c.line,
      currentStyle: c.style,
      dominantStyle: dominant
    }));
}
