/**
 * 프로젝트 진행 시 도움이 되는 명언 및 유머 데이터 모듈
 * 스네이크 케이스(snake_case) 규칙 준수 및 유지보수 용이하도록 구성
 */

export const project_quotes = [
  {
    id: 1,
    category: "개발의 비밀",
    quote: "코드가 작동하지 않을 때는 이유를 모르겠고, 코드가 잘 작동할 때도 이유를 모르겠다.",
    author: "익명의 프로그래머",
    icon: "💻"
  },
  {
    id: 2,
    category: "배포의 철칙",
    quote: "금요일 오후 5시 배포는 '용기'라 부르지 않고 '재앙'이라 부릅니다.",
    author: "데브옵스 엔지니어의 수칙",
    icon: "🚨"
  },
  {
    id: 3,
    category: "감리의 법칙",
    quote: "감리 위원님의 '이 부분 보완이 필요해 보이네요'라는 말은 '주말 잘 가라'는 따뜻한 인사입니다.",
    author: "ISMP 감리 현장 전설",
    icon: "📋"
  },
  {
    id: 4,
    category: "설계서의 진실",
    quote: "완벽한 설계서란 존재하지 않는다. 다만 오픈 직전에 급히 수정된 설계서가 있을 뿐이다.",
    author: "수석 아키텍트",
    icon: "📐"
  },
  {
    id: 5,
    category: "PM의 지혜",
    quote: "성공적인 프로젝트란 버그가 0개인 프로젝트가 아니라, 고객이 수긍하고 결재한 프로젝트다.",
    author: "10년차 프로젝트 매니저",
    icon: "💼"
  },
  {
    id: 6,
    category: "품질의 가치",
    quote: "오탈자 하나가 시스템 전체의 신뢰도를 떨어뜨립니다. 꼼꼼한 검수는 최고의 기술입니다.",
    author: "건강한 프로젝트 감리팀",
    icon: "🔍"
  },
  {
    id: 7,
    category: "휴식의 힘",
    quote: "휴식도 개발 과정의 일부입니다. 산책하며 생각한 10분이 밤샘 3시간보다 낫습니다.",
    author: "건강한 프로젝트 힐링 센터",
    icon: "🌿"
  },
  {
    id: 8,
    category: "요구사항의 변주",
    quote: "제일 무서운 한마디: '스펙은 그대로인데요, 간단하게 이거 하나만 추가해 주세요.'",
    author: "프론트엔드 개발자",
    icon: "😱"
  },
  {
    id: 9,
    category: "기술 부채",
    quote: "기술 부채는 고금리 대출과 같다. 이자가 붙어 프로젝트가 마비되기 전에 조기 상환하라.",
    author: "클린 코드 격언",
    icon: "💳"
  },
  {
    id: 10,
    category: "버전 관리",
    quote: "커밋 메시지를 'fix', 'real_fix', 'final_fix2'로 쓰지 마세요. 미래의 자신이 울게 됩니다.",
    author: "Git 마스터",
    icon: "🌿"
  },
  {
    id: 11,
    category: "DB의 명언",
    quote: "가장 빠른 쿼리는 호출하지 않는 쿼리이고, 가장 안전한 데이터는 백업된 데이터다.",
    author: "DBA 튜닝 일기",
    icon: "🗄️"
  },
  {
    id: 12,
    category: "환경의 법칙",
    quote: "'내 컴퓨터에서는 잘 되는데요?'라는 말은 운전자가 '제 차에서는 브레이크 잘 잡히는데요?'하는 것과 같다.",
    author: "QA 팀장의 한마디",
    icon: "🖥️"
  },
  {
    id: 13,
    category: "팀워크",
    quote: "혼자 가면 빠르게 가지만, 함께 가면 멀리 간다. 팀원이 서로를 도울 때 프로젝트는 건강해집니다.",
    author: "아프리카 속담 & 애자일 가이드",
    icon: "🤝"
  },
  {
    id: 14,
    category: "커피와 에너지를 바꿀 때",
    quote: "오늘 마신 아메리카노는 내일의 에너지를 미리 당겨쓴 것입니다. 오늘은 정시 퇴근하고 푹 쉬세요!",
    author: "프로젝트 응원단",
    icon: "☕"
  },
  {
    id: 15,
    category: "AI 활용 팁",
    quote: "AI는 개발자를 대체하는 것이 아니라, AI를 잘 쓰는 개발자가 AI를 쓰지 않는 개발자를 대체합니다.",
    author: "인공지능 시대의 격언",
    icon: "✨"
  }
];

/**
 * 프로젝트 명언/유머 중 무작위 1개를 추첨하는 함수 (스네이크 케이스 사용)
 */
export function get_random_quote(previous_id = null) {
  const filtered_quotes = previous_id 
    ? project_quotes.filter(q => q.id !== previous_id) 
    : project_quotes;
  
  const random_index = Math.floor(Math.random() * filtered_quotes.length);
  return filtered_quotes[random_index];
}
