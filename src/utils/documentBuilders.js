import PptxGenJS from 'pptxgenjs';
import * as docx from 'docx';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

// ── 5대 표준 컨설팅 장표 레이아웃 타입 상수 ──────────────────────────
export const LAYOUT_TYPES = {
    TYPE_A_TABLE: 'TYPE_A_TABLE',               // 표준 요약 박스 + 데이터 테이블형
    TYPE_B_BEFORE_AFTER: 'TYPE_B_BEFORE_AFTER', // As-Is vs To-Be 대칭 비교형
    TYPE_C_PROCESS: 'TYPE_C_PROCESS',           // 4단계 가로 프로세스 로드맵형
    TYPE_D_2COL_CARDS: 'TYPE_D_2COL_CARDS',     // 좌/우 2단 핵심 분석 프레임워크형
    TYPE_E_MATRIX: 'TYPE_E_MATRIX'              // 2x2 사분면 SWOT / 연관분석 매트릭스형
};

export const LAYOUT_LABELS = {
    TYPE_A_TABLE: 'A형: 요약+표준 테이블',
    TYPE_B_BEFORE_AFTER: 'B형: As-Is vs To-Be 대칭 비교',
    TYPE_C_PROCESS: 'C형: 4단계 가로 프로세스',
    TYPE_D_2COL_CARDS: 'D형: 좌우 2단 분석 카드',
    TYPE_E_MATRIX: 'E형: 2x2 SWOT/매트릭스'
};

// ── 1. PPTX 프레젠테이션 빌더 (PptxGenJS) ──────────────────────────
export async function buildPptxDeck(slidesData, tokens, options = {}) {
    const pptx = new PptxGenJS();
    const { colors, typography, pageLayout } = tokens;
    const is16x9 = (pageLayout.aspectRatio || '16:9') === '16:9';

    // 슬라이드 크기 설정
    if (is16x9) {
        pptx.layout = 'LAYOUT_16x9'; // 10 x 5.625 or 13.33 x 7.5
    } else {
        pptx.layout = 'LAYOUT_4x3';  // 10 x 7.5
    }

    const slideW = is16x9 ? 13.33 : 10.0;
    const slideH = 7.5;

    // 슬라이드 루프
    slidesData.forEach((slideItem, index) => {
        const slide = pptx.addSlide();
        const pageNum = index + 1;

        // A. [공통 마스터] 상단 브레드크럼 및 프로젝트/요구사항 헤더
        const breadcrumbText = [slideItem.chapter, slideItem.section, slideItem.subsection]
            .filter(Boolean)
            .join('  >  ') || 'Ⅰ. 현황 분석 및 시사점';

        // 상단 얇은 바/브레드크럼
        slide.addText(breadcrumbText, {
            x: 0.6,
            y: 0.35,
            w: 7.5,
            h: 0.3,
            fontFace: typography.fontFamily,
            fontSize: 9.5,
            color: colors.textSecondary,
            bold: true,
            valign: 'middle'
        });

        // 우측 상단 프로젝트명 및 요구사항 ID 뱃지
        const badgeText = slideItem.reqId ? `[${slideItem.reqId}]` : '';
        const projectInfo = `${pageLayout.projectName || ''}  ${badgeText}`.trim();
        slide.addText(projectInfo, {
            x: slideW - 5.5,
            y: 0.35,
            w: 4.9,
            h: 0.3,
            fontFace: typography.fontFamily,
            fontSize: 9.5,
            color: colors.primary,
            bold: true,
            align: 'right',
            valign: 'middle'
        });

        // 장표 제목 (Title)
        slide.addText(slideItem.title || '표준 장표 제목', {
            x: 0.6,
            y: 0.68,
            w: slideW - 1.2,
            h: 0.45,
            fontFace: typography.fontFamily,
            fontSize: typography.titleSize || 18,
            color: colors.primary,
            bold: true,
            valign: 'middle'
        });

        // B. [공통 마스터] 헤드메시지(Governing Message) 박스
        const headMessage = slideItem.headMessage || '본 장표는 표준화된 정보화전략계획(ISP) 체계에 따라 수립된 핵심 이행방안을 제시합니다.\n현행 시스템의 병목 현상을 진단하고 목표 모델로의 단계적 전환 방향을 정의합니다.';
        
        slide.addShape(pptx.ShapeType.rect, {
            x: 0.6,
            y: 1.18,
            w: slideW - 1.2,
            h: 0.82,
            fill: { color: colors.keyPointBg || 'EFF6FF' },
            line: { color: colors.secondary || '2563EB', width: 1.2 }
        });

        // 헤드메시지 좌측 굵은 포인트 바
        slide.addShape(pptx.ShapeType.rect, {
            x: 0.6,
            y: 1.18,
            w: 0.12,
            h: 0.82,
            fill: { color: colors.primary },
            line: { type: 'none' }
        });

        slide.addText(headMessage, {
            x: 0.85,
            y: 1.22,
            w: slideW - 1.6,
            h: 0.74,
            fontFace: typography.fontFamily,
            fontSize: typography.headMessageSize || 11.5,
            color: colors.textPrimary,
            bold: true,
            valign: 'middle',
            margin: 0
        });

        // C. 5대 레이아웃 타입별 본문 렌더링
        const layoutType = slideItem.layoutType || LAYOUT_TYPES.TYPE_A_TABLE;
        renderPptxLayoutContent(pptx, slide, layoutType, slideItem, tokens, slideW, slideH);

        // D. [공통 마스터] 하단 페이지 번호
        const pageStr = (pageLayout.pageNumberFormat || 'Ⅰ - {page}').replace('{page}', String(pageNum));
        slide.addText(pageStr, {
            x: slideW - 2.5,
            y: slideH - 0.45,
            w: 1.9,
            h: 0.3,
            fontFace: typography.fontFamily,
            fontSize: 9,
            color: colors.textSecondary,
            align: 'right',
            valign: 'middle'
        });

        // 하단 기준선
        slide.addShape(pptx.ShapeType.line, {
            x: 0.6,
            y: slideH - 0.48,
            w: slideW - 1.2,
            h: 0,
            line: { color: colors.borderColor, width: 0.6 }
        });
    });

    const blob = await pptx.write({ outputType: 'blob' });
    return blob;
}

// ── PPTX 5대 레이아웃 내부 렌더러 ────────────────────────────────
function renderPptxLayoutContent(pptx, slide, layoutType, item, tokens, slideW, slideH) {
    const { colors, typography } = tokens;
    const contentY = 2.15;
    const contentW = slideW - 1.2;
    const contentH = slideH - contentY - 0.65;

    switch (layoutType) {
        // ─────────────────────────────────────────────────────────────
        // 1) TYPE_A_TABLE: Key Point 요약 박스 + 표준 데이터 테이블
        case LAYOUT_TYPES.TYPE_A_TABLE: {
            const keyPoints = item.keyPoints || [
                '공공보건의료 데이터 통합 연계를 위한 표준 EMR 인터페이스 규격 수립 필요',
                '지방의료원 35개소 중 28개소가 개별 레거시 EMR을 운영하여 상호운용성 결여',
                '국가 표준 진료정보교류(FHIR) 준수율을 95% 이상으로 상향하여 데이터 고립 방지'
            ];

            // 상단 Key Point 요약 박스 (h: 1.0)
            slide.addShape(pptx.ShapeType.roundRect, {
                x: 0.6,
                y: contentY,
                w: contentW,
                h: 0.95,
                rectRadius: 0.06,
                fill: { color: colors.bgLight || 'F8FAFC' },
                line: { color: colors.borderColor, width: 1 }
            });

            slide.addText(
                keyPoints.map(kp => ({
                    text: `  •  ${kp}`,
                    options: { breakLine: true, fontSize: 10, color: colors.textPrimary, fontFace: typography.fontFamily }
                })),
                {
                    x: 0.8,
                    y: contentY + 0.08,
                    w: contentW - 0.4,
                    h: 0.8,
                    valign: 'top',
                    margin: 0
                }
            );

            // 하단 표준 데이터 테이블 (y: contentY + 1.1)
            const tableData = item.tableData || {
                headers: ['구분', '현행(As-Is) 수준', '목표(To-Be) 모델', '표준 적용 방안', '비고'],
                rows: [
                    ['데이터 표준', '의료원별 독자 코드체계 사용', '보건의료표준용어 및 FHIR 적용', '국가 표준 코드 매핑 테이블 구축', '필수준수'],
                    ['연계 방식', '수기 추출 및 배치 파일 전송', '실시간 API 게이트웨이 연계', 'RESTful API 및 EAI 표준화', '무중단'],
                    ['보안 거버넌스', '원내 자체 보안 기준 적용', '국가정보원 및 개인정보보호법 엄수', '가명정보 처리 및 전송구간 암호화', '규제대응'],
                    ['유지관리 체계', '개별 유지보수 계약 체결', '중앙 통합 관제 및 표준 SLA 적용', '통합 헬프데스크 및 기술지원', '비용절감']
                ]
            };

            const pptxRows = [
                tableData.headers.map(h => ({
                    text: h,
                    options: {
                        fill: { color: colors.tableHeaderBg },
                        color: colors.tableHeaderText,
                        bold: true,
                        align: 'center',
                        valign: 'middle',
                        fontFace: typography.fontFamily,
                        fontSize: typography.tableTextSize || 9.5
                    }
                })),
                ...tableData.rows.map((row, rIdx) =>
                    row.map((cell, cIdx) => ({
                        text: cell,
                        options: {
                            fill: { color: rIdx % 2 === 1 ? colors.bgLight : colors.white },
                            color: colors.textPrimary,
                            align: cIdx === 0 ? 'center' : 'left',
                            valign: 'middle',
                            fontFace: typography.fontFamily,
                            fontSize: typography.tableTextSize || 9
                        }
                    }))
                )
            ];

            slide.addTable(pptxRows, {
                x: 0.6,
                y: contentY + 1.15,
                w: contentW,
                rowH: 0.6,
                border: { pt: 0.5, color: colors.borderColor },
                autoPage: false
            });
            break;
        }

        // ─────────────────────────────────────────────────────────────
        // 2) TYPE_B_BEFORE_AFTER: As-Is vs To-Be 대칭 비교 + 하단 기대효과 바
        case LAYOUT_TYPES.TYPE_B_BEFORE_AFTER: {
            const cardW = (contentW - 0.7) / 2;
            const cardH = contentH - 0.9;

            // 좌측 As-Is 카드
            slide.addShape(pptx.ShapeType.roundRect, {
                x: 0.6,
                y: contentY,
                w: cardW,
                h: cardH,
                rectRadius: 0.08,
                fill: { color: 'FFF5F5' }, // Red 50
                line: { color: colors.danger || 'DC2626', width: 1.2 }
            });

            slide.addText('현행 업무 및 시스템 한계점 (As-Is)', {
                x: 0.8,
                y: contentY + 0.15,
                w: cardW - 0.4,
                h: 0.35,
                fontFace: typography.fontFamily,
                fontSize: 12,
                color: colors.danger || 'DC2626',
                bold: true
            });

            const asIsBullets = item.asIsBullets || [
                '의료원별 데이터 단절로 환자 전원 시 진료기록 수기 복사 전달',
                '수기 통계 집계로 인한 데이터 집계 지연(평균 3주 소요)',
                '의료 AI 모델을 현장에 적용할 수 있는 표준 인프라 부재',
                '응급 이송 시 의료원 간 수용 가능 병상 실시간 확인 불가'
            ];

            slide.addText(
                asIsBullets.map(b => ({
                    text: `  [문제점]  ${b}`,
                    options: { breakLine: true, fontSize: 9.5, color: colors.textPrimary, fontFace: typography.fontFamily }
                })),
                { x: 0.8, y: contentY + 0.6, w: cardW - 0.4, h: cardH - 0.8, valign: 'top' }
            );

            // 중앙 화살표 기호
            slide.addText('➔', {
                x: 0.6 + cardW,
                y: contentY + cardH / 2 - 0.3,
                w: 0.7,
                h: 0.6,
                fontSize: 24,
                color: colors.secondary,
                align: 'center',
                valign: 'middle'
            });

            // 우측 To-Be 카드
            slide.addShape(pptx.ShapeType.roundRect, {
                x: 0.6 + cardW + 0.7,
                y: contentY,
                w: cardW,
                h: cardH,
                rectRadius: 0.08,
                fill: { color: 'F0FDF4' }, // Green 50
                line: { color: colors.success || '16A34A', width: 1.2 }
            });

            slide.addText('목표 혁신 모델 및 개선 방향 (To-Be)', {
                x: 0.6 + cardW + 0.9,
                y: contentY + 0.15,
                w: cardW - 0.4,
                h: 0.35,
                fontFace: typography.fontFamily,
                fontSize: 12,
                color: colors.success || '16A34A',
                bold: true
            });

            const toBeBullets = item.toBeBullets || [
                '국가 보건의료 표준 FHIR 기반 실시간 진료정보교류 자동화',
                '중앙 데이터 레이크 구축으로 전국 공공의료 통계 실시간 시각화',
                '클라우드 기반 AI CDSS(진료지원) 알고리즘 표준 탑재 및 배포',
                '실시간 응급의료 관제 대시보드 구축으로 골든타임 40분 단축'
            ];

            slide.addText(
                toBeBullets.map(b => ({
                    text: `  [혁신안]  ${b}`,
                    options: { breakLine: true, fontSize: 9.5, color: colors.textPrimary, fontFace: typography.fontFamily }
                })),
                { x: 0.6 + cardW + 0.9, y: contentY + 0.6, w: cardW - 0.4, h: cardH - 0.8, valign: 'top' }
            );

            // 하단 기대효과 (Delta) 요약 바
            slide.addShape(pptx.ShapeType.rect, {
                x: 0.6,
                y: contentY + cardH + 0.15,
                w: contentW,
                h: 0.65,
                fill: { color: colors.keyPointBg },
                line: { color: colors.secondary, width: 1 }
            });

            const deltaText = item.deltaText || '🎯 [기대성과] 환자 재검사율 28% 감소 | 응급환자 골든타임 확보율 35% 향상 | 연간 행정비용 45억원 절감 달성';
            slide.addText(deltaText, {
                x: 0.8,
                y: contentY + cardH + 0.17,
                w: contentW - 0.4,
                h: 0.6,
                fontFace: typography.fontFamily,
                fontSize: 10.5,
                color: colors.primary,
                bold: true,
                valign: 'middle'
            });
            break;
        }

        // ─────────────────────────────────────────────────────────────
        // 3) TYPE_C_PROCESS: 단계별 4단계 가로 프로세스
        case LAYOUT_TYPES.TYPE_C_PROCESS: {
            const steps = item.steps || [
                { num: 'Step 1', title: '기반 분석 및 진단', period: '1~2개월', desc: '• 레거시 EMR 연계성 평가\n• 개인정보 규제 요건 분석\n• 표준화 대상 항목 도출' },
                { num: 'Step 2', title: '목표 아키텍처 설계', period: '3~4개월', desc: '• 클라우드 인프라 설계\n• FHIR 데이터 모델 정의\n• 보안 거버넌스 수립' },
                { num: 'Step 3', title: '플랫폼 구축 및 실증', period: '5~6개월', desc: '• AI CDSS 알고리즘 탑재\n• 응급관제 대시보드 구축\n• 5개 선도병원 시범 실증' },
                { num: 'Step 4', title: '전국 확산 및 고도화', period: '7~8개월', desc: '• 전국 35개 의료원 연계\n• 운영 자동화 및 감리\n• 최종 산출물 완성' }
            ];

            const stepW = (contentW - 0.9) / 4;
            const stepH = contentH;

            steps.forEach((s, sIdx) => {
                const stepX = 0.6 + sIdx * (stepW + 0.3);

                // 상단 스텝 헤더 배너
                slide.addShape(pptx.ShapeType.roundRect, {
                    x: stepX,
                    y: contentY,
                    w: stepW,
                    h: 0.7,
                    rectRadius: 0.05,
                    fill: { color: sIdx === 0 ? colors.primary : sIdx === 1 ? colors.secondary : sIdx === 2 ? colors.accent : '059669' },
                    line: { type: 'none' }
                });

                slide.addText(`${s.num} : ${s.title}`, {
                    x: stepX + 0.1,
                    y: contentY + 0.08,
                    w: stepW - 0.2,
                    h: 0.35,
                    fontFace: typography.fontFamily,
                    fontSize: 10.5,
                    color: 'FFFFFF',
                    bold: true,
                    align: 'center'
                });

                slide.addText(s.period, {
                    x: stepX + 0.1,
                    y: contentY + 0.38,
                    w: stepW - 0.2,
                    h: 0.25,
                    fontFace: typography.fontFamily,
                    fontSize: 8.5,
                    color: 'E2E8F0',
                    align: 'center'
                });

                // 본문 카드
                slide.addShape(pptx.ShapeType.roundRect, {
                    x: stepX,
                    y: contentY + 0.8,
                    w: stepW,
                    h: stepH - 0.8,
                    rectRadius: 0.05,
                    fill: { color: colors.bgLight },
                    line: { color: colors.borderColor, width: 1 }
                });

                slide.addText(s.desc, {
                    x: stepX + 0.15,
                    y: contentY + 0.95,
                    w: stepW - 0.3,
                    h: stepH - 1.1,
                    fontFace: typography.fontFamily,
                    fontSize: 9.5,
                    color: colors.textPrimary,
                    valign: 'top',
                    lineSpacingMultiple: 1.2
                });

                // 스텝 사이 화살표
                if (sIdx < 3) {
                    slide.addText('▶', {
                        x: stepX + stepW,
                        y: contentY + 0.2,
                        w: 0.3,
                        h: 0.35,
                        fontSize: 11,
                        color: colors.secondary,
                        align: 'center',
                        valign: 'middle'
                    });
                }
            });
            break;
        }

        // ─────────────────────────────────────────────────────────────
        // 4) TYPE_D_2COL_CARDS: 좌우 2단 핵심 분석 프레임워크 카드
        case LAYOUT_TYPES.TYPE_D_2COL_CARDS: {
            const cardW = (contentW - 0.4) / 2;
            const leftCard = item.leftCard || {
                title: '전략 방향 1 : 공공의료 데이터 거버넌스 및 연계 기반',
                items: [
                    '전국 지방의료원 및 국립병원 EMR 표준 연계 인터페이스(API) 단일화',
                    '보건의료 마이데이터 연계 표준화 및 개인정보보호법 제28조 가명처리 엄수',
                    '국가 의료정보 교류망을 통한 전원 환자 영상·검사결과 무중단 연계'
                ]
            };
            const rightCard = item.rightCard || {
                title: '전략 방향 2 : 지능형 AI 임상결정지원(CDSS) 서비스 구축',
                items: [
                    '급성 중증 환자 조기 예측 및 패혈증·심정지 위험도 실시간 모니터링',
                    '지방의료원 전공의 및 전문의 부족을 보완하는 AI 흉부 X-ray 자동 판독',
                    '응급의료센터 간 병상 자원 및 의료진 실시간 배정 지능화'
                ]
            };

            [leftCard, rightCard].forEach((card, idx) => {
                const cX = 0.6 + idx * (cardW + 0.4);

                // 헤더
                slide.addShape(pptx.ShapeType.roundRect, {
                    x: cX,
                    y: contentY,
                    w: cardW,
                    h: 0.6,
                    rectRadius: 0.06,
                    fill: { color: idx === 0 ? colors.primary : colors.secondary },
                    line: { type: 'none' }
                });

                slide.addText(card.title, {
                    x: cX + 0.2,
                    y: contentY + 0.1,
                    w: cardW - 0.4,
                    h: 0.4,
                    fontFace: typography.fontFamily,
                    fontSize: 11,
                    color: 'FFFFFF',
                    bold: true,
                    valign: 'middle'
                });

                // 본문
                slide.addShape(pptx.ShapeType.roundRect, {
                    x: cX,
                    y: contentY + 0.7,
                    w: cardW,
                    h: contentH - 0.7,
                    rectRadius: 0.06,
                    fill: { color: colors.bgLight },
                    line: { color: colors.borderColor, width: 1 }
                });

                slide.addText(
                    card.items.map((it, itIdx) => ({
                        text: `  [전략과제 1-${itIdx + 1}]  ${it}`,
                        options: { breakLine: true, fontSize: 10, color: colors.textPrimary, fontFace: typography.fontFamily }
                    })),
                    { x: cX + 0.25, y: contentY + 0.9, w: cardW - 0.5, h: contentH - 1.1, valign: 'top' }
                );
            });
            break;
        }

        // ─────────────────────────────────────────────────────────────
        // 5) TYPE_E_MATRIX: 2x2 사분면 SWOT / 연관분석 매트릭스
        case LAYOUT_TYPES.TYPE_E_MATRIX: {
            const quadW = (contentW - 0.3) / 2;
            const quadH = (contentH - 0.3) / 2;

            const matrixData = item.matrixData || [
                { title: '강점 (Strength)', color: '1E3A8A', items: ['전국 공공의료원 네트워크 기구축', '풍부한 임상 코호트 데이터 보유'] },
                { title: '약점 (Weakness)', color: 'DC2626', items: ['의료원별 EMR 버전 상이 및 비표준', '원내 전산 인력 및 AI 전문성 부족'] },
                { title: '기회 (Opportunity)', color: '16A34A', items: ['정부 디지털 헬스케어 육성 정책', '클라우드 도입 가이드라인 완화'] },
                { title: '위협 (Threat)', color: 'D97706', items: ['개인정보 유출 등 보안 규제 리스크', '의료계 이해관계자 간 반발 가능성'] }
            ];

            matrixData.forEach((m, idx) => {
                const col = idx % 2;
                const row = Math.floor(idx / 2);
                const qX = 0.6 + col * (quadW + 0.3);
                const qY = contentY + row * (quadH + 0.3);

                slide.addShape(pptx.ShapeType.roundRect, {
                    x: qX,
                    y: qY,
                    w: quadW,
                    h: quadH,
                    rectRadius: 0.06,
                    fill: { color: colors.bgLight },
                    line: { color: m.color, width: 1.5 }
                });

                // 사분면 타이틀 뱃지
                slide.addShape(pptx.ShapeType.rect, {
                    x: qX,
                    y: qY,
                    w: quadW,
                    h: 0.45,
                    fill: { color: m.color },
                    line: { type: 'none' }
                });

                slide.addText(m.title, {
                    x: qX + 0.15,
                    y: qY + 0.05,
                    w: quadW - 0.3,
                    h: 0.35,
                    fontFace: typography.fontFamily,
                    fontSize: 10.5,
                    color: 'FFFFFF',
                    bold: true,
                    valign: 'middle'
                });

                slide.addText(
                    m.items.map(it => ({
                        text: `  •  ${it}`,
                        options: { breakLine: true, fontSize: 9.5, color: colors.textPrimary, fontFace: typography.fontFamily }
                    })),
                    { x: qX + 0.15, y: qY + 0.55, w: quadW - 0.3, h: quadH - 0.7, valign: 'top' }
                );
            });
            break;
        }

        default:
            break;
    }
}

// ── 2. Word (DOCX) 보고서 빌더 ────────────────────────────────────
export async function buildDocxDocument(slidesData, tokens, options = {}) {
    const { colors, typography, pageLayout } = tokens;

    const sections = [];

    slidesData.forEach((slideItem, index) => {
        const pageNum = index + 1;
        const breadcrumbText = [slideItem.chapter, slideItem.section, slideItem.subsection]
            .filter(Boolean)
            .join(' > ') || 'Ⅰ. 현황 분석 및 시사점';

        const children = [
            // 상단 브레드크럼 & 요구사항 뱃지
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: breadcrumbText,
                        bold: true,
                        size: 19, // 9.5pt (half-points)
                        color: colors.textSecondary
                    }),
                    new docx.TextRun({
                        text: `\t${pageLayout.projectName} [${slideItem.reqId || 'CNR-001'}]`,
                        bold: true,
                        size: 19,
                        color: colors.primary
                    })
                ],
                tabStops: [
                    {
                        type: docx.TabStopType.RIGHT,
                        position: docx.TabStopPosition.MAX
                    }
                ],
                spacing: { after: 120 }
            }),

            // 장표 대제목
            new docx.Paragraph({
                text: slideItem.title || '표준 장표 제목',
                heading: docx.HeadingLevel.HEADING_1,
                spacing: { after: 200 }
            }),

            // 거버닝 헤드메시지 콜아웃 박스
            new docx.Table({
                width: { size: 100, type: docx.WidthType.PERCENTAGE },
                rows: [
                    new docx.TableRow({
                        children: [
                            new docx.TableCell({
                                children: [
                                    new docx.Paragraph({
                                        children: [
                                            new docx.TextRun({
                                                text: '💡 [핵심 방향] ',
                                                bold: true,
                                                color: colors.primary
                                            }),
                                            new docx.TextRun({
                                                text: slideItem.headMessage || '본 장표는 표준화된 정보화전략계획(ISP) 체계에 따라 수립된 핵심 이행방안을 제시합니다.',
                                                bold: true,
                                                color: colors.textPrimary
                                            })
                                        ]
                                    })
                                ],
                                shading: { fill: colors.keyPointBg || 'EFF6FF' },
                                borders: {
                                    left: { style: docx.BorderStyle.SINGLE, size: 24, color: colors.primary },
                                    top: { style: docx.BorderStyle.SINGLE, size: 4, color: colors.borderColor },
                                    right: { style: docx.BorderStyle.SINGLE, size: 4, color: colors.borderColor },
                                    bottom: { style: docx.BorderStyle.SINGLE, size: 4, color: colors.borderColor }
                                },
                                margins: { top: 140, bottom: 140, left: 180, right: 180 }
                            })
                        ]
                    })
                ]
            }),

            new docx.Paragraph({ text: '', spacing: { after: 200 } }),

            // 본문 내용 렌더링
            ...renderDocxLayoutChildren(slideItem, tokens),

            // 하단 페이지 나눔 (마지막 장표 제외)
            ...(index < slidesData.length - 1 ? [new docx.Paragraph({ children: [new docx.PageBreak()] })] : [])
        ];

        sections.push(...children);
    });

    const doc = new docx.Document({
        sections: [
            {
                properties: {
                    page: {
                        margin: { top: 1200, bottom: 1200, left: 1200, right: 1200 }
                    }
                },
                children: sections
            }
        ]
    });

    const blob = await docx.Packer.toBlob(doc);
    return blob;
}

// ── DOCX 레이아웃별 Paragraph/Table 생성기 ────────────────────────
function renderDocxLayoutChildren(slideItem, tokens) {
    const { colors } = tokens;
    const layoutType = slideItem.layoutType || LAYOUT_TYPES.TYPE_A_TABLE;
    const children = [];

    if (layoutType === LAYOUT_TYPES.TYPE_A_TABLE) {
        const tableData = slideItem.tableData || {
            headers: ['구분', '현행(As-Is) 수준', '목표(To-Be) 모델', '표준 적용 방안'],
            rows: [
                ['데이터 표준', '의료원별 독자 코드체계 사용', '보건의료표준용어 및 FHIR 적용', '국가 표준 코드 매핑 테이블 구축'],
                ['연계 방식', '수기 추출 및 배치 파일 전송', '실시간 API 게이트웨이 연계', 'RESTful API 및 EAI 표준화'],
                ['보안 거버넌스', '원내 자체 보안 기준 적용', '국가정보원 및 개인정보보호법 엄수', '가명정보 처리 및 전송구간 암호화']
            ]
        };

        const docxRows = [
            new docx.TableRow({
                children: tableData.headers.map(h => new docx.TableCell({
                    children: [new docx.Paragraph({ text: h, alignment: docx.AlignmentType.CENTER, bold: true })],
                    shading: { fill: colors.tableHeaderBg },
                    margins: { top: 100, bottom: 100, left: 100, right: 100 }
                }))
            }),
            ...tableData.rows.map((row, rIdx) => new docx.TableRow({
                children: row.map((cell, cIdx) => new docx.TableCell({
                    children: [new docx.Paragraph({ text: cell, alignment: cIdx === 0 ? docx.AlignmentType.CENTER : docx.AlignmentType.LEFT })],
                    shading: { fill: rIdx % 2 === 1 ? colors.bgLight : 'FFFFFF' },
                    margins: { top: 80, bottom: 80, left: 100, right: 100 }
                }))
            }))
        ];

        children.push(new docx.Table({
            width: { size: 100, type: docx.WidthType.PERCENTAGE },
            rows: docxRows
        }));
    } else {
        // 기타 레이아웃(B, C, D, E) 불릿 포인트 리스트 변환
        const bullets = slideItem.keyPoints || slideItem.toBeBullets || [
            '핵심 세부 이행방안 1 : 표준 인터페이스 규격 및 클라우드 아키텍처 수립',
            '핵심 세부 이행방안 2 : 실시간 응급환자 이송 관제 모니터링 시스템 구축',
            '핵심 세부 이행방안 3 : 국가 표준 보건의료용어체계 매핑 및 품질 검증 완결'
        ];

        bullets.forEach(b => {
            children.push(new docx.Paragraph({
                text: `• ${b}`,
                bullet: { level: 0 },
                spacing: { after: 120 }
            }));
        });
    }

    return children;
}

// ── 3. Excel (XLSX) 워크북 빌더 (ExcelJS) ─────────────────────────
export async function buildXlsxWorkbook(slidesData, tokens, options = {}) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = '건강한 프로젝트 (Uniform Document Generator)';
    workbook.created = new Date();

    const { colors, pageLayout } = tokens;

    // 시트 1: 산출물 목차 및 장표 명세서
    const sheet1 = workbook.addWorksheet('장표_스캐폴딩_목차명세');
    sheet1.views = [{ showGridLines: true }];

    sheet1.columns = [
        { header: 'No.', width: 8 },
        { header: '대장', width: 22 },
        { header: '중절', width: 22 },
        { header: '소항', width: 22 },
        { header: '장표 제목', width: 35 },
        { header: '요구사항 ID', width: 15 },
        { header: '레이아웃 타입', width: 25 },
        { header: '핵심 헤드메시지', width: 45 },
        { header: '담당자', width: 14 }
    ];

    const headRow = sheet1.getRow(1);
    headRow.height = 28;
    headRow.eachCell(cell => {
        cell.font = { name: '맑은 고딕', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${colors.primary}` } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    slidesData.forEach((s, idx) => {
        const r = sheet1.addRow([
            idx + 1,
            s.chapter || '-',
            s.section || '-',
            s.subsection || '-',
            s.title || '-',
            s.reqId || '-',
            LAYOUT_LABELS[s.layoutType] || s.layoutType || '-',
            s.headMessage || '-',
            s.owner || '-'
        ]);
        r.height = 24;
        r.alignment = { vertical: 'middle' };
        r.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };
        r.getCell(6).alignment = { vertical: 'middle', horizontal: 'center' };
    });

    // 엑셀 버퍼 출력
    const buffer = await workbook.xlsx.writeBuffer();
    return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
