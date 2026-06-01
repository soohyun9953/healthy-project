import React, { useState, useRef, useEffect } from 'react';
import { 
    Presentation, FileSpreadsheet, Upload, X, Loader2, Info, Settings, 
    Download, Play, FileDown, Sparkles, MousePointer2, CheckCircle2,
    ChevronDown, ChevronUp, Layers, Sliders, ToggleLeft, ToggleRight, Box
} from 'lucide-react';
import { 
    parseExcelData, generatePptFromTemplate, processPptBatch, 
    addSmartAnimationsToPpt, saveFileWithLocationPicker, getPptSlideCount 
} from '../utils/pptExporter';
import { convertPptxToHwpx, fusePptToHwpxTemplate } from '../utils/hwpxConverter';
import JSZip from 'jszip';

export default function PptGenerator() {
    const [activeTab, setActiveTab] = useState('excel_mapping'); // 'excel_mapping' or 'batch_edit'
    
    // 엑셀 매핑 관련 State
    const [excelFile, setExcelFile] = useState(null);
    const [pptTemplate, setPptTemplate] = useState(null);
    const [excelDataPreview, setExcelDataPreview] = useState(null);
    const [templateLabel, setTemplateLabel] = useState('');
    const [isParsing, setIsParsing] = useState(false);
    const [isGenerating, setIsGenerating] = useState(false);
    const [generationMode, setGenerationMode] = useState('single');
    const [chunkSize, setChunkSize] = useState(10);
    const [isDraggingExcel, setIsDraggingExcel] = useState(false);
    const [isDraggingTemplate, setIsDraggingTemplate] = useState(false);
    const excelInputRef = useRef(null);
    const pptInputRef = useRef(null);

    // PPT 일괄 편집 (단어 수정 + 디자인 변경) 관련 State
    const [batchPptFiles, setBatchPptFiles] = useState([]);
    const [replaceRules, setReplaceRules] = useState(() => {
        try {
            return localStorage.getItem('ppt_replace_rules') || '';
        } catch (e) {
            return '';
        }
    });
    const [fontRules, setFontRules] = useState('');
    const [fontSize, setFontSize] = useState('');
    const [applyDesignChecked, setApplyDesignChecked] = useState(false);
    const [applyTableDesignChecked, setApplyTableDesignChecked] = useState(false);
    const [applyFirstRowHeaderStyle, setApplyFirstRowHeaderStyle] = useState(true); // 옵션 E 하위 옵션: 첫 행 특별 포맷팅 적용 여부
    const [designTargetText, setDesignTargetText] = useState('');
    const [applySpecialCharClean, setApplySpecialCharClean] = useState(false); // 옵션 F 활성화 여부
    const [replaceNbs, setReplaceNbs] = useState(true); // 하위 옵션 1: NBS 일반 공백 변환
    const [unifyBullets, setUnifyBullets] = useState(true); // 하위 옵션 2: 중간점 통일
    const [clean_vertical_tab, set_clean_vertical_tab] = useState(true); // 하위 옵션 3: 세로 탭 품질검토 오류 수정
    const [add_title_page_numbers, set_add_title_page_numbers] = useState(false); // 옵션 G: 동일 제목 일련번호 자동 추가
    const [add_space_before_parenthesis, set_add_space_before_parenthesis] = useState(true); // 옵션 G 하위 옵션: 제목 뒤 괄호 앞 공백 추가
    const [textColorRules, setTextColorRules] = useState(() => {
        try {
            return localStorage.getItem('ppt_textcolor_rules') || '';
        } catch (e) {
            return '';
        }
    }); // 옵션 H: 글자 색상 매핑 변경
    const [preventWordWrap, setPreventWordWrap] = useState(() => {
        try {
            return localStorage.getItem('ppt_prevent_word_wrap') === 'true';
        } catch (e) {
            return false;
        }
    }); // 옵션 I: 단락 한글 단어 잘림 방지
    const [isProcessingBatch, setIsProcessingBatch] = useState(false);
    const [isDraggingBatch, setIsDraggingBatch] = useState(false);
    const [batchReport, setBatchReport] = useState([]); // 📊 일괄 편집 결과 상세 피드백 리포트 리스트

    // PPT 스마트 애니메이션 관련 State
    const [animationPptFile, setAnimationPptFile] = useState(null);
    const [isProcessingAnimation, setIsProcessingAnimation] = useState(false);
    const [isDraggingAnimation, setIsDraggingAnimation] = useState(false);
    const [animationType, setAnimationType] = useState('transition'); // Default to transition
    const [useGrouping, setUseGrouping] = useState(true); 
    const [slideAnimations, setSlideAnimations] = useState([]); // [{ enabled: true, type: 'transition', useGrouping: true }, ...]
    const animInputRef = useRef(null);

    // PPT ➜ HWPX 스마트 변환 관련 State
    const [hwpxPptFile, setHwpxPptFile] = useState(null);
    const [isConvertingHwpx, setIsConvertingHwpx] = useState(false);
    const [isDraggingHwpx, setIsDraggingHwpx] = useState(false);
    const [hwpxResultStats, setHwpxResultStats] = useState(null);
    const hwpxInputRef = useRef(null);

    // PPT ➜ HWPX 양식 융합 관련 State
    const [hwpxFusionPptFile, setHwpxFusionPptFile] = useState(null);
    const [hwpxFusionTemplateFile, setHwpxFusionTemplateFile] = useState(null);
    const [isFusingHwpx, setIsFusingHwpx] = useState(false);
    const [isDraggingFusionPpt, setIsDraggingFusionPpt] = useState(false);
    const [isDraggingFusionTemplate, setIsDraggingFusionTemplate] = useState(false);
    const [fusionMergeMode, setFusionMergeMode] = useState(true); // true: 하나의 HWPX 병합, false: ZIP 압축 분할
    const [fusionResultStats, setFusionResultStats] = useState(null);
    const fusionPptInputRef = useRef(null);
    const fusionTemplateInputRef = useRef(null);

    const [errorMsg, setErrorMsg] = useState(null);
    const [successMsg, setSuccessMsg] = useState(null);

    useEffect(() => {
        try {
            localStorage.setItem('ppt_replace_rules', replaceRules);
        } catch (e) {
            console.error('Error saving replaceRules to localStorage:', e);
        }
    }, [replaceRules]);

    useEffect(() => {
        try {
            localStorage.setItem('ppt_textcolor_rules', textColorRules);
        } catch (e) {
            console.error('Error saving textColorRules to localStorage:', e);
        }
    }, [textColorRules]);

    useEffect(() => {
        try {
            localStorage.setItem('ppt_prevent_word_wrap', preventWordWrap);
        } catch (e) {
            console.error('Error saving preventWordWrap to localStorage:', e);
        }
    }, [preventWordWrap]);

    // 엑셀 매핑 핸들러들
    const processExcelFile = async (file) => {
        if (!file) return;
        setErrorMsg(null);
        setSuccessMsg(null);
        setExcelFile(file);
        setIsParsing(true);
        setExcelDataPreview(null);
        try {
            const data = await parseExcelData(file);
            if (data.length === 0) {
                setErrorMsg('엑셀 파일에 데이터가 없습니다.');
            } else {
                setExcelDataPreview(data);
            }
        } catch (err) {
            console.error(err);
            setErrorMsg('엑셀 파일 분석 중 오류가 발생했습니다. 올바른 파일인지 확인해주세요.');
        } finally {
            setIsParsing(false);
        }
    };

    const handleExcelChange = (e) => processExcelFile(e.target.files[0]);

    const processTemplateFile = (file) => {
        if (!file) return;
        const ext = file.name.split('.').pop().toLowerCase();
        if (ext !== 'pptx') {
            setErrorMsg('PPT 템플릿은 .pptx 확장자만 지원합니다.');
            return;
        }
        setErrorMsg(null);
        setSuccessMsg(null);
        setPptTemplate(file);
        setTemplateLabel(file.name);
    };

    const handleTemplateChange = (e) => processTemplateFile(e.target.files[0]);

    const handleExcelDragEvents = {
        onDragOver: (e) => { e.preventDefault(); e.stopPropagation(); },
        onDragEnter: (e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingExcel(true); },
        onDragLeave: (e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingExcel(false); },
        onDrop: (e) => {
            e.preventDefault(); e.stopPropagation(); setIsDraggingExcel(false);
            const file = e.dataTransfer.files[0];
            if (file) processExcelFile(file);
        }
    };

    const handleTemplateDragEvents = {
        onDragOver: (e) => { e.preventDefault(); e.stopPropagation(); },
        onDragEnter: (e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingTemplate(true); },
        onDragLeave: (e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingTemplate(false); },
        onDrop: (e) => {
            e.preventDefault(); e.stopPropagation(); setIsDraggingTemplate(false);
            const file = e.dataTransfer.files[0];
            if (file) processTemplateFile(file);
        }
    };

    const handleBatchPptDragEvents = {
        onDragOver: (e) => { e.preventDefault(); e.stopPropagation(); },
        onDragEnter: (e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingBatch(true); },
        onDragLeave: (e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingBatch(false); },
        onDrop: (e) => {
            e.preventDefault(); e.stopPropagation(); setIsDraggingBatch(false);
            const files = Array.from(e.dataTransfer.files);
            const validFiles = files.filter(f => f.name.toLowerCase().endsWith('.pptx'));
            if (validFiles.length > 0) {
                setBatchPptFiles(prev => [...prev, ...validFiles]);
            } else {
                setErrorMsg('PPT 파일(.pptx)만 지원합니다.');
            }
        }
    };

    const handleBatchFileChange = (e) => {
        const files = Array.from(e.target.files);
        const validFiles = files.filter(f => f.name.toLowerCase().endsWith('.pptx'));
        if (validFiles.length > 0) {
            setBatchPptFiles(prev => [...prev, ...validFiles]);
        }
    };

    const removeBatchFile = (indexToRemove) => {
        setBatchPptFiles(prev => prev.filter((_, idx) => idx !== indexToRemove));
    };

    const handleGenerate = async () => {
        if (!excelFile || !pptTemplate || !excelDataPreview) {
            setErrorMsg('엑셀 파일과 PPT 템플릿을 모두 등록해주세요.');
            return;
        }
        setErrorMsg(null);
        setSuccessMsg(null);
        setIsGenerating(true);
        await new Promise(r => setTimeout(r, 800));
        try {
            await generatePptFromTemplate(pptTemplate, excelDataPreview, generationMode, chunkSize);
            setSuccessMsg('성공적으로 PPT 파일이 생성되어 다운로드되었습니다.');
            setExcelFile(null);
            setPptTemplate(null);
            setExcelDataPreview(null);
            setTemplateLabel('');
        } catch (err) {
            console.error(err);
            if (err.message && err.message.includes("Can't find end of central directory")) {
                setErrorMsg('유효하지 않은 PPTX 파일입니다. 손상되었거나 암호가 걸려있을 수 있습니다.');
            } else {
                setErrorMsg(`변환 중 오류 발생: ${err.message || '알 수 없는 오류'}`);
            }
        } finally {
            setIsGenerating(false);
        }
    };

    // 스마트 애니메이션 핸들러
    const processAnimationFile = async (file) => {
        if (!file) return;
        if (!file.name.toLowerCase().endsWith('.pptx')) {
            setErrorMsg('PPTX 파일만 지원합니다.');
            return;
        }
        setErrorMsg(null);
        setSuccessMsg(null);
        setAnimationPptFile(file);
        
        // 슬라이드 개수 분석
        try {
            const count = await getPptSlideCount(file);
            setSlideAnimations(Array(count).fill(null).map(() => ({ enabled: true, type: 'transition', useGrouping: true })));
        } catch (err) {
            console.error('슬라이드 분석 오류:', err);
            setErrorMsg('PPT 파일을 분석하는 중 오류가 발생했습니다.');
        }
    };

    const handleAnimationProcess = async () => {
        if (!animationPptFile) {
            setErrorMsg('애니메이션을 적용할 PPT 파일을 등록해주세요.');
            return;
        }
        setErrorMsg(null);
        setSuccessMsg(null);
        setIsProcessingAnimation(true);
        try {
            const modifiedBlob = await addSmartAnimationsToPpt(animationPptFile, { 
                animationType,
                useGrouping,
                perSlideConfigs: slideAnimations 
            });
            await saveFileWithLocationPicker(modifiedBlob, `애니메이션_${animationPptFile.name}`);
            setSuccessMsg('성공적으로 스마트 애니메이션이 적용되어 저장되었습니다.');
            setAnimationPptFile(null);
            setSlideAnimations([]);
        } catch (err) {
            console.error(err);
            setErrorMsg(`애니메이션 처리 중 오류 발생: ${err.message || '알 수 없는 오류'}`);
        } finally {
            setIsProcessingAnimation(false);
        }
    };

    const updateAllSlideAnimations = (enabled, type, grouping) => {
        setSlideAnimations(prev => prev.map(anim => ({ 
            enabled: enabled !== null ? enabled : anim.enabled, 
            type: type !== null ? type : anim.type,
            useGrouping: grouping !== null ? grouping : anim.useGrouping
        })));
        if (type) setAnimationType(type);
        if (grouping !== null) setUseGrouping(grouping);
    };

    const updateSingleSlideAnimation = (index, updates) => {
        setSlideAnimations(prev => {
            const next = [...prev];
            next[index] = { ...next[index], ...updates };
            return next;
        });
    };

    const handleHwpxConversion = async () => {
        if (!hwpxPptFile) return;
        setIsConvertingHwpx(true);
        setErrorMsg(null);
        setSuccessMsg(null);
        setHwpxResultStats(null);
        try {
            const hwpxBlob = await convertPptxToHwpx(hwpxPptFile);
            
            const defaultFileName = `변환_${hwpxPptFile.name.replace(/\.[^/.]+$/, "")}.hwpx`;
            const saved = await saveFileWithLocationPicker(hwpxBlob, defaultFileName);
            
            if (saved) {
                setSuccessMsg(`성공적으로 HWPX 한글 표준 문서로 스마트 변환 및 저장을 마쳤습니다!`);
                setHwpxResultStats({
                    slidesCount: hwpxBlob.totalSlidesCount,
                    paragraphsCount: hwpxBlob.totalParagraphsCount,
                    tablesCount: hwpxBlob.totalTablesCount,
                    imagesCount: hwpxBlob.totalImagesCount
                });
            }
        } catch (err) {
            console.error("HWPX 변환 실패:", err);
            setErrorMsg(`HWPX 변환에 실패했습니다: ${err.message || 'PPT 내부 구조 분석 오류'}`);
        } finally {
            setIsConvertingHwpx(false);
        }
    };

    const handleHwpxFusion = async () => {
        if (!hwpxFusionPptFile || !hwpxFusionTemplateFile) {
            setErrorMsg('PPT 요건기술서 파일과 HWPX 양식 템플릿 파일을 모두 업로드해주세요.');
            return;
        }
        setIsFusingHwpx(true);
        setErrorMsg(null);
        setSuccessMsg(null);
        setFusionResultStats(null);
        try {
            const resultBlob = await fusePptToHwpxTemplate(hwpxFusionPptFile, hwpxFusionTemplateFile, fusionMergeMode);
            
            const defaultFileName = fusionMergeMode 
                ? `융합_${hwpxFusionPptFile.name.replace(/\.[^/.]+$/, "")}.hwpx`
                : `융합개별_${hwpxFusionPptFile.name.replace(/\.[^/.]+$/, "")}.zip`;
                
            const saved = await saveFileWithLocationPicker(resultBlob, defaultFileName);
            
            if (saved) {
                setSuccessMsg(fusionMergeMode 
                    ? `성공적으로 HWPX 표준 양식 융합 및 다운로드를 마쳤습니다! (1개 HWPX 파일 병합)`
                    : `성공적으로 HWPX 양식 분할 융합 및 ZIP 파일 다운로드를 마쳤습니다! (개별 고유번호별 분할)`
                );
                setFusionResultStats({
                    fusedCount: resultBlob.fusedCount,
                    requirementsList: resultBlob.requirementsList,
                    mode: resultBlob.mode
                });
            }
        } catch (err) {
            console.error("HWPX 융합 실패:", err);
            setErrorMsg(`HWPX 양식 융합에 실패했습니다: ${err.message || 'PPT/HWPX 내부 데이터 융합 오류'}`);
        } finally {
            setIsFusingHwpx(false);
        }
    };

    const buildBatchReportDetail = (modifiedBlob, options) => {
        const {
            applyTableDesignChecked,
            parsedRules,
            parsedFontRules,
            parsedFontSizeRules,
            applyDesignChecked,
            applySpecialCharClean,
            add_space_before_parenthesis,
            textColorRules,
            preventWordWrap
        } = options;

        const changes = [];

        // 1. 단어 수정
        if (parsedRules.length > 0) {
            const count = modifiedBlob.totalReplacedWords || 0;
            changes.push(`단어 수정 ${count}개`);
        }
        // 2. 폰트 변경
        if (parsedFontRules.length > 0) {
            const count = modifiedBlob.totalReplacedFonts || 0;
            changes.push(`폰트 변경 ${count}개`);
        }
        // 3. 폰트 크기 변경
        if (parsedFontSizeRules.length > 0) {
            const count = modifiedBlob.totalReplacedFontSizes || 0;
            changes.push(`폰트 크기 변경 ${count}개`);
        }
        // 4. 특수문자 일괄 정제
        if (applySpecialCharClean) {
            const count = modifiedBlob.totalSpecialCharsCleaned || 0;
            changes.push(`특수문자 정제 ${count}개`);
        }
        // 5. 외곽선/텍스트 디자인 변경
        if (applyDesignChecked) {
            const count = modifiedBlob.totalReplacedTextDesigns || 0;
            changes.push(`텍스트 디자인 변경 ${count}개`);
        }
        // 6. 표 표준화
        if (applyTableDesignChecked) {
            if (modifiedBlob.totalTablesCount > 0) {
                changes.push(`표 표준화 및 스타일 적용 ${modifiedBlob.totalTablesCount}개`);
            } else {
                changes.push(`표 없음(표 스타일 적용 제외됨)`);
            }
        }
        // 7. 제목 괄호 공백 추가
        if (add_space_before_parenthesis) {
            const count = modifiedBlob.totalTitleSpacesAdded || 0;
            changes.push(`제목 괄호 공백 추가 ${count}개`);
        }
        // 8. 글자 색상 변경
        if (textColorRules && textColorRules.trim()) {
            const count = modifiedBlob.totalTextColorReplaced || 0;
            changes.push(`글자 색상 변경 ${count}개`);
        }
        // 9. 단락 단어 잘림 방지
        if (preventWordWrap) {
            const count = modifiedBlob.totalWordWrapPrevented || 0;
            changes.push(`단어 잘림 방지 적용 ${count}개 단락`);
        }

        if (changes.length > 0) {
            // 실질적인 변경이 하나라도 존재하는지 확인 (표 없음 제외)
            const hasRealChanges = (modifiedBlob.totalReplacedWords || 0) > 0 ||
                                  (modifiedBlob.totalReplacedFonts || 0) > 0 ||
                                  (modifiedBlob.totalReplacedFontSizes || 0) > 0 ||
                                  (modifiedBlob.totalSpecialCharsCleaned || 0) > 0 ||
                                  (modifiedBlob.totalReplacedTextDesigns || 0) > 0 ||
                                  (add_space_before_parenthesis && (modifiedBlob.totalTitleSpacesAdded || 0) > 0) ||
                                  ((textColorRules && textColorRules.trim()) && (modifiedBlob.totalTextColorReplaced || 0) > 0) ||
                                  (preventWordWrap && (modifiedBlob.totalWordWrapPrevented || 0) > 0) ||
                                  (applyTableDesignChecked && (modifiedBlob.totalTablesCount || 0) > 0);

            if (hasRealChanges) {
                return `${changes.join(', ')} 완료`;
            } else {
                if (applyTableDesignChecked && (modifiedBlob.totalTablesCount || 0) === 0) {
                    return `⚠️ 표가 존재하지 않으며, 감지된 다른 일치 변경 대상(단어, 폰트, 크기, 특수문자, 제목 공백, 글자 색상, 단어 잘림 방지)이 없어 원본 그대로 저장했습니다.`;
                }
                return `ℹ️ 일치하는 단어, 폰트명, 폰트 크기 변경, 정제할 특수문자, 수정할 제목 괄호, 글자 색상, 또는 잘림 방지 대상 단락이 감지되지 않아 원본 그대로 저장했습니다.`;
            }
        }

        return '변경 사항 없음 (원본 그대로 저장 완료)';
    };

    // 일괄 편집 핸들러 (단어 수정 + 디자인 적용 + 다중 파일 + 폴더 지정)
    const handleBatchProcess = async () => {
        if (batchPptFiles.length === 0) {
            setErrorMsg('PPT 파일을 1개 이상 등록해주세요.');
            return;
        }

        let parsedRules = [];
        if (replaceRules.trim()) {
            const parts = replaceRules.split(',');
            for (const part of parts) {
                const trimmed = part.trim();
                if (!trimmed) continue;
                const match = trimmed.match(/^(.+?)\((.+?)\)$/);
                if (match) {
                    parsedRules.push({ oldWord: match[1].trim(), newWord: match[2].trim() });
                } else {
                    setErrorMsg(`규칙 형식이 올바르지 않습니다: "${trimmed}" (예: 기존단어(새단어))`);
                    return;
                }
            }
        }

        let parsedFontRules = [];
        if (fontRules.trim()) {
            const parts = fontRules.split(',');
            for (const part of parts) {
                const trimmed = part.trim();
                if (!trimmed) continue;
                const match = trimmed.match(/^(.+?)\((.+?)\)$/);
                if (match) {
                    parsedFontRules.push({ oldWord: match[1].trim(), newWord: match[2].trim() });
                } else {
                    setErrorMsg(`폰트 규칙 형식이 올바르지 않습니다: "${trimmed}" (예: Arial(나눔고딕))`);
                    return;
                }
            }
        }

        let parsedFontSizeRules = [];
        if (fontSize.trim()) {
            const parts = fontSize.split(',');
            for (const part of parts) {
                const trimmed = part.trim();
                if (!trimmed) continue;
                const match = trimmed.match(/^(.+?)\((.+?)\)$/);
                if (match) {
                    const oldSize = parseFloat(match[1].trim());
                    const newSize = parseFloat(match[2].trim());
                    if (isNaN(oldSize) || isNaN(newSize)) {
                        setErrorMsg(`폰트 크기 규칙의 숫자가 올바르지 않습니다: "${trimmed}"`);
                        return;
                    }
                    parsedFontSizeRules.push({ oldSize, newSize });
                } else {
                    // 단일 숫자 입력 시 전체 적용 (기존 기능 유지)
                    const size = parseFloat(trimmed);
                    if (isNaN(size)) {
                        setErrorMsg(`폰트 크기 형식이 올바르지 않습니다: "${trimmed}" (예: 7.9(10.0))`);
                        return;
                    }
                    parsedFontSizeRules.push({ oldSize: null, newSize: size });
                }
            }
        }

        if (parsedRules.length === 0 && parsedFontRules.length === 0 && !applyDesignChecked && parsedFontSizeRules.length === 0 && !applyTableDesignChecked && !applySpecialCharClean && !add_title_page_numbers && !textColorRules.trim() && !preventWordWrap) {
            setErrorMsg('적용할 단어 수정, 폰트 변경, 폰트 크기, 테이블 디자인 표준화, 텍스트 디자인 변경, 특수문자 일괄 정제, 동일 제목 일련번호 추가, 글자 색상 일괄 매핑, 또는 단락 단어 잘림 방지 중 하나 이상을 입력/선택해주세요.');
            return;
        }

        setErrorMsg(null);
        setSuccessMsg(null);
        setIsProcessingBatch(true);

        try {
            let successCount = 0;
            const reports = []; // 📊 실시간 파일별 처리 리포트 축적 배열

            // 💡 파일별 showSaveFilePicker 방식으로 전환:
            // showDirectoryPicker는 다운로드 폴더를 시스템 보호 폴더로 분류하여 브라우저가 차단합니다.
            // 대신 각 파일을 showSaveFilePicker로 개별 저장하거나, 다중 파일은 ZIP으로 묶어 다운로드합니다.

            if (batchPptFiles.length === 1 && 'showSaveFilePicker' in window) {
                // 단일 파일: showSaveFilePicker로 다운로드 폴더에서 저장 다이얼로그 열기
                const file = batchPptFiles[0];
                try {
                    const options = { 
                        replaceRules: parsedRules, 
                        fontRules: parsedFontRules,
                        fontSizeRules: parsedFontSizeRules,
                        applyDesign: applyDesignChecked, 
                        applyTableDesign: applyTableDesignChecked, 
                        applyFirstRowHeaderStyle: applyFirstRowHeaderStyle,
                        targetText: designTargetText,
                        applySpecialCharClean: applySpecialCharClean,
                        replaceNbs: replaceNbs,
                        unifyBullets: unifyBullets,
                        clean_vertical_tab: clean_vertical_tab,
                        add_title_page_numbers: add_title_page_numbers,
                        add_space_before_parenthesis: add_space_before_parenthesis,
                        textColorRulesStr: textColorRules,
                        preventWordWrap: preventWordWrap
                    };
                    const modifiedBlob = await processPptBatch(file, options);
                    const fileName = `수정_${file.name}`;

                    let saved = false;
                    try {
                        const handle = await window.showSaveFilePicker({
                            suggestedName: fileName,
                            startIn: 'downloads',
                            types: [{
                                description: 'PowerPoint Presentation',
                                accept: { 'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx'] },
                            }],
                        });
                        const writable = await handle.createWritable();
                        await writable.write(modifiedBlob);
                        await writable.close();
                        saved = true;
                    } catch (pickerErr) {
                        if (pickerErr.name === 'AbortError') {
                            setIsProcessingBatch(false);
                            return;
                        }
                        // 폴백: file-saver
                        const { saveAs } = await import('file-saver');
                        saveAs(modifiedBlob, fileName);
                        saved = true;
                    }

                    if (saved) {
                        const detailMsg = buildBatchReportDetail(modifiedBlob, {
                            applyTableDesignChecked,
                            parsedRules,
                            parsedFontRules,
                            parsedFontSizeRules,
                            applyDesignChecked,
                            applySpecialCharClean,
                            add_space_before_parenthesis,
                            textColorRules,
                            preventWordWrap
                        });
                        reports.push({ fileName: file.name, status: 'success', detail: detailMsg });
                        successCount++;
                    }
                } catch (fileErr) {
                    console.error(`Error processing ${file.name}:`, fileErr);
                    reports.push({ fileName: file.name, status: 'error', detail: `❌ 처리 실패: ${fileErr.message || 'PPT 내부 구조 파싱 에러'}` });
                }
            } else {
                // 다중 파일: 모두 처리 후 ZIP으로 일괄 다운로드
                const zip = new JSZip();

                for (const file of batchPptFiles) {
                    try {
                        const options = { 
                            replaceRules: parsedRules, 
                            fontRules: parsedFontRules,
                            fontSizeRules: parsedFontSizeRules,
                            applyDesign: applyDesignChecked, 
                            applyTableDesign: applyTableDesignChecked, 
                            applyFirstRowHeaderStyle: applyFirstRowHeaderStyle,
                            targetText: designTargetText,
                            applySpecialCharClean: applySpecialCharClean,
                            replaceNbs: replaceNbs,
                            unifyBullets: unifyBullets,
                            clean_vertical_tab: clean_vertical_tab,
                            add_title_page_numbers: add_title_page_numbers,
                            add_space_before_parenthesis: add_space_before_parenthesis,
                            textColorRulesStr: textColorRules,
                            preventWordWrap: preventWordWrap
                        };
                        const modifiedBlob = await processPptBatch(file, options);
                        const fileName = `수정_${file.name}`;
                        zip.file(fileName, modifiedBlob);

                        const detailMsg = buildBatchReportDetail(modifiedBlob, {
                            applyTableDesignChecked,
                            parsedRules,
                            parsedFontRules,
                            parsedFontSizeRules,
                            applyDesignChecked,
                            applySpecialCharClean,
                            add_space_before_parenthesis,
                            textColorRules,
                            preventWordWrap
                        });
                        reports.push({ fileName: file.name, status: 'success', detail: detailMsg });
                        successCount++;
                    } catch (fileErr) {
                        console.error(`Error processing ${file.name}:`, fileErr);
                        reports.push({ fileName: file.name, status: 'error', detail: `❌ 처리 실패: ${fileErr.message || 'PPT 내부 구조 파싱 에러'}` });
                    }
                }

                if (successCount > 0) {
                    const zipBlob = await zip.generateAsync({ type: 'blob' });
                    // 다중 파일 ZIP도 다운로드 폴더에 저장 다이얼로그 시도
                    let zipSaved = false;
                    if ('showSaveFilePicker' in window) {
                        try {
                            const handle = await window.showSaveFilePicker({
                                suggestedName: '수정_PPT_산출물_일괄다운로드.zip',
                                startIn: 'downloads',
                                types: [{
                                    description: 'ZIP 압축 파일',
                                    accept: { 'application/zip': ['.zip'] },
                                }],
                            });
                            const writable = await handle.createWritable();
                            await writable.write(zipBlob);
                            await writable.close();
                            zipSaved = true;
                        } catch (pickerErr) {
                            if (pickerErr.name !== 'AbortError') {
                                const { saveAs } = await import('file-saver');
                                saveAs(zipBlob, '수정_PPT_산출물_일괄다운로드.zip');
                                zipSaved = true;
                            }
                        }
                    }
                    if (!zipSaved) {
                        const { saveAs } = await import('file-saver');
                        saveAs(zipBlob, '수정_PPT_산출물_일괄다운로드.zip');
                    }
                }
            }

            setBatchReport(reports);

            if (successCount > 0) {
                if (batchPptFiles.length > 1) {
                    setSuccessMsg(`성공적으로 ${successCount}개의 파일을 처리하여 ZIP 압축 파일로 다운로드했습니다. 하단의 파일별 일괄 편집 상세 결과 리포트를 확인해 주세요.`);
                } else {
                    setSuccessMsg(`성공적으로 파일이 편집·저장되었습니다. 하단의 파일별 일괄 편집 상세 결과 리포트를 확인해 주세요.`);
                }
                setBatchPptFiles([]);
                setFontRules('');
                setFontSize('');
                setApplyDesignChecked(false);
                setApplyTableDesignChecked(false);
                setApplyFirstRowHeaderStyle(true);
                setDesignTargetText('');
                setApplySpecialCharClean(false);
                setReplaceNbs(true);
                setUnifyBullets(true);
                set_clean_vertical_tab(true);
                set_add_title_page_numbers(false);
                set_add_space_before_parenthesis(true);
            } else {
                setErrorMsg('처리된 파일이 없습니다. 변경 대상 텍스트나 디자인 요소가 존재하는지 확인해주세요.');
            }
        } catch (err) {
            console.error(err);
            setErrorMsg(`작업 중 오류 발생: ${err.message || '알 수 없는 오류'}`);
        } finally {
            setIsProcessingBatch(false);
        }
    };

    const columns = excelDataPreview && excelDataPreview.length > 0 ? Object.keys(excelDataPreview[0]) : [];

    return (
        <div style={{ display: 'flex', flexDirection: 'column' }}>
            <div className="glass-panel animate-slide-up" style={{ padding: '24px', display: 'flex', flexDirection: 'column', gap: '24px' }}>
                
                {/* 헤더 */}
                <div style={{ display: 'flex', alignItems: 'center', gap: '12px', borderBottom: '1px solid var(--panel-border)', paddingBottom: '16px' }}>
                    <div style={{ padding: '10px', background: 'rgba(59, 130, 246, 0.1)', borderRadius: '10px' }}>
                        <Presentation size={24} color="var(--accent-blue)" />
                    </div>
                    <div>
                        <h2 style={{ margin: 0, fontSize: '20px', color: 'var(--text-primary)', letterSpacing: '-0.5px' }}>PPT 스마트 편집기</h2>
                        <p style={{ margin: '4px 0 0', fontSize: '13.5px', color: 'var(--text-secondary)' }}>
                            데이터 매핑을 통한 PPT 생성 또는 텍스트 서식 일괄 변경 기능을 제공합니다.
                        </p>
                    </div>
                </div>

                {/* 탭 메뉴 */}
                <div style={{ display: 'flex', gap: '12px', borderBottom: '1px solid var(--panel-border)', paddingBottom: '12px' }}>
                    <button 
                        id="tab-excel-mapping"
                        onClick={() => { setActiveTab('excel_mapping'); setErrorMsg(null); setSuccessMsg(null); }}
                        className="interactive"
                        style={{
                            padding: '10px 20px', borderRadius: '8px', cursor: 'pointer',
                            background: activeTab === 'excel_mapping' ? 'rgba(59, 130, 246, 0.1)' : 'transparent',
                            border: '1px solid ' + (activeTab === 'excel_mapping' ? 'var(--accent-blue)' : 'transparent'),
                            color: activeTab === 'excel_mapping' ? 'var(--accent-blue)' : 'var(--text-secondary)',
                            fontWeight: 600, fontSize: '14px', transition: 'all 0.2s'
                        }}
                    >
                        엑셀 데이터 매핑 생성
                    </button>
                    <button 
                        id="tab-batch-edit"
                        onClick={() => { setActiveTab('batch_edit'); setErrorMsg(null); setSuccessMsg(null); }}
                        className="interactive"
                        style={{
                            padding: '10px 20px', borderRadius: '8px', cursor: 'pointer',
                            background: activeTab === 'batch_edit' ? 'rgba(236, 72, 153, 0.1)' : 'transparent',
                            border: '1px solid ' + (activeTab === 'batch_edit' ? '#ec4899' : 'transparent'),
                            color: activeTab === 'batch_edit' ? '#f472b6' : 'var(--text-secondary)',
                            fontWeight: 600, fontSize: '14px', transition: 'all 0.2s'
                        }}
                    >
                        PPT 텍스트/디자인 일괄 편집
                    </button>
                    <button 
                        id="tab-smart-animation"
                        onClick={() => { setActiveTab('smart_animation'); setErrorMsg(null); setSuccessMsg(null); }}
                        className="interactive"
                        style={{
                            padding: '10px 20px', borderRadius: '8px', cursor: 'pointer',
                            background: activeTab === 'smart_animation' ? 'rgba(168, 85, 247, 0.1)' : 'transparent',
                            border: '1px solid ' + (activeTab === 'smart_animation' ? '#a855f7' : 'transparent'),
                            color: activeTab === 'smart_animation' ? '#c084fc' : 'var(--text-secondary)',
                            fontWeight: 600, fontSize: '14px', transition: 'all 0.2s'
                        }}
                    >
                        PPT 스마트 애니메이션
                    </button>

                    <button 
                        id="tab-pptx-hwpx-fusion"
                        onClick={() => { setActiveTab('pptx_hwpx_fusion'); setErrorMsg(null); setSuccessMsg(null); setFusionResultStats(null); }}
                        className="interactive"
                        style={{
                            padding: '10px 20px', borderRadius: '8px', cursor: 'pointer',
                            background: activeTab === 'pptx_hwpx_fusion' ? 'rgba(59, 130, 246, 0.1)' : 'transparent',
                            border: '1px solid ' + (activeTab === 'pptx_hwpx_fusion' ? 'var(--accent-color)' : 'transparent'),
                            color: activeTab === 'pptx_hwpx_fusion' ? 'var(--accent-color)' : 'var(--text-secondary)',
                            fontWeight: 600, fontSize: '14px', transition: 'all 0.2s'
                        }}
                    >
                        PPT ➜ HWPX 양식 융합
                    </button>
                </div>

                {/* 에러/성공 메시지 공통 표시 */}
                {errorMsg && (
                    <div className="animate-fade-in" style={{ padding: '12px', background: 'rgba(239, 68, 68, 0.1)', color: 'var(--danger-color)', borderRadius: '8px', border: '1px solid rgba(239, 68, 68, 0.2)', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <X size={16} /> {errorMsg}
                    </div>
                )}
                {successMsg && (
                    <div className="animate-fade-in" style={{ padding: '12px', background: 'rgba(16, 185, 129, 0.1)', color: 'var(--success-color)', borderRadius: '8px', border: '1px solid rgba(16, 185, 129, 0.2)', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <FileDown size={16} /> {successMsg}
                    </div>
                )}

                {/* 탭 컨텐츠 */}
                {activeTab === 'excel_mapping' ? (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
                        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
                            {/* 데이터 엑셀 업로드 영역 */}
                            <div 
                                {...handleExcelDragEvents}
                                style={{ 
                                    border: `2px dashed ${isDraggingExcel ? 'var(--success-color)' : 'rgba(34, 197, 94, 0.3)'}`, 
                                    borderRadius: '12px', padding: '24px',
                                    background: isDraggingExcel ? 'rgba(34, 197, 94, 0.1)' : 'rgba(34, 197, 94, 0.05)', 
                                    display: 'flex', flexDirection: 'column', gap: '16px',
                                    transition: 'all 0.2s ease', cursor: isDraggingExcel ? 'copy' : 'default'
                                }}
                            >
                                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                    <FileSpreadsheet size={20} color="var(--success-color)" />
                                    <h3 style={{ margin: 0, fontSize: '16px', color: 'var(--text-primary)' }}>1. 엑셀 데이터 등록</h3>
                                </div>
                                
                                <input
                                    type="file"
                                    accept=".xlsx, .xls"
                                    onChange={handleExcelChange}
                                    ref={excelInputRef}
                                    style={{ display: 'none' }}
                                />
                                
                                <div style={{ display: 'flex', gap: '12px', alignItems: 'center' }}>
                                    <button
                                        onClick={() => excelInputRef.current?.click()}
                                        className="interactive"
                                        style={{
                                            display: 'flex', alignItems: 'center', gap: '8px',
                                            padding: '10px 16px', borderRadius: '8px', cursor: 'pointer',
                                            background: 'var(--bg-secondary)', border: '1px solid var(--panel-border)',
                                            color: 'var(--text-primary)', fontWeight: 600, fontSize: '13px'
                                        }}
                                    >
                                        {isParsing ? <Loader2 size={16} className="animate-spin" /> : <Upload size={16} />}
                                        엑셀 찾아보기 (.xlsx)
                                    </button>
                                    {excelFile && (
                                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', background: 'rgba(0,0,0,0.2)', padding: '6px 12px', borderRadius: '6px', fontSize: '13px', border: '1px solid var(--success-color)' }}>
                                            <span style={{ color: 'var(--success-color)' }}>✔</span>
                                            {excelFile.name}
                                        </div>
                                    )}
                                </div>

                                <div style={{ fontSize: '12px', color: 'var(--text-secondary)', lineHeight: '1.5' }}>
                                    💡 1행이 라벨(변수명)이 되고, 2행부터 실제 데이터로 간주합니다.
                                </div>
                            </div>

                            {/* PPT 양식 파일 업로드 영역 */}
                            <div 
                                {...handleTemplateDragEvents}
                                style={{ 
                                    border: `2px dashed ${isDraggingTemplate ? 'var(--accent-blue)' : 'rgba(59, 130, 246, 0.3)'}`, 
                                    borderRadius: '12px', padding: '24px',
                                    background: isDraggingTemplate ? 'rgba(59, 130, 246, 0.1)' : 'rgba(59, 130, 246, 0.05)', 
                                    display: 'flex', flexDirection: 'column', gap: '16px',
                                    transition: 'all 0.2s ease', cursor: isDraggingTemplate ? 'copy' : 'default'
                                }}
                            >
                                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                    <Presentation size={20} color="var(--accent-blue)" />
                                    <h3 style={{ margin: 0, fontSize: '16px', color: 'var(--text-primary)' }}>2. PPT 양식(.pptx) 등록</h3>
                                </div>
                                
                                <input
                                    type="file"
                                    accept=".pptx"
                                    onChange={handleTemplateChange}
                                    ref={pptInputRef}
                                    style={{ display: 'none' }}
                                />
                                
                                <div style={{ display: 'flex', gap: '12px', alignItems: 'center' }}>
                                    <button
                                        onClick={() => pptInputRef.current?.click()}
                                        className="interactive"
                                        style={{
                                            display: 'flex', alignItems: 'center', gap: '8px',
                                            padding: '10px 16px', borderRadius: '8px', cursor: 'pointer',
                                            background: 'var(--bg-secondary)', border: '1px solid var(--panel-border)',
                                            color: 'var(--text-primary)', fontWeight: 600, fontSize: '13px'
                                        }}
                                    >
                                        <Upload size={16} /> PPT 템플릿 찾기
                                    </button>
                                    {pptTemplate && (
                                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', background: 'rgba(0,0,0,0.2)', padding: '6px 12px', borderRadius: '6px', fontSize: '13px', border: '1px solid var(--accent-blue)' }}>
                                            <span style={{ color: 'var(--accent-blue)' }}>✔</span>
                                            {templateLabel}
                                        </div>
                                    )}
                                </div>

                                <div style={{ 
                                    fontSize: '13px', color: 'var(--text-secondary)', lineHeight: '1.6', 
                                    marginTop: '4px', background: 'rgba(255,255,255,0.03)', 
                                    padding: '14px 18px', borderRadius: '8px', borderLeft: '3px solid var(--accent-blue)'
                                }}>
                                    <div style={{ fontWeight: 700, color: 'var(--text-primary)', marginBottom: '8px', display: 'flex', alignItems: 'center', gap: '6px' }}>
                                        💡 PPT 템플릿 매핑 가이드
                                    </div>
                                    <p style={{ margin: 0 }}>
                                        PPT 내 텍스트 상자에 <code>{`{열이름}`}</code> (예: <code>{`{성명}`}</code>, <code>{`{부서}`}</code>) 형식으로 입력하세요.<br/>
                                        <strong style={{ color: 'var(--accent-blue)' }}>* 자동 슬라이드 복제:</strong> 엑셀 행 개수만큼 슬라이드가 자동으로 생성되며 각 행의 데이터가 각 슬라이드에 채워집니다.
                                    </p>
                                </div>
                            </div>
                        </div>

                        {/* 데이터 컬럼 매핑 프리뷰 */}
                        {excelDataPreview && (
                            <div className="animate-fade-in" style={{ padding: '20px', background: 'rgba(255,255,255,0.03)', border: '1px solid var(--panel-border)', borderRadius: '12px' }}>
                                <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '12px' }}>
                                    <Info size={18} color="var(--warning-color)" />
                                    <h4 style={{ margin: 0, fontSize: '15px' }}>사용 가능한 템플릿 태그 (총 {excelDataPreview.length}개 행 인식됨)</h4>
                                </div>
                                <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px' }}>
                                    {columns.map(col => (
                                        <div key={col} style={{ 
                                            padding: '4px 8px', background: 'rgba(168,85,247,0.1)', color: '#c084fc',
                                            borderRadius: '6px', fontSize: '13px', fontFamily: 'monospace', fontWeight: 600, border: '1px solid rgba(168,85,247,0.2)'
                                        }}>
                                            {`{${col}}`}
                                        </div>
                                    ))}
                                </div>
                            </div>
                        )}

                        <div style={{ marginTop: 'auto', paddingTop: '20px' }}>
                            <button
                                className="interactive"
                                onClick={handleGenerate}
                                disabled={!excelFile || !pptTemplate || isGenerating}
                                style={{
                                    width: '100%',
                                    padding: '16px',
                                    background: (!excelFile || !pptTemplate || isGenerating) ? 'rgba(255,255,255,0.05)' : 'linear-gradient(135deg, #0284c7, #3b82f6)',
                                    color: (!excelFile || !pptTemplate || isGenerating) ? 'var(--text-muted)' : 'white',
                                    border: 'none',
                                    borderRadius: '12px',
                                    fontSize: '16px',
                                    fontWeight: 700,
                                    cursor: (!excelFile || !pptTemplate || isGenerating) ? 'not-allowed' : 'pointer',
                                    display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '10px'
                                }}
                            >
                                {isGenerating ? (
                                    <><Loader2 size={20} className="animate-spin" /> 파워포인트 문서 자동 치환 및 생성 중...</>
                                ) : (
                                    <><Play size={20} fill={(!excelFile || !pptTemplate) ? 'none' : 'currentColor'} /> 엑셀 ↔ PPT 자동 매핑 및 파일 변환 시작</>
                                )}
                            </button>
                        </div>
                    </div>
                ) : activeTab === 'batch_edit' ? (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
                        {/* 일괄 편집 (다중 파일) UI */}
                        <div 
                            {...handleBatchPptDragEvents}
                            style={{ 
                                border: `2px dashed ${isDraggingBatch ? '#ec4899' : 'rgba(236, 72, 153, 0.3)'}`, 
                                borderRadius: '12px', padding: '24px',
                                background: isDraggingBatch ? 'rgba(236, 72, 153, 0.1)' : 'rgba(236, 72, 153, 0.05)', 
                                display: 'flex', flexDirection: 'column', gap: '16px',
                                transition: 'all 0.2s ease', cursor: isDraggingBatch ? 'copy' : 'default'
                            }}
                        >
                            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                <Presentation size={20} color="#f472b6" />
                                <h3 style={{ margin: 0, fontSize: '16px', color: 'var(--text-primary)' }}>1. 원본 PPT 파일(.pptx) 다중 등록</h3>
                            </div>
                            <input
                                type="file"
                                accept=".pptx"
                                multiple
                                onChange={handleBatchFileChange}
                                style={{ display: 'none' }}
                                id="batch-ppt-upload"
                            />
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                                <div>
                                    <button
                                        onClick={() => document.getElementById('batch-ppt-upload').click()}
                                        className="interactive"
                                        style={{
                                            display: 'flex', alignItems: 'center', gap: '8px',
                                            padding: '10px 16px', borderRadius: '8px', cursor: 'pointer',
                                            background: 'var(--bg-secondary)', border: '1px solid var(--panel-border)',
                                            color: 'var(--text-primary)', fontWeight: 600, fontSize: '13px'
                                        }}
                                    >
                                        <Upload size={16} /> PPT 파일(들) 찾기
                                    </button>
                                </div>
                                
                                {batchPptFiles.length > 0 && (
                                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px', marginTop: '8px' }}>
                                        {batchPptFiles.map((file, idx) => (
                                            <div key={idx} style={{ 
                                                display: 'flex', alignItems: 'center', gap: '6px', 
                                                background: 'rgba(0,0,0,0.2)', padding: '4px 10px', 
                                                borderRadius: '6px', fontSize: '12.5px', border: '1px solid #a855f7' 
                                            }}>
                                                <span style={{ color: '#c084fc' }}>✔</span>
                                                <span style={{ maxWidth: '150px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{file.name}</span>
                                                <button 
                                                    onClick={() => removeBatchFile(idx)}
                                                    style={{ background: 'none', border: 'none', color: 'var(--text-muted)', cursor: 'pointer', padding: '2px', display: 'flex' }}
                                                >
                                                    <X size={14} />
                                                </button>
                                            </div>
                                        ))}
                                    </div>
                                )}
                            </div>
                        </div>

                        <div style={{ 
                            border: '1px solid var(--panel-border)', 
                            borderRadius: '12px', padding: '24px',
                            background: 'rgba(255, 255, 255, 0.02)', 
                            display: 'flex', flexDirection: 'column', gap: '20px'
                        }}>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                <Settings size={20} color="#a855f7" />
                                <h3 style={{ margin: 0, fontSize: '16px', color: 'var(--text-primary)' }}>2. 일괄 편집 옵션 설정</h3>
                            </div>
                            
                            {/* 옵션 1: 단어 수정 */}
                            <div style={{ background: 'var(--bg-secondary)', padding: '16px', borderRadius: '8px', border: '1px solid var(--panel-border)' }}>
                                <div style={{ fontWeight: 600, fontSize: '14px', marginBottom: '8px', color: 'var(--text-primary)' }}>
                                    옵션 A: 수정할 단어 규칙 입력 (선택)
                                </div>
                                <input
                                    type="text"
                                    placeholder="예: 어플리케이션(애플리케이션), 컬럼(칼럼)"
                                    value={replaceRules}
                                    onChange={(e) => setReplaceRules(e.target.value)}
                                    style={{
                                        width: '100%', padding: '12px 16px', borderRadius: '8px',
                                        background: 'rgba(0,0,0,0.2)', border: '1px solid var(--panel-border)',
                                        color: 'var(--text-primary)', fontSize: '14px', marginBottom: '8px'
                                    }}
                                />
                                <div style={{ fontSize: '12.5px', color: 'var(--text-secondary)' }}>
                                    형식: <code>기존단어(새로운단어)</code> (복수는 쉼표로 구분)
                                </div>
                                <div style={{ fontSize: '12.5px', color: '#a855f7', marginTop: '6px', display: 'flex', alignItems: 'center', gap: '4px' }}>
                                    <CheckCircle2 size={13} style={{ flexShrink: 0 }} />
                                    <span>입력하신 단어 수정 규칙은 브라우저에 자동 영구 저장되어 언제든 재사용 가능합니다.</span>
                                </div>
                            </div>

                            {/* 옵션 B: 폰트 교체 */}
                            <div style={{ background: 'var(--bg-secondary)', padding: '16px', borderRadius: '8px', border: '1px solid var(--panel-border)' }}>
                                <div style={{ fontWeight: 600, fontSize: '14px', marginBottom: '8px', color: 'var(--text-primary)' }}>
                                    옵션 B: 일괄 변경할 폰트 입력 (선택)
                                </div>
                                <input
                                    type="text"
                                    placeholder="예: Arial(나눔고딕), Calibri(Pretendard)"
                                    value={fontRules}
                                    onChange={(e) => setFontRules(e.target.value)}
                                    style={{
                                        width: '100%', padding: '12px 16px', borderRadius: '8px',
                                        background: 'rgba(0,0,0,0.2)', border: '1px solid var(--panel-border)',
                                        color: 'var(--text-primary)', fontSize: '14px', marginBottom: '8px'
                                    }}
                                />
                                <div style={{ fontSize: '12.5px', color: 'var(--text-secondary)' }}>
                                    형식: <code>기존폰트(변경폰트)</code> (복수는 쉼표로 구분)
                                </div>
                            </div>

                            {/* 옵션 C: 폰트 크기 변경 */}
                            <div style={{ background: 'var(--bg-secondary)', padding: '16px', borderRadius: '8px', border: '1px solid var(--panel-border)' }}>
                                <div style={{ fontWeight: 600, fontSize: '14px', marginBottom: '8px', color: 'var(--text-primary)' }}>
                                    옵션 C: 변경할 폰트 크기 입력 (선택)
                                </div>
                                <input
                                    type="text"
                                    placeholder="예: 7.9(10.0), 12(14)"
                                    value={fontSize}
                                    onChange={(e) => setFontSize(e.target.value)}
                                    style={{
                                        width: '100%', padding: '12px 16px', borderRadius: '8px',
                                        background: 'rgba(0,0,0,0.2)', border: '1px solid var(--panel-border)',
                                        color: 'var(--text-primary)', fontSize: '14px', marginBottom: '8px'
                                    }}
                                />
                                <div style={{ fontSize: '12.5px', color: 'var(--text-secondary)' }}>
                                    형식: <code>기존크기(변경크기)</code> (단일 숫자 입력 시 전체 적용)
                                </div>
                            </div>

                            {/* 옵션 D: 텍스트 디자인 */}
                            <div style={{ background: 'var(--bg-secondary)', padding: '16px', borderRadius: '8px', border: '1px solid var(--panel-border)', display: 'flex', flexDirection: 'column', gap: '12px' }}>
                                <label style={{ display: 'flex', alignItems: 'center', gap: '10px', cursor: 'pointer', fontWeight: 600, fontSize: '14px', color: 'var(--text-primary)' }}>
                                    <input 
                                        type="checkbox" 
                                        checked={applyDesignChecked}
                                        onChange={(e) => setApplyDesignChecked(e.target.checked)}
                                        style={{ width: '18px', height: '18px', cursor: 'pointer', accentColor: '#a855f7' }}
                                    />
                                    옵션 D: 텍스트 윤곽선 디자인 일괄 변경 적용(흰색 실선, 투명도 100%, 너비 0.75)
                                </label>
                                
                                {applyDesignChecked && (
                                    <div className="animate-slide-up" style={{ paddingLeft: '28px', display: 'flex', flexDirection: 'column', gap: '8px' }}>
                                        <input
                                            type="text"
                                            placeholder="특정 단어가 포함된 텍스트 박스만 변경할 경우 입력 (비워두면 전체 적용)"
                                            value={designTargetText}
                                            onChange={(e) => setDesignTargetText(e.target.value)}
                                            style={{
                                                width: '100%', padding: '10px 14px', borderRadius: '8px',
                                                background: 'rgba(0,0,0,0.2)', border: '1px solid var(--panel-border)',
                                                color: 'var(--text-primary)', fontSize: '13.5px'
                                            }}
                                        />
                                    </div>
                                )}
                            </div>

                            {/* 옵션 E: 테이블 표준 디자인 */}
                            <div style={{ background: 'var(--bg-secondary)', padding: '16px', borderRadius: '8px', border: '1px solid var(--panel-border)', display: 'flex', flexDirection: 'column', gap: '12px' }}>
                                <label style={{ display: 'flex', alignItems: 'center', gap: '10px', cursor: 'pointer', fontWeight: 600, fontSize: '14px', color: 'var(--text-primary)' }}>
                                    <input 
                                        type="checkbox" 
                                        checked={applyTableDesignChecked}
                                        onChange={(e) => setApplyTableDesignChecked(e.target.checked)}
                                        style={{ width: '18px', height: '18px', cursor: 'pointer', accentColor: '#a855f7' }}
                                    />
                                    옵션 E: 테이블(표) 표준 디자인 일괄 변경 적용
                                </label>
                                <div style={{ paddingLeft: '28px', fontSize: '13px', color: 'var(--text-secondary)', lineHeight: '1.5' }}>
                                    💡 모든 표 테두리 실선 0.5pt (#7F7F7F)가 기본 적용되며, 첫 행의 특별 포맷팅을 하위 옵션으로 선택 제어할 수 있습니다.
                                </div>
                                
                                <div style={{ 
                                    paddingLeft: '28px', 
                                    display: 'flex', 
                                    flexDirection: 'column', 
                                    gap: '8px', 
                                    borderLeft: '2px solid ' + (applyTableDesignChecked ? 'rgba(168, 85, 247, 0.5)' : 'rgba(255, 255, 255, 0.1)'), 
                                    marginTop: '4px',
                                    opacity: applyTableDesignChecked ? 1 : 0.45,
                                    pointerEvents: applyTableDesignChecked ? 'auto' : 'none',
                                    transition: 'all 0.3s ease'
                                }}>
                                    <label style={{ 
                                        display: 'flex', 
                                        alignItems: 'center', 
                                        gap: '8px', 
                                        cursor: applyTableDesignChecked ? 'pointer' : 'not-allowed', 
                                        fontSize: '13px', 
                                        color: applyTableDesignChecked ? 'var(--text-primary)' : 'var(--text-muted)', 
                                        fontWeight: 600 
                                    }}>
                                        <input
                                            type="checkbox"
                                            checked={applyFirstRowHeaderStyle}
                                            disabled={!applyTableDesignChecked}
                                            onChange={(e) => setApplyFirstRowHeaderStyle(e.target.checked)}
                                            style={{ 
                                                width: '16px', 
                                                height: '16px', 
                                                cursor: applyTableDesignChecked ? 'pointer' : 'not-allowed', 
                                                accentColor: '#a855f7' 
                                            }}
                                        />
                                        첫 번째 행(헤더) 특별 포맷팅 적용
                                    </label>
                                    <div style={{ fontSize: '12px', color: 'var(--text-muted)', lineHeight: '1.4', paddingLeft: '24px' }}>
                                        (첫 행 배경 RGB(0,114,186), 글자 흰색 11pt KoPub돋움체 Bold, 첫 행 내부 실선만 흰색 적용)
                                    </div>
                                </div>
                            </div>

                            {/* 옵션 F: 특수문자 일괄 변경 */}
                            <div style={{ background: 'var(--bg-secondary)', padding: '16px', borderRadius: '8px', border: '1px solid var(--panel-border)', display: 'flex', flexDirection: 'column', gap: '12px' }}>
                                <label style={{ display: 'flex', alignItems: 'center', gap: '10px', cursor: 'pointer', fontWeight: 600, fontSize: '14px', color: 'var(--text-primary)' }}>
                                    <input 
                                        id="checkbox-option-f"
                                        type="checkbox" 
                                        checked={applySpecialCharClean}
                                        onChange={(e) => setApplySpecialCharClean(e.target.checked)}
                                        style={{ width: '18px', height: '18px', cursor: 'pointer', accentColor: '#a855f7' }}
                                    />
                                    옵션 F: 특수문자 일괄 변경 기능 적용
                                </label>
                                <div style={{ paddingLeft: '28px', fontSize: '13px', color: 'var(--text-secondary)', lineHeight: '1.5' }}>
                                    💡 PPT 본문 텍스트 내 Non-Breaking Space(줄바꿈 없는 공백)와 중간점(·, •, - 등)을 정제합니다.
                                </div>
                                
                                <div style={{ 
                                    paddingLeft: '28px', 
                                    display: 'flex', 
                                    flexDirection: 'column', 
                                    gap: '12px', 
                                    borderLeft: '2px solid ' + (applySpecialCharClean ? 'rgba(168, 85, 247, 0.5)' : 'rgba(255, 255, 255, 0.1)'), 
                                    marginTop: '4px',
                                    opacity: applySpecialCharClean ? 1 : 0.45,
                                    pointerEvents: applySpecialCharClean ? 'auto' : 'none',
                                    transition: 'all 0.3s ease'
                                }}>
                                    <label style={{ 
                                        display: 'flex', 
                                        alignItems: 'center', 
                                        gap: '8px', 
                                        cursor: applySpecialCharClean ? 'pointer' : 'not-allowed', 
                                        fontSize: '13px', 
                                        color: applySpecialCharClean ? 'var(--text-primary)' : 'var(--text-muted)', 
                                        fontWeight: 600 
                                    }}>
                                        <input
                                            type="checkbox"
                                            checked={replaceNbs}
                                            disabled={!applySpecialCharClean}
                                            onChange={(e) => setReplaceNbs(e.target.checked)}
                                            style={{ 
                                                width: '16px', 
                                                height: '16px', 
                                                cursor: applySpecialCharClean ? 'pointer' : 'not-allowed', 
                                                accentColor: '#a855f7' 
                                            }}
                                        />
                                        줄바꿈 없는 공백 (Non-Breaking Space, \xa0) 일반 공백으로 변환
                                    </label>
                                    
                                    <label style={{ 
                                        display: 'flex', 
                                        alignItems: 'center', 
                                        gap: '8px', 
                                        cursor: applySpecialCharClean ? 'pointer' : 'not-allowed', 
                                        fontSize: '13px', 
                                        color: applySpecialCharClean ? 'var(--text-primary)' : 'var(--text-muted)', 
                                        fontWeight: 600 
                                    }}>
                                        <input
                                            type="checkbox"
                                            checked={unifyBullets}
                                            disabled={!applySpecialCharClean}
                                            onChange={(e) => setUnifyBullets(e.target.checked)}
                                            style={{ 
                                                width: '16px', 
                                                height: '16px', 
                                                cursor: applySpecialCharClean ? 'pointer' : 'not-allowed', 
                                                accentColor: '#a855f7' 
                                            }}
                                        />
                                        중간점 혼용 (· vs • vs -)을 •로 통일 (텍스트 내부의 구분용)
                                    </label>
                                    
                                    <label style={{ 
                                        display: 'flex', 
                                        alignItems: 'center', 
                                        gap: '8px', 
                                        cursor: applySpecialCharClean ? 'pointer' : 'not-allowed', 
                                        fontSize: '13px', 
                                        color: applySpecialCharClean ? 'var(--text-primary)' : 'var(--text-muted)', 
                                        fontWeight: 600 
                                    }}>
                                        <input
                                            id="checkbox-clean-vt"
                                            type="checkbox"
                                            checked={clean_vertical_tab}
                                            disabled={!applySpecialCharClean}
                                            onChange={(e) => set_clean_vertical_tab(e.target.checked)}
                                            style={{ 
                                                width: '16px', 
                                                height: '16px', 
                                                cursor: applySpecialCharClean ? 'pointer' : 'not-allowed', 
                                                accentColor: '#a855f7' 
                                            }}
                                        />
                                        세로 탭(Vertical Tab) 품질검토 오류 수정 (\v → \n 변환)
                                    </label>
                                </div>
                            </div>

                            {/* 옵션 G: 동일 제목 일련번호 추가 */}
                            <div style={{ background: 'var(--bg-secondary)', padding: '16px', borderRadius: '8px', border: '1px solid var(--panel-border)', display: 'flex', flexDirection: 'column', gap: '12px' }}>
                                <label style={{ display: 'flex', alignItems: 'center', gap: '10px', cursor: 'pointer', fontWeight: 600, fontSize: '14px', color: 'var(--text-primary)' }}>
                                    <input 
                                        id="checkbox-option-g"
                                        type="checkbox" 
                                        checked={add_title_page_numbers}
                                        onChange={(e) => set_add_title_page_numbers(e.target.checked)}
                                        style={{ width: '18px', height: '18px', cursor: 'pointer', accentColor: '#a855f7' }}
                                    />
                                    옵션 G: 동일 제목 슬라이드 일련번호(페이지 수) 자동 추가
                                </label>
                                <div style={{ paddingLeft: '28px', fontSize: '13px', color: 'var(--text-secondary)', lineHeight: '1.5' }}>
                                    💡 PPT 상단 제목이 동일하게 여러 장 연속될 경우, 자동으로 제목 뒤에 일련번호를 주입합니다.<br />
                                    <span style={{ color: 'var(--accent-blue)', fontWeight: 'bold' }}>예: "2.1 예시 제목"이 10장 연속 시 ➜ "2.1 예시 제목 (1/10)" ... "2.1 예시 제목 (10/10)"</span>
                                </div>
                                
                                {/* 하위 옵션: 괄호 앞 공백 추가 */}
                                <div style={{ 
                                    paddingLeft: '28px', 
                                    display: 'flex', 
                                    flexDirection: 'column', 
                                    gap: '8px', 
                                    borderLeft: '2px solid rgba(168, 85, 247, 0.5)', 
                                    marginTop: '4px',
                                    transition: 'all 0.3s ease'
                                }}>
                                    <label style={{ 
                                        display: 'flex', 
                                        alignItems: 'center', 
                                        gap: '8px', 
                                        cursor: 'pointer', 
                                        fontSize: '13px', 
                                        color: 'var(--text-primary)', 
                                        fontWeight: 600 
                                    }}>
                                        <input
                                            type="checkbox"
                                            checked={add_space_before_parenthesis}
                                            onChange={(e) => set_add_space_before_parenthesis(e.target.checked)}
                                            style={{ 
                                                width: '16px', 
                                                height: '16px', 
                                                cursor: 'pointer', 
                                                accentColor: '#a855f7' 
                                            }}
                                        />
                                        제목 뒷부분 괄호 기호 앞에 공백(스페이스) 추가
                                    </label>
                                    <div style={{ fontSize: '12px', color: 'var(--text-muted)', lineHeight: '1.4', paddingLeft: '24px' }}>
                                        (예: "2.1 예시 제목(1/2)" ➜ <span style={{ color: 'var(--accent-blue)', fontWeight: 'bold' }}>"2.1 예시 제목 (1/2)"</span>로 자동 스페이스 보정)
                                    </div>
                                </div>
                            </div>

                            {/* 옵션 H: 글자 색상 일괄 매핑 변경 */}
                            <div style={{ background: 'var(--bg-secondary)', padding: '16px', borderRadius: '8px', border: '1px solid var(--panel-border)', display: 'flex', flexDirection: 'column', gap: '12px' }}>
                                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                    <div style={{ padding: '6px', background: 'rgba(236, 72, 153, 0.1)', borderRadius: '6px' }}>
                                        <Layers size={16} color="#ec4899" />
                                    </div>
                                    <h3 style={{ margin: 0, fontSize: '15px', color: 'var(--text-primary)', fontWeight: 600 }}>
                                        옵션 H: 글자 색상 일괄 매핑 변경 (선택)
                                    </h3>
                                </div>
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                                    <textarea
                                        id="input-textcolor-rules"
                                        value={textColorRules}
                                        onChange={(e) => setTextColorRules(e.target.value)}
                                        placeholder="기존색상(변경할색상) 형식으로 입력하세요.&#10;예: rgb(255,0,0)(rgb(0,0,255)) 또는 255,0,0(0,0,255)"
                                        style={{
                                            width: '100%', height: '80px', padding: '10px 12px',
                                            background: 'rgba(0,0,0,0.2)', border: '1px solid var(--panel-border)',
                                            borderRadius: '6px', color: 'var(--text-primary)', fontSize: '13px',
                                            fontFamily: 'monospace', resize: 'vertical', outline: 'none',
                                            lineHeight: '1.5'
                                        }}
                                    />
                                    <div style={{ fontSize: '12.5px', color: 'var(--text-secondary)', lineHeight: '1.5' }}>
                                        💡 PPT 내 텍스트의 글자 색상을 찾아 일괄 변경합니다. 다중 규칙은 **줄바꿈(엔터)** 또는 **세미콜론(<code>;</code>)**으로 구분합니다.<br />
                                        <span style={{ color: 'var(--accent-blue)', fontWeight: 700 }}>형식: rgb(255,0,0)(rgb(0,0,255)) 또는 255,0,0(0,0,255)</span>
                                    </div>
                                    <div style={{ display: 'flex', alignItems: 'center', gap: '6px', marginTop: '4px' }}>
                                        <span style={{ display: 'inline-block', width: '6px', height: '6px', borderRadius: '50%', background: 'var(--success-color)' }}></span>
                                        <span style={{ fontSize: '12px', color: 'var(--success-color)', fontWeight: 600 }}>브라우저 자동 영구 보존(Auto-save) 적용됨</span>
                                    </div>
                                </div>
                            </div>

                            {/* 옵션 I: 단락 한글 단어 잘림 방지 (Hangul Word Wrap) */}
                            <div style={{ background: 'var(--bg-secondary)', padding: '16px', borderRadius: '8px', border: '1px solid var(--panel-border)', display: 'flex', flexDirection: 'column', gap: '12px' }}>
                                <label style={{ display: 'flex', alignItems: 'center', gap: '10px', cursor: 'pointer', fontWeight: 600, fontSize: '14px', color: 'var(--text-primary)' }}>
                                    <input 
                                        id="checkbox-option-i"
                                        type="checkbox" 
                                        checked={preventWordWrap}
                                        onChange={(e) => setPreventWordWrap(e.target.checked)}
                                        style={{ width: '18px', height: '18px', cursor: 'pointer', accentColor: '#a855f7' }}
                                    />
                                    옵션 I: 단락 한글 단어 잘림 방지 (Hangul Word Wrap) 적용
                                </label>
                                <div style={{ paddingLeft: '28px', fontSize: '13px', color: 'var(--text-secondary)', lineHeight: '1.5' }}>
                                    💡 글자 끝에서 단어가 뚝 끊겨서 줄바꿈되는 현상을 해결합니다. 한글 단어가 온전한 상태로 가독성 높게 다음 줄로 정렬됩니다.
                                </div>
                                <div style={{ paddingLeft: '28px', display: 'flex', alignItems: 'center', gap: '6px', marginTop: '4px' }}>
                                    <span style={{ display: 'inline-block', width: '6px', height: '6px', borderRadius: '50%', background: 'var(--success-color)' }}></span>
                                    <span style={{ fontSize: '12px', color: 'var(--success-color)', fontWeight: 600 }}>브라우저 자동 영구 보존(Auto-save) 적용됨</span>
                                </div>
                            </div>
                        </div>

                        <div style={{ marginTop: 'auto', paddingTop: '20px' }}>
                            <button
                                id="btn-batch-process"
                                className="interactive"
                                onClick={handleBatchProcess}
                                disabled={batchPptFiles.length === 0 || isProcessingBatch || (!replaceRules.trim() && !fontRules.trim() && !fontSize.trim() && !applyDesignChecked && !applyTableDesignChecked && !applySpecialCharClean && !add_title_page_numbers && !add_space_before_parenthesis && !textColorRules.trim() && !preventWordWrap)}
                                style={{
                                    width: '100%',
                                    padding: '16px',
                                    background: (batchPptFiles.length === 0 || isProcessingBatch || (!replaceRules.trim() && !fontRules.trim() && !fontSize.trim() && !applyDesignChecked && !applyTableDesignChecked && !applySpecialCharClean && !add_title_page_numbers && !add_space_before_parenthesis && !textColorRules.trim() && !preventWordWrap)) ? 'rgba(255,255,255,0.05)' : 'linear-gradient(135deg, #a855f7, #3b82f6)',
                                    color: (batchPptFiles.length === 0 || isProcessingBatch || (!replaceRules.trim() && !fontRules.trim() && !fontSize.trim() && !applyDesignChecked && !applyTableDesignChecked && !applySpecialCharClean && !add_title_page_numbers && !add_space_before_parenthesis && !textColorRules.trim() && !preventWordWrap)) ? 'var(--text-muted)' : 'white',
                                    border: 'none',
                                    borderRadius: '12px',
                                    fontSize: '16px',
                                    fontWeight: 700,
                                    cursor: (batchPptFiles.length === 0 || isProcessingBatch || (!replaceRules.trim() && !fontRules.trim() && !fontSize.trim() && !applyDesignChecked && !applyTableDesignChecked && !applySpecialCharClean && !add_title_page_numbers && !add_space_before_parenthesis && !textColorRules.trim() && !preventWordWrap)) ? 'not-allowed' : 'pointer',
                                    display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '10px'
                                }}
                            >
                                {isProcessingBatch ? (
                                    <><Loader2 size={20} className="animate-spin" /> 폴더에 순차적으로 적용 및 저장 중...</>
                                ) : (
                                    <><Play size={20} /> 저장할 폴더 선택 및 일괄 편집 실행</>
                                )}
                            </button>
                            
                            {batchReport.length > 0 && (
                                <div className="animate-slide-up" style={{ 
                                    marginTop: '20px', 
                                    background: 'rgba(255,255,255,0.02)', 
                                    border: '1px solid var(--panel-border)', 
                                    borderRadius: '12px', 
                                    padding: '16px',
                                    display: 'flex',
                                    flexDirection: 'column',
                                    gap: '12px'
                                }}>
                                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px', borderBottom: '1px solid var(--panel-border)', paddingBottom: '8px' }}>
                                        <Info size={16} color="#a855f7" />
                                        <span style={{ fontWeight: 700, fontSize: '14px', color: 'var(--text-primary)' }}>일괄 편집 세부 처리 결과 리포트</span>
                                    </div>
                                    <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', maxHeight: '220px', overflowY: 'auto', paddingRight: '4px' }}>
                                        {batchReport.map((rep, idx) => (
                                            <div key={idx} style={{ 
                                                display: 'flex', 
                                                flexDirection: 'column', 
                                                gap: '6px', 
                                                padding: '10px', 
                                                borderRadius: '8px', 
                                                background: rep.status === 'success' ? 'rgba(16, 185, 129, 0.05)' : 'rgba(239, 68, 68, 0.05)',
                                                borderLeft: `3px solid ${rep.status === 'success' ? '#10b981' : '#ef4444'}`
                                            }}>
                                                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                                                    <span style={{ fontSize: '13px', fontWeight: 600, color: 'var(--text-primary)', wordBreak: 'break-all' }}>{rep.fileName}</span>
                                                    <span style={{ 
                                                        fontSize: '11px', 
                                                        fontWeight: 700, 
                                                        padding: '2px 6px', 
                                                        borderRadius: '4px',
                                                        background: rep.status === 'success' ? 'rgba(16, 185, 129, 0.15)' : 'rgba(239, 68, 68, 0.15)',
                                                        color: rep.status === 'success' ? '#10b981' : '#ef4444'
                                                    }}>
                                                        {rep.status === 'success' ? '수정 완료' : '실패'}
                                                    </span>
                                                </div>
                                                <div style={{ fontSize: '12px', color: 'var(--text-secondary)', lineHeight: '1.4' }}>
                                                    {rep.detail}
                                                </div>
                                            </div>
                                        ))}
                                    </div>
                                </div>
                            )}
                        </div>
                    </div>
                ) : activeTab === 'smart_animation' ? (
                    <div className="animate-fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
                        {/* 스마트 애니메이션 업로드 영역 */}
                        <div 
                            onDragOver={(e) => { e.preventDefault(); e.stopPropagation(); }}
                            onDragEnter={(e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingAnimation(true); }}
                            onDragLeave={(e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingAnimation(false); }}
                            onDrop={(e) => {
                                e.preventDefault(); e.stopPropagation(); setIsDraggingAnimation(false);
                                processAnimationFile(e.dataTransfer.files[0]);
                            }}
                            style={{ 
                                border: `2px dashed ${isDraggingAnimation ? '#a855f7' : 'rgba(168, 85, 247, 0.3)'}`, 
                                borderRadius: '12px', padding: '32px',
                                background: isDraggingAnimation ? 'rgba(168, 85, 247, 0.1)' : 'rgba(168, 85, 247, 0.05)', 
                                display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', gap: '16px',
                                transition: 'all 0.2s ease', cursor: isDraggingAnimation ? 'copy' : 'default',
                                minHeight: '200px'
                            }}
                        >
                            <div style={{ padding: '16px', background: 'rgba(168, 85, 247, 0.1)', borderRadius: '50%' }}>
                                <Presentation size={32} color="#a855f7" />
                            </div>
                            <div style={{ textAlign: 'center' }}>
                                <h3 style={{ margin: '0 0 8px 0', fontSize: '18px', color: 'var(--text-primary)' }}>원본 PPT 파일 등록</h3>
                                <p style={{ margin: 0, fontSize: '14px', color: 'var(--text-secondary)' }}>애니메이션을 추가할 .pptx 파일을 드래그하거나 선택하세요.</p>
                            </div>
                            
                            <input
                                type="file"
                                accept=".pptx"
                                onChange={(e) => processAnimationFile(e.target.files[0])}
                                ref={animInputRef}
                                style={{ display: 'none' }}
                            />
                            
                            <div style={{ display: 'flex', gap: '12px', alignItems: 'center' }}>
                                <button
                                    onClick={() => animInputRef.current?.click()}
                                    className="interactive"
                                    style={{
                                        display: 'flex', alignItems: 'center', gap: '8px',
                                        padding: '10px 20px', borderRadius: '8px', cursor: 'pointer',
                                        background: 'var(--bg-secondary)', border: '1px solid var(--panel-border)',
                                        color: 'var(--text-primary)', fontWeight: 600, fontSize: '14px'
                                    }}
                                >
                                    <Upload size={18} /> PPT 파일 찾기
                                </button>
                                <button 
                                    className={`interactive ${animationType === 'transition' ? 'active' : ''}`}
                                    onClick={() => updateAllSlideAnimations(null, 'transition', null)}
                                    style={{
                                        flex: 1, padding: '12px', borderRadius: '8px',
                                        background: animationType === 'transition' ? 'rgba(168, 85, 247, 0.2)' : 'rgba(255,255,255,0.05)',
                                        color: animationType === 'transition' ? '#c084fc' : 'var(--text-secondary)',
                                        cursor: 'pointer', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '8px',
                                        border: `1px solid ${animationType === 'transition' ? '#a855f7' : 'transparent'}`,
                                        transition: 'all 0.2s'
                                    }}
                                >
                                    <Presentation size={24} />
                                    <div style={{ fontSize: '13px', fontWeight: 600 }}>슬라이드 전환</div>
                                </button>
                                <button 
                                    className={`interactive ${animationType === 'fade' ? 'active' : ''}`}
                                    onClick={() => updateAllSlideAnimations(null, 'fade', null)}
                                    style={{
                                        flex: 1, padding: '12px', borderRadius: '8px',
                                        background: animationType === 'fade' ? 'rgba(168, 85, 247, 0.2)' : 'rgba(255,255,255,0.05)',
                                        color: animationType === 'fade' ? '#c084fc' : 'var(--text-secondary)',
                                        cursor: 'pointer', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '8px',
                                        border: `1px solid ${animationType === 'fade' ? '#a855f7' : 'transparent'}`,
                                        transition: 'all 0.2s'
                                    }}
                                >
                                    <Sparkles size={24} />
                                    <div style={{ fontSize: '13px', fontWeight: 600 }}>객체 페이드</div>
                                </button>
                                <button 
                                    className={`interactive ${animationType === 'appear' ? 'active' : ''}`}
                                    onClick={() => updateAllSlideAnimations(null, 'appear', null)}
                                    style={{
                                        flex: 1, padding: '12px', borderRadius: '8px',
                                        background: animationType === 'appear' ? 'rgba(168, 85, 247, 0.2)' : 'rgba(255,255,255,0.05)',
                                        color: animationType === 'appear' ? '#c084fc' : 'var(--text-secondary)',
                                        cursor: 'pointer', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '8px',
                                        border: `1px solid ${animationType === 'appear' ? '#a855f7' : 'transparent'}`,
                                        transition: 'all 0.2s'
                                    }}
                                >
                                    <MousePointer2 size={24} />
                                    <div style={{ fontSize: '13px', fontWeight: 600 }}>객체 나타나기</div>
                                </button>
                                {animationPptFile && (
                                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px', background: 'rgba(168, 85, 247, 0.1)', padding: '8px 16px', borderRadius: '8px', fontSize: '14px', border: '1px solid #a855f7', color: '#c084fc' }}>
                                        <CheckCircle2 size={16} />
                                        {animationPptFile.name}
                                    </div>
                                )}
                            </div>
                        </div>

                        {/* 옵션 설정 */}
                        <div style={{ 
                            border: '1px solid var(--panel-border)', 
                            borderRadius: '12px', padding: '24px',
                            background: 'rgba(255, 255, 255, 0.02)', 
                            display: 'flex', flexDirection: 'column', gap: '20px'
                        }}>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                <Sparkles size={20} color="#a855f7" />
                                <h3 style={{ margin: 0, fontSize: '16px', color: 'var(--text-primary)' }}>애니메이션 스타일 설정</h3>
                            </div>
                            
                            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
                                <button
                                    onClick={() => setAnimationType('fade')}
                                    style={{
                                        padding: '16px', borderRadius: '12px', cursor: 'pointer',
                                        background: animationType === 'fade' ? 'rgba(168, 85, 247, 0.15)' : 'rgba(0,0,0,0.1)',
                                        border: `1px solid ${animationType === 'fade' ? '#a855f7' : 'var(--panel-border)'}`,
                                        color: animationType === 'fade' ? '#c084fc' : 'var(--text-secondary)',
                                        display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '8px', transition: 'all 0.2s'
                                    }}
                                >
                                    <div style={{ fontWeight: 700, fontSize: '15px' }}>페이드 (Fade)</div>
                                    <div style={{ fontSize: '12px', opacity: 0.8 }}>객체가 부드럽게 나타납니다. (권장)</div>
                                </button>
                                <button
                                    onClick={() => setAnimationType('appear')}
                                    style={{
                                        padding: '16px', borderRadius: '12px', cursor: 'pointer',
                                        background: animationType === 'appear' ? 'rgba(168, 85, 247, 0.15)' : 'rgba(0,0,0,0.1)',
                                        border: `1px solid ${animationType === 'appear' ? '#a855f7' : 'var(--panel-border)'}`,
                                        color: animationType === 'appear' ? '#c084fc' : 'var(--text-secondary)',
                                        display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '8px', transition: 'all 0.2s'
                                    }}
                                >
                                    <div style={{ fontWeight: 700, fontSize: '15px' }}>나타나기 (Appear)</div>
                                    <div style={{ fontSize: '12px', opacity: 0.8 }}>객체가 즉시 나타납니다.</div>
                                </button>
                            </div>

                            {animationType !== 'transition' && (
                                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '12px', background: 'rgba(255,255,255,0.05)', borderRadius: '8px' }}>
                                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                        <Box size={18} color="#a855f7" />
                                        <div>
                                            <div style={{ fontSize: '14px', fontWeight: 600, color: 'var(--text-primary)' }}>객체 그룹화 (한번에 나타나기)</div>
                                            <div style={{ fontSize: '12px', color: 'var(--text-secondary)' }}>슬라이드의 모든 객체가 클릭 한 번에 동시에 나타납니다.</div>
                                        </div>
                                    </div>
                                    <button 
                                        onClick={() => updateAllSlideAnimations(null, null, !useGrouping)}
                                        style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 0, color: useGrouping ? '#a855f7' : 'var(--text-muted)' }}
                                    >
                                        {useGrouping ? <ToggleRight size={32} /> : <ToggleLeft size={32} />}
                                    </button>
                                </div>
                            )}

                            <div style={{ 
                                padding: '16px', background: 'rgba(59, 130, 246, 0.05)', 
                                borderLeft: '4px solid var(--accent-blue)', borderRadius: '4px',
                                fontSize: '13px', color: 'var(--text-secondary)', lineHeight: 1.6
                            }}>
                                {animationType === 'transition' ? (
                                    <>💡 <strong>전환 효과:</strong> 객체를 수정하지 않고 슬라이드가 넘어갈 때 부드러운 <strong>페이드 인</strong> 효과를 적용합니다. 가장 안전하고 깔끔한 방식입니다.</>
                                ) : (
                                    <>💡 <strong>작동 원리:</strong> <strong>'그룹화'</strong>를 켜면 슬라이드 전체가 한 번에 나타나며, 끄면 객체들이 <strong>상단 → 하단</strong> 순서대로 하나씩 나타납니다.</>
                                )}
                            </div>
                        </div>

                        {/* 페이지별 애니메이션 설정 (파일 업로드 후 노출) */}
                        {animationPptFile && slideAnimations.length > 0 && (
                            <div className="animate-fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
                                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                        <Layers size={18} color="#a855f7" />
                                        <h3 style={{ margin: 0, fontSize: '15px', color: 'var(--text-primary)' }}>페이지별 개별 설정 ({slideAnimations.length} 슬라이드)</h3>
                                    </div>
                                    <div style={{ display: 'flex', gap: '8px' }}>
                                        <button 
                                            onClick={() => updateAllSlideAnimations(true, null)}
                                            style={{ padding: '4px 8px', fontSize: '12px', background: 'rgba(168, 85, 247, 0.1)', border: '1px solid #a855f7', color: '#c084fc', borderRadius: '4px', cursor: 'pointer' }}
                                        >전체 켜기</button>
                                        <button 
                                            onClick={() => updateAllSlideAnimations(false, null)}
                                            style={{ padding: '4px 8px', fontSize: '12px', background: 'rgba(255,255,255,0.05)', border: '1px solid var(--panel-border)', color: 'var(--text-secondary)', borderRadius: '4px', cursor: 'pointer' }}
                                        >전체 끄기</button>
                                    </div>
                                </div>

                                <div style={{ 
                                    display: 'grid', 
                                    gridTemplateColumns: 'repeat(auto-fill, minmax(200px, 1fr))', 
                                    gap: '12px',
                                    maxHeight: '400px',
                                    overflowY: 'auto',
                                    padding: '4px'
                                }}>
                                    {slideAnimations.map((anim, idx) => (
                                        <div key={idx} style={{ 
                                            background: anim.enabled ? 'rgba(168, 85, 247, 0.05)' : 'rgba(0,0,0,0.1)', 
                                            border: `1px solid ${anim.enabled ? 'rgba(168, 85, 247, 0.3)' : 'var(--panel-border)'}`,
                                            borderRadius: '10px',
                                            padding: '12px',
                                            display: 'flex',
                                            flexDirection: 'column',
                                            gap: '10px',
                                            transition: 'all 0.2s'
                                        }}>
                                            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                                                <span style={{ fontSize: '13px', fontWeight: 700, color: anim.enabled ? '#c084fc' : 'var(--text-muted)' }}>Slide {idx + 1}</span>
                                                <button 
                                                    onClick={() => updateSingleSlideAnimation(idx, { enabled: !anim.enabled })}
                                                    style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 0, color: anim.enabled ? '#a855f7' : 'var(--text-muted)' }}
                                                >
                                                    {anim.enabled ? <ToggleRight size={24} /> : <ToggleLeft size={24} />}
                                                </button>
                                            </div>
                                            
                                            {anim.enabled && (
                                                <select 
                                                    value={anim.type}
                                                    onChange={(e) => updateSingleSlideAnimation(idx, { type: e.target.value })}
                                                    style={{ 
                                                        background: 'rgba(0,0,0,0.2)', 
                                                        color: 'white', 
                                                        border: '1px solid var(--panel-border)', 
                                                        borderRadius: '4px', 
                                                        fontSize: '12px', 
                                                        padding: '4px' 
                                                    }}
                                                >
                                                    <option value="transition">슬라이드 전환 (Transition)</option>
                                                    <option value="fade">객체 페이드 (Object Fade)</option>
                                                    <option value="appear">객체 나타나기 (Object Appear)</option>
                                                </select>
                                            )}

                                            {anim.enabled && anim.type !== 'transition' && (
                                                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginTop: '4px' }}>
                                                    <span style={{ fontSize: '11px', color: 'var(--text-secondary)' }}>그룹화</span>
                                                    <button 
                                                        onClick={() => updateSingleSlideAnimation(idx, { useGrouping: !anim.useGrouping })}
                                                        style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 0, color: anim.useGrouping ? '#a855f7' : 'var(--text-muted)' }}
                                                    >
                                                        {anim.useGrouping ? <ToggleRight size={20} /> : <ToggleLeft size={20} />}
                                                    </button>
                                                </div>
                                            )}
                                        </div>
                                    ))}
                                </div>
                            </div>
                        )}

                        <div style={{ marginTop: 'auto', paddingTop: '10px' }}>
                            <button
                                className="interactive"
                                onClick={handleAnimationProcess}
                                disabled={!animationPptFile || isProcessingAnimation}
                                style={{
                                    width: '100%',
                                    padding: '18px',
                                    background: (!animationPptFile || isProcessingAnimation) ? 'rgba(255,255,255,0.05)' : 'linear-gradient(135deg, #a855f7, #6366f1)',
                                    color: (!animationPptFile || isProcessingAnimation) ? 'var(--text-muted)' : 'white',
                                    border: 'none',
                                    borderRadius: '12px',
                                    fontSize: '17px',
                                    fontWeight: 800,
                                    cursor: (!animationPptFile || isProcessingAnimation) ? 'not-allowed' : 'pointer',
                                    display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '12px',
                                    boxShadow: (!animationPptFile || isProcessingAnimation) ? 'none' : '0 8px 20px rgba(168, 85, 247, 0.2)'
                                }}
                            >
                                {isProcessingAnimation ? (
                                    <><Loader2 size={22} className="animate-spin" /> 전 슬라이드 객체 분석 및 애니메이션 주입 중...</>
                                ) : (
                                    <><Play size={22} /> 스마트 애니메이션 적용 및 저장</>
                                )}
                            </button>
                        </div>
                    </div>
                ) : activeTab === 'pptx_hwpx_fusion' ? (
                    <div className="animate-fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
                        {/* 1. 이중 파일 업로드 드롭존 (PPTX + HWPX) */}
                        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
                            {/* PPTX 업로드 */}
                            <div style={{ background: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--panel-border)', display: 'flex', flexDirection: 'column', gap: '16px' }}>
                                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                    <div style={{ padding: '6px', background: 'rgba(59, 130, 246, 0.1)', borderRadius: '6px' }}>
                                        <Presentation size={16} color="var(--accent-color)" />
                                    </div>
                                    <h3 style={{ margin: 0, fontSize: '15px', color: 'var(--text-primary)', fontWeight: 700 }}>1. PPT 요건기술서 파일 (.pptx)</h3>
                                </div>
                                <div 
                                    onDragOver={(e) => { e.preventDefault(); setIsDraggingFusionPpt(true); }}
                                    onDragLeave={() => setIsDraggingFusionPpt(false)}
                                    onDrop={(e) => {
                                        e.preventDefault();
                                        setIsDraggingFusionPpt(false);
                                        const file = e.dataTransfer.files[0];
                                        if (file && file.name.toLowerCase().endsWith('.pptx')) {
                                            setHwpxFusionPptFile(file);
                                            setFusionResultStats(null);
                                            setErrorMsg(null);
                                            setSuccessMsg(null);
                                        } else {
                                            setErrorMsg('파워포인트 파일(.pptx)만 업로드할 수 있습니다.');
                                        }
                                    }}
                                    onClick={() => fusionPptInputRef.current?.click()}
                                    style={{
                                        border: `2px dashed ${isDraggingFusionPpt ? 'var(--accent-color)' : 'var(--panel-border)'}`,
                                        borderRadius: '10px',
                                        padding: '36px 20px',
                                        textAlign: 'center',
                                        cursor: 'pointer',
                                        background: isDraggingFusionPpt ? 'rgba(59, 130, 246, 0.05)' : 'rgba(0,0,0,0.15)',
                                        transition: 'all 0.2s',
                                        display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '12px'
                                    }}
                                >
                                    <input 
                                        type="file"
                                        ref={fusionPptInputRef}
                                        onChange={(e) => {
                                            const file = e.target.files[0];
                                            if (file) {
                                                setHwpxFusionPptFile(file);
                                                setFusionResultStats(null);
                                                setErrorMsg(null);
                                                setSuccessMsg(null);
                                            }
                                        }}
                                        accept=".pptx"
                                        style={{ display: 'none' }}
                                    />
                                    <div style={{ padding: '12px', background: 'rgba(255,255,255,0.03)', borderRadius: '50%' }}>
                                        <Upload size={30} color={hwpxFusionPptFile ? 'var(--accent-color)' : 'var(--text-muted)'} />
                                    </div>
                                    <div>
                                        <p style={{ margin: 0, fontSize: '13.5px', fontWeight: 600, color: 'var(--text-primary)' }}>
                                            {hwpxFusionPptFile ? hwpxFusionPptFile.name : 'PPTX 요건기술서 드래그 또는 클릭'}
                                        </p>
                                        <p style={{ margin: '4px 0 0', fontSize: '11.5px', color: 'var(--text-muted)' }}>
                                            {hwpxFusionPptFile ? `${(hwpxFusionPptFile.size / 1024 / 1024).toFixed(2)} MB` : '변환 데이터 소스가 될 PPTX 파일 등록'}
                                        </p>
                                    </div>
                                    {hwpxFusionPptFile && (
                                        <button 
                                            onClick={(e) => {
                                                e.stopPropagation();
                                                setHwpxFusionPptFile(null);
                                                setFusionResultStats(null);
                                                setErrorMsg(null);
                                                setSuccessMsg(null);
                                            }}
                                            className="interactive"
                                            style={{
                                                padding: '4px 8px', background: 'rgba(239,68,68,0.1)', color: 'var(--danger-color)',
                                                border: '1px solid rgba(239,68,68,0.2)', borderRadius: '4px', fontSize: '11.5px',
                                                cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '4px'
                                            }}
                                        >
                                            <X size={12} /> 파일 제거
                                        </button>
                                    )}
                                </div>
                            </div>

                            {/* HWPX 업로드 */}
                            <div style={{ background: 'var(--bg-secondary)', padding: '24px', borderRadius: '12px', border: '1px solid var(--panel-border)', display: 'flex', flexDirection: 'column', gap: '16px' }}>
                                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                    <div style={{ padding: '6px', background: 'rgba(16, 185, 129, 0.1)', borderRadius: '6px' }}>
                                        <Box size={16} color="var(--success-color)" />
                                    </div>
                                    <h3 style={{ margin: 0, fontSize: '15px', color: 'var(--text-primary)', fontWeight: 700 }}>2. HWPX 표준 양식 템플릿 (.hwpx)</h3>
                                </div>
                                <div 
                                    onDragOver={(e) => { e.preventDefault(); setIsDraggingFusionTemplate(true); }}
                                    onDragLeave={() => setIsDraggingFusionTemplate(false)}
                                    onDrop={(e) => {
                                        e.preventDefault();
                                        setIsDraggingFusionTemplate(false);
                                        const file = e.dataTransfer.files[0];
                                        if (file && file.name.toLowerCase().endsWith('.hwpx')) {
                                            setHwpxFusionTemplateFile(file);
                                            setFusionResultStats(null);
                                            setErrorMsg(null);
                                            setSuccessMsg(null);
                                        } else {
                                            setErrorMsg('한글 표준 문서(.hwpx)만 업로드할 수 있습니다.');
                                        }
                                    }}
                                    onClick={() => fusionTemplateInputRef.current?.click()}
                                    style={{
                                        border: `2px dashed ${isDraggingFusionTemplate ? 'var(--success-color)' : 'var(--panel-border)'}`,
                                        borderRadius: '10px',
                                        padding: '36px 20px',
                                        textAlign: 'center',
                                        cursor: 'pointer',
                                        background: isDraggingFusionTemplate ? 'rgba(16, 185, 129, 0.05)' : 'rgba(0,0,0,0.15)',
                                        transition: 'all 0.2s',
                                        display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '12px'
                                    }}
                                >
                                    <input 
                                        type="file"
                                        ref={fusionTemplateInputRef}
                                        onChange={(e) => {
                                            const file = e.target.files[0];
                                            if (file) {
                                                setHwpxFusionTemplateFile(file);
                                                setFusionResultStats(null);
                                                setErrorMsg(null);
                                                setSuccessMsg(null);
                                            }
                                        }}
                                        accept=".hwpx"
                                        style={{ display: 'none' }}
                                    />
                                    <div style={{ padding: '12px', background: 'rgba(255,255,255,0.03)', borderRadius: '50%' }}>
                                        <Upload size={30} color={hwpxFusionTemplateFile ? 'var(--success-color)' : 'var(--text-muted)'} />
                                    </div>
                                    <div>
                                        <p style={{ margin: 0, fontSize: '13.5px', fontWeight: 600, color: 'var(--text-primary)' }}>
                                            {hwpxFusionTemplateFile ? hwpxFusionTemplateFile.name : 'HWPX 템플릿 드래그 또는 클릭'}
                                        </p>
                                        <p style={{ margin: '4px 0 0', fontSize: '11.5px', color: 'var(--text-muted)' }}>
                                            {hwpxFusionTemplateFile ? `${(hwpxFusionTemplateFile.size / 1024).toFixed(2)} KB` : '데이터를 채워넣을 HWPX 양식 등록'}
                                        </p>
                                    </div>
                                    {hwpxFusionTemplateFile && (
                                        <button 
                                            onClick={(e) => {
                                                e.stopPropagation();
                                                setHwpxFusionTemplateFile(null);
                                                setFusionResultStats(null);
                                                setErrorMsg(null);
                                                setSuccessMsg(null);
                                            }}
                                            className="interactive"
                                            style={{
                                                padding: '4px 8px', background: 'rgba(239,68,68,0.1)', color: 'var(--danger-color)',
                                                border: '1px solid rgba(239,68,68,0.2)', borderRadius: '4px', fontSize: '11.5px',
                                                cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '4px'
                                            }}
                                        >
                                            <X size={12} /> 파일 제거
                                        </button>
                                    )}
                                </div>
                            </div>
                        </div>

                        {/* 2. 생성 제어 방식 설정 패널 */}
                        <div style={{ background: 'var(--bg-secondary)', padding: '20px', borderRadius: '12px', border: '1px solid var(--panel-border)', display: 'flex', flexDirection: 'column', gap: '12px' }}>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                <Settings size={16} color="var(--accent-color)" />
                                <span style={{ fontSize: '14.5px', fontWeight: 700, color: 'var(--text-primary)' }}>HWPX 표준산출물 빌드 옵션</span>
                            </div>
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px', paddingLeft: '4px' }}>
                                <label style={{ display: 'flex', alignItems: 'center', gap: '10px', cursor: 'pointer' }}>
                                    <input 
                                        type="radio" 
                                        name="fusion_merge_mode" 
                                        checked={fusionMergeMode === true} 
                                        onChange={() => setFusionMergeMode(true)}
                                        style={{ accentColor: 'var(--accent-color)', width: '16px', height: '16px' }}
                                    />
                                    <div>
                                        <span style={{ fontSize: '13.5px', fontWeight: 600, color: 'var(--text-primary)' }}>하나의 완결된 HWPX 문서로 통합 병합 (권장)</span>
                                        <p style={{ margin: '2px 0 0', fontSize: '12px', color: 'var(--text-muted)' }}>각 요구사항 건수를 한글 표로 각각 복제 및 쪽 나누기(Page Break)를 가동하여 1개의 완성도 높은 한글 파일로 출력합니다.</p>
                                    </div>
                                </label>
                                <div style={{ height: '1px', background: 'var(--panel-border)', margin: '4px 0' }} />
                                <label style={{ display: 'flex', alignItems: 'center', gap: '10px', cursor: 'pointer' }}>
                                    <input 
                                        type="radio" 
                                        name="fusion_merge_mode" 
                                        checked={fusionMergeMode === false} 
                                        onChange={() => setFusionMergeMode(false)}
                                        style={{ accentColor: 'var(--accent-color)', width: '16px', height: '16px' }}
                                    />
                                    <div>
                                        <span style={{ fontSize: '13.5px', fontWeight: 600, color: 'var(--text-primary)' }}>요구사항 고유번호별 개별 HWPX 파일들로 저장 (ZIP 압축)</span>
                                        <p style={{ margin: '2px 0 0', fontSize: '12px', color: 'var(--text-muted)' }}>고유번호별로 양식이 주입된 다수의 독립된 HWPX 파일을 생성하고 하나의 ZIP 파일로 압축하여 다운로드합니다.</p>
                                    </div>
                                </label>
                            </div>
                        </div>

                        {/* 3. 데이터 맵핑 및 스마트 안내 규칙 */}
                        <div style={{ background: 'rgba(250,204,21,0.03)', padding: '16px', borderRadius: '8px', border: '1px solid rgba(250,204,21,0.15)', display: 'flex', flexDirection: 'column', gap: '8px' }}>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', color: 'var(--accent-gold)' }}>
                                <Info size={16} />
                                <span style={{ fontSize: '13.5px', fontWeight: 700 }}>HWPX 표준산출물 매핑 정보 및 개행 가이드</span>
                            </div>
                            <ul style={{ margin: 0, paddingLeft: '20px', fontSize: '12.5px', color: 'var(--text-secondary)', lineHeight: '1.7', display: 'flex', flexDirection: 'column', gap: '4px' }}>
                                <li>**매핑 정보**: PPT 표의 **고유번호 ➜ 요구사항 고유번호**, **요구사항 명칭 ➜ 요구사항 명칭**, **정의 ➜ 정의**, **요건 세부내용 ➜ 세부내용**으로 지능 자동 매핑됩니다.</li>
                                <li>**가독성 보정**: 긴 개행 텍스트의 경우 HWPX 셀 내부에 여러 개의 단락(&lt;hp:p&gt;)을 동적 생성하여 한글 오피스 뷰어 상의 완전한 가폭 및 줄바꿈을 완벽히 보장합니다.</li>
                                <li>**스마트 필터**: 첫 번째 열이 `ECR-XXX`처럼 고유번호 패턴을 가지고 있는 4개 열 구조의 표 행만 지능적으로 골라내어 노이즈(헤더나 인트로)를 사전 격리합니다.</li>
                            </ul>
                        </div>

                        {/* 4. 융합 결과 통계 피드백 리포트 */}
                        {fusionResultStats && (
                            <div className="animate-scale-in" style={{ background: 'rgba(59,130,246,0.03)', padding: '20px', borderRadius: '12px', border: '1px solid rgba(59,130,246,0.2)', display: 'flex', flexDirection: 'column', gap: '12px' }}>
                                <h4 style={{ margin: 0, fontSize: '14.5px', color: 'var(--accent-color)', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '6px' }}>
                                    <CheckCircle2 size={16} /> HWPX 양식 융합 완료 통계 리포트
                                </h4>
                                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
                                    <div style={{ background: 'rgba(0,0,0,0.2)', padding: '12px', borderRadius: '8px', textAlign: 'center', border: '1px solid var(--panel-border)' }}>
                                        <div style={{ fontSize: '11.5px', color: 'var(--text-muted)' }}>융합된 요구사항 건수</div>
                                        <div style={{ fontSize: '20px', fontWeight: 800, color: 'var(--accent-color)', marginTop: '4px' }}>{fusionResultStats.fusedCount}개 건</div>
                                    </div>
                                    <div style={{ background: 'rgba(0,0,0,0.2)', padding: '12px', borderRadius: '8px', textAlign: 'center', border: '1px solid var(--panel-border)' }}>
                                        <div style={{ fontSize: '11.5px', color: 'var(--text-muted)' }}>최종 저장 방식</div>
                                        <div style={{ fontSize: '20px', fontWeight: 800, color: 'var(--success-color)', marginTop: '4px' }}>
                                            {fusionResultStats.mode === 'merge' ? '통합 HWPX 문서 병합' : '개별 HWPX 분할 ZIP 압축'}
                                        </div>
                                    </div>
                                </div>
                                <div style={{ marginTop: '8px' }}>
                                    <div style={{ fontSize: '12.5px', color: 'var(--text-secondary)', fontWeight: 600, marginBottom: '6px' }}>융합 요구사항 고유번호 목록:</div>
                                    <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px', maxHeight: '100px', overflowY: 'auto', padding: '6px', background: 'rgba(0,0,0,0.3)', borderRadius: '6px' }}>
                                        {fusionResultStats.requirementsList.map((req, idx) => (
                                            <span key={idx} style={{ padding: '2px 8px', background: 'rgba(59, 130, 246, 0.1)', color: 'var(--accent-color)', borderRadius: '4px', fontSize: '11px', fontWeight: 600 }}>
                                                {req.id}
                                            </span>
                                        ))}
                                    </div>
                                </div>
                            </div>
                        )}

                        {/* 5. 융합 실행 버튼 */}
                        <div style={{ marginTop: '10px' }}>
                            <button
                                className="interactive"
                                onClick={handleHwpxFusion}
                                disabled={!hwpxFusionPptFile || !hwpxFusionTemplateFile || isFusingHwpx}
                                style={{
                                    width: '100%',
                                    padding: '18px',
                                    background: (!hwpxFusionPptFile || !hwpxFusionTemplateFile || isFusingHwpx) ? 'rgba(255,255,255,0.05)' : 'linear-gradient(135deg, #3b82f6, #6366f1)',
                                    color: (!hwpxFusionPptFile || !hwpxFusionTemplateFile || isFusingHwpx) ? 'var(--text-muted)' : 'white',
                                    border: 'none',
                                    borderRadius: '12px',
                                    fontSize: '17px',
                                    fontWeight: 800,
                                    cursor: (!hwpxFusionPptFile || !hwpxFusionTemplateFile || isFusingHwpx) ? 'not-allowed' : 'pointer',
                                    display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '12px',
                                    boxShadow: (!hwpxFusionPptFile || !hwpxFusionTemplateFile || isFusingHwpx) ? 'none' : '0 8px 20px rgba(59, 130, 246, 0.2)'
                                }}
                            >
                                {isFusingHwpx ? (
                                    <><Loader2 size={22} className="animate-spin" /> 양식 딥 복제 및 데이터 융합 중...</>
                                ) : (
                                    <><Sparkles size={22} /> HWPX 표준 양식 융합 및 산출물 다운로드</>
                                )}
                            </button>
                        </div>
                    </div>
                ) : null}
            </div>
        </div>
    );
}

