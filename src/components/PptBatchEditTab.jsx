import React, { useState, useEffect } from 'react';
import { Presentation, Upload, X, Settings, CheckCircle2, Layers, Loader2, Sparkles, Info, FilePlus2, ArrowUpDown, ChevronLeft, ChevronRight, Trash2 } from 'lucide-react';
import { processPptBatch, mergeMultiplePptx } from '../utils/pptExporter';
import JSZip from 'jszip';

// 자연어(숫자 포함) 파일명 정렬 헬퍼 함수 (예: IV.4.1 -> IV.4.2 -> IV.4.10)
const sortFilesNaturally = (files, ascending = true) => {
    return [...files].sort((a, b) => {
        const res = a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' });
        return ascending ? res : -res;
    });
};

// PPT 일괄 편집(단어 수정 + 디자인 변경) 탭. PptGenerator.jsx의 batch_edit 서브탭에서 분리됨.
export default function PptBatchEditTab({ setErrorMsg, setSuccessMsg }) {
    const [batchPptFiles, setBatchPptFiles] = useState([]);
    const [replaceRules, setReplaceRules] = useState(() => {
        try {
            return localStorage.getItem('ppt_replace_rules') || '';
        } catch {
            return '';
        }
    });
    const [fontRules, setFontRules] = useState('');
    const [fontSize, setFontSize] = useState('');
    const [applyDesignChecked, setApplyDesignChecked] = useState(true);
    const [applyTableDesignChecked, setApplyTableDesignChecked] = useState(false);
    const [applyFirstRowHeaderStyle, setApplyFirstRowHeaderStyle] = useState(true); // 옵션 E 하위 옵션: 첫 행 특별 포맷팅 적용 여부 (기본 true)
    const [applyFirstColHeaderStyle, setApplyFirstColHeaderStyle] = useState(true); // 옵션 E 하위 옵션: 첫 열(2행부터) 특별 포맷팅 적용 여부 (기본 true)
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
        } catch {
            return '';
        }
    }); // 옵션 H: 글자 색상 매핑 변경
    const [preventWordWrap, setPreventWordWrap] = useState(() => {
        try {
            return localStorage.getItem('ppt_prevent_word_wrap') === 'true';
        } catch {
            return false;
        }
    }); // 옵션 I: 단락 한글 단어 잘림 방지
    const [clearAltText, setClearAltText] = useState(false); // 옵션 J: 대체 텍스트 일괄 제거 여부
    const [fixLangErrFlags, setFixLangErrFlags] = useState(false); // 옵션 K: 언어 태그/맞춤법 오류 표시 보정
    const [mergeToOneFile, setMergeToOneFile] = useState(false); // 옵션 L: 다중 PPT를 하나의 PPT 파일로 병합
    const [isProcessingBatch, setIsProcessingBatch] = useState(false);
    const [isDraggingBatch, setIsDraggingBatch] = useState(false);
    const [batchReport, setBatchReport] = useState([]); // 📊 일괄 편집 결과 상세 피드백 리포트 리스트

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

    const handleBatchPptDragEvents = {
        onDragOver: (e) => { e.preventDefault(); e.stopPropagation(); },
        onDragEnter: (e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingBatch(true); },
        onDragLeave: (e) => { e.preventDefault(); e.stopPropagation(); setIsDraggingBatch(false); },
        onDrop: (e) => {
            e.preventDefault(); e.stopPropagation(); setIsDraggingBatch(false);
            const files = Array.from(e.dataTransfer.files);
            const validFiles = files.filter(f => f.name.toLowerCase().endsWith('.pptx'));
            if (validFiles.length > 0) {
                // 💡 탐색기에서 드래그 시 마우스로 집은 파일이 1번으로 오는 OS 동작을 방지하고 번호순 자동 정렬
                setBatchPptFiles(prev => sortFilesNaturally([...prev, ...validFiles], true));
            } else {
                setErrorMsg('PPT 파일(.pptx)만 지원합니다.');
            }
        }
    };

    const handleBatchFileChange = (e) => {
        const files = Array.from(e.target.files);
        const validFiles = files.filter(f => f.name.toLowerCase().endsWith('.pptx'));
        if (validFiles.length > 0) {
            setBatchPptFiles(prev => sortFilesNaturally([...prev, ...validFiles], true));
        }
        e.target.value = '';
    };

    const removeBatchFile = (indexToRemove) => {
        setBatchPptFiles(prev => prev.filter((_, idx) => idx !== indexToRemove));
    };

    const moveFileOrder = (index, direction) => {
        setBatchPptFiles(prev => {
            const targetIndex = index + direction;
            if (targetIndex < 0 || targetIndex >= prev.length) return prev;
            const newArr = [...prev];
            const temp = newArr[index];
            newArr[index] = newArr[targetIndex];
            newArr[targetIndex] = temp;
            return newArr;
        });
    };

    const handleSortNaturally = () => {
        setBatchPptFiles(prev => sortFilesNaturally(prev, true));
    };

    const handleSortReverse = () => {
        setBatchPptFiles(prev => sortFilesNaturally(prev, false));
    };

    const handleClearAllFiles = () => {
        setBatchPptFiles([]);
    };

    const buildBatchReportDetail = (modifiedBlob, options) => {
        const {
            applyTableDesignChecked,
            applyFirstRowHeaderStyle,
            parsedRules,
            parsedFontRules,
            parsedFontSizeRules,
            applyDesignChecked,
            applySpecialCharClean,
            add_space_before_parenthesis,
            textColorRules,
            preventWordWrap,
            clearAltText,
            fixLangErrFlags
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
        // 6-1. 첫 행 헤더 특별 포맷팅 수치 표출
        if (applyFirstRowHeaderStyle !== false) {
            const count = modifiedBlob.totalHeaderRowsApplied || 0;
            if (count > 0) {
                changes.push(`🎯 첫 행 헤더 특별 포맷팅 적용 ${count}개 표`);
            } else {
                changes.push(`첫 행 헤더 스타일 적용 0개(대상 표 없음)`);
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
        // 10. 대체 텍스트 일괄 제거
        if (clearAltText) {
            changes.push(`대체 텍스트 일괄 제거 완료`);
        }
        // 11. 언어 태그/맞춤법 오류 표시 보정
        if (fixLangErrFlags) {
            const langCount = modifiedBlob.totalLangFixed || 0;
            const errCount = modifiedBlob.totalErrFlagsCleared || 0;
            changes.push(`언어 태그 보정 ${langCount}개, 맞춤법 오류 표시 제거 ${errCount}개`);
        }

        if (changes.length > 0) {
            // 실질적인 변경이 하나라도 존재하는지 확인 (표 없음 제외)
            const hasRealChanges = (modifiedBlob.totalReplacedWords || 0) > 0 ||
                                  (modifiedBlob.totalReplacedFonts || 0) > 0 ||
                                  (modifiedBlob.totalReplacedFontSizes || 0) > 0 ||
                                  (modifiedBlob.totalSpecialCharsCleaned || 0) > 0 ||
                                  (modifiedBlob.totalReplacedTextDesigns || 0) > 0 ||
                                  (modifiedBlob.totalHeaderRowsApplied || 0) > 0 ||
                                  (add_space_before_parenthesis && (modifiedBlob.totalTitleSpacesAdded || 0) > 0) ||
                                  ((textColorRules && textColorRules.trim()) && (modifiedBlob.totalTextColorReplaced || 0) > 0) ||
                                  (preventWordWrap && (modifiedBlob.totalWordWrapPrevented || 0) > 0) ||
                                  clearAltText ||
                                  (fixLangErrFlags && ((modifiedBlob.totalLangFixed || 0) > 0 || (modifiedBlob.totalErrFlagsCleared || 0) > 0)) ||
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

        if (parsedRules.length === 0 && parsedFontRules.length === 0 && !applyDesignChecked && parsedFontSizeRules.length === 0 && !applyTableDesignChecked && !applySpecialCharClean && !add_title_page_numbers && !textColorRules.trim() && !preventWordWrap && !clearAltText && !fixLangErrFlags) {
            setErrorMsg('적용할 단어 수정, 폰트 변경, 폰트 크기, 테이블 디자인 표준화, 텍스트 디자인 변경, 특수문자 일괄 정제, 동일 제목 일련번호 추가, 글자 색상 일괄 매핑, 단락 단어 잘림 방지, 대체 텍스트 일괄 제거, 또는 언어 태그/맞춤법 오류 표시 보정 중 하나 이상을 입력/선택해주세요.');
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

            if (batchPptFiles.length === 1) {
                // 단일 파일: showSaveFilePicker 혹은 file-saver 폴백으로 순수 PPTX 다운로드
                const file = batchPptFiles[0];
                try {
                    const options = {
                        replaceRules: parsedRules,
                        fontRules: parsedFontRules,
                        fontSizeRules: parsedFontSizeRules,
                        applyDesign: applyDesignChecked,
                        applyTableDesign: applyTableDesignChecked,
                        applyFirstRowHeaderStyle: applyTableDesignChecked && applyFirstRowHeaderStyle,
                        applyFirstColHeaderStyle: applyTableDesignChecked && applyFirstColHeaderStyle,
                        targetText: designTargetText,
                        applySpecialCharClean: applySpecialCharClean,
                        replaceNbs: replaceNbs,
                        unifyBullets: unifyBullets,
                        clean_vertical_tab: clean_vertical_tab,
                        add_title_page_numbers: add_title_page_numbers,
                        add_space_before_parenthesis: add_space_before_parenthesis,
                        textColorRulesStr: textColorRules,
                        preventWordWrap: preventWordWrap,
                        clearAltText: clearAltText,
                        fixLangErrFlags: fixLangErrFlags
                    };
                    const modifiedBlob = await processPptBatch(file, options);
                    const fileName = `수정_${file.name}`;

                    let saved = false;
                    if ('showSaveFilePicker' in window) {
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
                        }
                    }

                    if (!saved) {
                        // 폴백: file-saver
                        const { saveAs } = await import('file-saver');
                        saveAs(modifiedBlob, fileName);
                        saved = true;
                    }

                    if (saved) {
                        const detailMsg = buildBatchReportDetail(modifiedBlob, {
                            applyTableDesignChecked,
                            applyFirstRowHeaderStyle,
                            parsedRules,
                            parsedFontRules,
                            parsedFontSizeRules,
                            applyDesignChecked,
                            applySpecialCharClean,
                            add_space_before_parenthesis,
                            textColorRules,
                            preventWordWrap,
                            clearAltText,
                            fixLangErrFlags
                        });
                        reports.push({ fileName: file.name, status: 'success', detail: detailMsg });
                        successCount++;
                    }
                } catch (fileErr) {
                    console.error(`Error processing ${file.name}:`, fileErr);
                    reports.push({ fileName: file.name, status: 'error', detail: `❌ 처리 실패: ${fileErr.message || 'PPT 내부 구조 파싱 에러'}` });
                }
            } else {
                // 다중 파일: 모두 처리 후 개별 ZIP 압축 또는 단일 PPT 파일로 병합 다운로드
                const zip = new JSZip();
                const modifiedBlobs = [];

                for (const file of batchPptFiles) {
                    try {
                        const options = {
                            replaceRules: parsedRules,
                            fontRules: parsedFontRules,
                            fontSizeRules: parsedFontSizeRules,
                            applyDesign: applyDesignChecked,
                            applyTableDesign: applyTableDesignChecked,
                            applyFirstRowHeaderStyle: applyTableDesignChecked && applyFirstRowHeaderStyle,
                            applyFirstColHeaderStyle: applyTableDesignChecked && applyFirstColHeaderStyle,
                            targetText: designTargetText,
                            applySpecialCharClean: applySpecialCharClean,
                            replaceNbs: replaceNbs,
                            unifyBullets: unifyBullets,
                            clean_vertical_tab: clean_vertical_tab,
                            add_title_page_numbers: add_title_page_numbers,
                            add_space_before_parenthesis: add_space_before_parenthesis,
                            textColorRulesStr: textColorRules,
                            preventWordWrap: preventWordWrap,
                            clearAltText: clearAltText,
                            fixLangErrFlags: fixLangErrFlags
                        };
                        const modifiedBlob = await processPptBatch(file, options);
                        modifiedBlobs.push(modifiedBlob);

                        if (!mergeToOneFile) {
                            const fileName = `수정_${file.name}`;
                            zip.file(fileName, modifiedBlob);
                        }

                        const detailMsg = buildBatchReportDetail(modifiedBlob, {
                            applyTableDesignChecked,
                            applyFirstRowHeaderStyle,
                            parsedRules,
                            parsedFontRules,
                            parsedFontSizeRules,
                            applyDesignChecked,
                            applySpecialCharClean,
                            add_space_before_parenthesis,
                            textColorRules,
                            preventWordWrap,
                            clearAltText,
                            fixLangErrFlags
                        });
                        reports.push({ fileName: file.name, status: 'success', detail: detailMsg });
                        successCount++;
                    } catch (fileErr) {
                        console.error(`Error processing ${file.name}:`, fileErr);
                        reports.push({ fileName: file.name, status: 'error', detail: `❌ 처리 실패: ${fileErr.message || 'PPT 내부 구조 파싱 에러'}` });
                    }
                }

                if (successCount > 0) {
                    if (mergeToOneFile) {
                        // 💡 [옵션 L] 여러 PPT 파일들을 하나의 단일 PPT 파일로 병합하여 저장
                        const mergedBlob = await mergeMultiplePptx(modifiedBlobs);
                        const mergedFileName = '통합_수정_PPT_산출물.pptx';

                        let mergedSaved = false;
                        if ('showSaveFilePicker' in window) {
                            try {
                                const handle = await window.showSaveFilePicker({
                                    suggestedName: mergedFileName,
                                    startIn: 'downloads',
                                    types: [{
                                        description: 'PowerPoint Presentation',
                                        accept: { 'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx'] },
                                    }],
                                });
                                const writable = await handle.createWritable();
                                await writable.write(mergedBlob);
                                await writable.close();
                                mergedSaved = true;
                            } catch (pickerErr) {
                                if (pickerErr.name !== 'AbortError') {
                                    const { saveAs } = await import('file-saver');
                                    saveAs(mergedBlob, mergedFileName);
                                    mergedSaved = true;
                                }
                            }
                        }
                        if (!mergedSaved) {
                            const { saveAs } = await import('file-saver');
                            saveAs(mergedBlob, mergedFileName);
                        }
                    } else {
                        // 기본: ZIP 압축 파일로 다운로드
                        const zipBlob = await zip.generateAsync({ type: 'blob' });
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
            }

            setBatchReport(reports);

            if (successCount > 0) {
                if (batchPptFiles.length > 1) {
                    if (mergeToOneFile) {
                        setSuccessMsg(`성공적으로 ${successCount}개의 파일을 일괄 편집하고 1개의 통합 PPT 파일(통합_수정_PPT_산출물.pptx)로 병합하여 저장했습니다. 하단의 세부 리포트를 확인해 주세요.`);
                    } else {
                        setSuccessMsg(`성공적으로 ${successCount}개의 파일을 처리하여 ZIP 압축 파일로 다운로드했습니다. 하단의 파일별 일괄 편집 상세 결과 리포트를 확인해 주세요.`);
                    }
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
                setFixLangErrFlags(false);
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

    return (
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
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', marginTop: '10px' }}>
                            {/* 정렬 제어 툴바 */}
                            <div style={{
                                display: 'flex', alignItems: 'center', justifyContent: 'space-between',
                                flexWrap: 'wrap', gap: '8px', padding: '8px 12px',
                                background: 'rgba(0,0,0,0.25)', borderRadius: '8px', border: '1px solid var(--panel-border)'
                            }}>
                                <div style={{ fontSize: '13px', color: 'var(--text-secondary)', display: 'flex', alignItems: 'center', gap: '6px' }}>
                                    <span style={{ color: '#c084fc', fontWeight: 700 }}>총 {batchPptFiles.length}개</span> 파일 등록됨 (💡 순서대로 병합됩니다)
                                </div>
                                <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                                    <button
                                        type="button"
                                        onClick={handleSortNaturally}
                                        title="파일명의 번호/이름 순서대로 자동 정렬 (1, 2, ..., 10)"
                                        style={{
                                            display: 'flex', alignItems: 'center', gap: '4px',
                                            padding: '5px 10px', borderRadius: '6px', fontSize: '12px',
                                            background: 'rgba(168, 85, 247, 0.15)', border: '1px solid #a855f7',
                                            color: '#d8b4fe', cursor: 'pointer', fontWeight: 600
                                        }}
                                    >
                                        <ArrowUpDown size={13} /> 이름/번호순 정렬
                                    </button>
                                    <button
                                        type="button"
                                        onClick={handleSortReverse}
                                        title="역순으로 정렬"
                                        style={{
                                            display: 'flex', alignItems: 'center', gap: '4px',
                                            padding: '5px 9px', borderRadius: '6px', fontSize: '12px',
                                            background: 'rgba(255,255,255,0.05)', border: '1px solid var(--panel-border)',
                                            color: 'var(--text-secondary)', cursor: 'pointer'
                                        }}
                                    >
                                        역순
                                    </button>
                                    <button
                                        type="button"
                                        onClick={handleClearAllFiles}
                                        title="등록된 파일 모두 지우기"
                                        style={{
                                            display: 'flex', alignItems: 'center', gap: '4px',
                                            padding: '5px 9px', borderRadius: '6px', fontSize: '12px',
                                            background: 'rgba(239, 68, 68, 0.1)', border: '1px solid rgba(239, 68, 68, 0.3)',
                                            color: '#f87171', cursor: 'pointer'
                                        }}
                                    >
                                        <Trash2 size={13} /> 전체 삭제
                                    </button>
                                </div>
                            </div>

                            {/* 파일 태그 목록 */}
                            <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px' }}>
                                {batchPptFiles.map((file, idx) => (
                                    <div key={idx} style={{
                                        display: 'flex', alignItems: 'center', gap: '6px',
                                        background: 'rgba(0,0,0,0.3)', padding: '5px 8px',
                                        borderRadius: '8px', fontSize: '12.5px', border: '1px solid #a855f7',
                                        transition: 'all 0.15s ease'
                                    }}>
                                        <span style={{
                                            background: '#a855f7', color: 'white',
                                            borderRadius: '4px', padding: '1px 5px',
                                            fontSize: '11px', fontWeight: 700
                                        }}>
                                            {idx + 1}
                                        </span>
                                        <span
                                            title={file.name}
                                            style={{ maxWidth: '200px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', fontWeight: 500 }}
                                        >
                                            {file.name}
                                        </span>
                                        <div style={{ display: 'flex', alignItems: 'center', gap: '2px', marginLeft: '2px', borderLeft: '1px solid rgba(255,255,255,0.1)', paddingLeft: '4px' }}>
                                            <button
                                                type="button"
                                                onClick={() => moveFileOrder(idx, -1)}
                                                disabled={idx === 0}
                                                title="앞으로 이동"
                                                style={{
                                                    background: 'none', border: 'none',
                                                    color: idx === 0 ? 'rgba(255,255,255,0.2)' : 'var(--text-secondary)',
                                                    cursor: idx === 0 ? 'default' : 'pointer',
                                                    padding: '2px', display: 'flex', alignItems: 'center'
                                                }}
                                            >
                                                <ChevronLeft size={13} />
                                            </button>
                                            <button
                                                type="button"
                                                onClick={() => moveFileOrder(idx, 1)}
                                                disabled={idx === batchPptFiles.length - 1}
                                                title="뒤로 이동"
                                                style={{
                                                    background: 'none', border: 'none',
                                                    color: idx === batchPptFiles.length - 1 ? 'rgba(255,255,255,0.2)' : 'var(--text-secondary)',
                                                    cursor: idx === batchPptFiles.length - 1 ? 'default' : 'pointer',
                                                    padding: '2px', display: 'flex', alignItems: 'center'
                                                }}
                                            >
                                                <ChevronRight size={13} />
                                            </button>
                                            <button
                                                type="button"
                                                onClick={() => removeBatchFile(idx)}
                                                title="파일 삭제"
                                                style={{ background: 'none', border: 'none', color: '#f87171', cursor: 'pointer', padding: '2px', display: 'flex', alignItems: 'center', marginLeft: '2px' }}
                                            >
                                                <X size={13} />
                                            </button>
                                        </div>
                                    </div>
                                ))}
                            </div>
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
                        placeholder="예: Arial(KoPubDotum Bold), Calibri(Pretendard)"
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
                            onChange={(e) => {
                                const checked = e.target.checked;
                                setApplyTableDesignChecked(checked);
                                if (checked) {
                                    setApplyFirstRowHeaderStyle(true);
                                    setApplyFirstColHeaderStyle(true);
                                }
                            }}
                            style={{ width: '18px', height: '18px', cursor: 'pointer', accentColor: '#a855f7' }}
                        />
                        옵션 E: 테이블(표) 표준 디자인 일괄 변경 적용
                    </label>
                    <div style={{ paddingLeft: '28px', fontSize: '13px', color: 'var(--text-secondary)', lineHeight: '1.5' }}>
                        💡 모든 표 테두리 실선 0.5pt (#7F7F7F)가 기본 적용되며, 첫 행 및 첫 열의 특별 포맷팅을 하위 옵션으로 선택 제어할 수 있습니다.
                    </div>

                    <div style={{
                        paddingLeft: '28px',
                        display: 'flex',
                        flexDirection: 'column',
                        gap: '12px',
                        borderLeft: '2px solid ' + (applyTableDesignChecked ? 'rgba(168, 85, 247, 0.5)' : 'rgba(255, 255, 255, 0.1)'),
                        marginTop: '4px',
                        opacity: applyTableDesignChecked ? 1 : 0.45,
                        pointerEvents: applyTableDesignChecked ? 'auto' : 'none',
                        transition: 'all 0.3s ease'
                    }}>
                        {/* 하위 1: 첫 번째 행(헤더) 특별 포맷팅 */}
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
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
                                (첫 행 배경 RGB(0,114,186), 글자 흰색 11pt KoPubDotum Bold, 가운데 정렬, 첫 행 내부 실선 흰색 적용)
                            </div>
                        </div>

                        {/* 하위 2: 첫 번째 열(구분열) 특별 포맷팅 */}
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
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
                                    checked={applyFirstColHeaderStyle}
                                    disabled={!applyTableDesignChecked}
                                    onChange={(e) => setApplyFirstColHeaderStyle(e.target.checked)}
                                    style={{
                                        width: '16px',
                                        height: '16px',
                                        cursor: applyTableDesignChecked ? 'pointer' : 'not-allowed',
                                        accentColor: '#a855f7'
                                    }}
                                />
                                첫 번째 열(구분열) 특별 포맷팅 적용 (2번째 행부터 적용)
                            </label>
                            <div style={{ fontSize: '12px', color: 'var(--text-muted)', lineHeight: '1.4', paddingLeft: '24px' }}>
                                (첫 열 2번째 행부터 배경 #F2F2F2, 글자 검은색 11pt KoPubDotum Bold, 가운데 정렬)
                            </div>
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

                {/* 옵션 J: 대체 텍스트 일괄 제거 */}
                <div style={{ background: 'var(--bg-secondary)', padding: '16px', borderRadius: '8px', border: '1px solid var(--panel-border)', display: 'flex', flexDirection: 'column', gap: '12px' }}>
                    <label style={{ display: 'flex', alignItems: 'center', gap: '10px', cursor: 'pointer', fontWeight: 600, fontSize: '14px', color: 'var(--text-primary)' }}>
                        <input
                            id="checkbox-option-j"
                            type="checkbox"
                            checked={clearAltText}
                            onChange={(e) => setClearAltText(e.target.checked)}
                            style={{ width: '18px', height: '18px', cursor: 'pointer', accentColor: '#a855f7' }}
                        />
                        옵션 J: 개체(도형, 이미지 등)의 대체 텍스트 일괄 제거 적용
                    </label>
                    <div style={{ paddingLeft: '28px', fontSize: '13px', color: 'var(--text-secondary)', lineHeight: '1.5' }}>
                        💡 PPTX 파일 내의 이미지, 도형, 표 등 모든 개체에 설정된 대체 텍스트 제목 및 설명(descr, title)을 일괄 삭제합니다.
                    </div>
                </div>

                {/* 옵션 K: 언어 태그/맞춤법 오류 표시 보정 */}
                <div style={{ background: 'var(--bg-secondary)', padding: '16px', borderRadius: '8px', border: '1px solid var(--panel-border)', display: 'flex', flexDirection: 'column', gap: '12px' }}>
                    <label style={{ display: 'flex', alignItems: 'center', gap: '10px', cursor: 'pointer', fontWeight: 600, fontSize: '14px', color: 'var(--text-primary)' }}>
                        <input
                            id="checkbox-option-k"
                            type="checkbox"
                            checked={fixLangErrFlags}
                            onChange={(e) => setFixLangErrFlags(e.target.checked)}
                            style={{ width: '18px', height: '18px', cursor: 'pointer', accentColor: '#a855f7' }}
                        />
                        옵션 K: 언어 태그/맞춤법 오류 표시 보정
                    </label>
                    <div style={{ paddingLeft: '28px', fontSize: '13px', color: 'var(--text-secondary)', lineHeight: '1.5' }}>
                        💡 한글 텍스트가 영어(en-US)로 잘못 태깅되어 PowerPoint 맞춤법 검사기가 빨간 밑줄로 표시하는 오탐을 제거합니다. 한글이 포함된 텍스트런의 언어 태그를 ko-KR로 보정하고, 저장된 맞춤법 오류 표시(err) 플래그를 모두 지웁니다.
                    </div>
                </div>

                {/* 옵션 L: 다중 PPT 파일을 하나의 통합 PPT 파일로 병합 */}
                <div style={{ background: 'var(--bg-secondary)', padding: '16px', borderRadius: '8px', border: '1px solid var(--panel-border)', display: 'flex', flexDirection: 'column', gap: '12px' }}>
                    <label style={{ display: 'flex', alignItems: 'center', gap: '10px', cursor: 'pointer', fontWeight: 600, fontSize: '14px', color: 'var(--text-primary)' }}>
                        <input
                            id="checkbox-option-l"
                            type="checkbox"
                            checked={mergeToOneFile}
                            onChange={(e) => setMergeToOneFile(e.target.checked)}
                            style={{ width: '18px', height: '18px', cursor: 'pointer', accentColor: '#a855f7' }}
                        />
                        옵션 L: 다중 PPT 파일을 하나의 통합 PPT 파일로 합치기 (병합)
                    </label>
                    <div style={{ paddingLeft: '28px', fontSize: '13px', color: 'var(--text-secondary)', lineHeight: '1.5' }}>
                        💡 여러 개의 PPT 파일을 일괄 수정한 뒤, <strong>순서대로 결합하여 1개의 PPTX 파일로 통합 저장</strong>합니다. (미선택 시 기본 개별 ZIP 압축 파일로 다운로드됩니다.)
                    </div>
                </div>
            </div>

            <div style={{ marginTop: 'auto', paddingTop: '20px' }}>
                <button
                    id="btn-batch-process"
                    className="interactive"
                    onClick={handleBatchProcess}
                    disabled={batchPptFiles.length === 0 || isProcessingBatch || (!replaceRules.trim() && !fontRules.trim() && !fontSize.trim() && !applyDesignChecked && !applyTableDesignChecked && !applySpecialCharClean && !add_title_page_numbers && !add_space_before_parenthesis && !textColorRules.trim() && !preventWordWrap && !clearAltText && !fixLangErrFlags && !mergeToOneFile)}
                    style={{
                        width: '100%',
                        padding: '16px',
                        background: (batchPptFiles.length === 0 || isProcessingBatch || (!replaceRules.trim() && !fontRules.trim() && !fontSize.trim() && !applyDesignChecked && !applyTableDesignChecked && !applySpecialCharClean && !add_title_page_numbers && !add_space_before_parenthesis && !textColorRules.trim() && !preventWordWrap && !clearAltText && !fixLangErrFlags && !mergeToOneFile)) ? 'rgba(255,255,255,0.05)' : 'linear-gradient(135deg, #a855f7, #3b82f6)',
                        color: (batchPptFiles.length === 0 || isProcessingBatch || (!replaceRules.trim() && !fontRules.trim() && !fontSize.trim() && !applyDesignChecked && !applyTableDesignChecked && !applySpecialCharClean && !add_title_page_numbers && !add_space_before_parenthesis && !textColorRules.trim() && !preventWordWrap && !clearAltText && !fixLangErrFlags && !mergeToOneFile)) ? 'var(--text-muted)' : 'white',
                        border: 'none',
                        borderRadius: '12px',
                        fontSize: '16px',
                        fontWeight: 700,
                        cursor: (batchPptFiles.length === 0 || isProcessingBatch || (!replaceRules.trim() && !fontRules.trim() && !fontSize.trim() && !applyDesignChecked && !applyTableDesignChecked && !applySpecialCharClean && !add_title_page_numbers && !add_space_before_parenthesis && !textColorRules.trim() && !preventWordWrap && !clearAltText && !fixLangErrFlags && !mergeToOneFile)) ? 'not-allowed' : 'pointer',
                        display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '10px'
                    }}
                >
                    {isProcessingBatch ? (
                        <><Loader2 size={20} className="animate-spin" /> {mergeToOneFile && batchPptFiles.length > 1 ? '일괄 편집 및 하나의 PPT로 병합 중...' : '폴더에 순차적으로 적용 및 저장 중...'}</>
                    ) : (
                        <><Sparkles size={20} /> {mergeToOneFile && batchPptFiles.length > 1 ? '일괄 편집 및 하나의 PPT 파일로 병합 실행' : '저장할 폴더/위치 선택 및 일괄 편집 실행'}</>
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
    );
}
