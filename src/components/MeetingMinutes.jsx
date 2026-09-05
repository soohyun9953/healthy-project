import React, { useState, useRef, useCallback, useEffect } from 'react';
import { analyzeMeeting, askMeetingQuestion } from '../meetingAnalyzer';
import { processFile } from '../utils/fileExtractor';
import { FALLBACK_MODELS } from '../utils/geminiModels.js';
import {
  Mic, Upload, X, Plus, Trash2, FileAudio, ChevronDown, ChevronUp,
  CheckCircle2, AlertCircle, BookOpen, Download, Copy, Loader2,
  Users, ClipboardList, ListChecks, Sparkles, Tag, Clock,
  FileSearch, CheckSquare, Square, TrendingUp, Lightbulb, Calendar,
  Target, ArrowRight, Zap, BarChart3, MapPin, MessageSquare, Send, RotateCcw, HelpCircle
} from 'lucide-react';

// ── 전문 용어 사전 로컬스토리지 키 ─────────────
const TERM_STORAGE_KEY = 'meeting_terminology';

function loadTerminology() {
  try { return JSON.parse(localStorage.getItem(TERM_STORAGE_KEY) || '[]'); }
  catch { return []; }
}
function saveTerminology(terms) {
  localStorage.setItem(TERM_STORAGE_KEY, JSON.stringify(terms));
}

// ── 화자 색상 팔레트 ─────────────────────────
const SPEAKER_COLORS = {
  A: { bg: 'rgba(99,102,241,0.15)',  border: 'rgba(99,102,241,0.5)',  label: '#818cf8' },
  B: { bg: 'rgba(16,185,129,0.15)',  border: 'rgba(16,185,129,0.5)',  label: '#34d399' },
  C: { bg: 'rgba(245,158,11,0.15)',  border: 'rgba(245,158,11,0.5)',  label: '#fbbf24' },
  D: { bg: 'rgba(239,68,68,0.15)',   border: 'rgba(239,68,68,0.5)',   label: '#f87171' },
  E: { bg: 'rgba(236,72,153,0.15)',  border: 'rgba(236,72,153,0.5)',  label: '#f472b6' },
};
const getSpeakerColor = (sp) =>
  SPEAKER_COLORS[sp] || { bg: 'rgba(255,255,255,0.05)', border: 'var(--glass-border)', label: 'var(--text-secondary)' };

const fmtSize = (bytes) => {
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
};

function buildPlainText(result) {
  if (!result) return '';
  const lines = [];
  lines.push(`\u25a0 회의명: ${result.meetingTitle || ''}`);
  lines.push(`\uFEFF\u25a0 회의 목적 및 배경: ${result.meetingContext || ''}`);
  lines.push(`\uFEFF\u25a0 회의 안건: ${result.agenda || ''}`);
  lines.push('');
  lines.push('\uFEFF\u25a0 종합 요약');
  lines.push(`  ${result.summary || ''}`);
  lines.push('');
  if ((result.topicSummaries || []).length > 0) {
    lines.push('\uFEFF\u25a0 주제별 상세 논의 내용');
    (result.topicSummaries || []).forEach((t, i) => {
      lines.push(`  [주제 ${i + 1}] ${t.topic}`);
      lines.push(`  ${t.content}`);
      if (t.result) lines.push(`  → 결론: ${t.result}`);
    });
    lines.push('');
  }
  if ((result.keyInsights || []).length > 0) {
    lines.push('\uFEFF\u25a0 핵심 인사이트');
    (result.keyInsights || []).forEach((ins, i) => {
      lines.push(`  ${i + 1}. [${ins.type}] ${ins.content}`);
    });
    lines.push('');
  }
  lines.push('\uFEFF\u25a0 결정된 사항');
  (result.decisions || []).forEach((d, i) => lines.push(`  ${i + 1}. ${d}`));
  lines.push('');
  lines.push('\uFEFF\u25a0 할 일(Action Items) 및 담당자');
  (result.actionItems || []).forEach((a, i) => {
    const priority = a.priority ? ` [우선순위: ${a.priority}]` : '';
    lines.push(`  ${i + 1}. ${a.task} / 담당: ${a.owner || '-'}${a.deadline ? ` / 기한: ${a.deadline}` : ''}${priority}`);
  });
  lines.push('');
  const fp = result.futurePlans;
  if (fp && ((fp.shortTerm || []).length > 0 || (fp.longTerm || []).length > 0)) {
    lines.push('\uFEFF\u25a0 향후 계획');
    if ((fp.shortTerm || []).length > 0) {
      lines.push('  [단기 계획]');
      (fp.shortTerm || []).forEach((p, i) => {
        lines.push(`    ${i + 1}. ${p.plan}${p.owner ? ` / 담당: ${p.owner}` : ''}${p.targetDate ? ` / 목표: ${p.targetDate}` : ''}`);
      });
    }
    if ((fp.longTerm || []).length > 0) {
      lines.push('  [중장기 계획]');
      (fp.longTerm || []).forEach((p, i) => {
        lines.push(`    ${i + 1}. ${p.plan}${p.owner ? ` / 담당: ${p.owner}` : ''}${p.targetDate ? ` / 목표: ${p.targetDate}` : ''}`);
      });
    }
    lines.push('');
  }
  lines.push('\uFEFF\u25a0 화자별 발언록');
  (result.transcript || []).forEach(t => {
    lines.push(`  [${t.speaker}] ${t.text}${t.tag && t.tag !== '\uc815\uc0c1' ? ` (${t.tag})` : ''}`);
  });
  return lines.join('\n');
}

// ── Gemini API로 문서에서 전문 용어 추출 ────────
async function extractTermsFromText(text, apiKey) {
  const keys = String(apiKey).split(',').map(k => k.trim()).filter(k => k.match(/^(AIza|AQ\.)/));
  if (keys.length === 0) throw new Error('유효한 API 키가 없습니다.');

  const MODELS = FALLBACK_MODELS;
  const prompt = `아래 문서에서 자주 등장하거나 도메인에 특화된 전문 용어, 약어, 고유명사를 추출해줘.
일반적인 조사나 접속사, 너무 보편적인 단어(예: 회의, 업무, 내용 등)는 제외하고, 실제로 전문적이거나 프로젝트 고유의 맥락을 가진 단어만 추출해줘.

[출력 형식 - JSON만 출력, 다른 텍스트 절대 금지]
{
  "terms": [
    { "word": "<용어/약어>", "desc": "<한 줄 설명 또는 풀어쓰기. 불명확하면 빈 문자열>", "freq": <대략적 등장 빈도 추정 숫자> }
  ]
}

[분석할 문서]
${text.substring(0, 80000)}`;

  let currentKeyIndex = 0;
  let currentModelIndex = 0;

  while (currentModelIndex < MODELS.length) {
    const url = `https://generativelanguage.googleapis.com/v1beta/${MODELS[currentModelIndex]}:generateContent?key=${keys[currentKeyIndex]}`;
    const res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        contents: [{ role: 'user', parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.1 },
      }),
    });
    
    if (res.ok) {
      const data = await res.json();
      let content = data.candidates?.[0]?.content?.parts?.[0]?.text || '{}';
      if (content.includes('```')) {
        const m = content.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
        if (m?.[1]) content = m[1];
      }
      try {
        const parsed = JSON.parse(content);
        return (parsed.terms || []).sort((a, b) => (b.freq || 0) - (a.freq || 0));
      } catch (e) {
        return [];
      }
    }

    // 1. 에러 발생 시 항상 다음 API 키를 먼저 시도
    if (keys.length > 1 && (currentKeyIndex + 1) < keys.length) {
      currentKeyIndex++;
      continue;
    }

    // 2. 모든 키를 다 썼다면 모델 교체 시도
    currentKeyIndex = 0;
    currentModelIndex++;
  }
  throw new Error('용어 추출에 실패했습니다.');
}

// ── 인풋 공통 스타일 ──────────────────────────
const inputStyle = {
  padding: '7px 10px', borderRadius: '8px', border: '1px solid var(--glass-border)',
  background: 'rgba(255,255,255,0.04)', color: 'var(--text-primary)', fontSize: '12px', outline: 'none',
};

export default function MeetingMinutes({ apiKey }) {
  // ── 입력 모드 및 상태 ──────────────────────
  const [inputMode, setInputMode]   = useState('audio'); // 'audio' | 'text'
  const [textInput, setTextInput]   = useState('');
  
  // 오디오 파일
  const [audioFile, setAudioFile]   = useState(null);
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = useRef(null);
  const textFileInputRef = useRef(null);

  // 전문 용어 사전
  const [terminology, setTerminology]   = useState(loadTerminology);
  const [termWord, setTermWord]         = useState('');
  const [termDesc, setTermDesc]         = useState('');
  const [showTermPanel, setShowTermPanel] = useState(true);

  // 파일 기반 용어 추출
  const termFileRef = useRef(null);
  const [isExtractingTerms, setIsExtractingTerms] = useState(false);
  const [extractedTerms, setExtractedTerms]         = useState(null); // null | Term[]
  const [selectedTermIds, setSelectedTermIds]       = useState(new Set());
  const [extractError, setExtractError]             = useState('');
  const [isTermDragging, setIsTermDragging]         = useState(false);

  // 회의록 분석
  const [isLoading, setIsLoading] = useState(false);
  const [progress, setProgress]   = useState('');
  const [result, setResult]       = useState(null);
  const [error, setError]         = useState('');

  // UI
  const [showTranscript, setShowTranscript] = useState(true);
  const [copied, setCopied]                 = useState(false);

  // 💬 회의 내용 AI Q&A (질의응답) 상태
  const [qaMessages, setQaMessages]         = useState([]);
  const [qaInput, setQaInput]               = useState('');
  const [isQaLoading, setIsQaLoading]       = useState(false);
  const [qaError, setQaError]               = useState('');
  const [showQaPanel, setShowQaPanel]       = useState(true);
  const [copiedQaId, setCopiedQaId]         = useState(null);
  const qaChatEndRef = useRef(null);

  // 자동 스크롤
  useEffect(() => {
    if (qaMessages.length > 0) {
      qaChatEndRef.current?.scrollIntoView({ behavior: 'smooth' });
    }
  }, [qaMessages, isQaLoading]);

  // Q&A 질의 전송 핸들러
  const handleSendQa = async (questionText) => {
    const q = (typeof questionText === 'string' ? questionText : qaInput).trim();
    if (!q) return;

    if (!result && !textInput.trim()) {
      setQaError('회의록을 먼저 생성하거나 분석할 텍스트를 입력해 주세요.');
      return;
    }

    const keys = String(apiKey).split(',').map(k => k.trim()).filter(k => k.match(/^(AIza|AQ\.)/));
    if (keys.length === 0) {
      setQaError('상단 설정에서 Gemini API 키를 먼저 입력해 주세요.');
      return;
    }

    const userMsg = {
      id: Date.now(),
      role: 'user',
      text: q,
      time: new Date().toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit' })
    };

    const newHistory = [...qaMessages, userMsg];
    setQaMessages(newHistory);
    setQaInput('');
    setIsQaLoading(true);
    setQaError('');

    try {
      const contextData = result || textInput;
      // 최근 6개 대화 히스토리만 유지하여 토큰 최적화
      const historyForApi = newHistory.slice(-6).map(m => ({ role: m.role, text: m.text }));
      const res = await askMeetingQuestion(q, contextData, apiKey, historyForApi);

      const botMsg = {
        id: Date.now() + 1,
        role: 'model',
        text: res.answer,
        modelUsed: res.modelUsed,
        time: new Date().toLocaleTimeString('ko-KR', { hour: '2-digit', minute: '2-digit' })
      };
      setQaMessages([...newHistory, botMsg]);
    } catch (err) {
      setQaError(err.message || '답변 생성 중 오류가 발생했습니다.');
    } finally {
      setIsQaLoading(false);
    }
  };

  const handleCopyQa = (id, text) => {
    navigator.clipboard.writeText(text).then(() => {
      setCopiedQaId(id);
      setTimeout(() => setCopiedQaId(null), 2000);
    });
  };

  const handleClearQa = () => {
    if (window.confirm('질의응답 대화 기록을 초기화하시겠습니까?')) {
      setQaMessages([]);
      setQaError('');
    }
  };

  // ── 오디오 드래그앤드롭 ───────────────────────
  const handleDragOver  = useCallback((e) => { e.preventDefault(); setIsDragging(true); }, []);
  const handleDragLeave = useCallback(() => setIsDragging(false), []);
  const handleDrop = useCallback((e) => {
    e.preventDefault(); setIsDragging(false);
    const file = e.dataTransfer.files[0];
    if (file) handleFileSelect(file);
  }, []);

  const handleFileSelect = async (file) => {
    const ext = file.name.split('.').pop().toLowerCase();
    const audioExts = ['mp3', 'wav', 'm4a', 'ogg', 'flac', 'aac', 'webm', 'mp4'];
    
    if (audioExts.includes(ext) || file.type.startsWith('audio/')) {
      setInputMode('audio');
      setAudioFile(file);
      setTextInput('');
      setResult(null); setError('');
    } else {
      // 텍스트/문서 파일
      setInputMode('text');
      setIsLoading(true);
      try {
        const result = await processFile(file);
        const text = result.text;
        if (!text || text.trim() === '') throw new Error('텍스트를 추출할 수 없습니다.');
        setTextInput(text);
        setAudioFile(null);
        setResult(null); setError('');
      } catch (err) {
        setError(err.message || '파일 추출 중 오류가 발생했습니다.');
      } finally {
        setIsLoading(false);
      }
    }
  };

  // ── 수동 용어 추가 ──────────────────────────
  const handleAddTerm = () => {
    if (!termWord.trim()) return;
    const newTerms = [...terminology, { id: Date.now(), word: termWord.trim(), desc: termDesc.trim() }];
    setTerminology(newTerms); saveTerminology(newTerms);
    setTermWord(''); setTermDesc('');
  };

  const handleRemoveTerm = (id) => {
    const newTerms = terminology.filter(t => t.id !== id);
    setTerminology(newTerms); saveTerminology(newTerms);
  };

  // ── 용어 파일 처리 (클릭/드롭 공통) ──────────
  const handleTermFilePick = useCallback(async (file) => {
    const keys = String(apiKey).split(',').map(k => k.trim()).filter(k => k.match(/^(AIza|AQ\.)/));
    if (keys.length === 0) { setExtractError('설정에서 Gemini API 키를 먼저 입력하세요.'); return; }

    setIsExtractingTerms(true);
    setExtractedTerms(null);
    setSelectedTermIds(new Set());
    setExtractError('');

    try {
      const result = await processFile(file);
      const text = result.text;
      if (!text || text.trim().length < 10) throw new Error('파일에서 텍스트를 추출하지 못했습니다.');
      const terms = await extractTermsFromText(text, apiKey);
      const existing = new Set(terminology.map(t => t.word.toLowerCase()));
      const filtered = terms.filter(t => !existing.has(t.word.toLowerCase()));
      setExtractedTerms(filtered);
      setSelectedTermIds(new Set(filtered.map((_, i) => i)));
    } catch (e) {
      setExtractError(e.message || '용어 추출에 실패했습니다.');
    } finally {
      setIsExtractingTerms(false);
    }
  }, [apiKey, terminology]);

  // ── 용어 추출 파일: 드래그앤드롭 핸들러 ────────
  const handleTermDragOver = useCallback((e) => {
    e.preventDefault(); e.stopPropagation(); setIsTermDragging(true);
  }, []);
  const handleTermDragLeave = useCallback((e) => {
    e.preventDefault(); e.stopPropagation(); setIsTermDragging(false);
  }, []);
  const handleTermDrop = useCallback((e) => {
    e.preventDefault(); e.stopPropagation(); setIsTermDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) handleTermFilePick(file);
  }, [handleTermFilePick]);

  // input[type=file] onChange
  const handleTermFileChange = useCallback((e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    e.target.value = '';
    handleTermFilePick(file);
  }, [handleTermFilePick]);

  const toggleSelect = (idx) => {
    setSelectedTermIds(prev => {
      const next = new Set(prev);
      next.has(idx) ? next.delete(idx) : next.add(idx);
      return next;
    });
  };

  const toggleSelectAll = () => {
    if (!extractedTerms) return;
    if (selectedTermIds.size === extractedTerms.length) {
      setSelectedTermIds(new Set());
    } else {
      setSelectedTermIds(new Set(extractedTerms.map((_, i) => i)));
    }
  };

  const handleRegisterSelected = () => {
    if (!extractedTerms) return;
    const toAdd = extractedTerms
      .filter((_, i) => selectedTermIds.has(i))
      .map(t => ({ id: Date.now() + Math.random(), word: t.word, desc: t.desc || '' }));
    const newTerms = [...terminology, ...toAdd];
    setTerminology(newTerms); saveTerminology(newTerms);
    setExtractedTerms(null); setSelectedTermIds(new Set());
  };

  // ── 회의록 분석 ────────────────────────────
  const handleAnalyze = async () => {
    if (inputMode === 'audio' && !audioFile) { setError('오디오 파일을 먼저 업로드하세요.'); return; }
    if (inputMode === 'text' && !textInput.trim()) { setError('분석할 텍스트를 먼저 입력하세요.'); return; }
    
    const keys = String(apiKey).split(',').map(k => k.trim()).filter(k => k.match(/^(AIza|AQ\.)/));
    if (keys.length === 0) { setError('설정에서 Gemini API 키를 먼저 입력하세요.'); return; }
    
    setIsLoading(true); setError(''); setResult(null);
    try {
      const inputData = inputMode === 'audio' ? audioFile : textInput;
      const data = await analyzeMeeting(inputData, inputMode, apiKey, terminology, (msg) => setProgress(msg));
      setResult(data);
    } catch (e) {
      setError(e.message || '분석 중 오류가 발생했습니다.');
    } finally {
      setIsLoading(false); setProgress('');
    }
  };

  const handleCopy = () => {
    navigator.clipboard.writeText(buildPlainText(result)).then(() => {
      setCopied(true); setTimeout(() => setCopied(false), 2000);
    });
  };

  const handleDownload = () => {
    const text = buildPlainText(result);
    const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `회의록_${(result?.meetingTitle || 'meeting').replace(/\s+/g, '_')}.txt`;
    a.click(); URL.revokeObjectURL(url);
  };

  // ────────────────────────────────────────────
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: '20px', maxWidth: '1100px', margin: '0 auto', padding: '8px 0' }}>

      {/* 헤더 */}
      <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
        <div style={{ width: '40px', height: '40px', borderRadius: '12px', background: 'linear-gradient(135deg, #6366f1, #8b5cf6)', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
          <Mic size={20} color="white" />
        </div>
        <div>
          <h2 style={{ fontSize: '18px', fontWeight: 700, color: 'var(--text-primary)', margin: 0, display: 'flex', alignItems: 'center', gap: '6px' }}>
            AI 회의록 생성기
            <Sparkles size={16} style={{ color: '#38bdf8', filter: 'drop-shadow(0 0 2px rgba(56, 189, 248, 0.5))' }} />
          </h2>
          <p style={{ fontSize: '12px', color: 'var(--text-muted)', margin: 0 }}>음성 파일 업로드 → 화자 분류 → 결정사항·액션아이템 자동 추출</p>
        </div>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 330px', gap: '16px', alignItems: 'start' }}>

        {/* ── 왼쪽: 업로드 + 분석 ── */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: '14px' }}>
        
          {/* 모드 전환 탭 */}
          <div style={{ display: 'flex', background: 'rgba(0,0,0,0.2)', padding: '4px', borderRadius: '10px' }}>
            <button
              onClick={() => setInputMode('audio')}
              style={{
                flex: 1, padding: '8px', borderRadius: '8px', fontSize: '13px', fontWeight: 600, border: 'none',
                background: inputMode === 'audio' ? 'rgba(99,102,241,0.2)' : 'transparent',
                color: inputMode === 'audio' ? '#818cf8' : 'var(--text-muted)',
                cursor: 'pointer', transition: 'all 0.2s',
              }}
            >
              오디오 파일 분석
            </button>
            <button
              onClick={() => setInputMode('text')}
              style={{
                flex: 1, padding: '8px', borderRadius: '8px', fontSize: '13px', fontWeight: 600, border: 'none',
                background: inputMode === 'text' ? 'rgba(16,185,129,0.2)' : 'transparent',
                color: inputMode === 'text' ? '#34d399' : 'var(--text-muted)',
                cursor: 'pointer', transition: 'all 0.2s',
              }}
            >
              녹취록(문서) 직접 분석
            </button>
          </div>

          {/* 오디오 파일 입력 모드 */}
          {inputMode === 'audio' && (
            <>
              <div
                onDragOver={handleDragOver} onDragLeave={handleDragLeave} onDrop={handleDrop}
                onClick={() => !audioFile && fileInputRef.current?.click()}
                style={{
                  border: `2px dashed ${isDragging ? 'var(--accent-blue)' : audioFile ? 'rgba(99,102,241,0.5)' : 'var(--glass-border)'}`,
                  borderRadius: '14px', padding: '28px', textAlign: 'center',
                  cursor: audioFile ? 'default' : 'pointer',
                  background: isDragging ? 'rgba(99,102,241,0.08)' : audioFile ? 'rgba(99,102,241,0.05)' : 'rgba(255,255,255,0.02)',
                  transition: 'all 0.2s ease',
                }}
              >
                <input ref={fileInputRef} type="file" accept="audio/*,.mp3,.wav,.m4a,.ogg,.flac,.aac,.webm"
                  style={{ display: 'none' }} onChange={(e) => e.target.files[0] && handleFileSelect(e.target.files[0])} />
                {audioFile ? (
                  <div style={{ display: 'flex', alignItems: 'center', gap: '14px', justifyContent: 'center' }}>
                    <div style={{ width: '44px', height: '44px', borderRadius: '12px', background: 'rgba(99,102,241,0.2)', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                      <FileAudio size={22} color="#818cf8" />
                    </div>
                    <div style={{ textAlign: 'left' }}>
                      <div style={{ fontWeight: 700, fontSize: '14px', color: 'var(--text-primary)' }}>{audioFile.name}</div>
                      <div style={{ fontSize: '12px', color: 'var(--text-muted)' }}>{fmtSize(audioFile.size)} · {audioFile.size > 20 * 1024 * 1024 ? 'Files API 방식' : '인라인 방식'}</div>
                    </div>
                    <button onClick={(e) => { e.stopPropagation(); setAudioFile(null); setResult(null); }}
                      style={{ marginLeft: 'auto', background: 'rgba(239,68,68,0.1)', border: '1px solid rgba(239,68,68,0.3)', borderRadius: '8px', padding: '6px', cursor: 'pointer', color: '#ef4444', display: 'flex', alignItems: 'center' }}>
                      <X size={14} />
                    </button>
                  </div>
                ) : (
                  <>
                    <div style={{ width: '52px', height: '52px', borderRadius: '14px', background: 'rgba(99,102,241,0.1)', border: '1px solid rgba(99,102,241,0.2)', display: 'flex', alignItems: 'center', justifyContent: 'center', margin: '0 auto 12px' }}>
                      <Upload size={24} color="#818cf8" />
                    </div>
                    <p style={{ fontSize: '14px', color: 'var(--text-primary)', fontWeight: 600, margin: '0 0 4px' }}>오디오 파일을 드래그하거나 클릭하여 업로드</p>
                    <p style={{ fontSize: '12px', color: 'var(--text-muted)', margin: 0 }}>MP3, WAV, M4A, OGG, FLAC, AAC, WEBM · 최대 2GB</p>
                  </>
                )}
              </div>
              {audioFile && (
                <audio controls src={URL.createObjectURL(audioFile)} style={{ width: '100%', borderRadius: '10px', height: '40px' }} />
              )}
            </>
          )}

          {/* 텍스트/문서 파일 입력 모드 */}
          {inputMode === 'text' && (
            <div
              onDragOver={handleDragOver} onDragLeave={handleDragLeave} onDrop={handleDrop}
              style={{
                border: `2px dashed ${isDragging ? '#34d399' : 'var(--glass-border)'}`,
                borderRadius: '14px', padding: '16px',
                background: isDragging ? 'rgba(16,185,129,0.08)' : 'rgba(255,255,255,0.02)',
                transition: 'all 0.2s ease', position: 'relative',
              }}
            >
              <textarea
                value={textInput}
                onChange={(e) => setTextInput(e.target.value)}
                placeholder="추출된 회의 기록 화면 텍스트를 화면에 직접 붙여넣거나, 문서 파일(PDF, PPTX, TXT, DOXC 등)을 바로 여기에 드래그앤드롭 하세요."
                style={{
                  width: '100%', minHeight: '145px', background: 'rgba(0,0,0,0.2)', border: '1px solid var(--glass-border)',
                  color: 'var(--text-primary)', padding: '12px', borderRadius: '8px', resize: 'vertical', fontSize: '13px', lineHeight: 1.6
                }}
              />
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '10px' }}>
                <span style={{ fontSize: '11px', color: 'var(--text-muted)' }}>
                  {textInput.length > 0 ? `현재 ${textInput.length.toLocaleString()} 글자 입력됨` : '문서 드래그앤드롭 지원'}
                </span>
                <input ref={textFileInputRef} type="file" accept=".pdf,.txt,.md,.csv,.xlsx,.xls,.pptx,.hwpx,.json,.html,.xml"
                  style={{ display: 'none' }} onChange={(e) => e.target.files[0] && handleFileSelect(e.target.files[0])} />
                <button
                  onClick={() => textFileInputRef.current?.click()}
                  style={{ background: 'rgba(255,255,255,0.05)', border: '1px solid var(--glass-border)', borderRadius: '6px', padding: '4px 8px', fontSize: '11px', color: 'var(--text-secondary)', cursor: 'pointer' }}
                >
                  문서 업로드
                </button>
              </div>
            </div>
          )}

          <button onClick={handleAnalyze} disabled={isLoading || (inputMode === 'audio' ? !audioFile : (!textInput || !textInput.trim()))}
            style={{
              width: '100%', padding: '14px', borderRadius: '12px', border: 'none',
              background: isLoading || (inputMode === 'audio' ? !audioFile : !textInput?.trim()) ? 'rgba(255,255,255,0.05)' : (inputMode === 'audio' ? 'linear-gradient(135deg, #6366f1, #8b5cf6)' : 'linear-gradient(135deg, #10b981, #34d399)'),
              color: isLoading || (inputMode === 'audio' ? !audioFile : !textInput?.trim()) ? 'var(--text-muted)' : (inputMode === 'audio' ? 'white' : '#000'),
              fontWeight: 700, fontSize: '15px', cursor: isLoading || (inputMode === 'audio' ? !audioFile : !textInput?.trim()) ? 'not-allowed' : 'pointer',
              display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px',
              transition: 'all 0.2s ease',
              boxShadow: isLoading || (inputMode === 'audio' ? !audioFile : !textInput?.trim()) ? 'none' : (inputMode === 'audio' ? '0 4px 20px rgba(99,102,241,0.35)' : '0 4px 20px rgba(16,185,129,0.35)'),
            }}>
            {isLoading
              ? <><Loader2 size={18} style={{ animation: 'spin 1s linear infinite' }} /> 분석 중...</>
              : <><Sparkles size={18} /> 회의록 생성 시작</>}
          </button>

          {isLoading && progress && (
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', padding: '10px 14px', background: 'rgba(99,102,241,0.08)', borderRadius: '10px', border: '1px solid rgba(99,102,241,0.2)', fontSize: '13px', color: '#818cf8' }}>
              <Loader2 size={14} style={{ animation: 'spin 1s linear infinite', flexShrink: 0 }} />
              {progress}
            </div>
          )}

          {error && (
            <div style={{ display: 'flex', alignItems: 'flex-start', gap: '8px', padding: '12px 14px', background: 'rgba(239,68,68,0.08)', borderRadius: '10px', border: '1px solid rgba(239,68,68,0.25)', fontSize: '13px', color: '#f87171' }}>
              <AlertCircle size={16} style={{ flexShrink: 0, marginTop: '1px' }} />{error}
            </div>
          )}
        </div>

        {/* ── 오른쪽: 전문 용어 사전 ── */}
        <div style={{ background: 'var(--glass-bg)', border: '1px solid var(--glass-border)', borderRadius: '14px', overflow: 'hidden' }}>

          {/* 헤더 토글 */}
          <button onClick={() => setShowTermPanel(!showTermPanel)}
            style={{ width: '100%', padding: '14px 16px', background: 'none', border: 'none', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '8px', color: 'var(--text-primary)' }}>
            <BookOpen size={16} color="#fbbf24" />
            <span style={{ fontWeight: 700, fontSize: '13px', flex: 1, textAlign: 'left' }}>전문 용어 사전</span>
            <span style={{ fontSize: '11px', color: 'var(--text-muted)', marginRight: '4px' }}>{terminology.length}개</span>
            {showTermPanel ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
          </button>

          {showTermPanel && (
            <div style={{ padding: '0 14px 14px', display: 'flex', flexDirection: 'column', gap: '10px' }}>
              <p style={{ fontSize: '11px', color: 'var(--text-muted)', margin: 0 }}>등록한 단어는 회의록 작성 시 자동 적용됩니다.</p>

              {/* ── 파일에서 자동 추출 드롭존 ── */}
              <div
                onDragOver={handleTermDragOver}
                onDragLeave={handleTermDragLeave}
                onDrop={handleTermDrop}
                style={{
                  padding: '10px',
                  background: isTermDragging
                    ? 'rgba(251,191,36,0.12)'
                    : 'rgba(251,191,36,0.06)',
                  borderRadius: '10px',
                  border: isTermDragging
                    ? '2px dashed rgba(251,191,36,0.7)'
                    : '1px dashed rgba(251,191,36,0.3)',
                  transition: 'all 0.15s ease',
                }}
              >
                <p style={{ fontSize: '11px', color: '#fbbf24', fontWeight: 700, margin: '0 0 4px', display: 'flex', alignItems: 'center', gap: '5px' }}>
                  <FileSearch size={12} /> 파일에서 용어 자동 추출
                </p>
                <p style={{ fontSize: '10px', color: 'var(--text-muted)', margin: '0 0 8px', lineHeight: 1.5 }}>
                  문서를 이 영역에 <strong style={{ color: 'var(--text-secondary)' }}>드래그</strong>하거나 버튼으로 선택하세요.<br />
                  PDF·TXT·XLSX·PPTX·HWPX 지원
                </p>
                <input ref={termFileRef} type="file"
                  accept=".pdf,.txt,.md,.csv,.xlsx,.xls,.pptx,.hwpx,.json,.html,.xml"
                  style={{ display: 'none' }}
                  onChange={handleTermFileChange}
                />
                <button
                  onClick={() => termFileRef.current?.click()}
                  disabled={isExtractingTerms}
                  style={{
                    width: '100%', padding: '7px 10px', borderRadius: '8px',
                    background: isExtractingTerms ? 'rgba(255,255,255,0.03)' : 'rgba(251,191,36,0.12)',
                    border: '1px solid rgba(251,191,36,0.35)',
                    color: isExtractingTerms ? 'var(--text-muted)' : '#fbbf24',
                    cursor: isExtractingTerms ? 'not-allowed' : 'pointer',
                    display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '6px',
                    fontSize: '12px', fontWeight: 600,
                  }}
                >
                  {isExtractingTerms
                    ? <><Loader2 size={13} style={{ animation: 'spin 1s linear infinite' }} /> 분석 중...</>
                    : <><Upload size={13} /> 파일 선택하여 용어 추출</>}
                </button>
                {extractError && (
                  <p style={{ fontSize: '11px', color: '#f87171', margin: '6px 0 0' }}>{extractError}</p>
                )}
              </div>

              {/* ── 추출된 용어 목록 (체크박스 선택) ── */}
              {extractedTerms && extractedTerms.length > 0 && (
                <div style={{ background: 'rgba(255,255,255,0.02)', border: '1px solid var(--glass-border)', borderRadius: '10px', overflow: 'hidden' }}>
                  {/* 헤더 */}
                  <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '8px 10px', borderBottom: '1px solid var(--glass-border)' }}>
                    <button onClick={toggleSelectAll}
                      style={{ background: 'none', border: 'none', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '5px', color: 'var(--text-secondary)', fontSize: '11px', fontWeight: 600, padding: 0 }}>
                      {selectedTermIds.size === extractedTerms.length
                        ? <CheckSquare size={13} color="#818cf8" />
                        : <Square size={13} color="var(--text-muted)" />}
                      전체선택 ({selectedTermIds.size}/{extractedTerms.length})
                    </button>
                    <button onClick={handleRegisterSelected}
                      disabled={selectedTermIds.size === 0}
                      style={{
                        padding: '4px 10px', borderRadius: '6px', border: 'none',
                        background: selectedTermIds.size === 0 ? 'rgba(255,255,255,0.04)' : 'rgba(99,102,241,0.2)',
                        color: selectedTermIds.size === 0 ? 'var(--text-muted)' : '#818cf8',
                        cursor: selectedTermIds.size === 0 ? 'not-allowed' : 'pointer',
                        fontSize: '11px', fontWeight: 700,
                      }}>
                      선택 등록 ({selectedTermIds.size})
                    </button>
                  </div>

                  {/* 용어 체크박스 목록 */}
                  <div style={{ maxHeight: '200px', overflowY: 'auto' }}>
                    {extractedTerms.map((t, i) => (
                      <div key={i}
                        onClick={() => toggleSelect(i)}
                        style={{
                          display: 'flex', alignItems: 'center', gap: '7px',
                          padding: '6px 10px', cursor: 'pointer',
                          background: selectedTermIds.has(i) ? 'rgba(99,102,241,0.07)' : 'transparent',
                          borderBottom: '1px solid rgba(255,255,255,0.03)',
                          transition: 'background 0.15s',
                        }}>
                        {selectedTermIds.has(i)
                          ? <CheckSquare size={13} color="#818cf8" style={{ flexShrink: 0 }} />
                          : <Square size={13} color="var(--text-muted)" style={{ flexShrink: 0 }} />}
                        <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--text-primary)', minWidth: '50px' }}>{t.word}</span>
                        <span style={{ fontSize: '11px', color: 'var(--text-muted)', flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{t.desc || '-'}</span>
                        {t.freq > 0 && (
                          <span style={{ fontSize: '10px', color: 'var(--text-muted)', background: 'rgba(255,255,255,0.05)', padding: '1px 5px', borderRadius: '4px', flexShrink: 0 }}>{t.freq}회</span>
                        )}
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {extractedTerms && extractedTerms.length === 0 && (
                <p style={{ fontSize: '12px', color: 'var(--text-muted)', textAlign: 'center', padding: '8px 0', margin: 0 }}>
                  추출된 새 전문 용어가 없습니다. (이미 모두 등록됨)
                </p>
              )}

              {/* ── 구분선 ── */}
              <div style={{ borderTop: '1px solid var(--glass-border)', paddingTop: '10px' }}>
                <p style={{ fontSize: '11px', color: 'var(--text-muted)', margin: '0 0 6px', fontWeight: 600 }}>직접 추가</p>
              </div>

              {/* ── 수동 추가 ── */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                <input value={termWord} onChange={e => setTermWord(e.target.value)}
                  onKeyDown={e => e.key === 'Enter' && handleAddTerm()}
                  placeholder="단어 / 약어 (예: RFP)" style={inputStyle} />
                <input value={termDesc} onChange={e => setTermDesc(e.target.value)}
                  onKeyDown={e => e.key === 'Enter' && handleAddTerm()}
                  placeholder="설명 (예: 제안요청서)" style={inputStyle} />
                <button onClick={handleAddTerm}
                  style={{ padding: '7px', borderRadius: '8px', background: 'rgba(251,191,36,0.15)', border: '1px solid rgba(251,191,36,0.35)', color: '#fbbf24', cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '4px', fontSize: '12px', fontWeight: 600 }}>
                  <Plus size={14} /> 용어 추가
                </button>
              </div>

              {/* ── 등록된 용어 목록 ── */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: '5px', maxHeight: '200px', overflowY: 'auto' }}>
                {terminology.length === 0 && (
                  <div style={{ fontSize: '12px', color: 'var(--text-muted)', textAlign: 'center', padding: '12px 0' }}>등록된 용어가 없습니다.</div>
                )}
                {terminology.map(t => (
                  <div key={t.id} style={{ display: 'flex', alignItems: 'center', gap: '6px', padding: '6px 8px', background: 'rgba(255,255,255,0.03)', borderRadius: '7px', border: '1px solid var(--glass-border)' }}>
                    <Tag size={11} color="#fbbf24" style={{ flexShrink: 0 }} />
                    <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--text-primary)', minWidth: '50px' }}>{t.word}</span>
                    <span style={{ fontSize: '11px', color: 'var(--text-muted)', flex: 1 }}>{t.desc || '-'}</span>
                    <button onClick={() => handleRemoveTerm(t.id)}
                      style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#ef4444', padding: '2px', display: 'flex' }}>
                      <Trash2 size={12} />
                    </button>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      </div>

      {/* ── 결과 섹션 ── */}
      {result && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '16px', marginTop: '4px' }}>

          {/* 결과 헤더 */}
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '16px 20px', background: 'linear-gradient(135deg, rgba(99,102,241,0.1), rgba(139,92,246,0.08))', border: '1px solid rgba(99,102,241,0.25)', borderRadius: '14px' }}>
            <div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '4px' }}>
                <CheckCircle2 size={16} color="#34d399" />
                <span style={{ fontSize: '11px', color: '#34d399', fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.5px' }}>회의록 생성 완료</span>
              </div>
              <h3 style={{ margin: 0, fontSize: '17px', fontWeight: 800, color: 'var(--text-primary)' }}>{result.meetingTitle || '회의록'}</h3>
              {result.speakerCount > 0 && (
                <div style={{ display: 'flex', alignItems: 'center', gap: '6px', marginTop: '4px' }}>
                  <Users size={12} color="var(--text-muted)" />
                  <span style={{ fontSize: '12px', color: 'var(--text-muted)' }}>화자 {result.speakerCount}명 감지</span>
                </div>
              )}
            </div>
            <div style={{ display: 'flex', gap: '8px' }}>
              <button onClick={handleCopy}
                style={{ display: 'flex', alignItems: 'center', gap: '6px', padding: '8px 14px', borderRadius: '9px', background: 'rgba(255,255,255,0.06)', border: '1px solid var(--glass-border)', color: copied ? '#34d399' : 'var(--text-secondary)', cursor: 'pointer', fontSize: '12px', fontWeight: 600 }}>
                <Copy size={13} /> {copied ? '복사됨!' : '복사'}
              </button>
              <button onClick={handleDownload}
                style={{ display: 'flex', alignItems: 'center', gap: '6px', padding: '8px 14px', borderRadius: '9px', background: 'rgba(99,102,241,0.15)', border: '1px solid rgba(99,102,241,0.35)', color: '#818cf8', cursor: 'pointer', fontSize: '12px', fontWeight: 600 }}>
                <Download size={13} /> TXT 저장
              </button>
            </div>
          </div>

          {/* 1영역: 회의 대활 */}
          {result.meetingContext && (
            <div style={{ padding: '16px 20px', background: 'rgba(99,102,241,0.06)', border: '1px solid rgba(99,102,241,0.2)', borderRadius: '12px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '10px' }}>
                <MapPin size={15} color="#818cf8" />
                <span style={{ fontSize: '13px', fontWeight: 700, color: 'var(--text-primary)' }}>회의 목적 및 배경</span>
              </div>
              <p style={{ margin: 0, fontSize: '14px', color: 'var(--text-secondary)', lineHeight: 1.8 }}>{result.meetingContext}</p>
            </div>
          )}

          {/* 2영역: 안건 + 요약 */}
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '14px' }}>
            <div style={{ padding: '16px 20px', background: 'var(--glass-bg)', border: '1px solid var(--glass-border)', borderRadius: '12px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '10px' }}>
                <ClipboardList size={15} color="#818cf8" />
                <span style={{ fontSize: '13px', fontWeight: 700, color: 'var(--text-primary)' }}>회의 안건</span>
              </div>
              <p style={{ margin: 0, fontSize: '13px', color: 'var(--text-secondary)', lineHeight: 1.8 }}>{result.agenda || '-'}</p>
            </div>
            <div style={{ padding: '16px 20px', background: 'var(--glass-bg)', border: '1px solid var(--glass-border)', borderRadius: '12px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '10px' }}>
                <BarChart3 size={15} color="#34d399" />
                <span style={{ fontSize: '13px', fontWeight: 700, color: 'var(--text-primary)' }}>종합 요약</span>
              </div>
              <p style={{ margin: 0, fontSize: '13px', color: 'var(--text-secondary)', lineHeight: 1.8 }}>{result.summary || '-'}</p>
            </div>
          </div>

          {/* 3영역: 주제별 상세 논의 */}
          {(result.topicSummaries || []).length > 0 && (
            <div style={{ background: 'var(--glass-bg)', border: '1px solid var(--glass-border)', borderRadius: '12px', overflow: 'hidden' }}>
              <div style={{ padding: '14px 18px', borderBottom: '1px solid var(--glass-border)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                <BookOpen size={15} color="#a78bfa" />
                <span style={{ fontSize: '13px', fontWeight: 700, color: 'var(--text-primary)' }}>주제별 상세 논의 내용</span>
                <span style={{ fontSize: '11px', color: 'var(--text-muted)', marginLeft: '4px' }}>({result.topicSummaries.length}개 주제)</span>
              </div>
              <div style={{ padding: '14px 18px', display: 'flex', flexDirection: 'column', gap: '14px' }}>
                {result.topicSummaries.map((topic, i) => (
                  <div key={i} style={{ padding: '14px 16px', background: 'rgba(167,139,250,0.06)', borderRadius: '10px', borderLeft: '3px solid rgba(167,139,250,0.5)' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '8px' }}>
                      <span style={{ fontSize: '11px', fontWeight: 800, color: '#a78bfa', background: 'rgba(167,139,250,0.15)', padding: '2px 8px', borderRadius: '20px' }}>TOPIC {i + 1}</span>
                      <span style={{ fontSize: '13px', fontWeight: 700, color: 'var(--text-primary)' }}>{topic.topic}</span>
                    </div>
                    <p style={{ margin: '0 0 8px', fontSize: '13px', color: 'var(--text-secondary)', lineHeight: 1.8 }}>{topic.content}</p>
                    {topic.result && (
                      <div style={{ display: 'flex', alignItems: 'flex-start', gap: '6px', padding: '8px 10px', background: 'rgba(52,211,153,0.08)', borderRadius: '6px', border: '1px solid rgba(52,211,153,0.2)' }}>
                        <ArrowRight size={13} color="#34d399" style={{ flexShrink: 0, marginTop: '2px' }} />
                        <span style={{ fontSize: '12px', color: '#34d399', fontWeight: 600 }}>{topic.result}</span>
                      </div>
                    )}
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* 4영역: 핵심 인사이트 */}
          {(result.keyInsights || []).length > 0 && (
            <div style={{ background: 'var(--glass-bg)', border: '1px solid var(--glass-border)', borderRadius: '12px', overflow: 'hidden' }}>
              <div style={{ padding: '14px 18px', borderBottom: '1px solid var(--glass-border)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                <Lightbulb size={15} color="#fbbf24" />
                <span style={{ fontSize: '13px', fontWeight: 700, color: 'var(--text-primary)' }}>핵심 인사이트</span>
              </div>
              <div style={{ padding: '12px 16px', display: 'flex', flexDirection: 'column', gap: '8px' }}>
                {result.keyInsights.map((ins, i) => {
                  const typeConfig = {
                    '기회': { color: '#34d399', bg: 'rgba(52,211,153,0.08)', border: 'rgba(52,211,153,0.25)' },
                    '리스크': { color: '#f87171', bg: 'rgba(248,113,113,0.08)', border: 'rgba(248,113,113,0.25)' },
                    '통찰': { color: '#818cf8', bg: 'rgba(129,140,248,0.08)', border: 'rgba(129,140,248,0.25)' },
                    '우선순위': { color: '#fbbf24', bg: 'rgba(251,191,36,0.08)', border: 'rgba(251,191,36,0.25)' },
                  };
                  const cfg = typeConfig[ins.type] || typeConfig['통찰'];
                  return (
                    <div key={i} style={{ display: 'flex', alignItems: 'flex-start', gap: '10px', padding: '10px 12px', background: cfg.bg, borderRadius: '8px', border: `1px solid ${cfg.border}` }}>
                      <span style={{ fontSize: '10px', fontWeight: 800, color: cfg.color, background: `${cfg.border}`, padding: '2px 7px', borderRadius: '10px', flexShrink: 0, marginTop: '1px' }}>{ins.type}</span>
                      <span style={{ fontSize: '13px', color: 'var(--text-secondary)', lineHeight: 1.6 }}>{ins.content}</span>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* 5영역: 결정 사항 + 액션 아이템 */}
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '14px' }}>
            <div style={{ padding: '16px 20px', background: 'var(--glass-bg)', border: '1px solid var(--glass-border)', borderRadius: '12px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '12px' }}>
                <CheckCircle2 size={15} color="#34d399" />
                <span style={{ fontSize: '13px', fontWeight: 700, color: 'var(--text-primary)' }}>결정된 사항</span>
              </div>
              {(result.decisions || []).length === 0
                ? <p style={{ fontSize: '13px', color: 'var(--text-muted)', margin: 0 }}>결정된 사항이 없습니다.</p>
                : (result.decisions || []).map((d, i) => (
                  <div key={i} style={{ display: 'flex', gap: '8px', padding: '8px 10px', marginBottom: '6px', background: 'rgba(16,185,129,0.06)', borderRadius: '8px', borderLeft: '3px solid rgba(16,185,129,0.5)' }}>
                    <span style={{ color: '#34d399', fontWeight: 700, fontSize: '12px', flexShrink: 0, marginTop: '1px' }}>•</span>
                    <span style={{ fontSize: '13px', color: 'var(--text-secondary)', lineHeight: 1.6 }}>{d}</span>
                  </div>
                ))}
            </div>

            <div style={{ padding: '16px 20px', background: 'var(--glass-bg)', border: '1px solid var(--glass-border)', borderRadius: '12px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '12px' }}>
                <ListChecks size={15} color="#fbbf24" />
                <span style={{ fontSize: '13px', fontWeight: 700, color: 'var(--text-primary)' }}>할 일 및 담당자</span>
              </div>
              {(result.actionItems || []).length === 0
                ? <p style={{ fontSize: '13px', color: 'var(--text-muted)', margin: 0 }}>추출된 액션 아이템이 없습니다.</p>
                : (result.actionItems || []).map((a, i) => {
                  const priorityColor = a.priority === '높음' ? '#f87171' : a.priority === '낙음' ? '#6ee7b7' : '#fbbf24';
                  const priorityBg = a.priority === '높음' ? 'rgba(248,113,113,0.1)' : a.priority === '낙음' ? 'rgba(110,231,183,0.1)' : 'rgba(251,191,36,0.08)';
                  return (
                    <div key={i} style={{ padding: '9px 12px', marginBottom: '6px', background: 'rgba(245,158,11,0.06)', borderRadius: '8px', borderLeft: '3px solid rgba(245,158,11,0.5)' }}>
                      <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', gap: '6px' }}>
                        <div style={{ fontSize: '13px', color: 'var(--text-secondary)', fontWeight: 500, lineHeight: 1.5, flex: 1 }}>{a.task}</div>
                        {a.priority && (
                          <span style={{ fontSize: '10px', fontWeight: 700, color: priorityColor, background: priorityBg, padding: '1px 6px', borderRadius: '8px', flexShrink: 0 }}>{a.priority}</span>
                        )}
                      </div>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginTop: '4px' }}>
                        <span style={{ fontSize: '11px', color: '#fbbf24', fontWeight: 700 }}>👤 {a.owner || '미정'}</span>
                        {a.deadline && <span style={{ fontSize: '11px', color: 'var(--text-muted)', display: 'flex', alignItems: 'center', gap: '3px' }}><Clock size={10} /> {a.deadline}</span>}
                      </div>
                    </div>
                  );
                })}
            </div>
          </div>

          {/* 6영역: 향후 계획 */}
          {result.futurePlans && ((result.futurePlans.shortTerm || []).length > 0 || (result.futurePlans.longTerm || []).length > 0) && (
            <div style={{ background: 'var(--glass-bg)', border: '1px solid var(--glass-border)', borderRadius: '12px', overflow: 'hidden' }}>
              <div style={{ padding: '14px 18px', borderBottom: '1px solid var(--glass-border)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                <TrendingUp size={15} color="#60a5fa" />
                <span style={{ fontSize: '13px', fontWeight: 700, color: 'var(--text-primary)' }}>향후 계획</span>
              </div>
              <div style={{ padding: '14px 18px', display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '14px' }}>
                {/* 단기 계획 */}
                {(result.futurePlans.shortTerm || []).length > 0 && (
                  <div>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '6px', marginBottom: '10px' }}>
                      <Zap size={13} color="#34d399" />
                      <span style={{ fontSize: '12px', fontWeight: 700, color: '#34d399' }}>단기 계획 (즉시 ~ 1개월)</span>
                    </div>
                    {result.futurePlans.shortTerm.map((p, i) => (
                      <div key={i} style={{ padding: '9px 12px', marginBottom: '6px', background: 'rgba(52,211,153,0.06)', borderRadius: '8px', border: '1px solid rgba(52,211,153,0.2)' }}>
                        <div style={{ fontSize: '13px', color: 'var(--text-secondary)', lineHeight: 1.6, marginBottom: '4px' }}>{p.plan}</div>
                        <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
                          {p.owner && <span style={{ fontSize: '11px', color: '#34d399', fontWeight: 600 }}>👤 {p.owner}</span>}
                          {p.targetDate && <span style={{ fontSize: '11px', color: 'var(--text-muted)', display: 'flex', alignItems: 'center', gap: '3px' }}><Calendar size={10} />{p.targetDate}</span>}
                        </div>
                      </div>
                    ))}
                  </div>
                )}
                {/* 중장기 계획 */}
                {(result.futurePlans.longTerm || []).length > 0 && (
                  <div>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '6px', marginBottom: '10px' }}>
                      <Target size={13} color="#60a5fa" />
                      <span style={{ fontSize: '12px', fontWeight: 700, color: '#60a5fa' }}>중장기 계획 (1개월 이상)</span>
                    </div>
                    {result.futurePlans.longTerm.map((p, i) => (
                      <div key={i} style={{ padding: '9px 12px', marginBottom: '6px', background: 'rgba(96,165,250,0.06)', borderRadius: '8px', border: '1px solid rgba(96,165,250,0.2)' }}>
                        <div style={{ fontSize: '13px', color: 'var(--text-secondary)', lineHeight: 1.6, marginBottom: '4px' }}>{p.plan}</div>
                        <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
                          {p.owner && <span style={{ fontSize: '11px', color: '#60a5fa', fontWeight: 600 }}>👤 {p.owner}</span>}
                          {p.targetDate && <span style={{ fontSize: '11px', color: 'var(--text-muted)', display: 'flex', alignItems: 'center', gap: '3px' }}><Calendar size={10} />{p.targetDate}</span>}
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          )}

          {/* 키워드 */}
          {(result.keywords || []).length > 0 && (
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px', padding: '12px 16px', background: 'var(--glass-bg)', border: '1px solid var(--glass-border)', borderRadius: '10px' }}>
              <span style={{ fontSize: '11px', color: 'var(--text-muted)', fontWeight: 600, marginRight: '4px', alignSelf: 'center' }}>핵심 키워드</span>
              {result.keywords.map((kw, i) => (
                <span key={i} style={{ fontSize: '12px', padding: '3px 10px', background: 'rgba(99,102,241,0.1)', border: '1px solid rgba(99,102,241,0.25)', borderRadius: '20px', color: '#818cf8', fontWeight: 500 }}>{kw}</span>
              ))}
            </div>
          )}

          {/* 발언록 */}
          <div style={{ background: 'var(--glass-bg)', border: '1px solid var(--glass-border)', borderRadius: '12px', overflow: 'hidden' }}>
            <button onClick={() => setShowTranscript(!showTranscript)}
              style={{ width: '100%', padding: '14px 18px', background: 'none', border: 'none', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '8px', color: 'var(--text-primary)' }}>
              <Users size={15} color="#818cf8" />
              <span style={{ fontWeight: 700, fontSize: '13px', flex: 1, textAlign: 'left' }}>화자별 발언록 ({(result.transcript || []).length}건)</span>
              {showTranscript ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
            </button>
            {showTranscript && (
              <div style={{ padding: '0 16px 16px', display: 'flex', flexDirection: 'column', gap: '6px', maxHeight: '420px', overflowY: 'auto' }}>
                {(result.transcript || []).map((t, i) => {
                  const color = getSpeakerColor(t.speaker);
                  return (
                    <div key={i} style={{ display: 'flex', gap: '10px', padding: '10px 12px', background: color.bg, borderRadius: '9px', border: `1px solid ${color.border}` }}>
                      <span style={{ fontWeight: 800, fontSize: '13px', color: color.label, flexShrink: 0, minWidth: '20px' }}>{t.speaker}</span>
                      <span style={{ fontSize: '13px', color: 'var(--text-secondary)', lineHeight: 1.6, flex: 1 }}>
                        {t.text}
                        {t.tag && t.tag !== '\uc815\uc0c1' && (
                          <span style={{ marginLeft: '6px', fontSize: '10px', color: '#f87171', background: 'rgba(239,68,68,0.1)', padding: '1px 6px', borderRadius: '4px' }}>{t.tag}</span>
                        )}
                      </span>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
          {/* 7영역: 💬 회의 내용 AI 질의응답 (Q&A 어시스턴트) */}
          <div style={{
            background: 'linear-gradient(135deg, rgba(30, 27, 75, 0.4), rgba(15, 23, 42, 0.6))',
            border: '1px solid rgba(139, 92, 246, 0.3)',
            borderRadius: '14px',
            overflow: 'hidden',
            boxShadow: '0 8px 32px rgba(0, 0, 0, 0.25)'
          }}>
            {/* Q&A 헤더 */}
            <div style={{
              padding: '16px 20px',
              borderBottom: '1px solid rgba(139, 92, 246, 0.2)',
              display: 'flex',
              justifyContent: 'space-between',
              alignItems: 'center',
              background: 'rgba(139, 92, 246, 0.08)'
            }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                <div style={{ padding: '8px', background: 'rgba(139, 92, 246, 0.2)', borderRadius: '8px' }}>
                  <MessageSquare size={18} color="#c084fc" />
                </div>
                <div>
                  <div style={{ fontSize: '14px', fontWeight: 800, color: 'var(--text-primary)', display: 'flex', alignItems: 'center', gap: '6px' }}>
                    회의 내용 AI 질의응답 (Q&A)
                    <span style={{ fontSize: '11px', padding: '2px 8px', borderRadius: '10px', background: 'rgba(168, 85, 247, 0.2)', color: '#c084fc', fontWeight: 600 }}>
                      Gemini Multimodal
                    </span>
                  </div>
                  <div style={{ fontSize: '12px', color: 'var(--text-secondary)', marginTop: '2px' }}>
                    생성된 회의록 및 상세 발언록을 기반으로 궁금한 점을 자유롭게 질문하세요.
                  </div>
                </div>
              </div>

              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                {qaMessages.length > 0 && (
                  <button
                    onClick={handleClearQa}
                    title="대화 내역 초기화"
                    style={{
                      display: 'flex', alignItems: 'center', gap: '4px',
                      padding: '6px 12px', borderRadius: '6px',
                      background: 'rgba(239, 68, 68, 0.1)', border: '1px solid rgba(239, 68, 68, 0.25)',
                      color: '#f87171', fontSize: '11.5px', fontWeight: 600, cursor: 'pointer'
                    }}
                  >
                    <RotateCcw size={13} /> 초기화
                  </button>
                )}
                <button
                  onClick={() => setShowQaPanel(!showQaPanel)}
                  style={{
                    background: 'none', border: 'none', color: 'var(--text-muted)',
                    cursor: 'pointer', padding: '4px'
                  }}
                >
                  {showQaPanel ? <ChevronUp size={18} /> : <ChevronDown size={18} />}
                </button>
              </div>
            </div>

            {showQaPanel && (
              <div style={{ padding: '20px', display: 'flex', flexDirection: 'column', gap: '16px' }}>
                {/* 💡 추천 질문 프롬프트 칩 */}
                <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                  <div style={{ fontSize: '12px', color: 'var(--text-muted)', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '4px' }}>
                    <Sparkles size={13} color="#c084fc" /> 빠른 추천 질문:
                  </div>
                  <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
                    {[
                      '이번 회의의 핵심 쟁점과 결론은 무엇인가요?',
                      '담당자별 액션 아이템과 기한을 표 형태로 정리해줘',
                      '회의 중 제기된 주요 리스크와 대응 방안은?',
                      '차기 회의 전까지 완료해야 할 우선순위 과제는?'
                    ].map((qPrompt, idx) => (
                      <button
                        key={idx}
                        disabled={isQaLoading}
                        onClick={() => handleSendQa(qPrompt)}
                        style={{
                          padding: '6px 12px', borderRadius: '20px',
                          background: 'rgba(139, 92, 246, 0.1)',
                          border: '1px solid rgba(139, 92, 246, 0.3)',
                          color: '#ddd6fe', fontSize: '12px', fontWeight: 500,
                          cursor: isQaLoading ? 'not-allowed' : 'pointer',
                          display: 'flex', alignItems: 'center', gap: '6px',
                          transition: 'all 0.2s',
                          opacity: isQaLoading ? 0.6 : 1
                        }}
                      >
                        💡 {qPrompt}
                      </button>
                    ))}
                  </div>
                </div>

                {/* 대화 메시지 목록 창 */}
                <div style={{
                  minHeight: qaMessages.length > 0 ? '160px' : '90px',
                  maxHeight: '480px',
                  overflowY: 'auto',
                  display: 'flex',
                  flexDirection: 'column',
                  gap: '14px',
                  padding: '14px',
                  background: 'rgba(0, 0, 0, 0.25)',
                  borderRadius: '10px',
                  border: '1px solid rgba(255, 255, 255, 0.05)'
                }}>
                  {qaMessages.length === 0 ? (
                    <div style={{ padding: '30px 20px', textAlign: 'center', color: 'var(--text-muted)', fontSize: '13px' }}>
                      <HelpCircle size={28} color="#8b5cf6" style={{ margin: '0 auto 8px', opacity: 0.6 }} />
                      회의 내용에 대해 질문하시면 AI가 회의록과 발언록을 정밀 분석하여 즉시 답변해 드립니다.
                    </div>
                  ) : (
                    qaMessages.map((msg) => (
                      <div
                        key={msg.id}
                        style={{
                          display: 'flex',
                          flexDirection: 'column',
                          alignItems: msg.role === 'user' ? 'flex-end' : 'flex-start',
                          gap: '4px'
                        }}
                      >
                        <div style={{
                          display: 'flex',
                          alignItems: 'center',
                          gap: '6px',
                          fontSize: '11px',
                          color: 'var(--text-muted)'
                        }}>
                          <span>{msg.role === 'user' ? '👤 나의 질문' : '🤖 회의록 AI 어시스턴트'}</span>
                          {msg.modelUsed && (
                            <span style={{ fontSize: '10px', color: '#a78bfa', background: 'rgba(167,139,250,0.15)', padding: '1px 5px', borderRadius: '4px' }}>
                              {msg.modelUsed}
                            </span>
                          )}
                          <span>{msg.time}</span>
                        </div>

                        <div style={{
                          maxWidth: '88%',
                          padding: '12px 16px',
                          borderRadius: msg.role === 'user' ? '14px 14px 2px 14px' : '14px 14px 14px 2px',
                          background: msg.role === 'user'
                            ? 'linear-gradient(135deg, #6366f1, #4f46e5)'
                            : 'rgba(30, 41, 59, 0.85)',
                          border: msg.role === 'user'
                            ? '1px solid rgba(99, 102, 241, 0.4)'
                            : '1px solid rgba(148, 163, 184, 0.15)',
                          color: '#f8fafc',
                          fontSize: '13px',
                          lineHeight: '1.7',
                          whiteSpace: 'pre-wrap',
                          wordBreak: 'break-word',
                          boxShadow: '0 2px 8px rgba(0,0,0,0.2)'
                        }}>
                          {msg.text}
                          {msg.role === 'model' && (
                            <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: '8px', paddingTop: '6px', borderTop: '1px solid rgba(255,255,255,0.08)' }}>
                              <button
                                onClick={() => handleCopyQa(msg.id, msg.text)}
                                style={{
                                  background: 'none', border: 'none',
                                  color: copiedQaId === msg.id ? 'var(--success-color)' : 'var(--text-muted)',
                                  fontSize: '11px', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: '4px'
                                }}
                              >
                                {copiedQaId === msg.id ? <CheckCircle2 size={12} /> : <Copy size={12} />}
                                {copiedQaId === msg.id ? '복사됨' : '답변 복사'}
                              </button>
                            </div>
                          )}
                        </div>
                      </div>
                    ))
                  )}

                  {isQaLoading && (
                    <div style={{ display: 'flex', alignItems: 'flex-start', gap: '10px' }}>
                      <div style={{
                        padding: '12px 16px',
                        borderRadius: '14px 14px 14px 2px',
                        background: 'rgba(30, 41, 59, 0.85)',
                        border: '1px solid rgba(148, 163, 184, 0.15)',
                        color: '#c084fc',
                        fontSize: '13px',
                        display: 'flex', alignItems: 'center', gap: '8px'
                      }}>
                        <Loader2 size={16} className="animate-spin" />
                        회의록 및 상세 발언록을 분석하여 답변을 생성하는 중입니다...
                      </div>
                    </div>
                  )}

                  <div ref={qaChatEndRef} />
                </div>

                {/* 에러 메시지 */}
                {qaError && (
                  <div style={{ padding: '10px 14px', background: 'rgba(239,68,68,0.1)', border: '1px solid rgba(239,68,68,0.3)', borderRadius: '8px', color: '#f87171', fontSize: '12.5px', display: 'flex', alignItems: 'center', gap: '6px' }}>
                    <AlertCircle size={15} /> {qaError}
                  </div>
                )}

                {/* 질문 입력창 */}
                <div style={{ display: 'flex', gap: '8px' }}>
                  <input
                    type="text"
                    value={qaInput}
                    onChange={(e) => setQaInput(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter' && !e.shiftKey) {
                        e.preventDefault();
                        handleSendQa();
                      }
                    }}
                    placeholder="회의 내용에 대해 질문하세요 (예: 다음 주까지 누가 어떤 작업을 해야 하나요?)"
                    disabled={isQaLoading}
                    style={{
                      flex: 1, padding: '12px 16px', borderRadius: '10px',
                      background: 'rgba(0,0,0,0.3)', border: '1px solid rgba(139, 92, 246, 0.3)',
                      color: 'var(--text-primary)', fontSize: '13.5px', outline: 'none'
                    }}
                  />
                  <button
                    onClick={() => handleSendQa()}
                    disabled={isQaLoading || !qaInput.trim()}
                    style={{
                      padding: '0 20px', borderRadius: '10px',
                      background: !qaInput.trim() || isQaLoading
                        ? 'rgba(255,255,255,0.05)'
                        : 'linear-gradient(135deg, #8b5cf6, #6366f1)',
                      border: 'none', color: 'white', fontWeight: 700, fontSize: '13.5px',
                      cursor: !qaInput.trim() || isQaLoading ? 'not-allowed' : 'pointer',
                      display: 'flex', alignItems: 'center', gap: '6px',
                      opacity: !qaInput.trim() || isQaLoading ? 0.5 : 1,
                      boxShadow: qaInput.trim() && !isQaLoading ? '0 4px 12px rgba(139, 92, 246, 0.3)' : 'none'
                    }}
                  >
                    <Send size={15} /> 전송
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}
