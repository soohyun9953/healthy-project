/**
 * clientRagIndexer.js
 * 브라우저 내장 IndexedDB 기반 가이드라인 문서 인덱서 및 BM25 하이브리드 검색기
 * Vercel 배포 환경 및 로컬 환경 모두에서 서버/파이썬 의존성 없이 100% 브라우저 내에서 작동
 */
import { extractStructuredFileContent, getFileExtension } from './fileExtractor.js';

const DB_NAME = 'healthy_project_rag_db';
const DB_VERSION = 1;
const STORE_DOCS = 'rag_documents';
const STORE_CHUNKS = 'rag_chunks';

// ── 1. IndexedDB 초기화 및 헬퍼 ────────────────────────────────────
function openRagDB() {
    return new Promise((resolve, reject) => {
        if (typeof window === 'undefined' || !window.indexedDB) {
            reject(new Error('IndexedDB를 지원하지 않는 브라우저 환경입니다.'));
            return;
        }

        const request = window.indexedDB.open(DB_NAME, DB_VERSION);

        request.onupgradeneeded = (event) => {
            const db = event.target.result;
            // 1) 문서 메타데이터 저장소
            if (!db.objectStoreNames.contains(STORE_DOCS)) {
                db.createObjectStore(STORE_DOCS, { keyPath: 'id' });
            }
            // 2) 청크 저장소
            if (!db.objectStoreNames.contains(STORE_CHUNKS)) {
                const chunkStore = db.createObjectStore(STORE_CHUNKS, { keyPath: 'id' });
                chunkStore.createIndex('docId', 'docId', { unique: false });
                chunkStore.createIndex('filename', 'filename', { unique: false });
            }
        };

        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
    });
}

// ── 2. 스마트 텍스트 청킹 (800자 단위, 150자 오버랩, 문장/단락 경계 인식) ──────
export function chunkDocumentText(text, docId, filename, options = {}) {
    const chunkSize = options.chunkSize || 800;
    const overlap = options.overlap || 150;
    const chunks = [];
    if (!text || text.trim().length === 0) return chunks;

    const cleanText = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    let start = 0;
    let chunkSeq = 1;

    while (start < cleanText.length) {
        let end = start + chunkSize;

        if (end >= cleanText.length) {
            end = cleanText.length;
        } else {
            // 자연스러운 문장/단락 경계 탐색 (오버랩 범위 내)
            const minSplit = Math.max(start + chunkSize - overlap, start + Math.floor(chunkSize * 0.5));
            const searchWindow = cleanText.substring(minSplit, Math.min(cleanText.length, end + 80));

            // 1순위: 단락 구분 (\n\n)
            const pBreak = searchWindow.lastIndexOf('\n\n');
            if (pBreak !== -1) {
                end = minSplit + pBreak + 2;
            } else {
                // 2순위: 줄바꿈 (\n)
                const lineBreak = searchWindow.lastIndexOf('\n');
                if (lineBreak !== -1) {
                    end = minSplit + lineBreak + 1;
                } else {
                    // 3순위: 문장 마침표 (. , ? , !)
                    const match = searchWindow.match(/([.?!]\s+)(?=[^\s])/);
                    if (match && match.index !== undefined) {
                        end = minSplit + match.index + match[0].length;
                    }
                }
            }
        }

        const chunkText = cleanText.substring(start, end).trim();
        if (chunkText.length >= 25) {
            chunks.push({
                id: `${docId}_chk_${chunkSeq}`,
                docId,
                filename,
                chunkIndex: chunkSeq,
                text: chunkText,
                length: chunkText.length,
                charStart: start,
                charEnd: end,
                createdAt: new Date().toISOString()
            });
            chunkSeq++;
        }

        if (end >= cleanText.length) break;
        // 다음 시작점: 오버랩 반영 (반드시 전진 보장)
        const nextStart = Math.max(end - overlap, start + 1);
        start = nextStart;
    }

    return chunks;
}

// ── 3. 단일 파일 인덱싱 (PDF/HWPX/DOCX/XLSX 브라우저 직접 파싱 및 IndexedDB 저장) ──
export async function indexDocumentFile(file, onProgress = () => {}) {
    onProgress({ percent: 20, message: `[${file.name}] 텍스트 추출 중...` });
    const fileContent = await extractStructuredFileContent(file);
    const text = fileContent.text || '';

    if (!text || text.trim().length === 0) {
        throw new Error(`파일 [${file.name}]에서 추출된 텍스트가 없습니다.`);
    }

    onProgress({ percent: 50, message: `[${file.name}] 800자 스마트 청킹 중...` });
    const docId = `doc_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`;
    const chunks = chunkDocumentText(text, docId, file.name, { chunkSize: 800, overlap: 150 });

    const docRecord = {
        id: docId,
        filename: file.name,
        fileSize: file.size,
        ext: getFileExtension(file.name),
        uploadedAt: new Date().toISOString(),
        chunkCount: chunks.length,
        totalLength: text.length,
        pages: fileContent.pages || 1
    };

    onProgress({ percent: 80, message: `[${file.name}] IndexedDB에 ${chunks.length}개 청크 저장 중...` });
    const db = await openRagDB();

    await new Promise((resolve, reject) => {
        const tx = db.transaction([STORE_DOCS, STORE_CHUNKS], 'readwrite');
        const docStore = tx.objectStore(STORE_DOCS);
        const chunkStore = tx.objectStore(STORE_CHUNKS);

        docStore.put(docRecord);
        chunks.forEach(chunk => chunkStore.put(chunk));

        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
    });

    onProgress({ percent: 100, message: `[${file.name}] 인덱싱 완료 (${chunks.length}개 청크)` });

    return {
        docId,
        filename: file.name,
        chunkCount: chunks.length,
        totalLength: text.length
    };
}

// ── 4. IndexedDB 문서 및 청크 목록 조회 / 삭제 ─────────────────────
export async function getAllIndexedDocs() {
    try {
        const db = await openRagDB();
        return new Promise((resolve, reject) => {
            const tx = db.transaction([STORE_DOCS], 'readonly');
            const store = tx.objectStore(STORE_DOCS);
            const req = store.getAll();
            req.onsuccess = () => resolve(req.result || []);
            req.onerror = () => reject(req.error);
        });
    } catch (e) {
        console.warn('IndexedDB 조회 실패:', e);
        return [];
    }
}

export async function deleteIndexedDoc(docId) {
    const db = await openRagDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction([STORE_DOCS, STORE_CHUNKS], 'readwrite');
        const docStore = tx.objectStore(STORE_DOCS);
        const chunkStore = tx.objectStore(STORE_CHUNKS);

        docStore.delete(docId);

        const index = chunkStore.index('docId');
        const req = index.openCursor(IDBKeyRange.only(docId));
        req.onsuccess = (e) => {
            const cursor = e.target.result;
            if (cursor) {
                cursor.delete();
                cursor.continue();
            }
        };

        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
    });
}

export async function clearAllIndexedDocs() {
    const db = await openRagDB();
    return new Promise((resolve, reject) => {
        const tx = db.transaction([STORE_DOCS, STORE_CHUNKS], 'readwrite');
        tx.objectStore(STORE_DOCS).clear();
        tx.objectStore(STORE_CHUNKS).clear();
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
    });
}

// ── 5. 기존 public/rag_data.json 사전 적재 캐시 ─────────────────────
let preloadedChunksCache = null;

export async function loadPreloadedRagChunks() {
    if (preloadedChunksCache) return preloadedChunksCache;

    try {
        const res = await fetch('/rag_data.json');
        if (!res.ok) return [];
        const rawDocs = await res.json();
        if (!Array.isArray(rawDocs)) return [];

        const chunks = [];
        rawDocs.forEach((doc, dIdx) => {
            const filename = doc.filename || doc.title || `사전문서_${dIdx + 1}.pdf`;
            const docId = `preloaded_${doc.id || dIdx + 1}`;

            // 페이지별 내용이 있으면 페이지 단위 취합
            if (Array.isArray(doc.pages) && doc.pages.length > 0) {
                doc.pages.forEach((p, pIdx) => {
                    const text = (p.content || '').trim();
                    if (text.length >= 25) {
                        chunks.push({
                            id: `${docId}_p${p.page || pIdx + 1}`,
                            docId,
                            filename,
                            chunkIndex: p.page || pIdx + 1,
                            text,
                            length: text.length,
                            isPreloaded: true
                        });
                    }
                });
            } else if (doc.content) {
                // 단일 텍스트면 스마트 청킹
                const docChunks = chunkDocumentText(doc.content, docId, filename);
                chunks.push(...docChunks.map(c => ({ ...c, isPreloaded: true })));
            }
        });

        preloadedChunksCache = chunks;
        return chunks;
    } catch (e) {
        console.warn('사전 rag_data.json 로드 건너뜀:', e);
        return [];
    }
}

// ── 6. BM25 + 키워드 하이브리드 검색기 ──────────────────────────────
/**
 * 한국어 단어 및 2-gram 토큰화
 */
function tokenizeKorean(text) {
    if (!text) return [];
    const clean = text.toLowerCase().replace(/[^a-zA-Z0-9가-힣\s]/g, ' ');
    const words = clean.split(/\s+/).filter(w => w.length > 1);

    const tokens = new Set(words);

    // 복합명사 인식을 위한 2-gram 추가
    words.forEach(word => {
        if (/^[가-힣]+$/.test(word) && word.length >= 3 && word.length <= 8) {
            for (let i = 0; i < word.length - 1; i++) {
                tokens.add(word.substring(i, i + 2));
            }
        }
    });

    return Array.from(tokens);
}

/**
 * BM25 알고리즘 기반 상위 K개 청크 검색
 */
export async function searchRagHybrid(query, { topK = 5, docFilter = null } = {}) {
    if (!query || query.trim().length === 0) return [];

    // 1) IndexedDB 청크 로드
    let userChunks = [];
    try {
        const db = await openRagDB();
        userChunks = await new Promise((resolve) => {
            const tx = db.transaction([STORE_CHUNKS], 'readonly');
            const req = tx.objectStore(STORE_CHUNKS).getAll();
            req.onsuccess = () => resolve(req.result || []);
            req.onerror = () => resolve([]);
        });
    } catch (_) {}

    // 2) 사전 적재된 rag_data.json 청크 로드
    const preloadedChunks = await loadPreloadedRagChunks();

    // 3) 전체 청크 통합
    let allChunks = [...userChunks, ...preloadedChunks];

    if (docFilter) {
        allChunks = allChunks.filter(c => c.filename === docFilter || c.docId === docFilter);
    }

    if (allChunks.length === 0) return [];

    // 4) BM25 파라미터 및 인덱스 계산
    const queryTokens = tokenizeKorean(query);
    if (queryTokens.length === 0) return [];

    const N = allChunks.length;
    let totalLength = 0;
    const docTermFreqs = [];
    const docFreq = {};

    queryTokens.forEach(t => { docFreq[t] = 0; });

    allChunks.forEach((chunk, idx) => {
        const text = chunk.text || '';
        totalLength += text.length;

        const termMap = {};
        const lowerText = text.toLowerCase();

        queryTokens.forEach(term => {
            if (lowerText.includes(term)) {
                // 출현 빈도 카운트
                const count = (lowerText.match(new RegExp(term, 'g')) || []).length;
                termMap[term] = count;
                docFreq[term] = (docFreq[term] || 0) + 1;
            }
        });

        docTermFreqs[idx] = termMap;
    });

    const avgdl = totalLength / (N || 1);
    const k1 = 1.2;
    const b = 0.75;

    // 5) 각 청크별 점수 계산
    const scoredChunks = allChunks.map((chunk, idx) => {
        let score = 0;
        const termMap = docTermFreqs[idx] || {};
        const docLen = (chunk.text || '').length;

        queryTokens.forEach(term => {
            const freq = termMap[term] || 0;
            if (freq > 0) {
                const df = docFreq[term] || 1;
                // IDF 계산
                const idf = Math.log(1 + (N - df + 0.5) / (df + 0.5));
                // BM25 점수 가산
                const tfPart = (freq * (k1 + 1)) / (freq + k1 * (1 - b + b * (docLen / (avgdl || 1))));
                score += idf * tfPart;
            }
        });

        // 보너스 가중치: 파일명 매칭 또는 검색어 전체 구문 매칭
        const lowerText = (chunk.text || '').toLowerCase();
        const lowerFilename = (chunk.filename || '').toLowerCase();
        const trimmedQuery = query.toLowerCase().trim();

        if (trimmedQuery.length >= 4 && lowerText.includes(trimmedQuery)) {
            score += 4.0; // 정확한 구문 일치 보너스
        }
        if (queryTokens.some(t => lowerFilename.includes(t))) {
            score += 1.5; // 문서 제목/파일명 일치 보너스
        }
        if (!chunk.isPreloaded) {
            score += 0.5; // 사용자가 직접 올린 최신 가이드라인 가산점
        }

        return {
            id: chunk.id,
            docId: chunk.docId,
            filename: chunk.filename,
            chunkIndex: chunk.chunkIndex,
            text: chunk.text,
            length: chunk.length,
            isPreloaded: !!chunk.isPreloaded,
            score: Math.round(score * 100) / 100
        };
    });

    // 6) 정렬 후 Top-K 반환
    const filtered = scoredChunks.filter(c => c.score > 0.1);
    filtered.sort((a, b) => b.score - a.score);

    return filtered.slice(0, topK);
}

// ── 7. 프롬프트 주입용 포맷터 ──────────────────────────────────────
export function formatTopChunksForPrompt(topChunks) {
    if (!topChunks || topChunks.length === 0) return '';

    const lines = [
        `[지침 기준문서 / RAG 지식베이스 관련 발췌 (상위 ${topChunks.length}건)]:`
    ];

    topChunks.forEach((chk, i) => {
        lines.push(`\n▶ [참조 #${i + 1} | 문서: ${chk.filename} (청크 #${chk.chunkIndex})]`);
        lines.push(chk.text);
    });

    lines.push('\n------------------------------------------------------------');
    return lines.join('\n');
}
