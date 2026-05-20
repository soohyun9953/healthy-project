/**
 * RAG Knowledge Base Service
 * Handles loading and searching of indexed ISMP documents.
 */

let ragDataCache = null;

/**
 * Load RAG data from the public folder
 */
export const loadRagData = async (forceRefresh = false) => {
    if (ragDataCache && !forceRefresh) return ragDataCache;
    
    try {
        const response = await fetch('/rag_data.json');
        if (!response.ok) throw new Error('RAG 데이터를 불러오는데 실패했습니다.');
        const data = await response.json();
        ragDataCache = data;
        return data;
    } catch (error) {
        console.error('Error loading RAG data:', error);
        return [];
    }
};

/**
 * Search RAG data for relevant snippets
 * @param {string} query Search query
 * @param {number} limit Number of results to return
 */
export const searchRag = async (query, limit = 5) => {
    const data = await loadRagData();
    if (!data || data.length === 0) return [];
    
    if (!query || query.trim() === '') return data.slice(0, limit);
    
    const terms = query.toLowerCase().split(/\s+/).filter(t => t.length > 0);
    
    const results = data.map(doc => {
        let score = 0;
        const title = doc.title.toLowerCase();
        const content = doc.content.toLowerCase();
        
        terms.forEach(term => {
            if (title.includes(term)) score += 10;
            if (content.includes(term)) {
                // Count occurrences for better scoring
                const regex = new RegExp(term.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
                const matches = content.match(regex);
                score += (matches ? matches.length : 0);
            }
        });
        
        return { ...doc, score };
    })
    .filter(res => res.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, limit);
    
    return results;
};

/**
 * Get context string for LLM analysis
 * @param {string} query Search query or document summary
 */
export const getRagContext = async (query) => {
    const results = await searchRag(query, 3);
    if (results.length === 0) return "";
    
    let context = "\n\n--- [참고: 관련 ISMP 산출물 지식베이스] ---\n";
    results.forEach((res, idx) => {
        context += `\n[관련 문서 ${idx + 1}: ${res.title}]\n`;
        // Limit context per document to avoid token overflow
        context += res.content.substring(0, 2000) + (res.content.length > 2000 ? "..." : "") + "\n";
    });
    
    return context;
};
