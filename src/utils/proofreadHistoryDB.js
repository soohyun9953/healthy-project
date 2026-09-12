// proofreadHistoryDB.js
// AI 교정교열(TypoValidator) 결과를 브라우저 IndexedDB에 저장해 세션이 끝나도 이력을 유지한다.

const DB_NAME = 'ProofreadHistoryDB';
const DB_VERSION = 1;
const STORE_NAME = 'history';
const MAX_RECORDS = 50; // 과도한 누적 방지

class ProofreadHistoryDB {
    open() {
        return new Promise((resolve, reject) => {
            const request = indexedDB.open(DB_NAME, DB_VERSION);

            request.onupgradeneeded = (event) => {
                const db = event.target.result;
                if (!db.objectStoreNames.contains(STORE_NAME)) {
                    const store = db.createObjectStore(STORE_NAME, { keyPath: 'id', autoIncrement: true });
                    store.createIndex('createdAt', 'createdAt', { unique: false });
                }
            };

            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
        });
    }

    async saveRecord(result) {
        const db = await this.open();
        const record = {
            fileName: result.artifactFileName || '(제목 없음)',
            score: typeof result.score === 'number' ? result.score : null,
            typoCount: Array.isArray(result.typos) ? result.typos.length : 0,
            summary: String(result.summary || ''),
            typos: Array.isArray(result.typos) ? result.typos : [],
            correctedFullText: String(result.correctedFullText || ''),
            createdAt: Date.now()
        };

        await new Promise((resolve, reject) => {
            const tx = db.transaction([STORE_NAME], 'readwrite');
            const store = tx.objectStore(STORE_NAME);
            store.add(record);
            tx.oncomplete = () => resolve();
            tx.onerror = () => reject(tx.error);
        });

        await this._pruneOldRecords(db);
    }

    async _pruneOldRecords(db) {
        return new Promise((resolve, reject) => {
            const tx = db.transaction([STORE_NAME], 'readwrite');
            const store = tx.objectStore(STORE_NAME);
            const index = store.index('createdAt');
            const request = index.openCursor(null, 'prev'); // 최신순
            let count = 0;

            request.onsuccess = (event) => {
                const cursor = event.target.result;
                if (!cursor) return;
                count++;
                if (count > MAX_RECORDS) {
                    cursor.delete();
                }
                cursor.continue();
            };
            tx.oncomplete = () => resolve();
            tx.onerror = () => reject(tx.error);
        });
    }

    async getAll() {
        const db = await this.open();
        return new Promise((resolve, reject) => {
            const tx = db.transaction([STORE_NAME], 'readonly');
            const store = tx.objectStore(STORE_NAME);
            const request = store.getAll();
            request.onsuccess = () => {
                const records = request.result || [];
                records.sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0));
                resolve(records);
            };
            request.onerror = () => reject(request.error);
        });
    }

    async deleteRecord(id) {
        const db = await this.open();
        return new Promise((resolve, reject) => {
            const tx = db.transaction([STORE_NAME], 'readwrite');
            tx.objectStore(STORE_NAME).delete(id);
            tx.oncomplete = () => resolve(true);
            tx.onerror = () => reject(tx.error);
        });
    }

    async clearAll() {
        const db = await this.open();
        return new Promise((resolve, reject) => {
            const tx = db.transaction([STORE_NAME], 'readwrite');
            tx.objectStore(STORE_NAME).clear();
            tx.oncomplete = () => resolve(true);
            tx.onerror = () => reject(tx.error);
        });
    }
}

export const proofreadHistoryDB = new ProofreadHistoryDB();
