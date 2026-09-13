// pptValidatorHistoryDB.js
// 표준산출물 검증(PptValidator) 결과를 브라우저 IndexedDB에 저장해 세션이 끝나도 이력을 유지한다.

const DB_NAME = 'PptValidatorHistoryDB';
const DB_VERSION = 1;
const STORE_NAME = 'history';
const MAX_RECORDS = 30; // 과도한 누적 방지 (결과 배열을 통째로 저장하므로 교정이력보다 낮게 설정)

class PptValidatorHistoryDB {
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
            fileNames: Array.isArray(result.fileNames) ? result.fileNames : [],
            stats: Array.isArray(result.stats) ? result.stats : [],
            typoResults: Array.isArray(result.typoResults) ? result.typoResults : [],
            numberingResults: Array.isArray(result.numberingResults) ? result.numberingResults : [],
            altTextResults: Array.isArray(result.altTextResults) ? result.altTextResults : [],
            forbiddenResults: Array.isArray(result.forbiddenResults) ? result.forbiddenResults : [],
            engKoMixedResults: Array.isArray(result.engKoMixedResults) ? result.engKoMixedResults : [],
            duplicateResults: Array.isArray(result.duplicateResults) ? result.duplicateResults : [],
            pageRangeResults: Array.isArray(result.pageRangeResults) ? result.pageRangeResults : [],
            macImageResults: Array.isArray(result.macImageResults) ? result.macImageResults : [],
            piiResults: Array.isArray(result.piiResults) ? result.piiResults : [],
            overflowResults: Array.isArray(result.overflowResults) ? result.overflowResults : [],
            fontResults: Array.isArray(result.fontResults) ? result.fontResults : [],
            parenResults: Array.isArray(result.parenResults) ? result.parenResults : [],
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

export const pptValidatorHistoryDB = new PptValidatorHistoryDB();
