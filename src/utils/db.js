const DB_NAME = 'RFP_Ref_DB';
const DB_VERSION = 1;
const STORE_NAME = 'documents';

class ReferenceDB {
    open() {
        return new Promise((resolve, reject) => {
            const request = indexedDB.open(DB_NAME, DB_VERSION);
            
            request.onupgradeneeded = (event) => {
                const db = event.target.result;
                if (!db.objectStoreNames.contains(STORE_NAME)) {
                    db.createObjectStore(STORE_NAME, { keyPath: 'id' });
                }
            };
            
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
        });
    }

    async getAllDocs() {
        const db = await this.open();
        return new Promise((resolve, reject) => {
            const transaction = db.transaction([STORE_NAME], 'readonly');
            const store = transaction.objectStore(STORE_NAME);
            const request = store.getAll();
            
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
        });
    }

    async getDocsMetadata() {
        const db = await this.open();
        return new Promise((resolve, reject) => {
            const transaction = db.transaction([STORE_NAME], 'readonly');
            const store = transaction.objectStore(STORE_NAME);
            const request = store.openCursor();
            const results = [];
            
            request.onsuccess = (event) => {
                const cursor = event.target.result;
                if (cursor) {
                    const { id, title, filename, type, ext, size, createdAt } = cursor.value;
                    results.push({ id, title, filename, type, ext, size, createdAt });
                    cursor.continue();
                } else {
                    resolve(results);
                }
            };
            request.onerror = () => reject(request.error);
        });
    }

    async getDocById(id) {
        const db = await this.open();
        return new Promise((resolve, reject) => {
            const transaction = db.transaction([STORE_NAME], 'readonly');
            const store = transaction.objectStore(STORE_NAME);
            const request = store.get(id);
            
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
        });
    }

    async saveDoc(doc) {
        const db = await this.open();
        return new Promise((resolve, reject) => {
            const transaction = db.transaction([STORE_NAME], 'readwrite');
            const store = transaction.objectStore(STORE_NAME);
            const request = store.put(doc);
            
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
        });
    }

    async deleteDoc(id) {
        const db = await this.open();
        return new Promise((resolve, reject) => {
            try {
                const transaction = db.transaction([STORE_NAME], 'readwrite');
                const store = transaction.objectStore(STORE_NAME);
                
                // IndexedDB ID는 타입에 민감하므로 숫자 변환 시도
                const numericId = Number(id);
                const finalId = !isNaN(numericId) ? numericId : id;
                
                const request = store.delete(finalId);

                transaction.oncomplete = () => {
                    console.log('IndexedDB delete transaction complete for:', finalId);
                    resolve(true);
                };
                transaction.onerror = (event) => {
                    console.error('IndexedDB transaction error:', event.target.error);
                    reject(event.target.error);
                };
                request.onerror = (event) => {
                    console.error('IndexedDB request error:', event.target.error);
                    reject(event.target.error);
                };
            } catch (err) {
                console.error('IndexedDB delete catch error:', err);
                reject(err);
            }
        });
    }


    async clearAllDocs() {
        const db = await this.open();
        return new Promise((resolve, reject) => {
            try {
                const transaction = db.transaction([STORE_NAME], 'readwrite');
                const store = transaction.objectStore(STORE_NAME);
                store.clear();

                transaction.oncomplete = () => {
                    console.log('IndexedDB clear all transaction complete');
                    resolve(true);
                };
                transaction.onerror = (event) => reject(event.target.error);
            } catch (err) {
                reject(err);
            }
        });
    }

}

export const refDB = new ReferenceDB();
