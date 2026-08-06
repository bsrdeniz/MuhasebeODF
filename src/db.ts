export interface DocumentItem {
  id: string;
  name: string;
  fileName: string;
  fileSize: number;
  fileType: string;
  docType: string; // 'Maaş Bordrosu' | 'TOBB Belgesi' | 'Fatura' | 'Sözleşme' | 'Diğer'
  date: string;
  category: string;
  description: string;
  fileData: Blob;
  uploadedAt: string;
}

export interface DocumentRow {
  id: string;
  docId: string;
  rowNumber: number;
  data: Record<string, any>;
}

const DB_NAME = 'MuhasebeBelgeDB';
const DB_VERSION = 2;

export function initDB(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onerror = () => reject(request.error);
    request.onsuccess = () => resolve(request.result);

    request.onupgradeneeded = () => {
      const db = request.result;
      
      if (!db.objectStoreNames.contains('documents')) {
        db.createObjectStore('documents', { keyPath: 'id' });
      }
      
      if (!db.objectStoreNames.contains('document_rows')) {
        const rowStore = db.createObjectStore('document_rows', { keyPath: 'id' });
        rowStore.createIndex('docId', 'docId', { unique: false });
      }
    };
  });
}

export async function saveDocumentAndRows(
  doc: DocumentItem,
  rows: DocumentRow[]
): Promise<void> {
  const db = await initDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(['documents', 'document_rows'], 'readwrite');
    
    transaction.onerror = () => reject(transaction.error);
    transaction.oncomplete = () => resolve();

    const docStore = transaction.objectStore('documents');
    docStore.put(doc);

    const rowStore = transaction.objectStore('document_rows');
    for (const row of rows) {
      rowStore.put(row);
    }
  });
}

export async function getDocuments(): Promise<DocumentItem[]> {
  const db = await initDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction('documents', 'readonly');
    const store = transaction.objectStore('documents');
    const request = store.getAll();

    request.onerror = () => reject(request.error);
    request.onsuccess = () => resolve(request.result || []);
  });
}

export async function getDocumentRows(docId: string): Promise<DocumentRow[]> {
  const db = await initDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction('document_rows', 'readonly');
    const store = transaction.objectStore('document_rows');
    const index = store.index('docId');
    const request = index.getAll(IDBKeyRange.only(docId));

    request.onerror = () => reject(request.error);
    request.onsuccess = () => resolve(request.result || []);
  });
}

export async function deleteDocument(id: string): Promise<void> {
  const db = await initDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(['documents', 'document_rows'], 'readwrite');
    
    transaction.onerror = () => reject(transaction.error);
    transaction.oncomplete = () => resolve();

    const docStore = transaction.objectStore('documents');
    docStore.delete(id);

    // Delete associated rows
    const rowStore = transaction.objectStore('document_rows');
    const index = rowStore.index('docId');
    const request = index.openCursor(IDBKeyRange.only(id));

    request.onsuccess = (event) => {
      const cursor = (event.target as IDBRequest<IDBCursorWithValue | null>).result;
      if (cursor) {
        cursor.delete();
        cursor.continue();
      }
    };
  });
}
