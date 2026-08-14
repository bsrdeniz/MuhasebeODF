import { encryptPayload, decryptPayload, blobToBase64, base64ToBlob } from './utils/crypto';

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
  rows: DocumentRow[],
  key: CryptoKey | null
): Promise<void> {
  const db = await initDB();

  let encryptedDoc: any;
  let encryptedRows: any[];

  if (key) {
    // Encrypt document details
    const fileDataBase64 = await blobToBase64(doc.fileData);
    const docPayload = {
      name: doc.name,
      fileName: doc.fileName,
      fileSize: doc.fileSize,
      fileType: doc.fileType,
      docType: doc.docType,
      date: doc.date,
      category: doc.category,
      description: doc.description,
      fileDataBase64,
      uploadedAt: doc.uploadedAt
    };
    const encryptedDocPayload = await encryptPayload(docPayload, key);
    encryptedDoc = {
      id: doc.id,
      encryptedPayload: encryptedDocPayload,
      uploadedAt: doc.uploadedAt
    };

    // Encrypt row data
    encryptedRows = await Promise.all(rows.map(async (row) => {
      const encryptedRowPayload = await encryptPayload(row.data, key);
      return {
        id: row.id,
        docId: row.docId,
        rowNumber: row.rowNumber,
        encryptedPayload: encryptedRowPayload
      };
    }));
  } else {
    // Non-encrypted fallback
    encryptedDoc = doc;
    encryptedRows = rows;
  }

  return new Promise((resolve, reject) => {
    const transaction = db.transaction(['documents', 'document_rows'], 'readwrite');
    
    transaction.onerror = () => reject(transaction.error);
    transaction.oncomplete = () => resolve();

    const docStore = transaction.objectStore('documents');
    docStore.put(encryptedDoc);

    const rowStore = transaction.objectStore('document_rows');
    for (const erow of encryptedRows) {
      rowStore.put(erow);
    }
  });
}

export async function getDocuments(key: CryptoKey | null): Promise<DocumentItem[]> {
  const db = await initDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction('documents', 'readonly');
    const store = transaction.objectStore('documents');
    const request = store.getAll();

    request.onerror = () => reject(request.error);
    request.onsuccess = async () => {
      const results = request.result || [];
      try {
        const decryptedDocs = await Promise.all(results.map(async (edoc: any) => {
          if (!edoc.encryptedPayload) {
            // Returns legacy plaintext document if any exists
            return edoc as DocumentItem;
          }
          if (!key) {
            throw new Error("Veritabanı şifrelenmiş fakat anahtar verilmedi!");
          }
          const payload = await decryptPayload(edoc.encryptedPayload, key);
          const fileData = base64ToBlob(payload.fileDataBase64, payload.fileType === 'PDF' ? 'application/pdf' : 'image/png');
          return {
            id: edoc.id,
            name: payload.name,
            fileName: payload.fileName,
            fileSize: payload.fileSize,
            fileType: payload.fileType,
            docType: payload.docType,
            date: payload.date,
            category: payload.category,
            description: payload.description,
            fileData,
            uploadedAt: payload.uploadedAt
          };
        }));
        resolve(decryptedDocs);
      } catch (err) {
        reject(err);
      }
    };
  });
}

export async function getDocumentRows(docId: string, key: CryptoKey | null): Promise<DocumentRow[]> {
  const db = await initDB();
  return new Promise((resolve, reject) => {
    const transaction = db.transaction('document_rows', 'readonly');
    const store = transaction.objectStore('document_rows');
    const index = store.index('docId');
    const request = index.getAll(IDBKeyRange.only(docId));

    request.onerror = () => reject(request.error);
    request.onsuccess = async () => {
      const results = request.result || [];
      try {
        const decryptedRows = await Promise.all(results.map(async (erow: any) => {
          if (!erow.encryptedPayload) {
            // Returns legacy plaintext row if any exists
            return erow as DocumentRow;
          }
          if (!key) {
            throw new Error("Veritabanı şifrelenmiş fakat anahtar verilmedi!");
          }
          const data = await decryptPayload(erow.encryptedPayload, key);
          return {
            id: erow.id,
            docId: erow.docId,
            rowNumber: erow.rowNumber,
            data
          };
        }));
        resolve(decryptedRows);
      } catch (err) {
        reject(err);
      }
    };
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
