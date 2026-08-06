import * as XLSX from 'xlsx';

export interface ParsedSheet {
  name: string;
  headers: string[];
  rows: Record<string, any>[];
}

export interface AutoMapping {
  dateKey: string;
  descriptionKey: string;
  debitKey: string;
  creditKey: string;
  amountKey: string;
  categoryKey: string;
}

// Common keywords for auto-detecting column mappings
const KEYWORDS = {
  date: ['tarih', 'date', 'gün', 'gun', 'tarihi', 'zaman', 'time'],
  description: ['açıklama', 'aciklama', 'description', 'tanım', 'tanim', 'detay', 'detayi', 'işlem', 'islem', 'açiklama'],
  debit: ['borç', 'borc', 'debit', 'gider', 'ödeme', 'odeme', 'çıkan', 'cikan', 'çekilen', 'cekilen', 'maliyet'],
  credit: ['alacak', 'credit', 'gelir', 'tahsilat', 'giren', 'yatırılan', 'yatirilan'],
  amount: ['tutar', 'amount', 'toplam', 'net', 'bakiye', 'değer', 'deger'],
  category: ['kategori', 'category', 'tür', 'tur', 'grup', 'sınıf', 'sinif', 'departman']
};

function matchKeyword(header: string, list: string[]): boolean {
  const normalized = header.toLowerCase().replace(/[^a-z0-9ıışşğğüüööçç]/g, '');
  return list.some(k => normalized.includes(k) || k.includes(normalized));
}

export function autoDetectMapping(headers: string[]): AutoMapping {
  const mapping: AutoMapping = {
    dateKey: '',
    descriptionKey: '',
    debitKey: '',
    creditKey: '',
    amountKey: '',
    categoryKey: ''
  };

  headers.forEach(header => {
    const lower = header.toLowerCase();
    if (!mapping.dateKey && matchKeyword(lower, KEYWORDS.date)) {
      mapping.dateKey = header;
    } else if (!mapping.descriptionKey && matchKeyword(lower, KEYWORDS.description)) {
      mapping.descriptionKey = header;
    } else if (!mapping.debitKey && matchKeyword(lower, KEYWORDS.debit)) {
      mapping.debitKey = header;
    } else if (!mapping.creditKey && matchKeyword(lower, KEYWORDS.credit)) {
      mapping.creditKey = header;
    } else if (!mapping.amountKey && matchKeyword(lower, KEYWORDS.amount)) {
      mapping.amountKey = header;
    } else if (!mapping.categoryKey && matchKeyword(lower, KEYWORDS.category)) {
      mapping.categoryKey = header;
    }
  });

  // Fallbacks if not detected
  if (!mapping.descriptionKey && headers.length > 0) {
    mapping.descriptionKey = headers[0];
  }

  return mapping;
}

export function parseOdfFile(file: File): Promise<ParsedSheet[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        
        const sheets: ParsedSheet[] = [];

        workbook.SheetNames.forEach(sheetName => {
          const sheet = workbook.Sheets[sheetName];
          // convert to JSON array of objects
          const rows = XLSX.utils.sheet_to_json<Record<string, any>>(sheet, { defval: '' });
          
          if (rows.length > 0) {
            // Get all headers from the sheet
            const headers = Array.from(
              new Set(rows.flatMap(row => Object.keys(row)))
            );

            sheets.push({
              name: sheetName,
              headers,
              rows
            });
          }
        });

        resolve(sheets);
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}
