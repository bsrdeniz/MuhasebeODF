import * as XLSX from 'xlsx';

export interface ExtractedMetadata {
  name: string;
  docType: string;
  date: string;
  category: string;
  description: string;
}

const TURKISH_MONTHS: Record<string, string> = {
  ocak: '01', subat: '02', şubat: '02', mart: '03', nisan: '04', mayis: '05', mayıs: '05',
  haziran: '06', temmuz: '07', agustos: '08', ağustos: '08', eylul: '09', eylül: '09',
  ekim: '10', kasim: '11', kasım: '11', aralik: '12', aralık: '12'
};

const DATE_PATTERNS = [
  /\b(0?[1-9]|[12]\d|3[01])[./-](0?[1-9]|1[0-2])[./-](20\d{2})\b/, // DD.MM.YYYY
  /\b(20\d{2})[./-](0?[1-9]|1[0-2])[./-](0?[1-9]|[12]\d|3[01])\b/, // YYYY.MM.DD
];

function extractDate(text: string): string {
  for (const pattern of DATE_PATTERNS) {
    const match = text.match(pattern);
    if (match) {
      let day = '';
      let month = '';
      let year = '';

      if (pattern.source.startsWith('\\b(0?')) {
        day = match[1].padStart(2, '0');
        month = match[2].padStart(2, '0');
        year = match[3];
      } else {
        year = match[1];
        month = match[2].padStart(2, '0');
        day = match[3].padStart(2, '0');
      }
      return `${year}-${month}-${day}`;
    }
  }

  const lowerText = text.toLowerCase();
  for (const [monthName, monthNum] of Object.entries(TURKISH_MONTHS)) {
    const monthIndex = lowerText.indexOf(monthName);
    if (monthIndex !== -1) {
      const textAfterMonth = lowerText.substring(monthIndex, monthIndex + 25);
      const yearMatch = textAfterMonth.match(/\b(20\d{2})\b/);
      if (yearMatch) {
        return `${yearMatch[1]}-${monthNum}-01`;
      }
    }
  }

  return new Date().toISOString().split('T')[0];
}

function analyzeText(text: string, fileName: string): ExtractedMetadata {
  const combinedText = `${fileName} ${text}`.toLowerCase();
  
  let docType = 'TOBB Belgesi';
  let category = 'TOBB Evrak';
  let name = '';
  let description = '';

  if (
    combinedText.includes('bordro') || 
    combinedText.includes('maaş') || 
    combinedText.includes('maas') || 
    combinedText.includes('kazanç') || 
    combinedText.includes('net ödenen') || 
    combinedText.includes('sgk')
  ) {
    docType = 'Maaş Bordrosu';
    category = 'Personel';
  } else if (
    combinedText.includes('fatura') || 
    combinedText.includes('invoice') || 
    combinedText.includes('kdv') || 
    combinedText.includes('tutar') || 
    combinedText.includes('fiş') || 
    combinedText.includes('fis') || 
    combinedText.includes('ödeme')
  ) {
    docType = 'Fatura';
    category = 'Finans';
  } else if (
    combinedText.includes('sözleşme') || 
    combinedText.includes('sozlesme') || 
    combinedText.includes('anlaşma') || 
    combinedText.includes('protokol') || 
    combinedText.includes('maddesi')
  ) {
    docType = 'Sözleşme';
    category = 'Hukuk';
  } else if (
    combinedText.includes('tobb') || 
    combinedText.includes('odalar') || 
    combinedText.includes('borsa') || 
    combinedText.includes('sicil') || 
    combinedText.includes('faaliyet belgesi')
  ) {
    docType = 'TOBB Belgesi';
    category = 'TOBB Evrak';
  } else {
    docType = 'Diğer';
    category = 'Genel';
  }

  const date = extractDate(combinedText);
  const yearMonth = date.substring(0, 7);
  const formattedDate = new Date(date).toLocaleDateString('tr-TR', { month: 'long', year: 'numeric' });

  const cleanFileName = fileName.substring(0, fileName.lastIndexOf('.')) || fileName;
  
  if (docType === 'Maaş Bordrosu') {
    name = `${formattedDate} Maaş Bordrosu`;
    description = `Sistem tarafından otomatik çözümlenen ${formattedDate} dönemine ait personel maaş bordrosu belgesi.`;
  } else if (docType === 'TOBB Belgesi') {
    if (combinedText.includes('faaliyet')) {
      name = `TOBB Faaliyet Belgesi (${yearMonth.split('-')[0]})`;
      description = `TOBB Faaliyet Belgesi evrakı.`;
    } else {
      name = `TOBB Resmi Yazı - ${cleanFileName}`;
      description = `TOBB'dan gelen resmi evrak / yazışma.`;
    }
  } else if (docType === 'Fatura') {
    name = `Fatura - ${cleanFileName}`;
    description = `Arşivlenen fatura belgesi.`;
  } else if (docType === 'Sözleşme') {
    name = `Sözleşme - ${cleanFileName}`;
    description = `Taraflar arasında imzalanan sözleşme evrakı.`;
  } else {
    name = cleanFileName;
    description = `Arşivlenen genel nitelikli belge.`;
  }

  return {
    name,
    docType,
    date,
    category,
    description
  };
}

// Ensures PDFJS script is loaded inside DOM
async function ensurePdfJsLoaded(): Promise<any> {
  if ((window as any).pdfjsLib) {
    return (window as any).pdfjsLib;
  }
  
  return new Promise((resolve) => {
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
    document.head.appendChild(script);

    script.onload = () => {
      const pdfjsLib = (window as any).pdfjsLib;
      pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
      resolve(pdfjsLib);
    };
  });
}

// Extracts plain text from PDF
async function extractPdfText(file: File): Promise<string> {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const arrayBuffer = e.target?.result as ArrayBuffer;
        const pdfjsLib = await ensurePdfJsLoaded();

        const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
        const pdf = await loadingTask.promise;
        
        let text = '';
        const maxPages = pdf.numPages;
        
        for (let i = 1; i <= maxPages; i++) {
          const page = await pdf.getPage(i);
          const textContent = await page.getTextContent();
          const pageText = textContent.items.map((item: any) => item.str).join(' ');
          text += pageText + ' ';
        }
        
        resolve(text);
      } catch (err) {
        console.error('PDF metin çıkarma hatası:', err);
        resolve('');
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

// Extracts spreadsheet text
function extractSpreadsheetText(file: File): Promise<string> {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        let text = '';
        workbook.SheetNames.slice(0, 3).forEach(sheetName => {
          const sheet = workbook.Sheets[sheetName];
          for (const key in sheet) {
            if (key[0] !== '!' && sheet[key] && sheet[key].v !== undefined) {
              text += String(sheet[key].v) + ' ';
            }
          }
        });
        
        resolve(text);
      } catch (err) {
        console.error('Hesap tablosu metin çıkarma hatası:', err);
        resolve('');
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

export async function autoExtractMetadata(file: File): Promise<ExtractedMetadata> {
  const fileExt = file.name.split('.').pop()?.toLowerCase();
  let extractedText = '';

  if (fileExt === 'pdf') {
    extractedText = await extractPdfText(file);
  } else if (['ods', 'xlsx', 'xls'].includes(fileExt || '')) {
    extractedText = await extractSpreadsheetText(file);
  } else if (['txt', 'csv', 'json'].includes(fileExt || '')) {
    extractedText = await new Promise<string>((resolve) => {
      const r = new FileReader();
      r.onload = () => resolve(r.result as string);
      r.readAsText(file);
    });
  }

  return analyzeText(extractedText, file.name);
}

interface PdfTextItem {
  text: string;
  x: number;
  y: number;
  width: number;
}

function parseTobbBordroRow(rowItems: PdfTextItem[]): Record<string, any> | null {
  // Sort items from left to right
  const sorted = [...rowItems].sort((a, b) => a.x - b.x);
  
  // Flatten and split tokens
  const tokens: string[] = [];
  sorted.forEach(item => {
    const text = item.text.trim();
    
    // Split TC + Name
    const tcMatch = text.match(/^(\d{11})\s*(.*)$/);
    if (tcMatch) {
      tokens.push(tcMatch[1]);
      if (tcMatch[2].trim()) {
        tokens.push(tcMatch[2].trim());
      }
      return;
    }

    // Split space-separated numbers
    if (text.includes(' ') && /\d+[\.,]\d+/.test(text)) {
      const parts = text.split(/\s+/).map(p => p.trim()).filter(p => p !== '');
      tokens.push(...parts);
      return;
    }

    tokens.push(text);
  });

  // Check if it starts with an 11-digit TC
  if (tokens.length === 0 || !/^\d{11}$/.test(tokens[0])) {
    return null; // not a personnel data row
  }

  // Initialize target row
  const rowMap: Record<string, any> = {
    "TC Kimlik No": tokens[0],
    "Adı": "",
    "Soyadı": "",
    "Sicil No": "",
    "Statü": "",
    "P.E.K.": "",
    "Ü.A.": "",
    "T.P.": "",
    "T.K.": "",
    "Ç.G.S.": "",
    "Açıklama": ""
  };

  let tokenIdx = 1;

  // 1. Parse Name (Adı) - first alphabetic item
  if (tokenIdx < tokens.length) {
    rowMap["Adı"] = tokens[tokenIdx];
    tokenIdx++;
  }

  // 2. Parse Surname (Soyadı) - second alphabetic item
  if (tokenIdx < tokens.length) {
    rowMap["Soyadı"] = tokens[tokenIdx];
    tokenIdx++;
  }

  // 3. Parse Sicil No & Statü
  if (tokenIdx < tokens.length) {
    const nextToken = tokens[tokenIdx];
    if (nextToken === "Çalısan" || nextToken === "Çalışan" || nextToken.toLowerCase().includes("çalışan") || nextToken.toLowerCase().includes("calisan")) {
      rowMap["Statü"] = nextToken;
      tokenIdx++;
    } else {
      rowMap["Sicil No"] = nextToken;
      tokenIdx++;
      if (tokenIdx < tokens.length) {
        rowMap["Statü"] = tokens[tokenIdx];
        tokenIdx++;
      }
    }
  }

  // 4. Parse money columns: P.E.K., Ü.A., T.P., T.K.
  const moneyColumns = ["P.E.K.", "Ü.A.", "T.P.", "T.K."];
  moneyColumns.forEach(col => {
    if (tokenIdx < tokens.length) {
      rowMap[col] = tokens[tokenIdx];
      tokenIdx++;
    }
  });

  // 5. Parse Ç.G.S.
  if (tokenIdx < tokens.length) {
    rowMap["Ç.G.S."] = tokens[tokenIdx];
    tokenIdx++;
  }

  // 6. Parse Açıklama
  if (tokenIdx < tokens.length) {
    rowMap["Açıklama"] = tokens.slice(tokenIdx).join(' ');
  }

  return rowMap;
}

// Core Coordinate-Based PDF Table Parsing Algorithm
export function parsePdfTableRows(file: File): Promise<Record<string, any>[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const arrayBuffer = e.target?.result as ArrayBuffer;
        const pdfjsLib = await ensurePdfJsLoaded();
        const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
        const pdf = await loadingTask.promise;

        let allPageRows: Record<string, any>[] = [];
        let isTobbBordro = false;
        const pageDataList: any[] = [];

        // Scan all pages in the PDF
        const pagesToScan = pdf.numPages;
        
        for (let pageNum = 1; pageNum <= pagesToScan; pageNum++) {
          const page = await pdf.getPage(pageNum);
          const textContent = await page.getTextContent();
          
          const items: PdfTextItem[] = textContent.items
            .map((it: any) => ({
              text: it.str.trim(),
              x: it.transform[4],
              y: it.transform[5],
              width: it.width || 0
            }))
            .filter((it: PdfTextItem) => it.text !== '');

          if (items.length === 0) continue;

          // Group items into rows using vertical (Y) coordinate threshold (±6 pixels)
          const sortedItems = [...items].sort((a, b) => b.y - a.y);
          const groupedRows: PdfTextItem[][] = [];
          
          sortedItems.forEach(item => {
            const threshold = 6;
            const existingRow = groupedRows.find(row => 
              row.some(rowItem => Math.abs(rowItem.y - item.y) <= threshold)
            );

            if (existingRow) {
              existingRow.push(item);
            } else {
              groupedRows.push([item]);
            }
          });

          // Sort items in each row by X ascending (left to right)
          groupedRows.forEach(row => row.sort((a, b) => a.x - b.x));
          
          // Sort grouped rows by average Y descending (top to bottom)
          const rowsWithAvgY = groupedRows.map(row => {
            const avgY = row.reduce((sum, item) => sum + item.y, 0) / row.length;
            return { items: row, y: avgY };
          });
          rowsWithAvgY.sort((a, b) => b.y - a.y);

          // Horizontal word-merge: Merge close horizontal chunks (gap < 4px)
          const mergedRows = rowsWithAvgY.map(rowObj => {
            const rawItems = rowObj.items;
            if (rawItems.length <= 1) return { items: rawItems, y: rowObj.y };

            const merged: PdfTextItem[] = [rawItems[0]];
            for (let i = 1; i < rawItems.length; i++) {
              const current = rawItems[i];
              const previous = merged[merged.length - 1];
              const gap = current.x - (previous.x + previous.width);
              
              if (gap >= -2 && gap < 4) {
                previous.text += (previous.text.endsWith(' ') || current.text.startsWith(' ') ? '' : ' ') + current.text;
                previous.width += current.width + gap;
              } else {
                merged.push(current);
              }
            }
            return { items: merged, y: rowObj.y };
          });

          // Identify if it's a TOBB structured bordro
          const headerKeywords = ['tc', 'kimlik', 'ad', 'soyad', 'sicil', 'statü', 'statu', 'p.e.k', 't.p.', 't.k.', 'ç.g.s'];
          let headerRowIndex = -1;
          
          for (let i = 0; i < mergedRows.length; i++) {
            const row = mergedRows[i].items;
            const matchCount = row.filter(item => 
              headerKeywords.some(k => item.text.toLowerCase().includes(k))
            ).length;
            
            if (matchCount >= 3) {
              headerRowIndex = i;
              isTobbBordro = true;
              break;
            }
          }

          pageDataList.push({
            mergedRows,
            headerRowIndex
          });
        }

        // Dual-Mode Processing:
        if (!isTobbBordro) {
          // MODE B: General layout document - Convert text coordinates directly to generic columns
          pageDataList.forEach((pageData) => {
            pageData.mergedRows.forEach((rowObj: any) => {
              const rowMap: Record<string, any> = {};
              let hasText = false;
              
              rowObj.items.forEach((item: any, colIdx: number) => {
                const text = item.text.trim();
                if (text !== '') {
                  // Format check: Convert formatted numbers like '11.793,00' or decimals to standard floats
                  const cleanText = text.replace(/\./g, '').replace(/,/g, '.');
                  const num = parseFloat(cleanText);
                  
                  if (!isNaN(num) && /^[\d\.,]+$/.test(text) && !text.includes('/') && !text.includes('-')) {
                    if (text.length >= 9 && /^\d+$/.test(text)) {
                      rowMap[`Sütun ${colIdx + 1}`] = text; // Keep TC/IDs as plain string
                    } else {
                      rowMap[`Sütun ${colIdx + 1}`] = num;
                    }
                  } else {
                    rowMap[`Sütun ${colIdx + 1}`] = text;
                  }
                  hasText = true;
                }
              });
              
              if (hasText) {
                allPageRows.push(rowMap);
              }
            });
          });
        } else {
          // MODE A: Structured TOBB Payroll document
          pageDataList.forEach((pageData) => {
            const { mergedRows, headerRowIndex } = pageData;
            if (headerRowIndex === -1) return;

            const headerY = mergedRows[headerRowIndex].y;
            
            const postSplitRows = mergedRows.map((rowObj: any) => {
              const splitItems: PdfTextItem[] = [];
              rowObj.items.forEach((item: any) => {
                const tcMatch = item.text.match(/^(\d{11})\s*(.*)$/);
                if (tcMatch) {
                  const tcStr = tcMatch[1];
                  const nameStr = tcMatch[2].trim();
                  const tcWidth = item.width * (tcStr.length / item.text.length);
                  
                  splitItems.push({
                    text: tcStr,
                    x: item.x,
                    y: item.y,
                    width: tcWidth
                  });

                  if (nameStr) {
                    const nameWidth = item.width - tcWidth;
                    splitItems.push({
                      text: nameStr,
                      x: item.x + tcWidth + 4,
                      y: item.y,
                      width: nameWidth
                    });
                  }
                  return;
                }

                if (item.text.includes(' ') && /\d+[\.,]\d+/.test(item.text)) {
                  const parts = item.text.split(/\s+/).map((p: any) => p.trim()).filter((p: any) => p !== '');
                  if (parts.length > 1) {
                    const partWidth = item.width / parts.length;
                    parts.forEach((part: any, idx: number) => {
                      splitItems.push({
                        text: part,
                        x: item.x + idx * partWidth,
                        y: item.y,
                        width: partWidth
                      });
                    });
                    return;
                  }
                }

                splitItems.push(item);
              });
              return splitItems.sort((a, b) => a.x - b.x);
            });

            const dataRows = postSplitRows.filter((_: any, idx: number) => mergedRows[idx].y < headerY);

            dataRows.forEach((row: any) => {
              const rowMap = parseTobbBordroRow(row);
              if (rowMap) {
                allPageRows.push(rowMap);
              }
            });
          });
        }

        resolve(allPageRows);
      } catch (err) {
        reject(err);
      }
    };

    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}
