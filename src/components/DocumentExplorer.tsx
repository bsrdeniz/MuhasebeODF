import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { Search, Trash2, Eye, Download, Calendar, FileText, ChevronLeft, ChevronRight, X, ExternalLink, Clipboard, Check, Sheet, FileSpreadsheet } from 'lucide-react';
import { getDocuments, deleteDocument, getDocumentRows, saveDocumentAndRows, type DocumentItem, type DocumentRow } from '../db';
import { parsePdfTableRows, autoExtractMetadata } from '../utils/extractor';

interface DocumentExplorerProps {
  cryptoKey: CryptoKey | null;
}

export const DocumentExplorer: React.FC<DocumentExplorerProps> = ({ cryptoKey }) => {
  const [documents, setDocuments] = useState<DocumentItem[]>([]);
  const [filteredDocs, setFilteredDocs] = useState<DocumentItem[]>([]);
  
  // Filters for Documents
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedType, setSelectedType] = useState('all');
  const [selectedCategory, setSelectedCategory] = useState('all');
  const [sortOrder, setSortOrder] = useState('date-desc'); // date-desc, date-asc, name-asc

  // Pagination for Documents
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 8;

  // Selected Document details & Preview
  const [previewDoc, setPreviewDoc] = useState<DocumentItem | null>(null);
  const [previewUrl, setPreviewUrl] = useState<string | null>(null);
  const [activePreviewTab, setActivePreviewTab] = useState<'info' | 'preview' | 'table'>('info');
  
  // Tabular spreadsheet data states
  const [docRows, setDocRows] = useState<DocumentRow[]>([]);
  const [filteredDocRows, setFilteredDocRows] = useState<DocumentRow[]>([]);
  const [rowSearchTerm, setRowSearchTerm] = useState('');
  const [copySuccess, setCopySuccess] = useState(false);

  // Cell Selection States (Excel-like Click & Drag)
  const [selectionStart, setSelectionStart] = useState<{ r: number; c: number } | null>(null);
  const [selectionEnd, setSelectionEnd] = useState<{ r: number; c: number } | null>(null);
  const [isSelecting, setIsSelecting] = useState(false);
  const [selectionActive, setSelectionActive] = useState(false);

  // Loading
  const [isLoading, setIsLoading] = useState(true);

  const formatCellValue = (val: any, headerKey: string) => {
    if (val === undefined || val === null) return '-';
    if (typeof val === 'number') {
      const lowerKey = headerKey.toLowerCase();
      if (
        lowerKey.includes('tc') || 
        lowerKey.includes('kimlik') || 
        lowerKey.includes('sicil') || 
        lowerKey.includes('no') || 
        lowerKey.includes('tckn')
      ) {
        return val.toFixed(0);
      }
      return val.toLocaleString('tr-TR', { minimumFractionDigits: 2 });
    }
    return String(val);
  };

  const formatCopyValue = (val: any, headerKey: string) => {
    if (val === undefined || val === null) return '';
    if (typeof val === 'number') {
      const lowerKey = headerKey.toLowerCase();
      if (
        lowerKey.includes('tc') || 
        lowerKey.includes('kimlik') || 
        lowerKey.includes('sicil') || 
        lowerKey.includes('no') || 
        lowerKey.includes('tckn')
      ) {
        return val.toFixed(0);
      }
      return val.toLocaleString('tr-TR', { minimumFractionDigits: 2, useGrouping: false });
    }
    return String(val);
  };

  const normalizeNumber = (val: any) => {
    if (typeof val === 'string' && val.trim() !== '') {
      const normalized = val.replace(/\s/g, '');
      const isNumeric = /^-?\d{1,3}(,\d{3})*(\.\d+)?$/.test(normalized) || /^-?\d{1,3}(\.\d{3})*(,\d+)?$/.test(normalized) || /^-?\d+(\.\d+)?$/.test(normalized);
      if (isNumeric) {
        let num = NaN;
        if (normalized.includes(',') && normalized.includes('.')) {
          if (normalized.indexOf(',') > normalized.indexOf('.')) {
            num = parseFloat(normalized.replace(/\./g, '').replace(',', '.'));
          } else {
            num = parseFloat(normalized.replace(/,/g, ''));
          }
        } else if (normalized.includes(',')) {
          const parts = normalized.split(',');
          if (parts[1].length <= 2) {
            num = parseFloat(normalized.replace(',', '.'));
          } else {
            num = parseFloat(normalized.replace(/,/g, ''));
          }
        } else {
          num = parseFloat(normalized);
        }
        if (!isNaN(num)) return num;
      }
    }
    return val;
  };

  const loadDocuments = async () => {
    setIsLoading(true);
    try {
      const docs = await getDocuments(cryptoKey);
      setDocuments(docs);
    } catch (err) {
      console.error('Belgeler yüklenirken hata:', err);
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    loadDocuments();
  }, [cryptoKey]);

  // Get categories list dynamically
  const categories = Array.from(new Set(documents.map(d => d.category))).filter(Boolean);

  // Filter and Sort Documents
  useEffect(() => {
    let result = [...documents];

    if (searchTerm) {
      const term = searchTerm.toLowerCase();
      result = result.filter(d => 
        d.name.toLowerCase().includes(term) || 
        d.description.toLowerCase().includes(term) ||
        d.fileName.toLowerCase().includes(term)
      );
    }

    if (selectedType !== 'all') {
      result = result.filter(d => d.docType === selectedType);
    }

    if (selectedCategory !== 'all') {
      result = result.filter(d => d.category === selectedCategory);
    }

    result.sort((a, b) => {
      if (sortOrder === 'date-desc') {
        return b.date.localeCompare(a.date);
      } else if (sortOrder === 'date-asc') {
        return a.date.localeCompare(b.date);
      } else if (sortOrder === 'name-asc') {
        return a.name.localeCompare(b.name);
      }
      return 0;
    });

    setFilteredDocs(result);
    setCurrentPage(1);
  }, [documents, searchTerm, selectedType, selectedCategory, sortOrder]);

  // Load spreadsheet rows when a document is selected (with instant auto-reparse fallback)
  useEffect(() => {
    const loadRows = async () => {
      if (previewDoc) {
        try {
          let rows = await getDocumentRows(previewDoc.id, cryptoKey);

          // AUTO-REPARSE: If 0 rows are found, parse the PDF on the fly and save to DB
          if (rows.length === 0 && previewDoc.fileData) {
            try {
              let parsedRows: Record<string, any>[] = [];
              if (previewDoc.fileType === 'PDF') {
                parsedRows = await parsePdfTableRows(previewDoc.fileData);
              }
              if (parsedRows.length > 0) {
                const cleaned = parsedRows.map(r => {
                  const cleanRow: Record<string, any> = {};
                  Object.entries(r).forEach(([k, v]) => {
                    cleanRow[k] = normalizeNumber(v);
                  });
                  return cleanRow;
                });
                
                const newDocRows: DocumentRow[] = cleaned.map((r, idx) => ({
                  id: `row_${previewDoc.id}_${idx}_${Math.random().toString(36).substring(2, 9)}`,
                  docId: previewDoc.id,
                  rowNumber: idx + 1,
                  data: r
                }));

                const meta = await autoExtractMetadata(previewDoc.fileData);
                const updatedDoc = {
                  ...previewDoc,
                  name: (previewDoc.name === previewDoc.fileName || previewDoc.docType === 'Diğer') ? meta.name : previewDoc.name,
                  docType: meta.docType,
                  category: meta.category,
                  description: meta.description
                };

                await saveDocumentAndRows(updatedDoc, newDocRows, cryptoKey);
                rows = newDocRows;
                setPreviewDoc(updatedDoc);
                loadDocuments();
              }
            } catch (reErr) {
              console.error('Otomatik reparse hatası:', reErr);
            }
          }

          setDocRows(rows);
          setFilteredDocRows(rows);
          setRowSearchTerm('');
          clearSelection();
          
          if (rows.length > 0) {
            setActivePreviewTab('table');
          } else {
            setActivePreviewTab('info');
          }
        } catch (err) {
          console.error('Satır verileri yüklenirken hata:', err);
          setDocRows([]);
          setFilteredDocRows([]);
          setActivePreviewTab('info');
        }
      } else {
        setDocRows([]);
        setFilteredDocRows([]);
        clearSelection();
      }
    };
    loadRows();
  }, [previewDoc?.id, cryptoKey]);

  // Filter document rows inside "Tablo Verileri" tab
  useEffect(() => {
    if (rowSearchTerm) {
      const term = rowSearchTerm.toLowerCase();
      const filtered = docRows.filter(row => {
        return Object.values(row.data).some(val => 
          String(val).toLowerCase().includes(term)
        );
      });
      setFilteredDocRows(filtered);
      clearSelection();
    } else {
      setFilteredDocRows(docRows);
    }
  }, [rowSearchTerm, docRows]);

  // Excel-like selection handlers
  const handleCellMouseDown = (r: number, c: number) => {
    setSelectionStart({ r, c });
    setSelectionEnd({ r, c });
    setIsSelecting(true);
    setSelectionActive(true);
  };

  const handleCellMouseEnter = (r: number, c: number) => {
    if (isSelecting) {
      setSelectionEnd({ r, c });
    }
  };

  const handleCellMouseUp = () => {
    setIsSelecting(false);
  };

  const isCellSelected = (r: number, c: number) => {
    if (!selectionStart || !selectionEnd) return false;
    const minR = Math.min(selectionStart.r, selectionEnd.r);
    const maxR = Math.max(selectionStart.r, selectionEnd.r);
    const minC = Math.min(selectionStart.c, selectionEnd.c);
    const maxC = Math.max(selectionStart.c, selectionEnd.c);
    return r >= minR && r <= maxR && c >= minC && c <= maxC;
  };

  const handleColumnHeaderClick = (colIndex: number) => {
    setSelectionStart({ r: 0, c: colIndex });
    setSelectionEnd({ r: filteredDocRows.length - 1, c: colIndex });
    setSelectionActive(true);
  };

  const handleRowSelect = (rowIndex: number) => {
    const headers = docRows.length > 0 ? Object.keys(docRows[0].data) : [];
    setSelectionStart({ r: rowIndex, c: 0 });
    setSelectionEnd({ r: rowIndex, c: headers.length - 1 });
    setSelectionActive(true);
  };

  const clearSelection = () => {
    setSelectionStart(null);
    setSelectionEnd(null);
    setSelectionActive(false);
  };

  // Global mouseUp listener
  useEffect(() => {
    const handleGlobalMouseUp = () => {
      setIsSelecting(false);
    };
    window.addEventListener('mouseup', handleGlobalMouseUp);
    return () => window.removeEventListener('mouseup', handleGlobalMouseUp);
  }, []);

  const spreadsheetHeaders = docRows.length > 0 ? Object.keys(docRows[0].data) : [];

  // Capture standard Ctrl+C Keyboard Event to copy selected cells
  useEffect(() => {
    const handleCopy = (e: ClipboardEvent) => {
      if (!selectionStart || !selectionEnd) return;
      
      e.preventDefault();
      
      const minR = Math.min(selectionStart.r, selectionEnd.r);
      const maxR = Math.max(selectionStart.r, selectionEnd.r);
      const minC = Math.min(selectionStart.c, selectionEnd.c);
      const maxC = Math.max(selectionStart.c, selectionEnd.c);

      const copyLines: string[] = [];
      for (let r = minR; r <= maxR; r++) {
        const row = filteredDocRows[r];
        if (!row) continue;
        
        const lineCells: string[] = [];
        for (let c = minC; c <= maxC; c++) {
          const headerKey = spreadsheetHeaders[c];
          const val = row.data[headerKey];
          lineCells.push(formatCopyValue(val, headerKey));
        }
        copyLines.push(lineCells.join('\t'));
      }

      const tsvString = copyLines.join('\r\n');
      e.clipboardData?.setData('text/plain', tsvString);
      
      setCopySuccess(true);
      setTimeout(() => setCopySuccess(false), 3000);
    };

    window.addEventListener('copy', handleCopy);
    return () => window.removeEventListener('copy', handleCopy);
  }, [selectionStart, selectionEnd, filteredDocRows, spreadsheetHeaders]);

  const handleOpenTableInNewTab = () => {
    if (filteredDocRows.length === 0) return;

    const newWindow = window.open('', '_blank');
    if (!newWindow) {
      alert('Yeni sekme açılması tarayıcınız tarafından engellendi. Lütfen pop-up engelleyicisini kaldırın.');
      return;
    }

    const headers = Object.keys(filteredDocRows[0].data);
    
    const tableHeadersHTML = `
      <tr>
        <th style="border: 1px solid #cbd5e1; padding: 10px; background-color: #f1f5f9; text-align: center; user-select: none;">#</th>
        ${headers.map((h, colIdx) => `<th class="selectable-header-col" data-col="${colIdx}" style="border: 1px solid #cbd5e1; padding: 10px; background-color: #f1f5f9; text-align: left;">${h}</th>`).join('')}
      </tr>
    `;

    const tableRowsHTML = filteredDocRows.map((row, rowIdx) => `
      <tr>
        <td class="selectable-header-row" data-row="${rowIdx}" style="border: 1px solid #e2e8f0; padding: 8px; font-weight: bold; text-align: center; background-color: #f8fafc;">${row.rowNumber}</td>
        ${headers.map((h, colIdx) => {
          const val = row.data[h];
          const displayVal = formatCellValue(val, h);
          const copyVal = formatCopyValue(val, h);
          
          const lowerKey = h.toLowerCase();
          const isNumericMeasure = typeof val === 'number' && 
            !(lowerKey.includes('tc') || lowerKey.includes('kimlik') || lowerKey.includes('sicil') || lowerKey.includes('no') || lowerKey.includes('tckn'));
            
          const style = isNumericMeasure ? 'text-align: right; font-weight: bold;' : 'text-align: left;';
          return `<td class="selectable-cell" data-row="${rowIdx}" data-col="${colIdx}" data-raw="${copyVal}" style="border: 1px solid #e2e8f0; padding: 8px; ${style}">${displayVal}</td>`;
        }).join('')}
      </tr>
    `).join('');

    const pageHTML = `
      <!DOCTYPE html>
      <html lang="tr">
      <head>
        <meta charset="UTF-8">
        <title>${previewDoc?.name || 'Tablo Verileri'}</title>
        <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;700&display=swap" rel="stylesheet">
        <!-- SheetJS Excel Library -->
        <script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>
        <style>
          body {
            font-family: 'Plus Jakarta Sans', sans-serif;
            margin: 0;
            padding: 30px;
            background-color: #f8fafc;
            color: #0f172a;
          }
          .container {
            max-width: 100%;
            margin: 0 auto;
            background: #ffffff;
            border-radius: 12px;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
            padding: 24px;
            border: 1px solid #e2e8f0;
          }
          h1 {
            font-size: 1.5rem;
            margin-top: 0;
            margin-bottom: 8px;
          }
          .subtitle {
            font-size: 0.85rem;
            color: #64748b;
            margin-bottom: 20px;
          }
          table {
            width: 100%;
            border-collapse: collapse;
            font-size: 0.82rem;
            margin-top: 15px;
          }
          th, td {
            border: 1px solid #e2e8f0;
            padding: 10px;
            white-space: nowrap;
          }
          .selectable-cell {
            user-select: none;
            -webkit-user-select: none;
            cursor: cell;
          }
          .selected-cell {
            background-color: rgba(79, 70, 229, 0.16) !important;
            outline: 1.5px solid #4f46e5;
            outline-offset: -1.5px;
          }
          .selectable-header-col, .selectable-header-row {
            cursor: pointer;
            user-select: none;
          }
          .selectable-header-col:hover, .selectable-header-row:hover {
            background-color: #cbd5e1 !important;
          }
          .btn {
            background-color: #4f46e5;
            color: white;
            border: none;
            padding: 8px 16px;
            font-weight: 600;
            border-radius: 6px;
            cursor: pointer;
            font-size: 0.85rem;
          }
          .btn:hover {
            background-color: #4338ca;
          }
          #toast {
            position: fixed;
            bottom: 20px;
            right: 20px;
            background-color: #10b981;
            color: white;
            padding: 12px 24px;
            border-radius: 8px;
            font-weight: 600;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
            z-index: 1000;
            display: none;
            font-size: 0.85rem;
          }
        </style>
      </head>
      <body>
        <div id="toast"></div>
        <div class="container">
          <div style="display: flex; justify-content: space-between; align-items: start;">
            <div>
              <h1>${previewDoc?.name || 'Tablo Verileri'}</h1>
              <div class="subtitle">
                Kaynak Dosya: ${previewDoc?.fileName} • Toplam ${filteredDocRows.length} Satır 
                <span style="margin-left: 10px; color: #4f46e5; font-weight: 600;">(Excel olarak indirmek için butona basın)</span>
              </div>
            </div>
            <div style="display: flex; gap: 8px;">
              <button class="btn" style="background-color: #10b981;" onclick="downloadExcel()">Excel İndir (.xls)</button>
              <button class="btn" onclick="window.print()">Yazdır / PDF Kaydet</button>
            </div>
          </div>
          <div style="overflow-x: auto;">
            <table data-max-row="${filteredDocRows.length - 1}" data-max-col="${headers.length - 1}">
              <thead>${tableHeadersHTML}</thead>
              <tbody>${tableRowsHTML}</tbody>
            </table>
          </div>
        </div>

        <script>
          console.log("DMS Tablo Seçim Motoru Başlatılıyor...");
          
          window.downloadExcel = function() {
            if (typeof XLSX === 'undefined') {
              alert('Excel kütüphanesi yükleniyor, lütfen 1 saniye sonra tekrar deneyin.');
              return;
            }
            const table = document.querySelector('table');
            const wb = XLSX.utils.table_to_book(table, { raw: true });
            XLSX.writeFile(wb, '${previewDoc?.name || 'tablo'}_verileri.xls', { bookType: 'xls' });
          };

          const checkExist = setInterval(() => {
            const cells = document.querySelectorAll('.selectable-cell');
            if (cells.length > 0) {
              clearInterval(checkExist);
              initializeSelection(cells);
            }
          }, 30);

          function initializeSelection(cells) {
            console.log("DMS Tablo Seçim Motoru Aktifleşti. Hücre Sayısı:", cells.length);
            let selectionStart = null;
            let selectionEnd = null;
            let isSelecting = false;

            const colHeaders = document.querySelectorAll('.selectable-header-col');
            const rowHeaders = document.querySelectorAll('.selectable-header-row');

            function getCoords(cell) {
              return {
                r: parseInt(cell.getAttribute('data-row') || '0'),
                c: parseInt(cell.getAttribute('data-col') || '0')
              };
            }

            function updateSelection() {
              if (!selectionStart || !selectionEnd) {
                cells.forEach(el => el.classList.remove('selected-cell'));
                return;
              }
              const minR = Math.min(selectionStart.r, selectionEnd.r);
              const maxR = Math.max(selectionStart.r, selectionEnd.r);
              const minC = Math.min(selectionStart.c, selectionEnd.c);
              const maxC = Math.max(selectionStart.c, selectionEnd.c);

              cells.forEach(el => {
                const r = parseInt(el.getAttribute('data-row') || '0');
                const c = parseInt(el.getAttribute('data-col') || '0');
                if (r >= minR && r <= maxR && c >= minC && c <= maxC) {
                  el.classList.add('selected-cell');
                } else {
                  el.classList.remove('selected-cell');
                }
              });
            }

            cells.forEach(cell => {
              cell.addEventListener('mousedown', (e) => {
                e.preventDefault(); // Prevents standard browser text selection
                const coords = getCoords(cell);
                selectionStart = coords;
                selectionEnd = coords;
                isSelecting = true;
                updateSelection();
              });

              cell.addEventListener('mouseenter', (e) => {
                if (isSelecting) {
                  selectionEnd = getCoords(cell);
                  updateSelection();
                }
              });
            });

            let autoScrollInterval = null;
            let mouseX = 0;
            let mouseY = 0;

            document.addEventListener('mousemove', (e) => {
              mouseX = e.clientX;
              mouseY = e.clientY;
              if (isSelecting) {
                startAutoScroll();
              }
            });

            function startAutoScroll() {
              if (autoScrollInterval) return;
              autoScrollInterval = setInterval(() => {
                const scrollSpeed = 15;
                const threshold = 60;
                const winHeight = window.innerHeight;
                const winWidth = window.innerWidth;
                let scrollX = 0;
                let scrollY = 0;

                if (mouseY > winHeight - threshold) {
                  scrollY = scrollSpeed;
                } else if (mouseY < threshold) {
                  scrollY = -scrollSpeed;
                }

                if (mouseX > winWidth - threshold) {
                  scrollX = scrollSpeed;
                } else if (mouseX < threshold) {
                  scrollX = -scrollSpeed;
                }

                if (scrollX !== 0 || scrollY !== 0) {
                  window.scrollBy(scrollX, scrollY);
                  // Find cell currently under the cursor and update selectionEnd
                  const element = document.elementFromPoint(mouseX, mouseY);
                  if (element && element.classList.contains('selectable-cell')) {
                    selectionEnd = getCoords(element);
                    updateSelection();
                  }
                }
              }, 30);
            }

            function stopAutoScroll() {
              if (autoScrollInterval) {
                clearInterval(autoScrollInterval);
                autoScrollInterval = null;
              }
            }

            document.addEventListener('mouseup', () => {
              isSelecting = false;
              stopAutoScroll();
            });

            // Prevent default drag and selection behaviors to allow smooth cell dragging
            document.addEventListener('dragstart', (e) => {
              e.preventDefault();
            });
            document.addEventListener('selectstart', (e) => {
              e.preventDefault();
            });

            colHeaders.forEach(h => {
              h.addEventListener('click', () => {
                const colIdx = parseInt(h.getAttribute('data-col') || '0');
                const table = document.querySelector('table');
                const maxRow = parseInt(table.getAttribute('data-max-row') || '0');
                selectionStart = { r: 0, c: colIdx };
                selectionEnd = { r: maxRow, c: colIdx };
                updateSelection();
              });
            });

            rowHeaders.forEach(rh => {
              rh.addEventListener('click', () => {
                const rowIdx = parseInt(rh.getAttribute('data-row') || '0');
                const table = document.querySelector('table');
                const maxCol = parseInt(table.getAttribute('data-max-col') || '0');
                selectionStart = { r: rowIdx, c: 0 };
                selectionEnd = { r: rowIdx, c: maxCol };
                updateSelection();
              });
            });

            function showToast(msg) {
              const toast = document.getElementById('toast');
              if (toast) {
                toast.innerText = msg;
                toast.style.display = 'block';
                setTimeout(() => {
                  toast.style.display = 'none';
                }, 2500);
              }
            }

            document.addEventListener('copy', (e) => {
              if (!selectionStart || !selectionEnd) return;
              e.preventDefault();

              const minR = Math.min(selectionStart.r, selectionEnd.r);
              const maxR = Math.max(selectionStart.r, selectionEnd.r);
              const minC = Math.min(selectionStart.c, selectionEnd.c);
              const maxC = Math.max(selectionStart.c, selectionEnd.c);

              const rows = document.querySelectorAll('table tbody tr');
              const copyLines = [];

              for (let r = minR; r <= maxR; r++) {
                const tr = rows[r];
                if (!tr) continue;
                const cellsInRow = tr.querySelectorAll('.selectable-cell');
                const lineCells = [];
                
                cellsInRow.forEach(td => {
                  const c = parseInt(td.getAttribute('data-col') || '0');
                  if (c >= minC && c <= maxC) {
                    const val = td.getAttribute('data-raw') || td.innerText || '';
                    lineCells.push(val);
                  }
                });
                copyLines.push(lineCells.join('\\t'));
              }

              const tsv = copyLines.join('\\r\\n');
              e.clipboardData.setData('text/plain', tsv);
              showToast("✓ Seçilen hücreler kopyalandı! Excel'e yapıştırabilirsiniz.");
            });
          }
        </script>
      </body>
      </html>
    `;

    newWindow.document.write(pageHTML);
    newWindow.document.close();
  };

  const handleDelete = async (id: string) => {
    if (window.confirm('Bu belgeyi ve ilişkili tüm tablo verilerini arşivden tamamen silmek istediğinizden emin misiniz?')) {
      await deleteDocument(id);
      loadDocuments();
      if (previewDoc?.id === id) {
        closePreview();
      }
    }
  };

  const handleDownload = (doc: DocumentItem) => {
    const url = URL.createObjectURL(doc.fileData);
    const link = window.document.createElement('a');
    link.href = url;
    link.download = doc.fileName;
    window.document.body.appendChild(link);
    link.click();
    window.document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const handlePreview = (doc: DocumentItem) => {
    if (previewUrl) {
      URL.revokeObjectURL(previewUrl);
    }
    const url = URL.createObjectURL(doc.fileData);
    setPreviewDoc(doc);
    setPreviewUrl(url);
  };

  const closePreview = () => {
    if (previewUrl) {
      URL.revokeObjectURL(previewUrl);
    }
    setPreviewDoc(null);
    setPreviewUrl(null);
  };

  // Button Click Copy (copies selection if active, otherwise whole table)
  const handleCopyToClipboard = () => {
    let tsvString = '';
    
    if (selectionStart && selectionEnd) {
      const minR = Math.min(selectionStart.r, selectionEnd.r);
      const maxR = Math.max(selectionStart.r, selectionEnd.r);
      const minC = Math.min(selectionStart.c, selectionEnd.c);
      const maxC = Math.max(selectionStart.c, selectionEnd.c);

      const copyLines: string[] = [];
      for (let r = minR; r <= maxR; r++) {
        const row = filteredDocRows[r];
        if (!row) continue;
        const lineCells: string[] = [];
        for (let c = minC; c <= maxC; c++) {
          const headerKey = spreadsheetHeaders[c];
          const val = row.data[headerKey];
          lineCells.push(formatCopyValue(val, headerKey));
        }
        copyLines.push(lineCells.join('\t'));
      }
      tsvString = copyLines.join('\r\n');
    } else {
      if (filteredDocRows.length === 0) return;
      const headers = Object.keys(filteredDocRows[0].data);
      const headerLine = headers.join('\t');
      const bodyLines = filteredDocRows.map(row => 
        headers.map(h => {
          const val = row.data[h];
          return formatCopyValue(val, h);
        }).join('\t')
      );
      tsvString = [headerLine, ...bodyLines].join('\r\n');
    }

    navigator.clipboard.writeText(tsvString)
      .then(() => {
        setCopySuccess(true);
        setTimeout(() => setCopySuccess(false), 3000);
      })
      .catch(err => {
        console.error('Kopyalama hatası:', err);
        alert('Panoya kopyalanamadı.');
      });
  };

  // CSV Export logic
  const handleDownloadCSV = () => {
    if (filteredDocRows.length === 0) return;

    const headers = Object.keys(filteredDocRows[0].data);
    
    const csvRows = [
      headers.join(';'),
      ...filteredDocRows.map(row => 
        headers.map(h => {
          const val = row.data[h];
          const valStr = formatCopyValue(val, h);
          return valStr.includes(';') ? `"${valStr}"` : valStr;
        }).join(';')
      )
    ];

    const csvContent = '\uFEFF' + csvRows.join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = window.document.createElement('a');
    link.href = url;
    link.download = `${previewDoc?.name || 'tablo'}_verileri.csv`;
    window.document.body.appendChild(link);
    link.click();
    window.document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  // Excel (.xls) Export logic using SheetJS (universally compatible format)
  const handleDownloadExcel = () => {
    if (filteredDocRows.length === 0) return;

    const headers = Object.keys(filteredDocRows[0].data);
    
    // Construct rows data: array of arrays, starting with headers
    const dataMatrix = [
      headers,
      ...filteredDocRows.map(row => 
        headers.map(h => {
          const val = row.data[h];
          return formatCopyValue(val, h);
        })
      )
    ];

    const worksheet = XLSX.utils.aoa_to_sheet(dataMatrix);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Tablo Verileri');

    // Generate universally compatible Excel 97-2003 binary format (.xls)
    const excelBuffer = XLSX.write(workbook, { bookType: 'xls', type: 'array' });
    const blob = new Blob([excelBuffer], { type: 'application/vnd.ms-excel' });
    const url = URL.createObjectURL(blob);
    
    const excelLink = window.document.createElement('a');
    excelLink.href = url;
    excelLink.download = `${previewDoc?.name || 'tablo'}_verileri.xls`;
    window.document.body.appendChild(excelLink);
    excelLink.click();
    window.document.body.removeChild(excelLink);
    URL.revokeObjectURL(url);
  };

  // Direct single-document Excel (.xls) download from document card
  const handleDirectDownloadDocExcel = async (doc: DocumentItem) => {
    try {
      let rows = await getDocumentRows(doc.id, cryptoKey);
      if (rows.length === 0 && doc.fileData) {
        try {
          let parsedRows: Record<string, any>[] = [];
          if (doc.fileType === 'PDF') {
            parsedRows = await parsePdfTableRows(doc.fileData);
          }
          if (parsedRows.length > 0) {
            const cleaned = parsedRows.map(r => {
              const cleanRow: Record<string, any> = {};
              Object.entries(r).forEach(([k, v]) => {
                cleanRow[k] = normalizeNumber(v);
              });
              return cleanRow;
            });
            const newDocRows: DocumentRow[] = cleaned.map((r, idx) => ({
              id: `row_${doc.id}_${idx}_${Math.random().toString(36).substring(2, 9)}`,
              docId: doc.id,
              rowNumber: idx + 1,
              data: r
            }));
            const meta = await autoExtractMetadata(doc.fileData);
            const updatedDoc = {
              ...doc,
              name: (doc.name === doc.fileName || doc.docType === 'Diğer') ? meta.name : doc.name,
              docType: meta.docType,
              category: meta.category,
              description: meta.description
            };
            await saveDocumentAndRows(updatedDoc, newDocRows, cryptoKey);
            rows = newDocRows;
            loadDocuments();
          }
        } catch (e) {
          console.error('Doğrudan reparse hatası:', e);
        }
      }

      if (rows.length === 0) {
        alert('Bu belgeden çözümlenmiş tablo verisi bulunamadı.');
        return;
      }

      const headers = Object.keys(rows[0].data);
      const dataMatrix = [
        headers,
        ...rows.map(row => 
          headers.map(h => formatCopyValue(row.data[h], h))
        )
      ];

      const worksheet = XLSX.utils.aoa_to_sheet(dataMatrix);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Tablo Verileri');

      const excelBuffer = XLSX.write(workbook, { bookType: 'xls', type: 'array' });
      const blob = new Blob([excelBuffer], { type: 'application/vnd.ms-excel' });
      const url = URL.createObjectURL(blob);
      const link = window.document.createElement('a');
      link.href = url;
      link.download = `${doc.name || 'tablo'}_verileri.xls`;
      window.document.body.appendChild(link);
      link.click();
      window.document.body.removeChild(link);
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error('Doğrudan Excel indirme hatası:', err);
      alert('Excel dosyası oluşturulurken hata oluştu.');
    }
  };

  // Merges and downloads ALL archived documents into a single combined Excel spreadsheet (.xls)
  const handleDownloadAllArchiveExcel = async () => {
    if (documents.length === 0) {
      alert('Arşivde birleştirilecek evrak bulunmuyor.');
      return;
    }

    try {
      const allRowsMatrix: any[] = [];
      let headers: string[] = [];

      for (const doc of documents) {
        let rows = await getDocumentRows(doc.id, cryptoKey);
        if (rows.length === 0 && doc.fileData) {
          try {
            let parsedRows: Record<string, any>[] = [];
            if (doc.fileType === 'PDF') {
              parsedRows = await parsePdfTableRows(doc.fileData);
            }
            if (parsedRows.length > 0) {
              const cleaned = parsedRows.map(r => {
                const cleanRow: Record<string, any> = {};
                Object.entries(r).forEach(([k, v]) => {
                  cleanRow[k] = normalizeNumber(v);
                });
                return cleanRow;
              });
              const newDocRows: DocumentRow[] = cleaned.map((r, idx) => ({
                id: `row_${doc.id}_${idx}_${Math.random().toString(36).substring(2, 9)}`,
                docId: doc.id,
                rowNumber: idx + 1,
                data: r
              }));
              await saveDocumentAndRows(doc, newDocRows, cryptoKey);
              rows = newDocRows;
            }
          } catch (e) {}
        }
        if (rows.length === 0) continue;

        // Determine spreadsheet headers on first document hit
        if (headers.length === 0) {
          // Put metadata headers first so user can filter by source file
          headers = ["Kaynak Evrak", "Belge Türü", "Tarih", ...Object.keys(rows[0].data)];
        }

        rows.forEach(row => {
          const rowData = headers.map(h => {
            if (h === "Kaynak Evrak") return doc.name;
            if (h === "Belge Türü") return doc.docType;
            if (h === "Tarih") return doc.date;
            const val = row.data[h];
            return formatCopyValue(val, h);
          });
          allRowsMatrix.push(rowData);
        });
      }

      if (allRowsMatrix.length === 0) {
        alert('Arşivdeki belgelerden çözümlenmiş tablo verisi bulunamadı.');
        return;
      }

      // Prepend headers to rows matrix
      allRowsMatrix.unshift(headers);

      const worksheet = XLSX.utils.aoa_to_sheet(allRowsMatrix);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Tüm Arşiv Verileri');

      // Export as universally compatible .xls format
      const excelBuffer = XLSX.write(workbook, { bookType: 'xls', type: 'array' });
      const blob = new Blob([excelBuffer], { type: 'application/vnd.ms-excel' });
      const url = URL.createObjectURL(blob);
      
      const link = window.document.createElement('a');
      link.href = url;
      link.download = `Tum_Arsiv_Bordro_Verileri.xls`;
      window.document.body.appendChild(link);
      link.click();
      window.document.body.removeChild(link);
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error('Arşivi birleştirme hatası:', err);
      alert('Arşiv verileri indirilirken hata oluştu.');
    }
  };

  // Combines all archived documents and opens them in a new tab for select/copy/download
  const handleOpenAllArchiveInNewTab = async () => {
    if (documents.length === 0) {
      alert('Arşivde birleştirilecek evrak bulunmuyor.');
      return;
    }

    try {
      const allMergedRows: any[] = [];
      let headers: string[] = [];

      for (const doc of documents) {
        let rows = await getDocumentRows(doc.id, cryptoKey);
        if (rows.length === 0 && doc.fileData) {
          try {
            let parsedRows: Record<string, any>[] = [];
            if (doc.fileType === 'PDF') {
              parsedRows = await parsePdfTableRows(doc.fileData);
            }
            if (parsedRows.length > 0) {
              const cleaned = parsedRows.map(r => {
                const cleanRow: Record<string, any> = {};
                Object.entries(r).forEach(([k, v]) => {
                  cleanRow[k] = normalizeNumber(v);
                });
                return cleanRow;
              });
              const newDocRows: DocumentRow[] = cleaned.map((r, idx) => ({
                id: `row_${doc.id}_${idx}_${Math.random().toString(36).substring(2, 9)}`,
                docId: doc.id,
                rowNumber: idx + 1,
                data: r
              }));
              await saveDocumentAndRows(doc, newDocRows, cryptoKey);
              rows = newDocRows;
            }
          } catch (e) {}
        }
        if (rows.length === 0) continue;

        if (headers.length === 0) {
          headers = ["Kaynak Evrak", "Belge Türü", "Tarih", ...Object.keys(rows[0].data)];
        }

        rows.forEach((row) => {
          const rowObj = {
            rowNumber: allMergedRows.length + 1,
            data: {} as Record<string, any>
          };
          headers.forEach(h => {
            if (h === "Kaynak Evrak") rowObj.data[h] = doc.name;
            else if (h === "Belge Türü") rowObj.data[h] = doc.docType;
            else if (h === "Tarih") rowObj.data[h] = doc.date;
            else rowObj.data[h] = row.data[h] === undefined || row.data[h] === null ? '' : row.data[h];
          });
          allMergedRows.push(rowObj);
        });
      }

      if (allMergedRows.length === 0) {
        alert('Arşivdeki belgelerden çözümlenmiş tablo verisi bulunamadı.');
        return;
      }

      const newWindow = window.open('', '_blank');
      if (!newWindow) {
        alert('Yeni sekme açılması tarayıcınız tarafından engellendi. Lütfen pop-up engelleyicisini kaldırın.');
        return;
      }

      const tableHeadersHTML = `
        <tr>
          <th style="border: 1px solid #cbd5e1; padding: 10px; background-color: #f1f5f9; text-align: center; user-select: none;">#</th>
          ${headers.map((h, colIdx) => `<th class="selectable-header-col" data-col="${colIdx}" style="border: 1px solid #cbd5e1; padding: 10px; background-color: #f1f5f9; text-align: left;">${h}</th>`).join('')}
        </tr>
      `;

      const tableRowsHTML = allMergedRows.map((row, rowIdx) => `
        <tr>
          <td class="selectable-header-row" data-row="${rowIdx}" style="border: 1px solid #e2e8f0; padding: 8px; font-weight: bold; text-align: center; background-color: #f8fafc;">${row.rowNumber}</td>
          ${headers.map((h, colIdx) => {
            const val = row.data[h];
            const displayVal = formatCellValue(val, h);
            const copyVal = formatCopyValue(val, h);
            
            const lowerKey = h.toLowerCase();
            const isNumericMeasure = typeof val === 'number' && 
              !(lowerKey.includes('tc') || lowerKey.includes('kimlik') || lowerKey.includes('sicil') || lowerKey.includes('no') || lowerKey.includes('tckn'));
              
            const style = isNumericMeasure ? 'text-align: right; font-weight: bold;' : 'text-align: left;';
            return `<td class="selectable-cell" data-row="${rowIdx}" data-col="${colIdx}" data-raw="${copyVal}" style="border: 1px solid #e2e8f0; padding: 8px; ${style}">${displayVal}</td>`;
          }).join('')}
        </tr>
      `).join('');

      const pageHTML = `
        <!DOCTYPE html>
        <html lang="tr">
        <head>
          <meta charset="UTF-8">
          <title>Birleştirilmiş Tüm Arşiv Verileri</title>
          <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;700&display=swap" rel="stylesheet">
          <script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>
          <style>
            body {
              font-family: 'Plus Jakarta Sans', sans-serif;
              margin: 0;
              padding: 30px;
              background-color: #f8fafc;
              color: #0f172a;
            }
            .container {
              max-width: 100%;
              margin: 0 auto;
              background: #ffffff;
              border-radius: 12px;
              box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
              padding: 24px;
              border: 1px solid #e2e8f0;
            }
            h1 {
              font-size: 1.5rem;
              margin-top: 0;
              margin-bottom: 8px;
            }
            .subtitle {
              font-size: 0.85rem;
              color: #64748b;
              margin-bottom: 20px;
            }
            table {
              width: 100%;
              border-collapse: collapse;
              font-size: 0.82rem;
              margin-top: 15px;
            }
            th, td {
              border: 1px solid #e2e8f0;
              padding: 10px;
              white-space: nowrap;
            }
            .selectable-cell {
              user-select: none;
              -webkit-user-select: none;
              cursor: cell;
            }
            .selected-cell {
              background-color: rgba(79, 70, 229, 0.16) !important;
              outline: 1.5px solid #4f46e5;
              outline-offset: -1.5px;
            }
            .selectable-header-col, .selectable-header-row {
              cursor: pointer;
              user-select: none;
            }
            .selectable-header-col:hover, .selectable-header-row:hover {
              background-color: #cbd5e1 !important;
            }
            .btn {
              background-color: #4f46e5;
              color: white;
              border: none;
              padding: 8px 16px;
              font-weight: 600;
              border-radius: 6px;
              cursor: pointer;
              font-size: 0.85rem;
            }
            .btn:hover {
              background-color: #4338ca;
            }
            #toast {
              position: fixed;
              bottom: 20px;
              right: 20px;
              background-color: #10b981;
              color: white;
              padding: 12px 24px;
              border-radius: 8px;
              font-weight: 600;
              box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
              z-index: 1000;
              display: none;
              font-size: 0.85rem;
            }
          </style>
        </head>
        <body>
          <div id="toast"></div>
          <div class="container">
            <div style="display: flex; justify-content: space-between; align-items: start;">
              <div>
                <h1>Birleştirilmiş Tüm Arşiv Verileri</h1>
                <div class="subtitle">
                  Toplam ${allMergedRows.length} Satır Yüklendi • 
                  <span style="color: #4f46e5; font-weight: 600;">Fareyle sürükleyerek hücre veya sütun seçip Ctrl+C ile Excel'e doğrudan yapıştırabilirsiniz.</span>
                </div>
              </div>
              <div style="display: flex; gap: 8px;">
                <button class="btn" style="background-color: #10b981;" onclick="downloadExcel()">Excel İndir (.xls)</button>
                <button class="btn" onclick="window.print()">Yazdır / PDF Kaydet</button>
              </div>
            </div>
            <div style="overflow-x: auto;">
              <table data-max-row="${allMergedRows.length - 1}" data-max-col="${headers.length - 1}">
                <thead>${tableHeadersHTML}</thead>
                <tbody>${tableRowsHTML}</tbody>
              </table>
            </div>
          </div>

          <script>
            window.downloadExcel = function() {
              if (typeof XLSX === 'undefined') {
                alert('Excel kütüphanesi yükleniyor, lütfen 1 saniye sonra tekrar deneyin.');
                return;
              }
              const table = document.querySelector('table');
              const wb = XLSX.utils.table_to_book(table, { raw: true });
              XLSX.writeFile(wb, 'Tum_Arsiv_Bordro_Verileri.xls', { bookType: 'xls' });
            };

            const checkExist = setInterval(() => {
              const cells = document.querySelectorAll('.selectable-cell');
              if (cells.length > 0) {
                clearInterval(checkExist);
                initializeSelection(cells);
              }
            }, 30);

            function initializeSelection(cells) {
              let selectionStart = null;
              let selectionEnd = null;
              let isSelecting = false;

              const colHeaders = document.querySelectorAll('.selectable-header-col');
              const rowHeaders = document.querySelectorAll('.selectable-header-row');

              function getCoords(cell) {
                return {
                  r: parseInt(cell.getAttribute('data-row') || '0'),
                  c: parseInt(cell.getAttribute('data-col') || '0')
                };
              }

              function updateSelection() {
                if (!selectionStart || !selectionEnd) {
                  cells.forEach(el => el.classList.remove('selected-cell'));
                  return;
                }
                const minR = Math.min(selectionStart.r, selectionEnd.r);
                const maxR = Math.max(selectionStart.r, selectionEnd.r);
                const minC = Math.min(selectionStart.c, selectionEnd.c);
                const maxC = Math.max(selectionStart.c, selectionEnd.c);

                cells.forEach(el => {
                  const r = parseInt(el.getAttribute('data-row') || '0');
                  const c = parseInt(el.getAttribute('data-col') || '0');
                  if (r >= minR && r <= maxR && c >= minC && c <= maxC) {
                    el.classList.add('selected-cell');
                  } else {
                    el.classList.remove('selected-cell');
                  }
                });
              }

              cells.forEach(cell => {
                cell.addEventListener('mousedown', (e) => {
                  e.preventDefault();
                  const coords = getCoords(cell);
                  selectionStart = coords;
                  selectionEnd = coords;
                  isSelecting = true;
                  updateSelection();
                });

                cell.addEventListener('mouseenter', (e) => {
                  if (isSelecting) {
                    selectionEnd = getCoords(cell);
                    updateSelection();
                  }
                });
              });

              let autoScrollInterval = null;
              let mouseX = 0;
              let mouseY = 0;

              document.addEventListener('mousemove', (e) => {
                mouseX = e.clientX;
                mouseY = e.clientY;
                if (isSelecting) {
                  startAutoScroll();
                }
              });

              function startAutoScroll() {
                if (autoScrollInterval) return;
                autoScrollInterval = setInterval(() => {
                  const scrollSpeed = 15;
                  const threshold = 60;
                  const winHeight = window.innerHeight;
                  const winWidth = window.innerWidth;
                  let scrollX = 0;
                  let scrollY = 0;

                  if (mouseY > winHeight - threshold) {
                    scrollY = scrollSpeed;
                  } else if (mouseY < threshold) {
                    scrollY = -scrollSpeed;
                  }

                  if (mouseX > winWidth - threshold) {
                    scrollX = scrollSpeed;
                  } else if (mouseX < threshold) {
                    scrollX = -scrollSpeed;
                  }

                  if (scrollX !== 0 || scrollY !== 0) {
                    window.scrollBy(scrollX, scrollY);
                    const element = document.elementFromPoint(mouseX, mouseY);
                    if (element && element.classList.contains('selectable-cell')) {
                      selectionEnd = getCoords(element);
                      updateSelection();
                    }
                  }
                }, 30);
              }

              function stopAutoScroll() {
                if (autoScrollInterval) {
                  clearInterval(autoScrollInterval);
                  autoScrollInterval = null;
                }
              }

              document.addEventListener('mouseup', () => {
                isSelecting = false;
                stopAutoScroll();
              });

              document.addEventListener('dragstart', (e) => e.preventDefault());
              document.addEventListener('selectstart', (e) => e.preventDefault());

              colHeaders.forEach(h => {
                h.addEventListener('click', () => {
                  const colIdx = parseInt(h.getAttribute('data-col') || '0');
                  const table = document.querySelector('table');
                  const maxRow = parseInt(table.getAttribute('data-max-row') || '0');
                  selectionStart = { r: 0, c: colIdx };
                  selectionEnd = { r: maxRow, c: colIdx };
                  updateSelection();
                });
              });

              rowHeaders.forEach(rh => {
                rh.addEventListener('click', () => {
                  const rowIdx = parseInt(rh.getAttribute('data-row') || '0');
                  const table = document.querySelector('table');
                  const maxCol = parseInt(table.getAttribute('data-max-col') || '0');
                  selectionStart = { r: rowIdx, c: 0 };
                  selectionEnd = { r: rowIdx, c: maxCol };
                  updateSelection();
                });
              });

              window.addEventListener('copy', (e) => {
                if (!selectionStart || !selectionEnd) return;
                e.preventDefault();
                const minR = Math.min(selectionStart.r, selectionEnd.r);
                const maxR = Math.max(selectionStart.r, selectionEnd.r);
                const minC = Math.min(selectionStart.c, selectionEnd.c);
                const maxC = Math.max(selectionStart.c, selectionEnd.c);

                const copyLines = [];
                const table = document.querySelector('table');

                for (let r = minR; r <= maxR; r++) {
                  const lineCells = [];
                  for (let c = minC; c <= maxC; c++) {
                    const cell = table.querySelector('td.selectable-cell[data-row="' + r + '"][data-col="' + c + '"]');
                    if (cell) {
                      lineCells.push(cell.getAttribute('data-raw') || cell.innerText);
                    } else {
                      lineCells.push('');
                    }
                  }
                  copyLines.push(lineCells.join('\\t'));
                }

                const tsvString = copyLines.join('\\r\\n');
                e.clipboardData.setData('text/plain', tsvString);

                const toast = document.getElementById('toast');
                toast.innerText = 'Seçilen ' + (maxR - minR + 1) + ' satır Excel\'e yapıştırılmak üzere kopyalandı!';
                toast.style.display = 'block';
                setTimeout(() => {
                  toast.style.display = 'none';
                }, 3000);
              });
            }
          </script>
        </body>
        </html>
      `;

      newWindow.document.write(pageHTML);
      newWindow.document.close();
    } catch (err) {
      console.error('Yeni sekmede birleştirme hatası:', err);
      alert('Arşiv verileri yeni sekmede açılırken hata oluştu.');
    }
  };

  // Pagination for Document list
  const indexOfLastItem = currentPage * itemsPerPage;
  const indexOfFirstItem = indexOfLastItem - itemsPerPage;
  const currentItems = filteredDocs.slice(indexOfFirstItem, indexOfLastItem);
  const totalPages = Math.ceil(filteredDocs.length / itemsPerPage);

  const getDocTypeBadgeClass = (type: string) => {
    switch (type) {
      case 'Maaş Bordrosu': return 'badge-success';
      case 'TOBB Belgesi': return 'badge-info';
      case 'Fatura': return 'badge-warning';
      case 'Sözleşme': return 'badge-danger';
      default: return 'badge-secondary';
    }
  };

  const getFileIcon = () => {
    return <FileText size={28} style={{ color: 'var(--primary)' }} />;
  };

  const formatFileSize = (bytes: number) => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const dm = 2;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
  };

  // Compute column statistics dynamically for numeric values in spreadsheet rows
  const getColumnStats = () => {
    if (docRows.length === 0) return [];
    
    const firstRowData = docRows[0].data;
    const numericKeys = Object.keys(firstRowData).filter(key => {
      return docRows.some(row => typeof row.data[key] === 'number');
    });

    return numericKeys.map(key => {
      const sum = filteredDocRows.reduce((acc, row) => {
        const val = row.data[key];
        return acc + (typeof val === 'number' ? val : 0);
      }, 0);

      return {
        key,
        sum
      };
    });
  };

  const columnStats = getColumnStats();

  return (
    <div className="animate-fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
      {/* Title */}
      <div>
        <h1 className="text-h1" style={{ marginBottom: '8px' }}>Belge Gezgini</h1>
        <p style={{ color: 'var(--text-secondary)' }}>
          Arşivlenen tüm resmi evrakları, maaş bordrolarını ve TOBB belgelerini arayın, filtreleyin ve görüntüleyin.
        </p>
      </div>

      {isLoading ? (
        <div style={{ textAlign: 'center', padding: '40px', color: 'var(--text-muted)' }}>
          Yükleniyor...
        </div>
      ) : (
        <div style={{
          display: 'grid',
          gridTemplateColumns: previewDoc ? '1fr 600px' : '1fr',
          gap: '24px',
          alignItems: 'start',
          transition: 'all 0.3s ease'
        }}>
          
          {/* MAIN DOCUMENT LIST */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
            {/* Filter toolbar */}
            <div className="glass-card" style={{ padding: '20px', backgroundColor: 'var(--bg-secondary)', display: 'flex', flexDirection: 'column', gap: '16px' }}>
              <div style={{ display: 'flex', gap: '12px', flexWrap: 'wrap' }}>
                {/* Search */}
                <div style={{ position: 'relative', flex: 1, minWidth: '240px' }}>
                  <Search size={16} style={{ position: 'absolute', left: '12px', top: '50%', transform: 'translateY(-50%)', color: 'var(--text-muted)' }} />
                  <input
                    type="text"
                    placeholder="Belge adı, açıklama veya dosya adı ara..."
                    className="form-input"
                    style={{ paddingLeft: '36px' }}
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                  />
                </div>

                {/* Doc Type Filter */}
                <select
                  className="form-input"
                  style={{ width: '180px' }}
                  value={selectedType}
                  onChange={(e) => setSelectedType(e.target.value)}
                >
                  <option value="all">Tüm Belge Türleri</option>
                  <option value="Maaş Bordrosu">Maaş Bordrosu</option>
                  <option value="TOBB Belgesi">TOBB Belgesi</option>
                  <option value="Fatura">Fatura</option>
                  <option value="Sözleşme">Sözleşme</option>
                  <option value="Diğer">Diğer</option>
                </select>

                {/* Category Filter */}
                <select
                  className="form-input"
                  style={{ width: '180px' }}
                  value={selectedCategory}
                  onChange={(e) => setSelectedCategory(e.target.value)}
                >
                  <option value="all">Tüm Kategoriler</option>
                  {categories.map(c => <option key={c} value={c}>{c}</option>)}
                </select>

                {/* Sorting */}
                <select
                  className="form-input"
                  style={{ width: '160px' }}
                  value={sortOrder}
                  onChange={(e) => setSortOrder(e.target.value)}
                >
                  <option value="date-desc">Tarih (Yeniye Doğru)</option>
                  <option value="date-asc">Tarih (Eskiye Doğru)</option>
                  <option value="name-asc">İsim (A-Z)</option>
                </select>

                {documents.length > 0 && (
                  <div style={{ display: 'flex', gap: '8px', marginLeft: 'auto' }}>
                    <button
                      onClick={handleOpenAllArchiveInNewTab}
                      className="btn"
                      style={{
                        height: '38px',
                        padding: '0 16px',
                        display: 'flex',
                        alignItems: 'center',
                        gap: '8px',
                        fontSize: '0.8rem',
                        fontWeight: 600,
                        backgroundColor: '#4f46e5',
                        color: '#ffffff',
                        border: 'none',
                        borderRadius: '6px',
                        cursor: 'pointer',
                        whiteSpace: 'nowrap'
                      }}
                      title="Arşivdeki tüm evrakların verilerini birleştirip yeni sekmede Excel formatında aç"
                    >
                      <ExternalLink size={14} /> Tüm Arşivi Yeni Sekmede Aç
                    </button>
                    
                    <button
                      onClick={handleDownloadAllArchiveExcel}
                      className="btn btn-success"
                      style={{
                        height: '38px',
                        padding: '0 16px',
                        display: 'flex',
                        alignItems: 'center',
                        gap: '8px',
                        fontSize: '0.8rem',
                        fontWeight: 600,
                        backgroundColor: '#10b981',
                        color: '#ffffff',
                        border: 'none',
                        borderRadius: '6px',
                        cursor: 'pointer',
                        whiteSpace: 'nowrap'
                      }}
                      title="Arşivdeki tüm evrakların verilerini birleştirip tek bir Excel dosyası (.xls) olarak indir"
                    >
                      <Download size={14} /> Tüm Arşivi Birleştir ve Excel İndir (.xls)
                    </button>
                  </div>
                )}
              </div>
            </div>

            {/* Document Card Grid */}
            {filteredDocs.length === 0 ? (
              <div className="glass-card" style={{ padding: '60px 20px', textAlign: 'center', color: 'var(--text-muted)' }}>
                Aradığınız kriterlere uygun arşivlenmiş belge bulunamadı.
              </div>
            ) : (
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(280px, 1fr))', gap: '20px' }}>
                {currentItems.map((doc) => (
                  <div
                    key={doc.id}
                    className="glass-card"
                    style={{
                      padding: '20px',
                      backgroundColor: 'var(--bg-secondary)',
                      display: 'flex',
                      flexDirection: 'column',
                      gap: '14px',
                      border: previewDoc?.id === doc.id ? '2px solid var(--primary)' : '1px solid var(--border-color)',
                      transform: previewDoc?.id === doc.id ? 'scale(1.01)' : 'none',
                    }}
                  >
                    {/* Header: Icon & Metadata */}
                    <div style={{ display: 'flex', gap: '12px', alignItems: 'start' }}>
                      <div style={{
                        padding: '10px',
                        backgroundColor: 'var(--bg-tertiary)',
                        borderRadius: '8px',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center'
                      }}>
                        {getFileIcon()}
                      </div>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <h3 className="text-h3" style={{ fontSize: '0.95rem', marginBottom: '4px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={doc.name}>
                          {doc.name}
                        </h3>
                        <p style={{ color: 'var(--text-muted)', fontSize: '0.75rem', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                          {doc.fileName}
                        </p>
                      </div>
                    </div>

                    {/* Badge row */}
                    <div style={{ display: 'flex', gap: '6px', flexWrap: 'wrap' }}>
                      <span className={`badge ${getDocTypeBadgeClass(doc.docType)}`}>
                        {doc.docType}
                      </span>
                      <span className="badge badge-info" style={{ backgroundColor: 'var(--primary-light)', color: 'var(--primary)' }}>
                        {doc.category}
                      </span>
                    </div>

                    {/* Stats & Description */}
                    {doc.description && (
                      <p style={{ fontSize: '0.8rem', color: 'var(--text-secondary)', display: '-webkit-box', WebkitLineClamp: 2, WebkitBoxOrient: 'vertical', overflow: 'hidden', minHeight: '36px' }}>
                        {doc.description}
                      </p>
                    )}
                    
                    <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.75rem', color: 'var(--text-muted)', borderTop: '1px solid var(--border-color)', paddingTop: '10px', marginTop: 'auto' }}>
                      <span style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                        <Calendar size={12} /> {doc.date}
                      </span>
                      <span>{formatFileSize(doc.fileSize)}</span>
                    </div>

                    {/* Action buttons */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', marginTop: '4px' }}>
                      <div style={{ display: 'flex', gap: '8px' }}>
                        <button
                          onClick={() => handlePreview(doc)}
                          className="btn btn-primary"
                          style={{ flex: 1, padding: '8px 10px', fontSize: '0.8rem', gap: '4px', backgroundColor: 'var(--primary)' }}
                        >
                          <Eye size={14} /> Detay & Önizleme
                        </button>
                        <button
                          onClick={() => handleDirectDownloadDocExcel(doc)}
                          className="btn"
                          style={{
                            padding: '8px 12px',
                            fontSize: '0.8rem',
                            gap: '4px',
                            backgroundColor: '#10b981',
                            color: '#ffffff',
                            borderRadius: '6px',
                            fontWeight: 600,
                            border: 'none',
                            cursor: 'pointer'
                          }}
                          title="Doğrudan Excel (.xls) Olarak İndir"
                        >
                          <FileSpreadsheet size={14} /> Excel İndir
                        </button>
                      </div>
                      <div style={{ display: 'flex', gap: '8px' }}>
                        <button
                          onClick={() => handleDownload(doc)}
                          className="btn btn-secondary"
                          style={{ flex: 1, padding: '6px 10px', fontSize: '0.75rem', gap: '4px' }}
                          title="Orijinal PDF Dosyasını İndir"
                        >
                          <Download size={13} /> Orijinal PDF
                        </button>
                        <button
                          onClick={() => handleDelete(doc.id)}
                          className="btn btn-danger btn-icon"
                          title="Sil"
                          style={{
                            backgroundColor: 'transparent',
                            borderColor: 'var(--border-color)',
                            color: 'var(--danger)',
                            padding: '6px 10px'
                          }}
                        >
                          <Trash2 size={14} />
                        </button>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            )}

            {/* Pagination */}
            {totalPages > 1 && (
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '16px', marginTop: '12px' }}>
                <button
                  className="btn btn-secondary"
                  disabled={currentPage === 1}
                  onClick={() => setCurrentPage(prev => Math.max(prev - 1, 1))}
                  style={{ padding: '8px 14px' }}
                >
                  <ChevronLeft size={16} /> Önceki
                </button>
                <span style={{ fontSize: '0.85rem', fontWeight: 600 }}>
                  {currentPage} / {totalPages}
                </span>
                <button
                  className="btn btn-secondary"
                  disabled={currentPage === totalPages}
                  onClick={() => setCurrentPage(prev => Math.min(prev + 1, totalPages))}
                  style={{ padding: '8px 14px' }}
                >
                  Sonraki <ChevronRight size={16} />
                </button>
              </div>
            )}
          </div>

          {/* SIDE PREVIEW DRAWER (600px Wide) */}
          {previewDoc && previewUrl && (
            <div className="glass-card animate-fade-in" style={{
              padding: '24px',
              backgroundColor: 'var(--bg-secondary)',
              position: 'sticky',
              top: '40px',
              height: 'calc(100vh - 80px)',
              display: 'flex',
              flexDirection: 'column',
              gap: '16px',
              boxShadow: 'var(--shadow-lg)'
            }}>
              {/* Preview Header */}
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'start', borderBottom: '1px solid var(--border-color)', paddingBottom: '12px' }}>
                <div style={{ minWidth: 0 }}>
                  <span className={`badge ${getDocTypeBadgeClass(previewDoc.docType)}`} style={{ marginBottom: '6px' }}>
                    {previewDoc.docType}
                  </span>
                  <h3 className="text-h3" style={{ fontSize: '1.05rem', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={previewDoc.name}>
                    {previewDoc.name}
                  </h3>
                </div>
                <button
                  onClick={closePreview}
                  style={{
                    background: 'transparent',
                    border: 'none',
                    color: 'var(--text-secondary)',
                    cursor: 'pointer',
                    padding: '4px',
                    borderRadius: '50%',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center'
                  }}
                  className="btn-secondary"
                >
                  <X size={18} />
                </button>
              </div>

              {/* Tab Selector */}
              <div style={{ display: 'flex', gap: '8px', borderBottom: '1px solid var(--border-color)', paddingBottom: '10px' }}>
                <button
                  onClick={() => setActivePreviewTab('info')}
                  className={`btn ${activePreviewTab === 'info' ? 'btn-primary' : 'btn-secondary'}`}
                  style={{ padding: '6px 12px', fontSize: '0.8rem', borderRadius: '6px' }}
                >
                  Genel Bilgiler
                </button>
                <button
                  onClick={() => setActivePreviewTab('preview')}
                  className={`btn ${activePreviewTab === 'preview' ? 'btn-primary' : 'btn-secondary'}`}
                  style={{ padding: '6px 12px', fontSize: '0.8rem', borderRadius: '6px' }}
                >
                  Önizleme
                </button>
                {docRows.length > 0 && (
                  <button
                    onClick={() => setActivePreviewTab('table')}
                    className={`btn ${activePreviewTab === 'table' ? 'btn-primary' : 'btn-secondary'}`}
                    style={{ padding: '6px 12px', fontSize: '0.8rem', borderRadius: '6px', gap: '6px' }}
                  >
                    <Sheet size={14} /> Tablo Verileri
                  </button>
                )}
              </div>

              {/* Tab Contents */}
              <div style={{ flex: 1, minHeight: 0, display: 'flex', flexDirection: 'column' }}>
                
                {/* 1. GENERAL INFORMATION TAB */}
                {activePreviewTab === 'info' && (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '16px', overflowY: 'auto' }}>
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', fontSize: '0.85rem' }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', paddingBottom: '8px', borderBottom: '1px dashed var(--border-color)' }}>
                        <span style={{ color: 'var(--text-muted)' }}>Dosya Adı:</span>
                        <span style={{ fontWeight: 600, wordBreak: 'break-all', textAlign: 'right', maxWidth: '75%' }}>{previewDoc.fileName}</span>
                      </div>
                      <div style={{ display: 'flex', justifyContent: 'space-between', paddingBottom: '8px', borderBottom: '1px dashed var(--border-color)' }}>
                        <span style={{ color: 'var(--text-muted)' }}>Belge Türü:</span>
                        <span style={{ fontWeight: 600 }}>{previewDoc.docType}</span>
                      </div>
                      <div style={{ display: 'flex', justifyContent: 'space-between', paddingBottom: '8px', borderBottom: '1px dashed var(--border-color)' }}>
                        <span style={{ color: 'var(--text-muted)' }}>Kategori:</span>
                        <span style={{ fontWeight: 600 }}>{previewDoc.category}</span>
                      </div>
                      <div style={{ display: 'flex', justifyContent: 'space-between', paddingBottom: '8px', borderBottom: '1px dashed var(--border-color)' }}>
                        <span style={{ color: 'var(--text-muted)' }}>Belge Tarihi:</span>
                        <span style={{ fontWeight: 600 }}>{previewDoc.date}</span>
                      </div>
                      <div style={{ display: 'flex', justifyContent: 'space-between', paddingBottom: '8px', borderBottom: '1px dashed var(--border-color)' }}>
                        <span style={{ color: 'var(--text-muted)' }}>Dosya Boyutu:</span>
                        <span style={{ fontWeight: 600 }}>{formatFileSize(previewDoc.fileSize)}</span>
                      </div>
                      <div style={{ display: 'flex', justifyContent: 'space-between', paddingBottom: '8px', borderBottom: '1px dashed var(--border-color)' }}>
                        <span style={{ color: 'var(--text-muted)' }}>Arşivlenme Tarihi:</span>
                        <span style={{ fontWeight: 600 }}>{previewDoc.uploadedAt}</span>
                      </div>
                    </div>

                    {previewDoc.description && (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                        <span style={{ color: 'var(--text-muted)', fontSize: '0.85rem', fontWeight: 600 }}>Açıklama & Notlar:</span>
                        <p style={{ padding: '12px', backgroundColor: 'var(--bg-primary)', borderRadius: '8px', fontSize: '0.85rem', color: 'var(--text-secondary)', lineHeight: '145%' }}>
                          {previewDoc.description}
                        </p>
                      </div>
                    )}

                    {docRows.length > 0 && (
                      <div className="glass-card" style={{ padding: '14px 18px', backgroundColor: 'var(--primary-light)', color: 'var(--primary)', borderRadius: '8px', display: 'flex', alignItems: 'center', gap: '10px', fontSize: '0.8rem', fontWeight: 600 }}>
                        <Sheet size={18} />
                        <span>Bu belge {docRows.length} satır çözümlenmiş tablo verisi barındırıyor. "Tablo Verileri" sekmesinden erişebilirsiniz.</span>
                      </div>
                    )}

                    <div style={{ display: 'flex', gap: '12px', marginTop: 'auto' }}>
                      <button className="btn btn-primary" onClick={() => handleDownload(previewDoc)} style={{ flex: 1 }}>
                        <Download size={16} /> Dosyayı İndir
                      </button>
                      <a
                        href={previewUrl}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="btn btn-secondary"
                        style={{ flex: 1, gap: '6px' }}
                      >
                        <ExternalLink size={16} /> Yeni Sekmede Aç
                      </a>
                    </div>
                  </div>
                )}

                {/* 2. FILE PREVIEW TAB */}
                {activePreviewTab === 'preview' && (
                  <div style={{ flex: 1, border: '1px solid var(--border-color)', borderRadius: '8px', overflow: 'hidden', backgroundColor: 'var(--bg-primary)', display: 'flex', alignItems: 'center', justifyContent: 'center', position: 'relative' }}>
                    {previewDoc.fileType === 'PDF' ? (
                      <iframe
                        src={`${previewUrl}#toolbar=0`}
                        title="PDF Önizleme"
                        width="100%"
                        height="100%"
                        style={{ border: 'none' }}
                      />
                    ) : ['PNG', 'JPG', 'JPEG', 'GIF', 'SVG'].includes(previewDoc.fileType) ? (
                      <img
                        src={previewUrl}
                        alt="Belge Önizleme"
                        style={{ maxWidth: '100%', maxHeight: '100%', objectFit: 'contain' }}
                      />
                    ) : (
                      <div style={{ textAlign: 'center', padding: '24px', display: 'flex', flexDirection: 'column', gap: '14px', alignItems: 'center' }}>
                        <FileText size={48} style={{ color: 'var(--text-muted)' }} />
                        <div>
                          <p style={{ fontWeight: 600, fontSize: '0.9rem' }}>Önizleme Desteklenmiyor</p>
                          <p style={{ fontSize: '0.75rem', color: 'var(--text-muted)', marginTop: '4px', maxWidth: '240px' }}>
                            .{previewDoc.fileType.toLowerCase()} dosyaları tarayıcı içinde doğrudan görüntülenemez.
                          </p>
                        </div>
                        <button className="btn btn-primary" onClick={() => handleDownload(previewDoc)} style={{ fontSize: '0.8rem', padding: '8px 16px' }}>
                          <Download size={14} /> Dosyayı İndir
                        </button>
                      </div>
                    )}
                  </div>
                )}

                {/* 3. STRUCTURED TABLE DATA TAB (Excel Copying & Calculations!) */}
                {activePreviewTab === 'table' && (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '16px', flex: 1, minHeight: 0 }}>
                    {/* Sum Statistics Row */}
                    {columnStats.length > 0 && (
                      <div style={{ display: 'flex', gap: '10px', overflowX: 'auto', paddingBottom: '6px' }}>
                        {columnStats.map(stat => (
                          <div
                            key={stat.key}
                            className="glass-card"
                            style={{
                              padding: '10px 14px',
                              backgroundColor: 'var(--bg-primary)',
                              border: '1px solid var(--border-color)',
                              borderRadius: '8px',
                              flexShrink: 0,
                              minWidth: '130px'
                            }}
                          >
                            <span style={{ fontSize: '0.7rem', color: 'var(--text-muted)', fontWeight: 700, display: 'block', textTransform: 'uppercase' }}>
                              Toplam {stat.key}
                            </span>
                            <span style={{ fontSize: '0.95rem', fontWeight: 800, color: 'var(--primary)', marginTop: '2px', display: 'block' }}>
                              {stat.sum.toLocaleString('tr-TR', { minimumFractionDigits: 2 })}
                            </span>
                          </div>
                        ))}
                      </div>
                    )}

                    {/* Table search & Excel Copy Bar */}
                    <div style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
                      <div style={{ position: 'relative', flex: 1 }}>
                        <Search size={14} style={{ position: 'absolute', left: '10px', top: '50%', transform: 'translateY(-50%)', color: 'var(--text-muted)' }} />
                        <input
                          type="text"
                          placeholder="Tablo içinde ara (örn: TC, Ad, Soyad)..."
                          className="form-input"
                          style={{ paddingLeft: '32px', height: '36px', fontSize: '0.8rem' }}
                          value={rowSearchTerm}
                          onChange={(e) => setRowSearchTerm(e.target.value)}
                        />
                      </div>
                      
                      {/* Copy to Clipboard (Excel compatible) */}
                      <button
                        onClick={handleCopyToClipboard}
                        className={`btn ${copySuccess ? 'btn-primary' : 'btn-secondary'}`}
                        style={{ height: '36px', padding: '0 12px', gap: '6px', fontSize: '0.8rem', whiteSpace: 'nowrap' }}
                        title={selectionActive ? "Seçilen hücreleri kopyala" : "Tüm tabloyu kopyala"}
                      >
                        {copySuccess ? <Check size={14} style={{ color: '#ffffff' }} /> : <Clipboard size={14} />}
                        {copySuccess ? 'Kopyalandı!' : selectionActive ? 'Seçileni Excel İçin Kopyala' : 'Tümünü Excel İçin Kopyala'}
                      </button>

                      {/* Download Excel (.xlsx) */}
                      <button
                        onClick={handleDownloadExcel}
                        className="btn btn-success"
                        style={{
                          height: '36px',
                          padding: '0 12px',
                          display: 'flex',
                          alignItems: 'center',
                          gap: '6px',
                          fontSize: '0.8rem',
                          fontWeight: 600,
                          backgroundColor: '#10b981',
                          color: '#ffffff',
                          border: 'none',
                          borderRadius: '6px',
                          cursor: 'pointer',
                          whiteSpace: 'nowrap'
                        }}
                        title="Excel dosyası (.xls) olarak indir"
                      >
                        <Download size={14} /> Excel İndir (.xls)
                      </button>

                      {/* Download CSV */}
                      <button
                        onClick={handleDownloadCSV}
                        className="btn btn-secondary btn-icon"
                        style={{ height: '36px', width: '36px' }}
                        title="CSV olarak indir"
                      >
                        <Download size={14} />
                      </button>

                      {/* Open in New Tab */}
                      <button
                        onClick={handleOpenTableInNewTab}
                        className="btn btn-secondary btn-icon"
                        style={{ height: '36px', width: '36px' }}
                        title="Tabloyu yeni sekmede tam ekran aç"
                      >
                        <ExternalLink size={14} />
                      </button>
                    </div>

                    {selectionActive && (
                      <div style={{
                        fontSize: '0.75rem',
                        color: 'var(--primary)',
                        backgroundColor: 'var(--primary-light)',
                        padding: '6px 12px',
                        borderRadius: '6px',
                        display: 'flex',
                        justifyContent: 'space-between',
                        alignItems: 'center'
                      }}>
                        <span>
                          🔲 <strong>Hücre Seçimi Aktif:</strong> Excel gibi kopyalamak için sürükleyip hücreleri seçin, ardından <strong>Ctrl+C</strong> tuşlarına basın veya yukarıdaki butona tıklayın.
                        </span>
                        <button
                          onClick={clearSelection}
                          style={{
                            background: 'transparent',
                            border: 'none',
                            color: 'var(--primary)',
                            cursor: 'pointer',
                            fontWeight: 700,
                            fontSize: '0.75rem'
                          }}
                        >
                          Seçimi Temizle
                        </button>
                      </div>
                    )}

                    {copySuccess && (
                      <div style={{
                        fontSize: '0.75rem',
                        color: 'var(--success)',
                        fontWeight: 600,
                        backgroundColor: 'var(--success-light)',
                        padding: '6px 12px',
                        borderRadius: '6px'
                      }}>
                        ✓ {selectionActive ? 'Seçilen hücreler' : 'Tüm tablo'} panoya kopyalandı! Excel'e gidip doğrudan <strong>Yapıştır (Ctrl+V)</strong> yapabilirsiniz.
                      </div>
                    )}

                    {/* Table Grid (Auto generated headers based on columns) */}
                    <div className="table-container" style={{ flex: 1, minHeight: 0, overflowY: 'auto' }}>
                      <table className="custom-table" style={{ fontSize: '0.78rem' }}>
                        <thead>
                          <tr style={{ position: 'sticky', top: 0, zIndex: 1 }}>
                            <th style={{ width: '40px', background: 'var(--bg-tertiary)', borderBottom: '2px solid var(--border-color)', userSelect: 'none' }}>#</th>
                            {spreadsheetHeaders.map((h, colIdx) => (
                              <th
                                key={h}
                                onClick={() => handleColumnHeaderClick(colIdx)}
                                className="selectable-header"
                                style={{ background: 'var(--bg-tertiary)', borderBottom: '2px solid var(--border-color)' }}
                                title={`${h} sütununun tamamını seçmek için tıklayın`}
                              >
                                {h}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {filteredDocRows.map((row, rowIdx) => (
                            <tr key={row.id}>
                              <td
                                onClick={() => handleRowSelect(rowIdx)}
                                className="selectable-header"
                                style={{ fontWeight: 600, color: 'var(--text-muted)', textAlign: 'center' }}
                                title={`${row.rowNumber}. satırın tamamını seçmek için tıklayın`}
                              >
                                {row.rowNumber}
                              </td>
                              {spreadsheetHeaders.map((h, colIdx) => {
                                const cellVal = row.data[h];
                                const isSelected = isCellSelected(rowIdx, colIdx);
                                return (
                                  <td
                                    key={h}
                                    className={`selectable-cell ${isSelected ? 'selected-cell' : ''}`}
                                    onMouseDown={() => handleCellMouseDown(rowIdx, colIdx)}
                                    onMouseEnter={() => handleCellMouseEnter(rowIdx, colIdx)}
                                    onMouseUp={handleCellMouseUp}
                                    style={{
                                      fontWeight: (typeof cellVal === 'number' && 
                                        !(h.toLowerCase().includes('tc') || h.toLowerCase().includes('kimlik') || h.toLowerCase().includes('sicil') || h.toLowerCase().includes('no') || h.toLowerCase().includes('tckn'))) ? 700 : 400,
                                      textAlign: (typeof cellVal === 'number' && 
                                        !(h.toLowerCase().includes('tc') || h.toLowerCase().includes('kimlik') || h.toLowerCase().includes('sicil') || h.toLowerCase().includes('no') || h.toLowerCase().includes('tckn'))) ? 'right' : 'left'
                                    }}
                                  >
                                    {formatCellValue(cellVal, h)}
                                  </td>
                                );
                              })}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>

                    <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.75rem', color: 'var(--text-muted)' }}>
                      <span>Toplam {docRows.length} satırdan {filteredDocRows.length} tanesi gösteriliyor</span>
                      <span>Seçmek için tıklayıp sürükleyin. Seçtikten sonra Ctrl+C tuşlarına basın.</span>
                    </div>
                  </div>
                )}

              </div>
            </div>
          )}

        </div>
      )}
    </div>
  );
};
export default DocumentExplorer;
