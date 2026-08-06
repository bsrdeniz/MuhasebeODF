import React, { useState, useRef } from 'react';
import { UploadCloud, CheckCircle2, FileText, AlertTriangle, Calendar, Tag, AlignLeft } from 'lucide-react';
import * as XLSX from 'xlsx';
import { saveDocumentAndRows, type DocumentItem, type DocumentRow } from '../db';
import { autoExtractMetadata, parsePdfTableRows } from '../utils/extractor';

interface DocumentUploaderProps {
  onUploadSuccess: () => void;
}

const DOCUMENT_TYPES = [
  'Maaş Bordrosu',
  'TOBB Belgesi',
  'Fatura',
  'Sözleşme',
  'Diğer'
];

export const DocumentUploader: React.FC<DocumentUploaderProps> = ({ onUploadSuccess }) => {
  const [dragActive, setDragActive] = useState(false);
  const [file, setFile] = useState<File | null>(null);
  const [isExtracting, setIsExtracting] = useState(false);
  const [spreadsheetRows, setSpreadsheetRows] = useState<Record<string, any>[]>([]);
  
  // Form States
  const [docName, setDocName] = useState('');
  const [docType, setDocType] = useState('TOBB Belgesi');
  const [docDate, setDocDate] = useState(new Date().toISOString().split('T')[0]);
  const [category, setCategory] = useState('Genel');
  const [description, setDescription] = useState('');
  
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  // Helper to normalize strings that look like numbers into floats
  const normalizeNumber = (val: any) => {
    if (typeof val === 'string' && val.trim() !== '') {
      const normalized = val.replace(/\s/g, ''); // strip spaces
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

        if (!isNaN(num)) {
          return num;
        }
      }
    }
    return val;
  };

  const processFile = async (selectedFile: File) => {
    setFile(selectedFile);
    setIsExtracting(true);
    setError(null);
    setSpreadsheetRows([]);

    try {
      // 1. Extract file metadata using helper
      const meta = await autoExtractMetadata(selectedFile);
      setDocName(meta.name);
      setDocType(meta.docType);
      setDocDate(meta.date);
      setCategory(meta.category);
      setDescription(meta.description);

      // 2. Tabular row parsing
      const fileExt = selectedFile.name.split('.').pop()?.toLowerCase();
      
      if (fileExt === 'pdf') {
        try {
          const pdfRows = await parsePdfTableRows(selectedFile);
          
          // Clean/normalize numerical values in PDF rows (like P.E.K, T.P, T.K, etc.)
          const cleaned = pdfRows.map(row => {
            const cleanRow: Record<string, any> = {};
            Object.entries(row).forEach(([key, val]) => {
              cleanRow[key] = normalizeNumber(val);
            });
            return cleanRow;
          });
          
          setSpreadsheetRows(cleaned);
        } catch (pdfErr) {
          console.error('PDF tablosu ayrıştırılamadı:', pdfErr);
        }
      } else if (['ods', 'xlsx', 'xls'].includes(fileExt || '')) {
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const data = new Uint8Array(e.target?.result as ArrayBuffer);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const rows = XLSX.utils.sheet_to_json<Record<string, any>>(worksheet, { defval: '' });
            
            // Clean headers and normalize numerical values
            const cleaned = rows.map(row => {
              const cleanRow: Record<string, any> = {};
              Object.entries(row).forEach(([key, val]) => {
                const cleanKey = String(key).trim();
                cleanRow[cleanKey] = normalizeNumber(val);
              });
              return cleanRow;
            });

            setSpreadsheetRows(cleaned);
          } catch (excelErr) {
            console.error('Spreadsheet satırları ayrıştırılamadı:', excelErr);
          }
        };
        reader.readAsArrayBuffer(selectedFile);
      }
    } catch (err: any) {
      console.error(err);
      // Fallback
      const lastDotIndex = selectedFile.name.lastIndexOf('.');
      const nameWithoutExt = lastDotIndex !== -1 
        ? selectedFile.name.substring(0, lastDotIndex) 
        : selectedFile.name;
      
      setDocName(nameWithoutExt);
      setDocType('TOBB Belgesi');
      setCategory('Genel');
    } finally {
      setIsExtracting(false);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);

    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      processFile(e.dataTransfer.files[0]);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      processFile(e.target.files[0]);
    }
  };

  const handleSave = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!file) return;

    setIsProcessing(true);
    setError(null);

    try {
      const docId = 'doc_' + Date.now() + '_' + Math.random().toString(36).substring(2, 9);
      
      const newDoc: DocumentItem = {
        id: docId,
        name: docName || file.name,
        fileName: file.name,
        fileSize: file.size,
        fileType: file.name.split('.').pop()?.toUpperCase() || 'UNKNOWN',
        docType,
        date: docDate,
        category: category || 'Genel',
        description,
        fileData: file, // Save File/Blob directly
        uploadedAt: new Date().toLocaleString('tr-TR')
      };

      const docRows: DocumentRow[] = spreadsheetRows.map((row, index) => ({
        id: `row_${docId}_${index}_${Math.random().toString(36).substring(2, 9)}`,
        docId,
        rowNumber: index + 1,
        data: row
      }));

      await saveDocumentAndRows(newDoc, docRows);
      onUploadSuccess();
    } catch (err: any) {
      setError('Belge kaydedilirken bir hata oluştu: ' + err.message);
    } finally {
      setIsProcessing(false);
    }
  };

  const resetUploader = () => {
    setFile(null);
    setDocName('');
    setDocType('TOBB Belgesi');
    setDocDate(new Date().toISOString().split('T')[0]);
    setCategory('Genel');
    setDescription('');
    setSpreadsheetRows([]);
    setError(null);
  };

  return (
    <div className="animate-fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
      <div>
        <h1 className="text-h1" style={{ marginBottom: '8px' }}>Yeni Belge Arşivle</h1>
        <p style={{ color: 'var(--text-secondary)' }}>
          TOBB'dan gelen yazıları, maaş bordrolarını veya diğer muhasebe evraklarını sisteme güvenli şekilde yükleyin.
        </p>
      </div>

      {error && (
        <div className="glass-card" style={{
          display: 'flex',
          alignItems: 'center',
          gap: '12px',
          padding: '16px 20px',
          backgroundColor: 'var(--danger-light)',
          borderColor: 'var(--danger)',
          color: 'var(--danger)',
          borderRadius: '10px'
        }}>
          <AlertTriangle size={20} />
          <span style={{ fontSize: '0.9rem', fontWeight: 600 }}>{error}</span>
        </div>
      )}

      {!file ? (
        /* DRAG AND DROP ZONE */
        <div
          onDragEnter={handleDrag}
          onDragOver={handleDrag}
          onDragLeave={handleDrag}
          onDrop={handleDrop}
          onClick={() => fileInputRef.current?.click()}
          style={{
            border: '2px dashed ' + (dragActive ? 'var(--primary)' : 'var(--border-color)'),
            borderRadius: '16px',
            padding: '80px 40px',
            textAlign: 'center',
            cursor: 'pointer',
            backgroundColor: dragActive ? 'var(--primary-light)' : 'var(--bg-secondary)',
            transition: 'all 0.2s ease',
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center',
            gap: '16px'
          }}
        >
          <input
            ref={fileInputRef}
            type="file"
            onChange={handleFileChange}
            style={{ display: 'none' }}
          />

          <div style={{
            width: '64px',
            height: '64px',
            borderRadius: '50%',
            backgroundColor: 'var(--primary-light)',
            color: 'var(--primary)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            marginBottom: '8px'
          }}>
            <UploadCloud size={32} />
          </div>

          <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
            <h3 className="text-h3">Yüklemek istediğiniz belgeyi sürükleyin</h3>
            <p style={{ color: 'var(--text-muted)', fontSize: '0.85rem' }}>
              veya bilgisayarınızdan seçmek için tıklayın (PDF, Excel, Görsel, Word vb.)
            </p>
          </div>
        </div>
      ) : isExtracting ? (
        /* AUTOMATIC AI METADATA EXTRACTION LOADING LOADER */
        <div className="glass-card animate-fade-in" style={{
          padding: '60px 40px',
          textAlign: 'center',
          backgroundColor: 'var(--bg-secondary)',
          display: 'flex',
          flexDirection: 'column',
          alignItems: 'center',
          justifyContent: 'center',
          gap: '20px',
          borderRadius: '16px'
        }}>
          <div style={{
            width: '48px',
            height: '48px',
            borderRadius: '50%',
            border: '4px solid var(--primary-light)',
            borderTopColor: 'var(--primary)',
            animation: 'spin 1s linear infinite'
          }} />
          <style>{`
            @keyframes spin {
              0% { transform: rotate(0deg); }
              100% { transform: rotate(360deg); }
            }
          `}</style>
          <div>
            <h3 className="text-h3" style={{ marginBottom: '6px' }}>Belge Çözümleniyor...</h3>
            <p style={{ color: 'var(--text-muted)', fontSize: '0.85rem', maxWidth: '380px', margin: '0 auto' }}>
              Belge içeriği taranıyor ve Başlık, Belge Türü, Kategori, Tarih ve Tablo Verileri otomatik olarak çıkartılıyor.
            </p>
          </div>
        </div>
      ) : (
        /* METADATA ENTRY FORM */
        <div className="glass-card" style={{ padding: '28px', backgroundColor: 'var(--bg-secondary)' }}>
          {/* Smart Extract Success Indicator */}
          <div style={{
            display: 'flex',
            alignItems: 'center',
            gap: '8px',
            padding: '12px 16px',
            backgroundColor: 'var(--success-light)',
            color: 'var(--success)',
            borderRadius: '10px',
            fontSize: '0.85rem',
            fontWeight: 600,
            marginBottom: '20px'
          }}>
            <CheckCircle2 size={18} />
            Belge içeriği başarıyla analiz edildi ve form alanları otomatik dolduruldu! {spreadsheetRows.length > 0 && `(${spreadsheetRows.length} Tablo Satırı Çıkartıldı)`}
          </div>

          <form onSubmit={handleSave} style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid var(--border-color)', paddingBottom: '16px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                <div style={{
                  backgroundColor: 'var(--primary-light)',
                  color: 'var(--primary)',
                  width: '40px',
                  height: '40px',
                  borderRadius: '8px',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center'
                }}>
                  <FileText size={20} />
                </div>
                <div>
                  <h4 style={{ fontWeight: 700, color: 'var(--text-primary)', fontSize: '0.95rem', margin: 0 }}>
                    {file.name}
                  </h4>
                  <span style={{ fontSize: '0.75rem', color: 'var(--text-muted)' }}>
                    {(file.size / (1024 * 1024)).toFixed(2)} MB • {file.type || 'Bilinmeyen Format'}
                  </span>
                </div>
              </div>
              <button type="button" className="btn btn-secondary" onClick={resetUploader} disabled={isProcessing}>
                Değiştir
              </button>
            </div>

            {/* Inputs Grid */}
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
              {/* Document Name */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                <label style={{ fontSize: '0.8rem', fontWeight: 700, color: 'var(--text-secondary)' }}>
                  Belge Adı *
                </label>
                <input
                  type="text"
                  required
                  className="form-input"
                  placeholder="Belgeye açıklayıcı bir isim verin"
                  value={docName}
                  onChange={(e) => setDocName(e.target.value)}
                  disabled={isProcessing}
                />
              </div>

              {/* Document Type */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                <label style={{ fontSize: '0.8rem', fontWeight: 700, color: 'var(--text-secondary)' }}>
                  Belge Türü *
                </label>
                <select
                  className="form-input"
                  value={docType}
                  onChange={(e) => setDocType(e.target.value)}
                  disabled={isProcessing}
                >
                  {DOCUMENT_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
                </select>
              </div>

              {/* Document Date */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                <label style={{ fontSize: '0.8rem', fontWeight: 700, color: 'var(--text-secondary)', display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <Calendar size={14} /> Belge Tarihi *
                </label>
                <input
                  type="date"
                  required
                  className="form-input"
                  value={docDate}
                  onChange={(e) => setDocDate(e.target.value)}
                  disabled={isProcessing}
                />
              </div>

              {/* Category */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                <label style={{ fontSize: '0.8rem', fontWeight: 700, color: 'var(--text-secondary)', display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <Tag size={14} /> Kategori *
                </label>
                <input
                  type="text"
                  required
                  className="form-input"
                  placeholder="Personel, Finans, Kararlar vb."
                  value={category}
                  onChange={(e) => setCategory(e.target.value)}
                  disabled={isProcessing}
                />
              </div>

              {/* Description */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: '6px', gridColumn: 'span 2' }}>
                <label style={{ fontSize: '0.8rem', fontWeight: 700, color: 'var(--text-secondary)', display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <AlignLeft size={14} /> Açıklama / Notlar
                </label>
                <textarea
                  className="form-input"
                  style={{ minHeight: '80px', resize: 'vertical' }}
                  placeholder="Belgeyle ilgili ek açıklamalar, notlar ekleyin..."
                  value={description}
                  onChange={(e) => setDescription(e.target.value)}
                  disabled={isProcessing}
                />
              </div>
            </div>

            {/* Form Actions */}
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '12px', marginTop: '10px' }}>
              <button type="button" className="btn btn-secondary" onClick={resetUploader} disabled={isProcessing}>
                Vazgeç
              </button>
              <button
                type="submit"
                className="btn btn-primary"
                disabled={isProcessing || !docName || !category}
                style={{ gap: '8px' }}
              >
                {isProcessing ? 'Belge Kaydediliyor...' : 'Belgeyi Sisteme Arşivle'}
                <CheckCircle2 size={16} />
              </button>
            </div>
          </form>
        </div>
      )}
    </div>
  );
};
export default DocumentUploader;
