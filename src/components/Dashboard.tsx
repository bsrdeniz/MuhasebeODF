import React, { useState, useEffect } from 'react';
import { FileText, UploadCloud, FolderOpen, ArrowUpRight, ShieldCheck, Tag } from 'lucide-react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Cell } from 'recharts';
import { getDocuments, type DocumentItem } from '../db';

interface DashboardProps {
  setActiveTab: (tab: string) => void;
}

export const Dashboard: React.FC<DashboardProps> = ({ setActiveTab }) => {
  const [documents, setDocuments] = useState<DocumentItem[]>([]);
  const [isLoading, setIsLoading] = useState(true);

  useEffect(() => {
    const loadData = async () => {
      try {
        const docs = await getDocuments();
        setDocuments(docs);
      } catch (error) {
        console.error('Gösterge paneli verileri yüklenirken hata:', error);
      } finally {
        setIsLoading(false);
      }
    };
    loadData();
  }, []);

  const totalDocuments = documents.length;
  const maasBordrosuCount = documents.filter(d => d.docType === 'Maaş Bordrosu').length;
  const tobbBelgesiCount = documents.filter(d => d.docType === 'TOBB Belgesi').length;
  const otherCount = totalDocuments - maasBordrosuCount - tobbBelgesiCount;

  // Process data: Document Type counts
  const getTypeChartData = () => {
    const types = ['Maaş Bordrosu', 'TOBB Belgesi', 'Fatura', 'Sözleşme', 'Diğer'];
    return types.map(t => ({
      name: t,
      'Belge Sayısı': documents.filter(d => d.docType === t).length
    }));
  };

  // Process data: Top Categories
  const getCategoryData = () => {
    const categories: Record<string, number> = {};
    documents.forEach(d => {
      categories[d.category] = (categories[d.category] || 0) + 1;
    });

    return Object.entries(categories)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 5);
  };

  const chartData = getTypeChartData();
  const categoryData = getCategoryData();

  const COLORS = ['#10b981', '#06b6d4', '#f59e0b', '#ef4444', '#6366f1'];

  if (isLoading) {
    return (
      <div style={{ textAlign: 'center', padding: '40px', color: 'var(--text-muted)' }}>
        Yükleniyor...
      </div>
    );
  }

  if (totalDocuments === 0) {
    return (
      <div className="animate-fade-in" style={{
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        gap: '24px',
        padding: '80px 20px',
        textAlign: 'center'
      }}>
        <div style={{
          width: '80px',
          height: '80px',
          borderRadius: '50%',
          backgroundColor: 'var(--primary-light)',
          color: 'var(--primary)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center'
        }}>
          <UploadCloud size={40} />
        </div>
        <div>
          <h2 className="text-h2" style={{ marginBottom: '8px' }}>Arşivde Belge Bulunmamaktadır</h2>
          <p style={{ color: 'var(--text-secondary)', maxWidth: '450px', margin: '0 auto' }}>
            TOBB belgelerini, maaş bordrolarını veya muhasebe evraklarını saklamak ve hızlıca erişmek için belge yüklemeye başlayın.
          </p>
        </div>
        <button className="btn btn-primary" onClick={() => setActiveTab('upload')} style={{ padding: '12px 24px', gap: '10px' }}>
          Yeni Belge Arşivle
          <ArrowUpRight size={18} />
        </button>
      </div>
    );
  }

  return (
    <div className="animate-fade-in" style={{ display: 'flex', flexDirection: 'column', gap: '28px' }}>
      {/* Header */}
      <div>
        <h1 className="text-h1" style={{ marginBottom: '8px' }}>Gösterge Paneli</h1>
        <p style={{ color: 'var(--text-secondary)' }}>
          Muhasebe birimi belge arşiv durumunu, kategorik kırılımları ve son belgeleri izleyin.
        </p>
      </div>

      {/* KPI Grid */}
      <div className="dashboard-grid">
        {/* Card 1: Total Docs */}
        <div className="glass-card" style={{ padding: '24px', display: 'flex', flexDirection: 'column', gap: '14px', position: 'relative', overflow: 'hidden' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <span style={{ fontSize: '0.85rem', fontWeight: 700, color: 'var(--text-secondary)' }}>Toplam Belge</span>
            <div style={{ padding: '8px', borderRadius: '10px', backgroundColor: 'var(--primary-light)', color: 'var(--primary)' }}>
              <FolderOpen size={20} />
            </div>
          </div>
          <div>
            <h2 className="text-h2" style={{ fontSize: '1.75rem', fontWeight: 800 }}>
              {totalDocuments} Adet
            </h2>
            <span style={{ fontSize: '0.75rem', color: 'var(--text-muted)', fontWeight: 500 }}>Sistemde arşivlenen toplam evrak</span>
          </div>
          <div style={{ position: 'absolute', bottom: 0, left: 0, right: 0, height: '4px', backgroundColor: 'var(--primary)' }} />
        </div>

        {/* Card 2: Salary Payrolls */}
        <div className="glass-card" style={{ padding: '24px', display: 'flex', flexDirection: 'column', gap: '14px', position: 'relative', overflow: 'hidden' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <span style={{ fontSize: '0.85rem', fontWeight: 700, color: 'var(--text-secondary)' }}>Maaş Bordroları</span>
            <div style={{ padding: '8px', borderRadius: '10px', backgroundColor: 'var(--success-light)', color: 'var(--success)' }}>
              <ShieldCheck size={20} />
            </div>
          </div>
          <div>
            <h2 className="text-h2" style={{ fontSize: '1.75rem', fontWeight: 800 }}>
              {maasBordrosuCount} Adet
            </h2>
            <span style={{ fontSize: '0.75rem', color: 'var(--text-muted)', fontWeight: 500 }}>Personel maaş & bordro evrakları</span>
          </div>
          <div style={{ position: 'absolute', bottom: 0, left: 0, right: 0, height: '4px', backgroundColor: 'var(--success)' }} />
        </div>

        {/* Card 3: TOBB Docs */}
        <div className="glass-card" style={{ padding: '24px', display: 'flex', flexDirection: 'column', gap: '14px', position: 'relative', overflow: 'hidden' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <span style={{ fontSize: '0.85rem', fontWeight: 700, color: 'var(--text-secondary)' }}>TOBB Evrakları</span>
            <div style={{ padding: '8px', borderRadius: '10px', backgroundColor: 'var(--info-light)', color: 'var(--info)' }}>
              <FileText size={20} />
            </div>
          </div>
          <div>
            <h2 className="text-h2" style={{ fontSize: '1.75rem', fontWeight: 800 }}>
              {tobbBelgesiCount} Adet
            </h2>
            <span style={{ fontSize: '0.75rem', color: 'var(--text-muted)', fontWeight: 500 }}>TOBB kaynaklı resmi yazışmalar</span>
          </div>
          <div style={{ position: 'absolute', bottom: 0, left: 0, right: 0, height: '4px', backgroundColor: 'var(--info)' }} />
        </div>

        {/* Card 4: Other Docs */}
        <div className="glass-card" style={{ padding: '24px', display: 'flex', flexDirection: 'column', gap: '14px', position: 'relative', overflow: 'hidden' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <span style={{ fontSize: '0.85rem', fontWeight: 700, color: 'var(--text-secondary)' }}>Diğer Belgeler</span>
            <div style={{ padding: '8px', borderRadius: '10px', backgroundColor: 'var(--warning-light)', color: 'var(--warning)' }}>
              <Tag size={20} />
            </div>
          </div>
          <div>
            <h2 className="text-h2" style={{ fontSize: '1.75rem', fontWeight: 800 }}>
              {otherCount} Adet
            </h2>
            <span style={{ fontSize: '0.75rem', color: 'var(--text-muted)', fontWeight: 500 }}>Fatura, sözleşme ve diğer dosyalar</span>
          </div>
          <div style={{ position: 'absolute', bottom: 0, left: 0, right: 0, height: '4px', backgroundColor: 'var(--warning)' }} />
        </div>
      </div>

      {/* CHARTS ROW */}
      <div style={{ display: 'grid', gridTemplateColumns: '2fr 1fr', gap: '24px' }}>
        {/* Document Type Distribution Chart */}
        <div className="glass-card" style={{ padding: '24px', backgroundColor: 'var(--bg-secondary)', display: 'flex', flexDirection: 'column', gap: '16px' }}>
          <div>
            <h3 className="text-h3" style={{ fontSize: '1.05rem', marginBottom: '4px' }}>Belge Türü Dağılımı</h3>
            <p style={{ color: 'var(--text-muted)', fontSize: '0.8rem' }}>Arşivdeki belgelerin türlerine göre genel dağılım oranları</p>
          </div>
          <div style={{ width: '100%', height: '300px', fontSize: '0.75rem' }}>
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={chartData} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="var(--border-color)" />
                <XAxis dataKey="name" stroke="var(--text-muted)" />
                <YAxis stroke="var(--text-muted)" allowDecimals={false} />
                <Tooltip
                  contentStyle={{
                    backgroundColor: 'var(--bg-secondary)',
                    borderColor: 'var(--border-color)',
                    color: 'var(--text-primary)',
                    borderRadius: '8px'
                  }}
                />
                <Bar dataKey="Belge Sayısı" radius={[4, 4, 0, 0]}>
                  {chartData.map((_, index) => (
                    <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                  ))}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* Top Categories */}
        <div className="glass-card" style={{ padding: '24px', backgroundColor: 'var(--bg-secondary)', display: 'flex', flexDirection: 'column', gap: '16px' }}>
          <div>
            <h3 className="text-h3" style={{ fontSize: '1.05rem', marginBottom: '4px' }}>Kategori Yoğunluğu</h3>
            <p style={{ color: 'var(--text-muted)', fontSize: '0.8rem' }}>En çok evrak içeren 5 ana kategori</p>
          </div>
          <div style={{ width: '100%', height: '300px', display: 'flex', flexDirection: 'column', justifyContent: 'center' }}>
            {categoryData.length === 0 ? (
              <div style={{ display: 'flex', height: '100%', alignItems: 'center', justifyContent: 'center', color: 'var(--text-muted)' }}>
                Kategori verisi yok.
              </div>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
                {categoryData.map((item, idx) => {
                  const percent = totalDocuments > 0 ? (item.value / totalDocuments) * 100 : 0;
                  return (
                    <div key={item.name} style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '0.8rem', fontWeight: 600 }}>
                        <span style={{ color: 'var(--text-primary)' }}>{item.name}</span>
                        <span style={{ color: 'var(--text-secondary)' }}>
                          {item.value} Adet ({percent.toFixed(1)}%)
                        </span>
                      </div>
                      <div style={{ width: '100%', height: '8px', backgroundColor: 'var(--bg-tertiary)', borderRadius: '4px', overflow: 'hidden' }}>
                        <div style={{
                          width: `${percent}%`,
                          height: '100%',
                          backgroundColor: COLORS[idx % COLORS.length],
                          borderRadius: '4px'
                        }} />
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        </div>
      </div>

      {/* RECENT DOCUMENTS TABLE */}
      <div className="glass-card" style={{ padding: '24px', backgroundColor: 'var(--bg-secondary)' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
          <div>
            <h3 className="text-h3" style={{ fontSize: '1.05rem', marginBottom: '4px' }}>Son Arşivlenen Belgeler</h3>
            <p style={{ color: 'var(--text-muted)', fontSize: '0.8rem' }}>En son yüklenen 5 evrak listesi</p>
          </div>
          <button className="btn btn-secondary" onClick={() => setActiveTab('explorer')}>
            Tümünü Gör
          </button>
        </div>

        <div className="table-container">
          <table className="custom-table">
            <thead>
              <tr>
                <th>Belge Adı</th>
                <th>Dosya Adı</th>
                <th>Tür</th>
                <th>Kategori</th>
                <th>Belge Tarihi</th>
                <th style={{ textAlign: 'right' }}>Yüklenme Zamanı</th>
              </tr>
            </thead>
            <tbody>
              {documents.slice(0, 5).map((doc) => (
                <tr key={doc.id}>
                  <td style={{ fontWeight: 600, color: 'var(--primary)' }}>{doc.name}</td>
                  <td style={{ color: 'var(--text-secondary)' }}>{doc.fileName}</td>
                  <td>
                    <span className="badge badge-info">{doc.docType}</span>
                  </td>
                  <td>{doc.category}</td>
                  <td>{doc.date}</td>
                  <td style={{ textAlign: 'right' }}>{doc.uploadedAt}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};
export default Dashboard;
