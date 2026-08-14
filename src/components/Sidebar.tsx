import React from 'react';
import { LayoutDashboard, UploadCloud, Table, Sun, Moon, FileSpreadsheet, Lock } from 'lucide-react';

interface SidebarProps {
  activeTab: string;
  setActiveTab: (tab: string) => void;
  darkMode: boolean;
  setDarkMode: (dark: boolean) => void;
  onLock?: () => void;
}

export const Sidebar: React.FC<SidebarProps> = ({
  activeTab,
  setActiveTab,
  darkMode,
  setDarkMode,
  onLock,
}) => {
  const menuItems = [
    { id: 'dashboard', name: 'Gösterge Paneli', icon: LayoutDashboard },
    { id: 'upload', name: 'Belge Arşivle', icon: UploadCloud },
    { id: 'explorer', name: 'Belge Gezgini', icon: Table },
  ];

  const toggleDarkMode = () => {
    const newMode = !darkMode;
    setDarkMode(newMode);
    if (newMode) {
      document.documentElement.classList.add('dark');
    } else {
      document.documentElement.classList.remove('dark');
    }
  };

  return (
    <aside className="sidebar glass-panel">
      {/* Brand Logo & Title */}
      <div style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '40px' }}>
        <div style={{
          background: 'linear-gradient(135deg, var(--primary) 0%, #a855f7 100%)',
          width: '40px',
          height: '40px',
          borderRadius: '10px',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          color: '#ffffff',
          boxShadow: '0 4px 12px rgba(79, 70, 229, 0.3)'
        }}>
          <FileSpreadsheet size={22} />
        </div>
        <div>
          <h2 style={{ fontSize: '1.15rem', fontWeight: 800, letterSpacing: '-0.02em', margin: 0 }}>
            Muhasebe ODF
          </h2>
          <span style={{ fontSize: '0.75rem', color: 'var(--text-muted)', fontWeight: 500 }}>
            Veri Yönetim Portalı
          </span>
        </div>
      </div>

      {/* Menu Navigation */}
      <nav style={{ display: 'flex', flexDirection: 'column', gap: '8px', flex: 1 }}>
        {menuItems.map((item) => {
          const Icon = item.icon;
          const isActive = activeTab === item.id;
          return (
            <button
              key={item.id}
              onClick={() => setActiveTab(item.id)}
              className="btn"
              style={{
                justifyContent: 'flex-start',
                width: '100%',
                padding: '12px 16px',
                backgroundColor: isActive ? 'var(--primary)' : 'transparent',
                color: isActive ? '#ffffff' : 'var(--text-secondary)',
                borderRadius: '10px',
                transition: 'all 0.2s ease',
                border: 'none',
              }}
            >
              <Icon size={18} style={{ opacity: isActive ? 1 : 0.7 }} />
              <span style={{ fontSize: '0.9rem', fontWeight: 600 }}>{item.name}</span>
            </button>
          );
        })}
      </nav>

      {/* Dark/Light Mode Switcher & Footer */}
      <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', paddingTop: '20px', borderTop: '1px solid var(--border-color)' }}>
        <button
          onClick={toggleDarkMode}
          className="btn btn-secondary"
          style={{ width: '100%', borderRadius: '10px', padding: '12px' }}
        >
          {darkMode ? <Sun size={18} /> : <Moon size={18} />}
          <span>{darkMode ? 'Açık Tema' : 'Koyu Tema'}</span>
        </button>

        {onLock && (
          <button
            onClick={onLock}
            className="btn btn-secondary"
            style={{ width: '100%', borderRadius: '10px', padding: '12px', color: '#ef4444', border: '1px solid rgba(239, 68, 68, 0.2)', backgroundColor: 'rgba(239, 68, 68, 0.05)' }}
          >
            <Lock size={18} />
            <span>Sistemi Kilitle</span>
          </button>
        )}

        <div style={{ fontSize: '0.7rem', color: 'var(--text-muted)', textAlign: 'center', fontWeight: 500 }}>
          © 2026 Muhasebe A.Ş. <br /> v1.0.0
        </div>
      </div>
    </aside>
  );
};
export default Sidebar;
