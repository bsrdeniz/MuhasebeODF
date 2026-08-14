import { useState } from 'react';
import Sidebar from './components/Sidebar';
import Dashboard from './components/Dashboard';
import DocumentUploader from './components/DocumentUploader';
import DocumentExplorer from './components/DocumentExplorer';
import { LockScreen } from './components/LockScreen';

function App() {
  const [activeTab, setActiveTab] = useState<string>('dashboard');
  const [darkMode, setDarkMode] = useState<boolean>(false);
  const [cryptoKey, setCryptoKey] = useState<CryptoKey | null>(null);

  if (!cryptoKey) {
    return <LockScreen onUnlock={(key) => setCryptoKey(key)} />;
  }

  const renderContent = () => {
    switch (activeTab) {
      case 'dashboard':
        return <Dashboard setActiveTab={setActiveTab} cryptoKey={cryptoKey} />;
      case 'upload':
        return <DocumentUploader onUploadSuccess={() => setActiveTab('explorer')} cryptoKey={cryptoKey} />;
      case 'explorer':
        return <DocumentExplorer cryptoKey={cryptoKey} />;
      default:
        return <Dashboard setActiveTab={setActiveTab} cryptoKey={cryptoKey} />;
    }
  };

  return (
    <div className="app-container">
      {/* Sidebar Navigation */}
      <Sidebar
        activeTab={activeTab}
        setActiveTab={setActiveTab}
        darkMode={darkMode}
        setDarkMode={setDarkMode}
        onLock={() => setCryptoKey(null)}
      />

      {/* Main Content Viewport */}
      <main className="main-content">
        {renderContent()}
      </main>
    </div>
  );
}

export default App;
