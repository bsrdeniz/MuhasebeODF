import { useState } from 'react';
import Sidebar from './components/Sidebar';
import Dashboard from './components/Dashboard';
import DocumentUploader from './components/DocumentUploader';
import DocumentExplorer from './components/DocumentExplorer';

function App() {
  const [activeTab, setActiveTab] = useState<string>('dashboard');
  const [darkMode, setDarkMode] = useState<boolean>(false);

  const renderContent = () => {
    switch (activeTab) {
      case 'dashboard':
        return <Dashboard setActiveTab={setActiveTab} />;
      case 'upload':
        return <DocumentUploader onUploadSuccess={() => setActiveTab('explorer')} />;
      case 'explorer':
        return <DocumentExplorer />;
      default:
        return <Dashboard setActiveTab={setActiveTab} />;
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
      />

      {/* Main Content Viewport */}
      <main className="main-content">
        {renderContent()}
      </main>
    </div>
  );
}

export default App;
