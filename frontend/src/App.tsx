import React, { useState } from 'react';
import DocumentUploader from './components/DocumentUploader';

function App() {
  const [activeTab, setActiveTab] = useState<'upload' | 'dashboard'>('upload');
  
  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-900 via-purple-900 to-slate-900">
      <div className="container mx-auto px-4 py-8">
        {/* Header */}
        <header className="text-center mb-12">
          <h1 className="text-5xl font-bold text-white mb-4 gradient-text">
            PDFCon
          </h1>
          <p className="text-xl text-gray-300">
            Government Financial Report Analysis System
          </p>
        </header>

        {/* Navigation Tabs */}
        <div className="max-w-4xl mx-auto mb-8">
          <div className="flex bg-gray-800/50 rounded-lg p-1">
            <button
              onClick={() => setActiveTab('upload')}
              className={`flex-1 py-3 px-4 rounded-md font-medium transition-colors ${
                activeTab === 'upload'
                  ? 'bg-purple-600 text-white'
                  : 'text-gray-400 hover:text-white'
              }`}
            >
              Upload Documents
            </button>
            <button
              onClick={() => setActiveTab('dashboard')}
              className={`flex-1 py-3 px-4 rounded-md font-medium transition-colors ${
                activeTab === 'dashboard'
                  ? 'bg-purple-600 text-white'
                  : 'text-gray-400 hover:text-white'
              }`}
            >
              Dashboard
            </button>
          </div>
        </div>

        {/* Main Content */}
        <main className="max-w-6xl mx-auto">
          {activeTab === 'upload' ? (
            <DocumentUploader />
          ) : (
            <div className="glass-dark rounded-lg p-8 text-center">
              <h2 className="text-2xl font-semibold text-white mb-4">
                Analysis Dashboard
              </h2>
              <p className="text-gray-300 mb-6">
                Analyze government financial reports with AI-powered insights and benchmarking against Norway's Sovereign Wealth Fund
              </p>
              <div className="grid md:grid-cols-3 gap-6 mt-8">
                <div className="glass-dark rounded-lg p-6">
                  <h3 className="text-xl font-semibold text-white mb-2">Financial Analysis</h3>
                  <p className="text-gray-400">
                    Extract key financial metrics and trends from government reports
                  </p>
                </div>
                <div className="glass-dark rounded-lg p-6">
                  <h3 className="text-xl font-semibold text-white mb-2">Comparative Benchmarking</h3>
                  <p className="text-gray-400">
                    Compare performance with Norway's Sovereign Wealth Fund benchmarks
                  </p>
                </div>
                <div className="glass-dark rounded-lg p-6">
                  <h3 className="text-xl font-semibold text-white mb-2">Risk Assessment</h3>
                  <p className="text-gray-400">
                    Identify financial risks and sustainability indicators
                  </p>
                </div>
              </div>
            </div>
          )}
        </main>

        {/* Footer */}
        <footer className="text-center mt-16 text-gray-500">
          <p>Built with ❤️ by Blu-Chips</p>
        </footer>
      </div>
    </div>
  );
}

export default App;