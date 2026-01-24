function App() {
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

        {/* Main Content */}
        <main className="max-w-4xl mx-auto">
          <div className="glass-dark rounded-lg p-8 text-center">
            <h2 className="text-2xl font-semibold text-white mb-4">
              Welcome to PDFCon
            </h2>
            <p className="text-gray-300 mb-6">
              AI-powered government financial report analysis with comparative benchmarking against Norway's Sovereign Wealth Fund
            </p>
            <div className="flex gap-4 justify-center">
              <button className="bg-purple-600 hover:bg-purple-700 text-white font-semibold py-3 px-6 rounded-lg transition-colors">
                Get Started
              </button>
              <button className="bg-transparent border border-purple-600 text-purple-400 hover:text-white hover:border-white font-semibold py-3 px-6 rounded-lg transition-colors">
                Learn More
              </button>
            </div>
          </div>

          {/* Features */}
          <div className="grid md:grid-cols-3 gap-6 mt-12">
            <div className="glass-dark rounded-lg p-6">
              <h3 className="text-xl font-semibold text-white mb-2">PDF Processing</h3>
              <p className="text-gray-400">
                Extract text, tables, and financial data from government reports
              </p>
            </div>
            <div className="glass-dark rounded-lg p-6">
              <h3 className="text-xl font-semibold text-white mb-2">AI Analysis</h3>
              <p className="text-gray-400">
                Automated financial analysis with actionable insights
              </p>
            </div>
            <div className="glass-dark rounded-lg p-6">
              <h3 className="text-xl font-semibold text-white mb-2">Benchmarking</h3>
              <p className="text-gray-400">
                Compare with Norway's Sovereign Wealth Fund performance
              </p>
            </div>
          </div>
        </main>

        {/* Footer */}
        <footer className="text-center mt-16 text-gray-500">
          <p>Built with ❤️ by Blu-Chips</p>
        </footer>
      </div>
    </div>
  )
}

export default App