# PDFCon - Government Financial Report Analysis System

An intelligent system that collects, analyzes, and visualizes government financial reports from Auditor General websites. Provides deep financial insights, comparative analysis with Norway's Sovereign Wealth Fund, and delivers McKinsey-quality reports through an ultra-modern, futuristic UI.

## 🌟 Features

- **Automated Document Collection**: Web scraping from Auditor General websites + manual upload support
- **Advanced PDF Processing**: Text extraction, OCR, table extraction using PyMuPDF, Tesseract, and Camelot
- **Financial Analysis Engine**: Automated indicator extraction, trend analysis, anomaly detection
- **Comparative Analysis**: Like-for-like benchmarking with Norway's Sovereign Wealth Fund
- **Strategic Recommendations**: AI-powered insights and best practices
- **Interactive Dashboards**: Real-time visualizations with stunning 3D graphics
- **Multi-format Reports**: PDF, HTML, Excel, and PowerPoint exports
- **AI Agent Orchestration**: Specialized agents for scraping, processing, analysis, and reporting

## 🏗️ Architecture

```
PDFCon/
├── backend/          # Python/FastAPI backend
│   ├── app/
│   │   ├── agents/           # AI agents
│   │   ├── analyzers/        # Financial analysis
│   │   ├── api/              # API routes
│   │   ├── core/             # Core configuration
│   │   ├── models/           # Database models
│   │   ├── processors/       # Document processing
│   │   ├── scrapers/         # Web scraping
│   │   └── services/         # Business logic
│   └── requirements.txt
├── frontend/         # React/TypeScript frontend
│   └── src/
│       ├── components/
│       │   ├── 3d/           # Three.js visualizations
│       │   ├── charts/       # Data charts
│       │   ├── dashboard/    # Dashboard components
│       │   └── ui/           # UI components
│       ├── pages/            # Page components
│       └── store/            # State management
├── data/             # Data files and templates
└── docker-compose.yml
```

## 🚀 Quick Start

### Prerequisites

- Docker and Docker Compose
- Python 3.11+ (for local development)
- Node.js 18+ (for local development)
- OpenAI API key (for AI features)

### Using Docker (Recommended)

1. **Clone the repository**
```bash
git clone https://github.com/Blu-Chips/PDFCon.git
cd PDFCon
```

2. **Configure environment variables**
```bash
cp .env.example .env
# Edit .env and add your OPENAI_API_KEY
```

3. **Start all services**
```bash
docker-compose up -d
```

4. **Access the application**
- Frontend: http://localhost:3000
- Backend API: http://localhost:8000
- API Documentation: http://localhost:8000/api/docs
- Celery Flower: http://localhost:5555
- MinIO Console: http://localhost:9001

### Local Development

#### Backend Setup

1. **Install Python dependencies**
```bash
cd backend
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install -r requirements.txt
```

2. **Install Playwright browsers**
```bash
playwright install chromium
```

3. **Configure environment**
```bash
cp ../.env.example .env
# Edit .env with your configuration
```

4. **Run the backend**
```bash
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

#### Frontend Setup

1. **Install Node.js dependencies**
```bash
cd frontend
npm install
```

2. **Run the frontend**
```bash
npm run dev
```

Access the frontend at http://localhost:5173

## 🛠️ Technology Stack

### Backend
- **Framework**: FastAPI
- **Database**: PostgreSQL, MongoDB, Redis
- **Document Processing**: PyMuPDF, Camelot, Tesseract OCR
- **Web Scraping**: Playwright, Selenium
- **Data Analysis**: Pandas, NumPy, SciPy
- **AI/ML**: LangChain, OpenAI GPT-4, SpaCy
- **Task Queue**: Celery
- **Storage**: MinIO (S3-compatible)

### Frontend
- **Framework**: React 18 + TypeScript
- **Build Tool**: Vite
- **UI Library**: shadcn/ui, Radix UI
- **Styling**: TailwindCSS
- **3D Graphics**: Three.js, React Three Fiber
- **Data Visualization**: Recharts, D3.js
- **State Management**: Zustand, React Query
- **Animations**: Framer Motion

## 📊 Usage

### Upload Documents

1. Navigate to the dashboard
2. Click "Upload Report"
3. Select a PDF file or specify a country/year for automatic scraping
4. Wait for processing

### Generate Analysis

1. Select a processed document
2. Choose analysis type (Financial, Compliance, Comparative)
3. Click "Generate Analysis"
4. Review interactive results

### Compare with Norway

1. Select a report
2. Click "Compare with Norway Sovereign Wealth Fund"
3. View like-for-like metrics and recommendations

### Export Reports

1. Navigate to the Reports section
2. Select desired format (PDF, HTML, Excel, PowerPoint)
3. Click "Export"

## 🔧 Configuration

### Environment Variables

Key variables in `.env`:

```env
# Database
DATABASE_URL=postgresql+asyncpg://user:pass@host:5432/dbname
MONGODB_URL=mongodb://host:27017/dbname

# AI/ML
OPENAI_API_KEY=your-api-key

# Storage
MINIO_ENDPOINT=localhost:9000
MINIO_ACCESS_KEY=minioadmin
MINIO_SECRET_KEY=minioadmin
```

## 🧪 Testing

### Backend Tests
```bash
cd backend
pytest tests/
```

### Frontend Tests
```bash
cd frontend
npm test
```

## 📈 Development Workflow

### Adding New Features

1. **Backend**: Create modules in appropriate directories under `backend/app/`
2. **Frontend**: Create components under `frontend/src/components/`
3. **API**: Add routes under `backend/app/api/`
4. **Tests**: Add tests to `backend/tests/` and `frontend/`

### Code Style

- Backend: Black, flake8, mypy
- Frontend: ESLint, Prettier
- Commit: Conventional Commits

## 🚢 Deployment

### Docker Swarm
```bash
docker stack deploy -c docker-compose.yml pdfcon
```

### Kubernetes
See `docs/deployment.md` for Kubernetes deployment instructions.

## 📚 Documentation

- [Architecture](docs/architecture.md)
- [API Documentation](docs/api.md)
- [Deployment Guide](docs/deployment.md)
- [User Guide](docs/user_guide.md)

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## 📄 License

This project is licensed under the MIT License.

## 👥 Team

- **Project Lead**: Blu-Chips
- **Architecture**: Enterprise-grade microservices
- **AI/ML**: LangChain + OpenAI integration
- **Frontend**: Ultra-modern React + Three.js

## 🎯 Roadmap

See [PROJECT_PLAN.md](PROJECT_PLAN.md) for the complete 18-week implementation roadmap.

### Current Status: Phase 1 - Foundation ✅

- [x] Project setup and architecture
- [x] Development environment configuration
- [x] Database schema design
- [x] API structure design
- [ ] Document collection module (upload only)
- [ ] Basic PDF extraction (PyMuPDF)
- [ ] Database models and migrations
- [ ] Basic API endpoints

## 📞 Support

For support and questions:
- GitHub Issues: https://github.com/Blu-Chips/PDFCon/issues
- Documentation: https://github.com/Blu-Chips/PDFCon/wiki

## 🙏 Acknowledgments

- Norway Sovereign Wealth Fund for benchmarking data
- OpenAI for GPT-4 API
- FastAPI and React communities
- All open-source contributors

---

**Built with ❤️ by Blu-Chips**

*Transforming government financial reporting with AI-powered insights*