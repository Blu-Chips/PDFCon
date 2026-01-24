"""
Application Configuration Module
"""
import os
from typing import Optional
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """Application settings"""
    
    # Application
    APP_NAME: str = "PDFCon - Government Financial Report Analysis System"
    APP_VERSION: str = "1.0.0"
    DEBUG: bool = False
    API_V1_PREFIX: str = "/api/v1"
    
    # Server
    HOST: str = "0.0.0.0"
    PORT: int = 8000
    
    # Database
    DATABASE_URL: str = "postgresql+asyncpg://postgres:postgres@localhost:5432/pdfcon"
    DATABASE_TEST_URL: str = "postgresql+asyncpg://postgres:postgres@localhost:5432/pdfcon_test"
    MONGODB_URL: str = "mongodb://localhost:27017/pdfcon"
    
    # Redis
    REDIS_URL: str = "redis://localhost:6379/0"
    REDIS_CELERY_DB: int = 1
    
    # Security
    SECRET_KEY: str = "your-secret-key-change-in-production"
    ALGORITHM: str = "HS256"
    ACCESS_TOKEN_EXPIRE_MINUTES: int = 30
    
    # File Storage
    UPLOAD_DIR: str = "uploads"
    MAX_FILE_SIZE: int = 100 * 1024 * 1024  # 100MB
    ALLOWED_FILE_TYPES: list = ["application/pdf"]
    
    # MinIO/S3
    MINIO_ENDPOINT: str = "localhost:9000"
    MINIO_ACCESS_KEY: str = "minioadmin"
    MINIO_SECRET_KEY: str = "minioadmin"
    MINIO_BUCKET: str = "pdfcon-documents"
    
    # AI/ML
    OPENAI_API_KEY: Optional[str] = None
    OPENAI_MODEL: str = "gpt-4-turbo-preview"
    OPENAI_TEMPERATURE: float = 0.3
    CEREBRAS_API_KEY: str = os.getenv("CEREBRAS_API_KEY", "")
    
    # Web Scraping
    PLAYWRIGHT_HEADLESS: bool = True
    SELENIUM_HEADLESS: bool = True
    USER_AGENT: str = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    
    # Celery
    CELERY_BROKER_URL: str = "redis://localhost:6379/1"
    CELERY_RESULT_BACKEND: str = "redis://localhost:6379/1"
    
    # Logging
    LOG_LEVEL: str = "INFO"
    LOG_FORMAT: str = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    
    # CORS
    CORS_ORIGINS: list = ["http://localhost:3000", "http://localhost:5173"]
    
    # Task Processing
    MAX_WORKERS: int = 4
    TASK_TIMEOUT: int = 3600  # 1 hour
    
    # Report Settings
    REPORT_CACHE_TTL: int = 86400  # 24 hours
    MAX_REPORT_AGE_DAYS: int = 30
    
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=True,
    )


settings = Settings()