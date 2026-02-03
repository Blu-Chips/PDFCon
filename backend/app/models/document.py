"""
Document Model
"""
from datetime import datetime
from enum import Enum
from sqlalchemy import Column, Integer, String, Text, DateTime, Enum as SQLEnum, Float
from sqlalchemy.dialects.postgresql import UUID
import uuid
from app.core.database import Base


class DocumentStatus(str, Enum):
    """Document processing status"""
    UPLOADED = "uploaded"
    PROCESSING = "processing"
    PROCESSED = "processed"
    FAILED = "failed"


class DocumentType(str, Enum):
    """Document type classification"""
    GOVERNMENT_REPORT = "government_report"
    FINANCIAL_STATEMENT = "financial_statement"
    BUDGET_DOCUMENT = "budget_document"
    AUDIT_REPORT = "audit_report"
    OTHER = "other"


class Document(Base):
    """Document model for storing uploaded documents"""
    __tablename__ = "documents"
    
    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    filename = Column(String(255), nullable=False)
    original_filename = Column(String(255), nullable=False)
    file_path = Column(Text, nullable=False)
    file_size = Column(Integer, nullable=False)
    mime_type = Column(String(100), nullable=False)
    document_type = Column(SQLEnum(DocumentType), default=DocumentType.OTHER)
    status = Column(SQLEnum(DocumentStatus), default=DocumentStatus.UPLOADED)
    
    # Metadata
    title = Column(String(500))
    description = Column(Text)
    author = Column(String(255))
    year = Column(Integer)
    country = Column(String(100))
    
    # Processing results
    extracted_text = Column(Text)
    word_count = Column(Integer)
    page_count = Column(Integer)
    processing_time = Column(Float)  # in seconds
    
    # Error handling
    error_message = Column(Text)
    
    # Timestamps
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    processed_at = Column(DateTime)
    
    def __repr__(self):
        return f"<Document(id={self.id}, filename='{self.filename}', status='{self.status}')>"
    
    @property
    def is_processed(self) -> bool:
        """Check if document is successfully processed"""
        return self.status == DocumentStatus.PROCESSED
    
    @property
    def processing_failed(self) -> bool:
        """Check if document processing failed"""
        return self.status == DocumentStatus.FAILED
    
    def to_dict(self) -> dict:
        """Convert document to dictionary representation"""
        return {
            "id": str(self.id),
            "filename": self.filename,
            "original_filename": self.original_filename,
            "file_size": self.file_size,
            "mime_type": self.mime_type,
            "document_type": self.document_type.value,
            "status": self.status.value,
            "title": self.title,
            "description": self.description,
            "author": self.author,
            "year": self.year,
            "country": self.country,
            "word_count": self.word_count,
            "page_count": self.page_count,
            "processing_time": self.processing_time,
            "created_at": self.created_at.isoformat() if self.created_at else None,
            "updated_at": self.updated_at.isoformat() if self.updated_at else None,
            "processed_at": self.processed_at.isoformat() if self.processed_at else None,
        }