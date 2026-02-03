"""
Document API Routes
"""
import os
import uuid
from datetime import datetime
from typing import List
from fastapi import APIRouter, UploadFile, File, Form, Depends, HTTPException, status
from fastapi.responses import FileResponse
from sqlalchemy.ext.asyncio import AsyncSession
import aiofiles
from app.core.database import get_db
from app.models.document import Document, DocumentStatus, DocumentType
from app.services.document_processor import DocumentProcessor
import logging

router = APIRouter(prefix="/documents", tags=["Documents"])
logger = logging.getLogger(__name__)

# Configuration
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)


@router.post("/", response_model=dict, status_code=status.HTTP_201_CREATED)
async def upload_document(
    file: UploadFile = File(...),
    title: str = Form(None),
    description: str = Form(None),
    author: str = Form(None),
    year: int = Form(None),
    country: str = Form(None),
    db: AsyncSession = Depends(get_db)
):
    """
    Upload a document for processing
    
    Args:
        file: The document file to upload
        title: Document title
        description: Document description
        author: Document author
        year: Document year
        country: Document country
        db: Database session
        
    Returns:
        Document information and processing status
    """
    # Validate file
    if not file:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="No file provided"
        )
    
    # Validate file type
    allowed_types = [
        "application/pdf",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "application/msword",
        "text/plain"
    ]
    
    if file.content_type not in allowed_types:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Unsupported file type: {file.content_type}. Allowed types: {allowed_types}"
        )
    
    # Validate file size (100MB limit)
    MAX_FILE_SIZE = 100 * 1024 * 1024  # 100MB
    content = await file.read()
    if len(content) > MAX_FILE_SIZE:
        raise HTTPException(
            status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
            detail=f"File too large. Maximum size is {MAX_FILE_SIZE / (1024*1024)}MB"
        )
    
    # Reset file pointer
    await file.seek(0)
    
    # Generate filename
    file_extension = file.filename.split('.')[-1] if '.' in file.filename else ''
    new_filename = f"{uuid.uuid4()}.{file_extension}"
    file_path = os.path.join(UPLOAD_DIR, new_filename)
    
    # Save file
    try:
        async with aiofiles.open(file_path, 'wb') as out_file:
            await out_file.write(content)
    except Exception as e:
        logger.error(f"Failed to save file {new_filename}: {e}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Failed to save file"
        )
    
    # Create document record
    document = Document(
        filename=new_filename,
        original_filename=file.filename,
        file_path=file_path,
        file_size=len(content),
        mime_type=file.content_type,
        title=title,
        description=description,
        author=author,
        year=year,
        country=country
    )
    
    try:
        db.add(document)
        await db.commit()
        await db.refresh(document)
        
        # Start asynchronous processing
        processor = DocumentProcessor()
        await processor.process_document(document.id, db)
        
        logger.info(f"Document uploaded successfully: {document.id}")
        
        return {
            "message": "Document uploaded successfully",
            "document": document.to_dict(),
            "processing_status": "started"
        }
        
    except Exception as e:
        await db.rollback()
        # Clean up file if database operation fails
        try:
            os.remove(file_path)
        except:
            pass
        logger.error(f"Failed to create document record: {e}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Failed to create document record"
        )


@router.get("/{document_id}", response_model=dict)
async def get_document(document_id: uuid.UUID, db: AsyncSession = Depends(get_db)):
    """
    Get document information by ID
    
    Args:
        document_id: Document UUID
        db: Database session
        
    Returns:
        Document information
    """
    document = await db.get(Document, document_id)
    if not document:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Document not found"
        )
    
    return document.to_dict()


@router.get("/", response_model=List[dict])
async def list_documents(
    skip: int = 0,
    limit: int = 100,
    status: DocumentStatus = None,
    db: AsyncSession = Depends(get_db)
):
    """
    List documents with optional filtering
    
    Args:
        skip: Number of records to skip
        limit: Maximum number of records to return
        status: Filter by document status
        db: Database session
        
    Returns:
        List of documents
    """
    from sqlalchemy import select
    
    query = select(Document)
    
    if status:
        query = query.where(Document.status == status)
    
    query = query.offset(skip).limit(limit)
    
    result = await db.execute(query)
    documents = result.scalars().all()
    
    return [doc.to_dict() for doc in documents]


@router.delete("/{document_id}", status_code=status.HTTP_204_NO_CONTENT)
async def delete_document(document_id: uuid.UUID, db: AsyncSession = Depends(get_db)):
    """
    Delete a document
    
    Args:
        document_id: Document UUID
        db: Database session
    """
    document = await db.get(Document, document_id)
    if not document:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Document not found"
        )
    
    # Delete file from filesystem
    try:
        if os.path.exists(document.file_path):
            os.remove(document.file_path)
    except Exception as e:
        logger.warning(f"Failed to delete file {document.file_path}: {e}")
    
    # Delete from database
    await db.delete(document)
    await db.commit()
    
    logger.info(f"Document deleted: {document_id}")


@router.get("/{document_id}/download")
async def download_document(document_id: uuid.UUID, db: AsyncSession = Depends(get_db)):
    """
    Download a document file
    
    Args:
        document_id: Document UUID
        db: Database session
        
    Returns:
        File response
    """
    document = await db.get(Document, document_id)
    if not document:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Document not found"
        )
    
    if not os.path.exists(document.file_path):
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Document file not found"
        )
    
    return FileResponse(
        path=document.file_path,
        filename=document.original_filename,
        media_type=document.mime_type
    )


@router.get("/{document_id}/processing-status")
async def get_processing_status(document_id: uuid.UUID, db: AsyncSession = Depends(get_db)):
    """
    Get document processing status
    
    Args:
        document_id: Document UUID
        db: Database session
        
    Returns:
        Processing status information
    """
    document = await db.get(Document, document_id)
    if not document:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Document not found"
        )
    
    return {
        "document_id": str(document_id),
        "status": document.status.value,
        "processed_at": document.processed_at.isoformat() if document.processed_at else None,
        "error_message": document.error_message,
        "word_count": document.word_count,
        "page_count": document.page_count,
        "processing_time": document.processing_time
    }