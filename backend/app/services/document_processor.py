"""
Document Processing Service
"""
import asyncio
import time
from datetime import datetime
import PyPDF2
from docx import Document as DocxDocument
# import textract  # Temporarily disabled due to installation issues
from sqlalchemy.ext.asyncio import AsyncSession
from app.models.document import Document, DocumentStatus
import logging

logger = logging.getLogger(__name__)


class DocumentProcessor:
    """Service for processing uploaded documents"""
    
    async def process_document(self, document_id: str, db: AsyncSession):
        """
        Process a document asynchronously
        
        Args:
            document_id: Document UUID
            db: Database session
        """
        # Get document from database
        document = await db.get(Document, document_id)
        if not document:
            logger.error(f"Document {document_id} not found")
            return
        
        # Update status to processing
        document.status = DocumentStatus.PROCESSING
        await db.commit()
        
        start_time = time.time()
        
        try:
            # Extract text based on file type
            extracted_text = await self._extract_text(document)
            
            # Update document with processing results
            document.extracted_text = extracted_text
            document.word_count = len(extracted_text.split()) if extracted_text else 0
            document.page_count = await self._count_pages(document)
            document.processing_time = time.time() - start_time
            document.status = DocumentStatus.PROCESSED
            document.processed_at = datetime.utcnow()
            
            await db.commit()
            logger.info(f"Document {document_id} processed successfully in {document.processing_time:.2f}s")
            
        except Exception as e:
            # Handle processing errors
            document.status = DocumentStatus.FAILED
            document.error_message = str(e)
            document.processing_time = time.time() - start_time
            
            await db.commit()
            logger.error(f"Failed to process document {document_id}: {e}")
    
    async def _extract_text(self, document: Document) -> str:
        """
        Extract text from document based on file type
        
        Args:
            document: Document object
            
        Returns:
            Extracted text content
        """
        file_path = document.file_path
        
        try:
            if document.mime_type == "application/pdf":
                return await self._extract_pdf_text(file_path)
            elif document.mime_type in [
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "application/msword"
            ]:
                return await self._extract_docx_text(file_path)
            elif document.mime_type == "text/plain":
                return await self._extract_txt_text(file_path)
            else:
                # Fallback to textract for other formats
                return await asyncio.get_event_loop().run_in_executor(
                    None, self._extract_with_textract, file_path
                )
        except Exception as e:
            logger.error(f"Text extraction failed for {document.filename}: {e}")
            raise
    
    async def _extract_pdf_text(self, file_path: str) -> str:
        """Extract text from PDF file"""
        def extract():
            text = ""
            with open(file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                for page in pdf_reader.pages:
                    text += page.extract_text() + "\n"
            return text.strip()
        
        return await asyncio.get_event_loop().run_in_executor(None, extract)
    
    async def _extract_docx_text(self, file_path: str) -> str:
        """Extract text from DOCX file"""
        def extract():
            doc = DocxDocument(file_path)
            text = ""
            for paragraph in doc.paragraphs:
                text += paragraph.text + "\n"
            return text.strip()
        
        return await asyncio.get_event_loop().run_in_executor(None, extract)
    
    async def _extract_txt_text(self, file_path: str) -> str:
        """Extract text from TXT file"""
        def extract():
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
                return file.read().strip()
        
        return await asyncio.get_event_loop().run_in_executor(None, extract)
    
    def _extract_with_textract(self, file_path: str) -> str:
        """Extract text using textract (fallback method) - temporarily disabled"""
        # try:
        #     text = textract.process(file_path)
        #     return text.decode('utf-8').strip()
        # except Exception as e:
        #     logger.warning(f"Textract failed for {file_path}: {e}")
        #     return ""
        logger.warning(f"Textract is disabled - returning empty text for {file_path}")
        return ""
    
    async def _count_pages(self, document: Document) -> int:
        """
        Count pages in document
        
        Args:
            document: Document object
            
        Returns:
            Page count
        """
        try:
            if document.mime_type == "application/pdf":
                def count_pdf_pages():
                    with open(document.file_path, 'rb') as file:
                        pdf_reader = PyPDF2.PdfReader(file)
                        return len(pdf_reader.pages)
                
                return await asyncio.get_event_loop().run_in_executor(None, count_pdf_pages)
            elif document.mime_type in [
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "application/msword"
            ]:
                # Approximate page count for Word documents
                # Rough estimate: ~500 words per page
                if document.word_count:
                    return max(1, document.word_count // 500)
                return 1
            else:
                return 1
        except Exception as e:
            logger.warning(f"Page counting failed for {document.filename}: {e}")
            return 1
    
    async def batch_process_documents(self, document_ids: list, db: AsyncSession):
        """
        Process multiple documents concurrently
        
        Args:
            document_ids: List of document UUIDs
            db: Database session
        """
        tasks = [
            self.process_document(doc_id, db) 
            for doc_id in document_ids
        ]
        await asyncio.gather(*tasks, return_exceptions=True)