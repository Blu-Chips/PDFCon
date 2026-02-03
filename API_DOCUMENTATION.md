# PDFCon API Documentation

## Overview
PDFCon is a Government Financial Report Analysis System that provides AI-powered analysis with comparative benchmarking against Norway's Sovereign Wealth Fund.

## Base URL
```
http://localhost:8000/api/v1
```

## Authentication
Currently no authentication required for testing purposes.

## Endpoints

### Health Check
```
GET /health
```
Returns application health status.

**Response:**
```json
{
  "status": "healthy",
  "version": "1.0.0"
}
```

### Document Management

#### Upload Document
```
POST /documents/
```
Upload a document for processing.

**Form Data:**
- `file` (required): The document file (PDF, DOCX, DOC, TXT)
- `title` (optional): Document title
- `description` (optional): Document description
- `author` (optional): Document author
- `year` (optional): Document year
- `country` (optional): Document country

**Response:**
```json
{
  "message": "Document uploaded successfully",
  "document": {
    "id": "uuid-string",
    "filename": "generated-filename.pdf",
    "original_filename": "user-file.pdf",
    "file_size": 123456,
    "mime_type": "application/pdf",
    "status": "uploaded",
    "title": "Document Title",
    "created_at": "2024-01-01T00:00:00"
  },
  "processing_status": "started"
}
```

#### Get Document
```
GET /documents/{document_id}
```
Retrieve document information by ID.

**Response:**
```json
{
  "id": "uuid-string",
  "filename": "generated-filename.pdf",
  "original_filename": "user-file.pdf",
  "file_size": 123456,
  "mime_type": "application/pdf",
  "status": "processed",
  "title": "Document Title",
  "word_count": 1250,
  "page_count": 5,
  "processing_time": 2.34,
  "created_at": "2024-01-01T00:00:00",
  "processed_at": "2024-01-01T00:00:05"
}
```

#### List Documents
```
GET /documents/
```
List all documents with optional filtering.

**Query Parameters:**
- `skip` (optional): Number of records to skip (default: 0)
- `limit` (optional): Maximum number of records (default: 100)
- `status` (optional): Filter by status (uploaded, processing, processed, failed)

**Response:**
```json
[
  {
    "id": "uuid-string",
    "filename": "generated-filename.pdf",
    "original_filename": "user-file.pdf",
    "status": "processed",
    "created_at": "2024-01-01T00:00:00"
  }
]
```

#### Delete Document
```
DELETE /documents/{document_id}
```
Delete a document and its associated file.

#### Download Document
```
GET /documents/{document_id}/download
```
Download the original document file.

#### Get Processing Status
```
GET /documents/{document_id}/processing-status
```
Get detailed processing status information.

**Response:**
```json
{
  "document_id": "uuid-string",
  "status": "processed",
  "processed_at": "2024-01-01T00:00:05",
  "error_message": null,
  "word_count": 1250,
  "page_count": 5,
  "processing_time": 2.34
}
```

## Document Status Values
- `uploaded`: Document has been uploaded but not processed
- `processing`: Document is currently being processed
- `processed`: Document has been successfully processed
- `failed`: Document processing failed

## Supported File Types
- PDF (.pdf)
- Microsoft Word (.docx, .doc)
- Plain Text (.txt)

## File Size Limit
Maximum file size: 100MB

## Error Responses
All endpoints return appropriate HTTP status codes:
- `200`: Success
- `201`: Created
- `400`: Bad Request
- `404`: Not Found
- `413`: Payload Too Large
- `500`: Internal Server Error

## Example Usage

### Uploading a Document
```bash
curl -X POST "http://localhost:8000/api/v1/documents/" \
  -F "file=@report.pdf" \
  -F "title=Annual Financial Report 2024" \
  -F "description=Government annual financial report" \
  -F "author=Finance Department" \
  -F "year=2024" \
  -F "country=Nigeria"
```

### Listing Documents
```bash
curl -X GET "http://localhost:8000/api/v1/documents/"
```

### Getting Document Details
```bash
curl -X GET "http://localhost:8000/api/v1/documents/{document-id}"
```

## Frontend Integration
The frontend is available at `http://localhost:5173` and provides:
- Drag-and-drop document upload interface
- Real-time upload progress tracking
- Document list with status indicators
- Tabbed navigation between upload and dashboard views

## Testing
Run the integration test suite:
```bash
python document_upload_test.py
```

## Development Notes
- Backend runs on port 8000
- Frontend runs on port 5173
- PostgreSQL database required for document storage
- Automatic document processing occurs after upload
- Text extraction supports PDF, Word, and text files