import React, { useState, useCallback } from 'react';
import { Upload, FileText, CheckCircle, AlertCircle, Loader } from 'lucide-react';

interface Document {
  id: string;
  filename: string;
  original_filename: string;
  file_size: number;
  status: 'uploaded' | 'processing' | 'processed' | 'failed';
  created_at: string;
  word_count?: number;
  page_count?: number;
  error_message?: string;
}

const DocumentUploader: React.FC = () => {
  const [dragActive, setDragActive] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [documents, setDocuments] = useState<Document[]>([]);
  const [uploadProgress, setUploadProgress] = useState(0);

  // Handle drag events
  const handleDrag = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  }, []);

  // Handle file drop
  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleFiles(e.dataTransfer.files);
    }
  }, []);

  // Handle file selection
  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    e.preventDefault();
    if (e.target.files && e.target.files[0]) {
      handleFiles(e.target.files);
    }
  };

  // Process uploaded files
  const handleFiles = async (files: FileList) => {
    const file = files[0];
    
    // Validate file type
    const allowedTypes = [
      'application/pdf',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'application/msword',
      'text/plain'
    ];
    
    if (!allowedTypes.includes(file.type)) {
      alert('Unsupported file type. Please upload PDF, DOCX, DOC, or TXT files.');
      return;
    }
    
    // Validate file size (100MB)
    if (file.size > 100 * 1024 * 1024) {
      alert('File too large. Maximum size is 100MB.');
      return;
    }
    
    setUploading(true);
    setUploadProgress(0);
    
    try {
      // Simulate progress
      const progressInterval = setInterval(() => {
        setUploadProgress(prev => {
          if (prev >= 90) {
            clearInterval(progressInterval);
            return 90;
          }
          return prev + 10;
        });
      }, 200);
      
      // Upload file
      const formData = new FormData();
      formData.append('file', file);
      formData.append('title', file.name.replace(/\.[^/.]+$/, ""));
      
      const response = await fetch('http://localhost:8000/api/v1/documents/', {
        method: 'POST',
        body: formData,
      });
      
      clearInterval(progressInterval);
      setUploadProgress(100);
      
      if (!response.ok) {
        throw new Error(`Upload failed: ${response.statusText}`);
      }
      
      const result = await response.json();
      
      // Add to documents list
      setDocuments(prev => [result.document, ...prev]);
      
      setTimeout(() => {
        setUploading(false);
        setUploadProgress(0);
      }, 1000);
      
    } catch (error) {
      console.error('Upload error:', error);
      alert('Upload failed. Please try again.');
      setUploading(false);
      setUploadProgress(0);
    }
  };

  // Format file size
  const formatFileSize = (bytes: number): string => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  // Get status icon
  const getStatusIcon = (status: Document['status']) => {
    switch (status) {
      case 'processed':
        return <CheckCircle className="w-5 h-5 text-green-500" />;
      case 'failed':
        return <AlertCircle className="w-5 h-5 text-red-500" />;
      case 'processing':
        return <Loader className="w-5 h-5 text-yellow-500 animate-spin" />;
      default:
        return <FileText className="w-5 h-5 text-blue-500" />;
    }
  };

  // Get status text
  const getStatusText = (status: Document['status']) => {
    switch (status) {
      case 'processed':
        return 'Processed';
      case 'failed':
        return 'Failed';
      case 'processing':
        return 'Processing';
      default:
        return 'Uploaded';
    }
  };

  return (
    <div className="max-w-4xl mx-auto">
      {/* Upload Area */}
      <div
        className={`border-2 border-dashed rounded-lg p-8 text-center transition-all duration-200 ${
          dragActive 
            ? 'border-purple-500 bg-purple-500/10' 
            : 'border-gray-600 hover:border-purple-400'
        }`}
        onDragEnter={handleDrag}
        onDragLeave={handleDrag}
        onDragOver={handleDrag}
        onDrop={handleDrop}
      >
        <input
          type="file"
          id="file-upload"
          className="hidden"
          onChange={handleChange}
          accept=".pdf,.docx,.doc,.txt"
        />
        
        {uploading ? (
          <div className="space-y-4">
            <Loader className="w-12 h-12 text-purple-500 animate-spin mx-auto" />
            <div>
              <p className="text-white font-medium">Uploading...</p>
              <div className="w-full bg-gray-700 rounded-full h-2 mt-2">
                <div 
                  className="bg-purple-600 h-2 rounded-full transition-all duration-300"
                  style={{ width: `${uploadProgress}%` }}
                ></div>
              </div>
              <p className="text-gray-400 text-sm mt-1">{uploadProgress}%</p>
            </div>
          </div>
        ) : (
          <div>
            <Upload className="w-12 h-12 text-gray-400 mx-auto mb-4" />
            <p className="text-white font-medium mb-2">
              Drag and drop your document here
            </p>
            <p className="text-gray-400 mb-4">
              or click to browse files
            </p>
            <label
              htmlFor="file-upload"
              className="bg-purple-600 hover:bg-purple-700 text-white font-semibold py-2 px-6 rounded-lg cursor-pointer transition-colors"
            >
              Select File
            </label>
            <p className="text-gray-500 text-sm mt-4">
              Supports PDF, DOCX, DOC, and TXT files (Max 100MB)
            </p>
          </div>
        )}
      </div>

      {/* Documents List */}
      {documents.length > 0 && (
        <div className="mt-8 glass-dark rounded-lg p-6">
          <h3 className="text-xl font-semibold text-white mb-4">Uploaded Documents</h3>
          <div className="space-y-3">
            {documents.map((doc) => (
              <div 
                key={doc.id} 
                className="flex items-center justify-between p-4 bg-gray-800/50 rounded-lg hover:bg-gray-800/70 transition-colors"
              >
                <div className="flex items-center space-x-3">
                  {getStatusIcon(doc.status)}
                  <div>
                    <p className="text-white font-medium">{doc.original_filename}</p>
                    <p className="text-gray-400 text-sm">
                      {formatFileSize(doc.file_size)}
                      {doc.word_count && ` • ${doc.word_count} words`}
                      {doc.page_count && ` • ${doc.page_count} pages`}
                      {doc.created_at && ` • ${new Date(doc.created_at).toLocaleDateString()}`}
                    </p>
                    {doc.error_message && (
                      <p className="text-red-400 text-sm mt-1">{doc.error_message}</p>
                    )}
                  </div>
                </div>
                <div className="flex items-center space-x-2">
                  <span className={`px-2 py-1 rounded text-xs font-medium ${
                    doc.status === 'processed' ? 'bg-green-900/50 text-green-400' :
                    doc.status === 'failed' ? 'bg-red-900/50 text-red-400' :
                    doc.status === 'processing' ? 'bg-yellow-900/50 text-yellow-400' :
                    'bg-blue-900/50 text-blue-400'
                  }`}>
                    {getStatusText(doc.status)}
                  </span>
                  <button className="text-purple-400 hover:text-purple-300 text-sm font-medium">
                    View Details
                  </button>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
};

export default DocumentUploader;