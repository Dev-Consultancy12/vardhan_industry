import { useState } from 'react'
import { Upload, FileText, CheckCircle, RefreshCcw, Download, Eye, EyeOff, FileSpreadsheet, Database } from 'lucide-react'
import TestReadingsEditor from './TestReadingsEditor'

function App() {
  const [file, setFile] = useState(null)
  const [status, setStatus] = useState('idle') // idle, uploading, processing, success, error
  const [errorMsg, setErrorMsg] = useState('')
  const [pdfUrl, setPdfUrl] = useState(null)
  const [isDragging, setIsDragging] = useState(false)
  const [showPreview, setShowPreview] = useState(true)
  const [totalPages, setTotalPages] = useState(0)
  const [activeTab, setActiveTab] = useState('generator') // generator, editor
  const [outputMode, setOutputMode] = useState('single') // single, group

  const handleDrag = (e) => {
    e.preventDefault()
    e.stopPropagation()
    if (e.type === 'dragenter' || e.type === 'dragover') {
      setIsDragging(true)
    } else if (e.type === 'dragleave') {
      setIsDragging(false)
    }
  }

  const handleDrop = (e) => {
    e.preventDefault()
    e.stopPropagation()
    setIsDragging(false)
    
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const droppedFile = e.dataTransfer.files[0]
      if (droppedFile.name.endsWith('.xlsx')) {
        setFile(droppedFile)
        handleUpload(droppedFile)
      } else {
        setErrorMsg('Please upload a valid .xlsx file')
        setStatus('error')
      }
    }
  }

  const handleFileSelect = (e) => {
    if (e.target.files && e.target.files[0]) {
      const selectedFile = e.target.files[0]
      if (selectedFile.name.endsWith('.xlsx')) {
        setFile(selectedFile)
        handleUpload(selectedFile)
      } else {
        setErrorMsg('Please upload a valid .xlsx file')
        setStatus('error')
      }
    }
  }

  const handleUpload = async (uploadFile) => {
    setStatus('processing')
    
    const formData = new FormData()
    formData.append('file', uploadFile)
    formData.append('output_mode', outputMode)

    try {
      // Assuming backend runs on port 8090
      const apiUrl = import.meta.env.VITE_API_URL || `http://${window.location.hostname}:8090`;
      const response = await fetch(`${apiUrl}/upload`, {
        method: 'POST',
        body: formData,
      })

      if (!response.ok) {
        const errorData = await response.json()
        throw new Error(errorData.detail || 'Failed to process file')
      }

      const data = await response.json()
      
      if (data.total_pages) setTotalPages(data.total_pages)
      const finalUrl = data.file_url.startsWith('http') 
        ? data.file_url 
        : `${apiUrl}${data.file_url}`
      setPdfUrl(finalUrl)
      setStatus('success')
    } catch (err) {
      setErrorMsg(err.message)
      setStatus('error')
    }
  }

  const resetState = () => {
    setFile(null)
    setStatus('idle')
    setErrorMsg('')
    setPdfUrl(null)
    setShowPreview(true)
    setTotalPages(0)
  }

  return (
    <div className={`app-wrapper ${(status === 'success' && showPreview && activeTab === 'generator') || activeTab === 'editor' ? 'expanded' : ''}`}>
      
      {/* Top Navigation Bar */}
      <nav className="top-nav">
        <div className="nav-container">
          <div className="nav-logo">
            <CheckCircle size={24} color="#10b981" />
            <span>Vardhan QC</span>
          </div>
          <div className="nav-tabs">
            <button 
              className={`nav-tab ${activeTab === 'generator' ? 'active' : ''}`}
              onClick={() => setActiveTab('generator')}
            >
              <FileSpreadsheet size={18} /> Report Generator
            </button>
            <button 
              className={`nav-tab ${activeTab === 'editor' ? 'active' : ''}`}
              onClick={() => setActiveTab('editor')}
            >
              <Database size={18} /> Master Data Editor
            </button>
          </div>
        </div>
      </nav>

      {activeTab === 'editor' ? (
        <TestReadingsEditor />
      ) : (
      <div className="glass-panel">
      <h1>Quality Inspection Automation</h1>
      <p className="subtitle">Upload your Packing Slip to perfectly generate the Master PDF</p>

      {status === 'idle' && (
        <div 
          className={`upload-zone ${isDragging ? 'drag-active' : ''}`}
          onDragEnter={handleDrag}
          onDragLeave={handleDrag}
          onDragOver={handleDrag}
          onDrop={handleDrop}
          onClick={() => document.getElementById('file-upload').click()}
        >
          <Upload size={48} className="upload-icon" />
          <div className="upload-text">Drag & drop your Packing Slip here</div>
          <div className="upload-subtext">Only .xlsx files are supported</div>
          <input 
            type="file" 
            id="file-upload" 
            accept=".xlsx" 
            style={{ display: 'none' }}
            onChange={handleFileSelect}
          />
          
          <div className="output-mode-selector" style={{ marginTop: '1.5rem', display: 'flex', gap: '1rem', justifyContent: 'center' }} onClick={(e) => e.stopPropagation()}>
            <label style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', color: 'var(--text-secondary)' }}>
                <input 
                  type="radio" 
                  value="single" 
                  checked={outputMode === 'single'} 
                  onChange={(e) => setOutputMode(e.target.value)} 
                />
                Single Master PDF
            </label>
            <label style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', color: 'var(--text-secondary)' }}>
                <input 
                  type="radio" 
                  value="group" 
                  checked={outputMode === 'group'} 
                  onChange={(e) => setOutputMode(e.target.value)} 
                />
                Group-Wise ZIP
            </label>
          </div>
        </div>
      )}

      {status === 'processing' && (
        <div className="loading-container">
          <div className="spinner"></div>
          <div className="loading-text">Processing your Packing Slip...</div>
          <div className="upload-subtext">Generating Excel sheets & contacting Cloud API</div>
        </div>
      )}

      {status === 'error' && (
        <div className="loading-container">
          <RefreshCcw size={48} color="#ef4444" style={{marginBottom: '1rem'}} />
          <div className="loading-text" style={{color: '#ef4444'}}>Error Occurred</div>
          <div className="upload-subtext" style={{color: '#f87171'}}>{errorMsg}</div>
          <button className="reset-btn" onClick={resetState}>Try Again</button>
        </div>
      )}

      {status === 'success' && (
        <div className={`result-container ${showPreview ? 'split-layout' : ''}`}>
          <div className="result-actions">
            <CheckCircle size={64} className="success-icon" />
            <div className="loading-text">Success! Your PDF is ready.</div>
            <div className="upload-subtext">
              100% pixel-perfect formatting preserved.
              {totalPages > 0 && <div>Generated exactly <strong>{totalPages}</strong> pages!</div>}
            </div>
            
            <div className="button-group">
              <a href={pdfUrl} target="_blank" rel="noreferrer" download className="download-btn">
                <Download size={20} /> Download {outputMode === 'group' ? 'ZIP Archive' : 'PDF'}
              </a>
              {outputMode === 'single' && (
                  <button 
                    className="toggle-preview-btn" 
                    onClick={() => setShowPreview(!showPreview)}
                  >
                    {showPreview ? <><EyeOff size={20} /> Close Preview</> : <><Eye size={20} /> View Preview</>}
                  </button>
              )}
            </div>
            
            <button className="reset-btn" onClick={resetState}>Process Another File</button>
          </div>

          {showPreview && outputMode === 'single' && (
            <div className="pdf-preview-container">
              <iframe 
                src={`${pdfUrl}#toolbar=0&navpanes=0&scrollbar=0`} 
                className="pdf-preview" 
                title="PDF Preview"
              />
            </div>
          )}
        </div>
      )}
    </div>
    )}
    </div>
  )
}

export default App
