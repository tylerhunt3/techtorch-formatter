'use client'

import { useState, useRef } from 'react'
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  HeadingLevel,
  AlignmentType,
  BorderStyle,
  WidthType,
  ShadingType,
  PageBreak,
  LevelFormat,
} from 'docx'
import mammoth from 'mammoth'
import { saveAs } from 'file-saver'

// ============================================================================
// TECHTORCH FORMATTING STANDARDS
// ============================================================================

const COLORS = {
  HEADING1: '1F4E79',
  HEADING2: '2E75B6',
  HEADING3: '404040',
  BODY: '000000',
  SECONDARY: '666666',
  CODE_TEXT: '2E2E2E',
  CODE_BG: 'F5F5F5',
  CODE_BORDER: 'BFBFBF',
  CODE_ACCENT: '1F4E79',
  TABLE_HEADER_BG: '1F4E79',
  TABLE_BORDER: 'CCCCCC',
  WHITE: 'FFFFFF',
}

const SIZES = {
  TITLE: 48,
  SUBTITLE: 28,
  HEADING1: 24,
  HEADING2: 22,
  HEADING3: 20,
  BODY: 18,
  CODE: 16,
}

// ============================================================================
// CONTENT EXTRACTION
// ============================================================================

interface ContentItem {
  type: 'title' | 'heading1' | 'heading2' | 'heading3' | 'paragraph' | 'bullet' | 'code_block' | 'table'
  text?: string
  lines?: string[]
  data?: string[][]
}

function isCodeBlock(text: string): boolean {
  const codeIndicators = ['SELECT', 'FROM', 'WHERE', 'INSERT', 'UPDATE', 'DELETE', 'VAR', 'RETURN']
  const upperText = text.toUpperCase()
  const matches = codeIndicators.filter(ind => upperText.includes(ind)).length
  const lines = text.split('\n')
  const indentedLines = lines.filter(line => line.startsWith('    ') || line.startsWith('\t')).length
  return matches >= 2 || indentedLines >= 3
}

function detectContentType(text: string, prevType: string | null): ContentItem['type'] {
  const trimmed = text.trim()
  
  if (/^\d+\.\d+\.\d+\s+[A-Z]/.test(trimmed)) return 'heading3'
  if (/^\d+\.\d+\s+[A-Z]/.test(trimmed)) return 'heading2'
  if (/^\d+\.\s+[A-Z]/.test(trimmed)) return 'heading1'
  
  const headerKeywords = ['Summary', 'Conclusion', 'Overview', 'Introduction', 'Final Summary', 'Next Steps', 'Background', 'Recommendations']
  for (const keyword of headerKeywords) {
    if (trimmed === keyword || trimmed.startsWith(keyword + ':')) return 'heading1'
  }
  
  if (trimmed.startsWith('•') || trimmed.startsWith('-') || trimmed.startsWith('*')) return 'bullet'
  if (trimmed.startsWith('Result:')) return 'bullet'
  
  return 'paragraph'
}

async function extractContent(file: File): Promise<ContentItem[]> {
  const arrayBuffer = await file.arrayBuffer()
  const result = await mammoth.extractRawText({ arrayBuffer })
  const rawText = result.value
  
  const lines = rawText.split('\n').filter(line => line.trim())
  const content: ContentItem[] = []
  let codeBuffer: string[] = []
  let inCodeBlock = false
  let prevType: string | null = null
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim()
    if (!line) continue
    
    const looksLikeCode = isCodeBlock(line) || 
      (line.toUpperCase().startsWith('SELECT') || 
       line.toUpperCase().startsWith('FROM') || 
       line.toUpperCase().startsWith('WHERE'))
    
    if (looksLikeCode && !inCodeBlock) {
      inCodeBlock = true
      codeBuffer = [line]
      continue
    }
    
    if (inCodeBlock) {
      const isStillCode = line.startsWith(' ') || 
        line.toUpperCase().startsWith('SELECT') ||
        line.toUpperCase().startsWith('FROM') ||
        line.toUpperCase().startsWith('WHERE') ||
        line.toUpperCase().startsWith('AND') ||
        line.toUpperCase().startsWith('OR') ||
        line.includes('__c') ||
        line.includes("'006") ||
        line.startsWith(')') ||
        /^[A-Za-z_]+,?$/.test(line) ||
        /^\s/.test(lines[i])
      
      if (isStillCode) {
        codeBuffer.push(line)
        continue
      } else {
        if (codeBuffer.length > 0) {
          content.push({ type: 'code_block', lines: codeBuffer })
        }
        inCodeBlock = false
        codeBuffer = []
      }
    }
    
    const contentType = detectContentType(line, prevType)
    
    if (contentType === 'bullet') {
      const cleanText = line.replace(/^[•\-*]\s*/, '')
      content.push({ type: 'bullet', text: cleanText })
    } else {
      content.push({ type: contentType, text: line })
    }
    
    prevType = contentType
  }
  
  if (codeBuffer.length > 0) {
    content.push({ type: 'code_block', lines: codeBuffer })
  }
  
  return content
}

// ============================================================================
// DOCUMENT CREATION
// ============================================================================

function createCodeBlockTable(lines: string[]): Table {
  const tableBorder = { style: BorderStyle.SINGLE, size: 1, color: COLORS.CODE_BORDER }
  const accentBorder = { style: BorderStyle.SINGLE, size: 12, color: COLORS.CODE_ACCENT }
  const noBorder = { style: BorderStyle.NIL, size: 0, color: COLORS.WHITE }
  
  return new Table({
    columnWidths: [9360],
    rows: lines.map((line, index) => 
      new TableRow({
        children: [
          new TableCell({
            width: { size: 9360, type: WidthType.DXA },
            shading: { fill: COLORS.CODE_BG, type: ShadingType.CLEAR },
            borders: {
              left: accentBorder,
              right: tableBorder,
              top: index === 0 ? tableBorder : noBorder,
              bottom: index === lines.length - 1 ? tableBorder : noBorder,
            },
            children: [
              new Paragraph({
                spacing: { before: 20, after: 20 },
                children: [
                  new TextRun({
                    text: line,
                    font: 'Consolas',
                    size: SIZES.CODE,
                    color: COLORS.CODE_TEXT,
                  }),
                ],
              }),
            ],
          }),
        ],
      })
    ),
  })
}

function createFormattedDocument(
  content: ContentItem[],
  docTitle: string,
  organization: string
): Document {
  const children: (Paragraph | Table)[] = []
  
  for (let i = 0; i < 4; i++) {
    children.push(new Paragraph({ children: [] }))
  }
  
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({
          text: docTitle,
          bold: true,
          font: 'Aptos',
          size: SIZES.TITLE,
          color: COLORS.BODY,
        }),
      ],
    })
  )
  
  let subtitleText: string | null = null
  let skipFirstHeading = false
  for (const item of content) {
    if (item.type === 'heading1' && item.text && item.text !== docTitle) {
      subtitleText = item.text
      skipFirstHeading = true
      break
    }
    if (item.type === 'paragraph' && item.text && item.text.includes(' - ')) {
      subtitleText = item.text
      break
    }
  }
  
  if (subtitleText) {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
        children: [
          new TextRun({
            text: subtitleText,
            font: 'Aptos',
            size: SIZES.SUBTITLE,
            color: COLORS.BODY,
          }),
        ],
      })
    )
  }
  
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [
        new TextRun({
          text: organization,
          font: 'Aptos',
          size: 24,
          italics: true,
          color: COLORS.BODY,
        }),
      ],
    })
  )
  
  const currentDate = new Date().toLocaleDateString('en-US', { month: 'long', year: 'numeric' })
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [
        new TextRun({
          text: `As of ${currentDate}`,
          font: 'Aptos',
          size: 20,
          color: COLORS.BODY,
        }),
      ],
    })
  )
  
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({
          text: 'Version 1.0',
          font: 'Aptos',
          size: SIZES.BODY,
          color: COLORS.SECONDARY,
        }),
      ],
    })
  )
  
  children.push(new Paragraph({ children: [new PageBreak()] }))
  
  let isFirstHeading = true
  
  for (const item of content) {
    switch (item.type) {
      case 'heading1':
        if (skipFirstHeading && isFirstHeading) {
          isFirstHeading = false
          continue
        }
        isFirstHeading = false
        children.push(
          new Paragraph({
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 300, after: 100 },
            children: [
              new TextRun({
                text: item.text || '',
                bold: true,
                font: 'Aptos',
                size: SIZES.HEADING1,
                color: COLORS.HEADING1,
              }),
            ],
          })
        )
        break
        
      case 'heading2':
        children.push(
          new Paragraph({
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 200, after: 80 },
            children: [
              new TextRun({
                text: item.text || '',
                bold: true,
                font: 'Aptos',
                size: SIZES.HEADING2,
                color: COLORS.HEADING2,
              }),
            ],
          })
        )
        break
        
      case 'heading3':
        children.push(
          new Paragraph({
            heading: HeadingLevel.HEADING_3,
            spacing: { before: 160, after: 60 },
            children: [
              new TextRun({
                text: item.text || '',
                bold: true,
                font: 'Aptos',
                size: SIZES.HEADING3,
                color: COLORS.HEADING3,
              }),
            ],
          })
        )
        break
        
      case 'paragraph':
        if (subtitleText && item.text === subtitleText) continue
        children.push(
          new Paragraph({
            spacing: { after: 160 },
            children: [
              new TextRun({
                text: item.text || '',
                font: 'Aptos',
                size: SIZES.BODY,
                color: COLORS.BODY,
              }),
            ],
          })
        )
        break
        
      case 'bullet':
        children.push(
          new Paragraph({
            bullet: { level: 0 },
            spacing: { after: 80 },
            children: [
              new TextRun({
                text: item.text || '',
                font: 'Aptos',
                size: SIZES.BODY,
                color: COLORS.BODY,
              }),
            ],
          })
        )
        break
        
      case 'code_block':
        if (item.lines && item.lines.length > 0) {
          children.push(new Paragraph({ children: [] }))
          children.push(createCodeBlockTable(item.lines))
          children.push(new Paragraph({ children: [] }))
        }
        break
    }
  }
  
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 400 },
      children: [
        new TextRun({
          text: 'End of Document',
          font: 'Aptos',
          size: SIZES.BODY,
          italics: true,
          color: COLORS.SECONDARY,
        }),
      ],
    })
  )
  
  return new Document({
    styles: {
      default: {
        document: {
          run: {
            font: 'Aptos',
            size: SIZES.BODY,
          },
        },
      },
      paragraphStyles: [
        {
          id: 'Heading1',
          name: 'Heading 1',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: {
            font: 'Aptos',
            size: SIZES.HEADING1,
            bold: true,
            color: COLORS.HEADING1,
          },
        },
        {
          id: 'Heading2',
          name: 'Heading 2',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: {
            font: 'Aptos',
            size: SIZES.HEADING2,
            bold: true,
            color: COLORS.HEADING2,
          },
        },
        {
          id: 'Heading3',
          name: 'Heading 3',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: {
            font: 'Aptos',
            size: SIZES.HEADING3,
            bold: true,
            color: COLORS.HEADING3,
          },
        },
      ],
    },
    numbering: {
      config: [
        {
          reference: 'bullet-list',
          levels: [
            {
              level: 0,
              format: LevelFormat.BULLET,
              text: '•',
              alignment: AlignmentType.LEFT,
              style: {
                paragraph: {
                  indent: { left: 720, hanging: 360 },
                },
              },
            },
          ],
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 1440,
              right: 1440,
              bottom: 1440,
              left: 1440,
            },
          },
        },
        children: children,
      },
    ],
  })
}

// ============================================================================
// REACT COMPONENT
// ============================================================================

export default function Home() {
  const [file, setFile] = useState<File | null>(null)
  const [docTitle, setDocTitle] = useState('')
  const [organization, setOrganization] = useState('TechTorch Inc.')
  const [status, setStatus] = useState<'idle' | 'processing' | 'success' | 'error'>('idle')
  const [statusMessage, setStatusMessage] = useState('')
  const [downloadReady, setDownloadReady] = useState(false)
  const [docBlob, setDocBlob] = useState<Blob | null>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)
  
  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0]
    if (selectedFile) {
      setFile(selectedFile)
      setStatus('idle')
      setDownloadReady(false)
    }
  }
  
  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault()
    const droppedFile = e.dataTransfer.files?.[0]
    if (droppedFile && droppedFile.name.endsWith('.docx')) {
      setFile(droppedFile)
      setStatus('idle')
      setDownloadReady(false)
    }
  }
  
  const handleFormat = async () => {
    if (!file || !docTitle) return
    
    setStatus('processing')
    setStatusMessage('Extracting content from document...')
    
    try {
      const content = await extractContent(file)
      
      setStatusMessage('Applying TechTorch formatting standards...')
      const doc = createFormattedDocument(content, docTitle, organization)
      
      setStatusMessage('Generating document...')
      const blob = await Packer.toBlob(doc)
      
      setDocBlob(blob)
      setDownloadReady(true)
      setStatus('success')
      
      const headings = content.filter(c => c.type.startsWith('heading')).length
      const bullets = content.filter(c => c.type === 'bullet').length
      const codeBlocks = content.filter(c => c.type === 'code_block').length
      
      setStatusMessage(`Formatted: ${headings} headings, ${bullets} bullet points, ${codeBlocks} code blocks`)
    } catch (err) {
      setStatus('error')
      setStatusMessage(`Error: ${err instanceof Error ? err.message : 'Unknown error'}`)
    }
  }
  
  const handleDownload = () => {
    if (docBlob) {
      const cleanTitle = docTitle.replace(/[^\w\s-]/g, '').replace(/\s+/g, '_')
      saveAs(docBlob, `${cleanTitle}_Formatted.docx`)
    }
  }
  
  return (
    <div className="container">
      <div className="card">
        <div className="header">
          <div className="logo">📄</div>
          <h1 className="title">TechTorch Document Formatter</h1>
          <p className="subtitle">Upload a Word document to apply professional formatting standards</p>
        </div>
        
        <div className="divider" />
        
        <div
          className={`upload-area ${file ? 'has-file' : ''}`}
          onClick={() => fileInputRef.current?.click()}
          onDrop={handleDrop}
          onDragOver={(e) => e.preventDefault()}
        >
          <div className="upload-icon">{file ? '✅' : '📁'}</div>
          <p className="upload-text">
            {file ? 'File selected' : 'Click or drag to upload'}
          </p>
          <p className="upload-hint">.docx files only</p>
          {file && <p className="file-name">{file.name}</p>}
          <input
            ref={fileInputRef}
            type="file"
            accept=".docx"
            onChange={handleFileSelect}
            className="hidden"
          />
        </div>
        
        <div className="divider" />
        
        <div className="form-group">
          <label className="label">Document Title</label>
          <input
            type="text"
            className="input"
            placeholder="Enter the document title..."
            value={docTitle}
            onChange={(e) => setDocTitle(e.target.value)}
          />
        </div>
        
        <div className="form-group">
          <label className="label">Organization</label>
          <input
            type="text"
            className="input"
            value={organization}
            onChange={(e) => setOrganization(e.target.value)}
          />
        </div>
        
        <button
          className="button"
          onClick={handleFormat}
          disabled={!file || !docTitle || status === 'processing'}
        >
          {status === 'processing' ? 'Processing...' : 'Format Document'}
        </button>
        
        {downloadReady && (
          <button className="button button-secondary" onClick={handleDownload}>
            Download Formatted Document
          </button>
        )}
        
        {status !== 'idle' && (
          <div className={`status ${status}`}>
            {status === 'processing' && '⏳ '}
            {status === 'success' && '✅ '}
            {status === 'error' && '❌ '}
            {statusMessage}
          </div>
        )}
        
        <div className="info-box">
          <p><strong>Applies TechTorch standards:</strong></p>
          <p>• Aptos font family with proper sizing hierarchy</p>
          <p>• Blue heading colors (#1F4E79, #2E75B6)</p>
          <p>• Formatted code blocks with accent borders</p>
          <p>• Professional title page layout</p>
        </div>
        
        <div className="footer">
          TechTorch Documentation Formatting Tool v2.0
        </div>
      </div>
    </div>
  )
}
