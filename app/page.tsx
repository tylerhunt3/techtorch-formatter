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
  TableOfContents,
  Footer,
  Header,
  PageNumber,
  NumberFormat,
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
// INTELLIGENT CONTENT DETECTION - Pattern Recognition & Heuristics
// ============================================================================

interface ContentItem {
  type: 'heading1' | 'heading2' | 'heading3' | 'paragraph' | 'bullet' | 'numbered' | 'code_block' | 'sub_bullet'
  text?: string
  lines?: string[]
  number?: number
  confidence?: number // How confident we are in this classification
}

// Header keywords that strongly indicate a section header
const HEADER_KEYWORDS = [
  'Introduction', 'Conclusion', 'Conclusions', 'Summary', 'Overview',
  'Background', 'Methodology', 'Results', 'Discussion', 'Recommendations',
  'Next Steps', 'Appendix', 'References', 'Glossary', 'Final Summary',
  'Development Tools', 'Development tools', 'Major Focus Areas',
  'Key Findings', 'Executive Summary', 'Scope', 'Objectives', 'Purpose',
  'Requirements', 'Implementation', 'Architecture', 'Configuration',
  'Troubleshooting', 'Best Practices', 'Limitations', 'Future Work',
  'Acknowledgments', 'Prerequisites', 'Dependencies', 'Setup', 'Installation',
  'Working Model', 'Initial Project Phase', 'Project Phase',
]

// Patterns that indicate the START of a list (next lines are likely bullets)
const LIST_INTRO_PATTERNS = [
  /the following[:\s]/i,
  /as follows[:\s]/i,
  /include[sd]?[:\s]/i,
  /such as[:\s]/i,
  /for example[:\s]/i,
  /listed below[:\s]/i,
  /steps? (are|were|is|was)[:\s]/i,
  /challenges? (are|were|is|was|included)[:\s]/i,
  /characteristics? (are|were|is|was|included)[:\s]/i,
  /reasons? (are|were|is|was|included)[:\s]/i,
  /benefits? (are|were|is|was|included)[:\s]/i,
  /features? (are|were|is|was|included)[:\s]/i,
  /requirements? (are|were|is|was|included)[:\s]/i,
  /components? (are|were|is|was|included)[:\s]/i,
  /:[:\s]*$/,  // Ends with colon
]

// Patterns that indicate code/technical content
const CODE_PATTERNS = [
  /^(SELECT|FROM|WHERE|INSERT|UPDATE|DELETE|CREATE|ALTER|DROP)\s/i,
  /\b(System\.debug|Console\.log|print\(|println\()/,
  /\b(try\s*\{|catch\s*\(|finally\s*\{|if\s*\(|else\s*\{|for\s*\(|while\s*\()/,
  /\b(new\s+\w+\s*\(|\.get\w+\s*\(|\.set\w+\s*\()/,
  /^\s*(Id|String|Integer|Boolean|List<|Map<|Set<|var|let|const)\s+\w+\s*=/,
  /(__c|__r|__mdt|__e)\b/,  // Salesforce custom fields
  /\b(Database\.|Schema\.|Apex|SOQL|SOSL)\b/i,
  /^\s*\}\s*(catch|finally|else)?\s*[\{]?\s*$/,
  /^\s*\/\/.*$/,  // Comments
  /^\s*\/\*.*\*\/\s*$/,  // Block comments
  /^\s*\*\s/,  // Javadoc style
  /=>\s*\{/,  // Arrow functions
  /\bfunction\s*\(/,
  /\bclass\s+\w+/,
  /\bimport\s+/,
  /\bexport\s+(default\s+)?/,
  /^\s*@\w+/,  // Decorators/annotations
  /\bSavepoint\b|\brollback\b/i,
  /\bDmlException\b|\bException\b/,
]

// Patterns for bullet-like content (even without bullet markers)
const BULLET_CONTENT_PATTERNS = [
  /^[A-Z][a-z]+ing\s/,  // Starts with gerund (e.g., "Identifying", "Processing")
  /^[A-Z][a-z]+ed\s/,   // Starts with past tense
  /^(How|What|When|Where|Why|Which)\s/i,  // Question words
  /^(The|A|An)\s+\w+\s+(was|were|is|are|has|have)\s/,  // Passive constructions
  /^(No|Missing|Incorrect|Invalid|Duplicate)\s/i,  // Status indicators
  /^(Enable|Disable|Configure|Set|Get|Create|Delete|Update|Review)\s/i,  // Action verbs
]

// Check if text looks like a header based on multiple signals
function analyzeHeaderProbability(text: string, prevText: string | null, nextText: string | null): number {
  let score = 0
  const trimmed = text.trim()
  
  // Length check - headers are usually short
  if (trimmed.length < 80) score += 15
  if (trimmed.length < 50) score += 10
  if (trimmed.length > 150) score -= 30
  
  // Numbered section pattern (strongest signal)
  if (/^\d+\.\d+\.\d+\s+[A-Z]/.test(trimmed)) return 95  // Definitely H3
  if (/^\d+\.\d+\s+[A-Z]/.test(trimmed)) return 95       // Definitely H2
  if (/^\d+\.\s+[A-Z]/.test(trimmed)) return 95          // Definitely H1
  
  // Keyword match
  for (const keyword of HEADER_KEYWORDS) {
    if (trimmed === keyword) return 90
    if (trimmed.startsWith(keyword + ':')) return 85
    if (trimmed.startsWith(keyword + ' ')) score += 20
    if (trimmed.toLowerCase().includes(keyword.toLowerCase())) score += 10
  }
  
  // Title case pattern (Most Words Capitalized)
  const words = trimmed.split(/\s+/)
  const capitalizedWords = words.filter(w => /^[A-Z]/.test(w))
  if (words.length >= 2 && capitalizedWords.length >= words.length * 0.7) score += 15
  
  // No ending punctuation (headers usually don't end with period)
  if (!trimmed.endsWith('.') && !trimmed.endsWith(',')) score += 10
  if (trimmed.endsWith('.')) score -= 15
  
  // Previous line ends with colon or is empty (common before headers)
  if (prevText === null || prevText.trim() === '') score += 10
  
  // Next line is much longer (body text follows header)
  if (nextText && nextText.length > trimmed.length * 2) score += 10
  
  // Contains colon mid-text (like "Screen 1: Selection")
  if (/^[\w\s]+:\s*[\w\s]+$/.test(trimmed) && trimmed.length < 60) score += 20
  
  // All caps (sometimes used for headers)
  if (trimmed === trimmed.toUpperCase() && trimmed.length > 3 && trimmed.length < 50) score += 15
  
  return Math.min(score, 100)
}

// Check if text is likely part of a code block
function analyzeCodeProbability(text: string, context: string[]): number {
  let score = 0
  const trimmed = text.trim()
  
  // Direct pattern matches
  for (const pattern of CODE_PATTERNS) {
    if (pattern.test(trimmed)) score += 30
  }
  
  // Indentation (code is often indented)
  if (text.startsWith('    ') || text.startsWith('\t')) score += 20
  
  // Contains typical code characters
  if (/[{}\[\]();]/.test(trimmed)) score += 15
  if (/\b\w+\.\w+\(/.test(trimmed)) score += 15  // Method calls
  if (/=\s*['"]/.test(trimmed)) score += 10  // String assignment
  if (/=\s*\d+/.test(trimmed)) score += 10  // Number assignment
  
  // Salesforce specific
  if (/blng__|SBQQ__|npsp__/.test(trimmed)) score += 25
  
  // Very short lines that look like code fragments
  if (trimmed.length < 30 && /^[\w_]+[,;]?$/.test(trimmed)) score += 15
  
  // Check context - if surrounded by code-like lines
  const codeNeighbors = context.filter(c => CODE_PATTERNS.some(p => p.test(c))).length
  if (codeNeighbors >= 2) score += 20
  
  return Math.min(score, 100)
}

// Check if text should be a bullet point
function analyzeBulletProbability(text: string, prevText: string | null, isAfterListIntro: boolean): number {
  let score = 0
  const trimmed = text.trim()
  
  // Explicit bullet markers
  if (/^[•\-\*\u2022\u2023\u25E6\u2043\u2219]\s/.test(trimmed)) return 95
  
  // If previous line introduced a list
  if (isAfterListIntro) score += 40
  
  // Short, self-contained statement
  if (trimmed.length < 100 && trimmed.length > 10) score += 10
  
  // Starts with action verb or gerund (common in bullets)
  for (const pattern of BULLET_CONTENT_PATTERNS) {
    if (pattern.test(trimmed)) score += 15
  }
  
  // Doesn't start with "The" or "This" (more likely body text)
  if (/^(The|This|That|These|Those|It)\s/.test(trimmed) && trimmed.length > 80) score -= 20
  
  // Multiple short lines in sequence are likely bullets
  if (prevText && prevText.length < 100 && trimmed.length < 100) score += 10
  
  // Ends without period but has content
  if (!trimmed.endsWith('.') && trimmed.length > 20) score += 5
  
  return Math.min(score, 100)
}

// Determine heading level based on numbering pattern
function getHeadingLevel(text: string): 'heading1' | 'heading2' | 'heading3' | null {
  const trimmed = text.trim()
  
  if (/^\d+\.\d+\.\d+\s/.test(trimmed)) return 'heading3'
  if (/^\d+\.\d+\s/.test(trimmed)) return 'heading2'
  if (/^\d+\.\s/.test(trimmed)) return 'heading1'
  
  // Check for letter-based numbering
  if (/^[a-z]\)\s/i.test(trimmed)) return 'heading3'
  if (/^[ivxIVX]+\.\s/.test(trimmed)) return 'heading2'
  
  return null
}

// Main content extraction with intelligent classification
async function extractContent(file: File): Promise<ContentItem[]> {
  const arrayBuffer = await file.arrayBuffer()
  
  // Try HTML conversion first for better structure
  let lines: string[] = []
  try {
    const htmlResult = await mammoth.convertToHtml({ arrayBuffer })
    const html = htmlResult.value
    lines = parseHtmlToLines(html)
  } catch {
    // Fallback to raw text
    const textResult = await mammoth.extractRawText({ arrayBuffer })
    lines = textResult.value.split('\n')
  }
  
  // Filter empty lines but track where they were (for context)
  const nonEmptyLines = lines.map((line, i) => ({ text: line.trim(), originalIndex: i, wasEmpty: line.trim() === '' }))
    .filter(l => l.text !== '')
  
  const content: ContentItem[] = []
  let i = 0
  let numberedListCounter = 0
  let lastListIntroIndex = -10  // Track when we saw a list introduction
  
  while (i < nonEmptyLines.length) {
    const current = nonEmptyLines[i]
    const text = current.text
    const prev = i > 0 ? nonEmptyLines[i - 1].text : null
    const next = i < nonEmptyLines.length - 1 ? nonEmptyLines[i + 1].text : null
    
    // Get context (surrounding lines)
    const context = nonEmptyLines.slice(Math.max(0, i - 3), Math.min(nonEmptyLines.length, i + 4)).map(l => l.text)
    
    // Check if this line introduces a list
    const isListIntro = LIST_INTRO_PATTERNS.some(p => p.test(text))
    if (isListIntro) lastListIntroIndex = i
    
    const isAfterListIntro = i - lastListIntroIndex <= 5 && i > lastListIntroIndex
    
    // === CODE BLOCK DETECTION ===
    const codeProb = analyzeCodeProbability(text, context)
    if (codeProb >= 60) {
      // Collect consecutive code lines
      const codeLines: string[] = [text]
      let j = i + 1
      while (j < nonEmptyLines.length) {
        const nextLine = nonEmptyLines[j].text
        const nextCodeProb = analyzeCodeProbability(nextLine, context)
        // Continue if it's code OR if it's a short line between code lines
        if (nextCodeProb >= 40 || (nextLine.length < 30 && j + 1 < nonEmptyLines.length && analyzeCodeProbability(nonEmptyLines[j + 1].text, context) >= 50)) {
          codeLines.push(nextLine)
          j++
        } else {
          break
        }
      }
      
      if (codeLines.length >= 2 || codeProb >= 80) {
        content.push({ type: 'code_block', lines: codeLines, confidence: codeProb })
        i = j
        numberedListCounter = 0
        continue
      }
    }
    
    // === HEADER DETECTION ===
    const headerProb = analyzeHeaderProbability(text, prev, next)
    const headingLevel = getHeadingLevel(text)
    
    if (headingLevel) {
      content.push({ type: headingLevel, text, confidence: 95 })
      i++
      numberedListCounter = 0
      continue
    }
    
    if (headerProb >= 70) {
      // Determine level based on context and patterns
      let level: 'heading1' | 'heading2' | 'heading3' = 'heading1'
      
      // If it's a sub-topic keyword, make it H2
      const subTopicKeywords = ['Example', 'Screen', 'Step', 'Phase', 'Part', 'Section', 'Case', 'Scenario']
      if (subTopicKeywords.some(k => text.startsWith(k))) level = 'heading2'
      
      // If very short and after another header, might be H3
      if (text.length < 30 && prev && analyzeHeaderProbability(prev, null, text) >= 50) level = 'heading3'
      
      content.push({ type: level, text, confidence: headerProb })
      i++
      numberedListCounter = 0
      continue
    }
    
    // === NUMBERED LIST DETECTION ===
    const numberedMatch = text.match(/^(\d+)[.)]\s+(.+)/)
    if (numberedMatch) {
      const num = parseInt(numberedMatch[1])
      // Check if this continues a sequence or starts a new one
      if (num === 1 || num === numberedListCounter + 1) {
        numberedListCounter = num
        content.push({ type: 'numbered', text: numberedMatch[2], number: num })
        i++
        continue
      }
    }
    
    // === BULLET DETECTION ===
    const bulletProb = analyzeBulletProbability(text, prev, isAfterListIntro)
    
    // Check for explicit bullet markers
    const bulletMatch = text.match(/^[•\-\*\u2022]\s*(.+)/)
    if (bulletMatch) {
      content.push({ type: 'bullet', text: bulletMatch[1], confidence: 95 })
      i++
      continue
    }
    
    // Check for sub-bullets (indented or with different markers)
    const subBulletMatch = text.match(/^[\u25E6\u2023\u2043\u2219\-]\s*(.+)/)
    if (subBulletMatch && content.length > 0 && content[content.length - 1].type === 'bullet') {
      content.push({ type: 'sub_bullet', text: subBulletMatch[1], confidence: 90 })
      i++
      continue
    }
    
    // Infer bullet from context
    if (bulletProb >= 60 && text.length < 120) {
      content.push({ type: 'bullet', text, confidence: bulletProb })
      i++
      continue
    }
    
    // === DEFAULT: PARAGRAPH ===
    // Reset numbered counter if we hit a paragraph
    numberedListCounter = 0
    content.push({ type: 'paragraph', text, confidence: 100 - headerProb - bulletProb })
    i++
  }
  
  // Post-processing: Fix sequences that should be bullets
  return postProcessContent(content)
}

// Parse HTML to extract lines with some structure awareness
function parseHtmlToLines(html: string): string[] {
  const lines: string[] = []
  
  // Simple regex-based extraction (works in browser without DOM)
  // Remove scripts and styles
  let cleaned = html.replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
  cleaned = cleaned.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
  
  // Convert block elements to newlines
  cleaned = cleaned.replace(/<\/(p|div|h[1-6]|li|tr|br)[^>]*>/gi, '\n')
  cleaned = cleaned.replace(/<(p|div|h[1-6]|li|tr|br)[^>]*>/gi, '\n')
  
  // Handle list items specially - add bullet marker
  cleaned = cleaned.replace(/<li[^>]*>/gi, '\n• ')
  
  // Remove remaining tags
  cleaned = cleaned.replace(/<[^>]+>/g, '')
  
  // Decode HTML entities
  cleaned = cleaned.replace(/&amp;/g, '&')
  cleaned = cleaned.replace(/&lt;/g, '<')
  cleaned = cleaned.replace(/&gt;/g, '>')
  cleaned = cleaned.replace(/&quot;/g, '"')
  cleaned = cleaned.replace(/&#39;/g, "'")
  cleaned = cleaned.replace(/&nbsp;/g, ' ')
  
  // Split into lines
  return cleaned.split('\n').map(l => l.trim()).filter(l => l)
}

// Post-process to fix common patterns
function postProcessContent(content: ContentItem[]): ContentItem[] {
  const result: ContentItem[] = []
  
  for (let i = 0; i < content.length; i++) {
    const item = content[i]
    const prev = i > 0 ? result[result.length - 1] : null
    const next = i < content.length - 1 ? content[i + 1] : null
    
    // If we have multiple short paragraphs in sequence after a list intro, convert to bullets
    if (item.type === 'paragraph' && item.text && item.text.length < 100) {
      // Check if previous item ends with colon or is a list intro
      if (prev && prev.text && LIST_INTRO_PATTERNS.some(p => p.test(prev.text || ''))) {
        result.push({ ...item, type: 'bullet' })
        continue
      }
      
      // Check if surrounded by bullets
      if (prev?.type === 'bullet' && next?.type === 'bullet') {
        result.push({ ...item, type: 'bullet' })
        continue
      }
    }
    
    // If a "paragraph" is very short and follows a heading, might be a subtitle - keep as paragraph
    // but we handle this in document creation
    
    result.push(item)
  }
  
  return result
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
  subtitle: string,
  organization: string,
  author: string
): Document {
  const children: (Paragraph | Table)[] = []
  
  // ---- TITLE PAGE ----
  for (let i = 0; i < 6; i++) {
    children.push(new Paragraph({ children: [] }))
  }
  
  // Document title
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 240 },
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
  
  // Subtitle
  if (subtitle) {
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 240 },
        children: [
          new TextRun({
            text: subtitle,
            font: 'Aptos',
            size: SIZES.SUBTITLE,
            color: COLORS.BODY,
          }),
        ],
      })
    )
  }
  
  // Organization
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 120 },
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
  
  // Date
  const currentDate = new Date().toLocaleDateString('en-US', { month: 'long', year: 'numeric' })
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 120 },
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
  
  // Version with author
  const versionText = author ? `Version 1.0 (${author})` : 'Version 1.0'
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({
          text: versionText,
          font: 'Aptos',
          size: SIZES.BODY,
          color: COLORS.SECONDARY,
        }),
      ],
    })
  )
  
  // Page break after title
  children.push(new Paragraph({ children: [new PageBreak()] }))
  
  // ---- TABLE OF CONTENTS ----
  children.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 300 },
      children: [
        new TextRun({
          text: 'Table of Contents',
          bold: true,
          font: 'Aptos',
          size: SIZES.SUBTITLE,
          color: COLORS.HEADING1,
        }),
      ],
    })
  )
  
  children.push(
    new TableOfContents("Table of Contents", {
      hyperlink: true,
      headingStyleRange: "1-3",
    })
  )
  
  children.push(new Paragraph({ children: [new PageBreak()] }))
  
  // ---- MAIN CONTENT ----
  let lastType: string | null = null
  
  for (const item of content) {
    switch (item.type) {
      case 'heading1':
        // Add extra space before H1 if coming from body content
        if (lastType && lastType !== 'heading1') {
          children.push(new Paragraph({ spacing: { before: 200 }, children: [] }))
        }
        children.push(
          new Paragraph({
            heading: HeadingLevel.HEADING_1,
            spacing: { before: 300, after: 120 },
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
            spacing: { before: 240, after: 100 },
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
            spacing: { before: 200, after: 80 },
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
            indent: { left: 720, hanging: 360 },
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
        
      case 'sub_bullet':
        children.push(
          new Paragraph({
            bullet: { level: 1 },
            spacing: { after: 60 },
            indent: { left: 1080, hanging: 360 },
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
        
      case 'numbered':
        children.push(
          new Paragraph({
            spacing: { after: 80 },
            indent: { left: 720, hanging: 360 },
            children: [
              new TextRun({
                text: `${item.number}. `,
                bold: true,
                font: 'Aptos',
                size: SIZES.BODY,
                color: COLORS.BODY,
              }),
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
          children.push(new Paragraph({ spacing: { before: 120, after: 60 }, children: [] }))
          children.push(createCodeBlockTable(item.lines))
          children.push(new Paragraph({ spacing: { before: 60, after: 120 }, children: [] }))
        }
        break
    }
    
    lastType = item.type
  }
  
  // ---- END OF DOCUMENT ----
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
          paragraph: {
            spacing: { before: 300, after: 120 },
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
          paragraph: {
            spacing: { before: 240, after: 100 },
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
          paragraph: {
            spacing: { before: 200, after: 80 },
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
              text: '\u2022',
              alignment: AlignmentType.LEFT,
              style: {
                paragraph: {
                  indent: { left: 720, hanging: 360 },
                },
              },
            },
            {
              level: 1,
              format: LevelFormat.BULLET,
              text: '\u25E6',
              alignment: AlignmentType.LEFT,
              style: {
                paragraph: {
                  indent: { left: 1080, hanging: 360 },
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
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    children: [PageNumber.CURRENT],
                    font: 'Aptos',
                    size: 18,
                    color: COLORS.SECONDARY,
                  }),
                ],
              }),
            ],
          }),
        },
        children: children,
      },
    ],
  })
}

// ============================================================================
// REACT COMPONENT
// ============================================================================

interface Stats {
  headings: number
  bullets: number
  numbered: number
  codeBlocks: number
  paragraphs: number
}

export default function Home() {
  const [file, setFile] = useState<File | null>(null)
  const [docTitle, setDocTitle] = useState('')
  const [subtitle, setSubtitle] = useState('')
  const [organization, setOrganization] = useState('')
  const [author, setAuthor] = useState('')
  const [status, setStatus] = useState<'idle' | 'processing' | 'success' | 'error'>('idle')
  const [statusMessage, setStatusMessage] = useState('')
  const [downloadReady, setDownloadReady] = useState(false)
  const [docBlob, setDocBlob] = useState<Blob | null>(null)
  const [stats, setStats] = useState<Stats | null>(null)
  const fileInputRef = useRef<HTMLInputElement>(null)
  
  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0]
    if (selectedFile) {
      setFile(selectedFile)
      setStatus('idle')
      setDownloadReady(false)
      setStats(null)
    }
  }
  
  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault()
    const droppedFile = e.dataTransfer.files?.[0]
    if (droppedFile && droppedFile.name.endsWith('.docx')) {
      setFile(droppedFile)
      setStatus('idle')
      setDownloadReady(false)
      setStats(null)
    }
  }
  
  const handleFormat = async () => {
    if (!file || !docTitle) return
    
    setStatus('processing')
    setStatusMessage('Analyzing document structure...')
    
    try {
      const content = await extractContent(file)
      
      // Calculate stats
      const newStats: Stats = {
        headings: content.filter(c => c.type.startsWith('heading')).length,
        bullets: content.filter(c => c.type === 'bullet' || c.type === 'sub_bullet').length,
        numbered: content.filter(c => c.type === 'numbered').length,
        codeBlocks: content.filter(c => c.type === 'code_block').length,
        paragraphs: content.filter(c => c.type === 'paragraph').length,
      }
      setStats(newStats)
      
      setStatusMessage('Applying TechTorch formatting standards...')
      const doc = createFormattedDocument(content, docTitle, subtitle, organization || 'TechTorch Inc.', author)
      
      setStatusMessage('Generating document...')
      const blob = await Packer.toBlob(doc)
      
      setDocBlob(blob)
      setDownloadReady(true)
      setStatus('success')
      setStatusMessage('Document formatted successfully!')
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
          <p className="subtitle">Intelligent document formatting with automatic structure detection</p>
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
          <label className="label">Document Title *</label>
          <input
            type="text"
            className="input"
            placeholder="e.g., Salesforce CPQ & Billing Stabilization"
            value={docTitle}
            onChange={(e) => setDocTitle(e.target.value)}
          />
        </div>
        
        <div className="form-group">
          <label className="label">Subtitle</label>
          <input
            type="text"
            className="input"
            placeholder="e.g., Technical Documentation and Operational Handoff"
            value={subtitle}
            onChange={(e) => setSubtitle(e.target.value)}
          />
        </div>
        
        <div className="form-row">
          <div className="form-group half">
            <label className="label">Client / Organization</label>
            <input
              type="text"
              className="input"
              placeholder="e.g., Calabrio Inc."
              value={organization}
              onChange={(e) => setOrganization(e.target.value)}
            />
          </div>
          
          <div className="form-group half">
            <label className="label">Author</label>
            <input
              type="text"
              className="input"
              placeholder="e.g., Kornel"
              value={author}
              onChange={(e) => setAuthor(e.target.value)}
            />
          </div>
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
        
        {stats && (
          <div className="stats-box">
            <p><strong>Detected Structure:</strong></p>
            <div className="stats-grid">
              <div className="stat-item">
                <span className="stat-number">{stats.headings}</span>
                <span className="stat-label">Headings</span>
              </div>
              <div className="stat-item">
                <span className="stat-number">{stats.bullets}</span>
                <span className="stat-label">Bullets</span>
              </div>
              <div className="stat-item">
                <span className="stat-number">{stats.numbered}</span>
                <span className="stat-label">Numbered</span>
              </div>
              <div className="stat-item">
                <span className="stat-number">{stats.codeBlocks}</span>
                <span className="stat-label">Code Blocks</span>
              </div>
              <div className="stat-item">
                <span className="stat-number">{stats.paragraphs}</span>
                <span className="stat-label">Paragraphs</span>
              </div>
            </div>
          </div>
        )}
        
        <div className="info-box">
          <p><strong>Intelligent Detection:</strong></p>
          <p>• Automatically identifies section headers vs body text</p>
          <p>• Detects bullet lists even without markers</p>
          <p>• Recognizes code blocks (SQL, Apex, JavaScript)</p>
          <p>• Infers numbered lists from context</p>
          <p>• Generates Table of Contents</p>
        </div>
        
        <div className="footer">
          TechTorch Document Formatter v4.0 - Intelligent Edition
        </div>
      </div>
    </div>
  )
}
