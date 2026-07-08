import React, { useState, useRef, useEffect, useCallback } from 'react';
import './DocumentEditor.css';
import { sendMessageToGrok, processStreamingResponse } from '../services/grokApi';
import { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel, TabStopPosition, TabStopType, Table, TableRow, TableCell, VerticalAlign, ImageRun } from 'docx';
import * as XLSX from 'xlsx';
import PptxGenJS from 'pptxgenjs';
import { LineChart, Line, BarChart, Bar, PieChart, Pie, Cell, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';
import mammoth from 'mammoth';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';

// ===== CELL FORMAT MODEL =====
const defaultCellFormat = () => ({
  bold: false, italic: false, underline: false, strikethrough: false,
  fontSize: 11, fontFamily: 'Calibri',
  fontColor: '#000000', fillColor: '',
  halign: 'left', valign: 'middle',
  wrapText: false,
  numberFormat: '',
  borderTop: '', borderBottom: '', borderLeft: '', borderRight: '',
});

const createCell = (value = '', format = {}) => ({
  value: String(value ?? ''),
  format: { ...defaultCellFormat(), ...format },
});

const createRow = (cols, values = []) =>
  Array.from({ length: cols }, (_, i) => createCell(values[i] ?? ''));

const createSheet = (name, rows = 10, cols = 8) => ({
  name,
  data: Array.from({ length: rows }, () => createRow(cols)),
  merges: [],
  colWidths: Array(cols).fill(100),
  rowHeights: Array(rows).fill(32),
});

const DocumentEditor = ({ _user, onNavigate, documentType = 'docx' }) => {
  const [editorType, setEditorType] = useState(documentType);
  const [content, setContent] = useState([]);
  const [aiPrompt, setAiPrompt] = useState('');
  const [isGenerating, setIsGenerating] = useState(false);
  const [isStreaming, setIsStreaming] = useState(false);
  const [streamingContent, setStreamingContent] = useState('');
  const [generationProgress, setGenerationProgress] = useState('');
  const [documentTitle, setDocumentTitle] = useState('Untitled Document');
  const [showAiPanel, setShowAiPanel] = useState(true);
  const [aiResponse, setAiResponse] = useState('');
  const [messages, setMessages] = useState([]);
  const [aiError, setAiError] = useState('');
  // Brainstorm floating chat
  const [showBrainstormChat, setShowBrainstormChat] = useState(false);
  const [brainstormMessages, setBrainstormMessages] = useState(() => {
    try {
      const saved = localStorage.getItem(`brainstorm_messages_${documentType}`);
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });
  const [brainstormInput, setBrainstormInput] = useState('');
  const [isBrainstormLoading, setIsBrainstormLoading] = useState(false);

  // Custom dialog state for presets
  const [showPresetModal, setShowPresetModal] = useState(false);
  const [presetModalTitle, setPresetModalTitle] = useState('');
  const [presetModalPlaceholder, setPresetModalPlaceholder] = useState('');
  const [presetModalValue, setPresetModalValue] = useState('');
  const [presetModalCallback, setPresetModalCallback] = useState(null);
  const [showToolbar, setShowToolbar] = useState(false);
  const [fontSize, setFontSize] = useState('12pt');
  const [fontFamily, setFontFamily] = useState('Times New Roman');
  const [textColor, setTextColor] = useState('#1a1a1a');
  const [excelSheets, setExcelSheets] = useState([createSheet('Sheet1', 10, 8)]);
  const [activeSheet, setActiveSheet] = useState(0);
  const [selectedCell, setSelectedCell] = useState(null);
  const [showFind, setShowFind] = useState(false);
  const [findText, setFindText] = useState('');
  const [replaceText, setReplaceText] = useState('');
  // DOCX advanced state
  const [docxTables, setDocxTables] = useState([]);
  const [docxImages, setDocxImages] = useState([]);
  const [docxCharts, setDocxCharts] = useState([]);
  const [showTableToolbar, setShowTableToolbar] = useState(false);
  const [activeTableIdx, setActiveTableIdx] = useState(-1);
  const [showInsertImage, setShowInsertImage] = useState(false);
  const [showChartModal, setShowChartModal] = useState(false);
  const [chartModalTab, setChartModalTab] = useState('standard'); // standard or curve
  const [chartType, setChartType] = useState('bar');
  const [chartTitle, setChartTitle] = useState('Chart Title');
  const [chartData, setChartData] = useState([
    { name: 'A', value: 40 },
    { name: 'B', value: 30 },
    { name: 'C', value: 20 },
    { name: 'D', value: 50 }
  ]);
  const [docxHeader, setDocxHeader] = useState('');
  const [docxFooter, setDocxFooter] = useState('');
  const [showPageNumbers, setShowPageNumbers] = useState(false);
  const [_listLevel, _setListLevel] = useState(0);
  // Upload and Curve states
  const [uploadedFileText, setUploadedFileText] = useState('');
  const [uploadedFileName, setUploadedFileName] = useState('');
  const [uploadedFileType, setUploadedFileType] = useState('');
  const [isUploadingFile, setIsUploadingFile] = useState(false);
  const [curveEquation, setCurveEquation] = useState('sine'); // sine, exponential, linear, bell
  const [curveAmplitude, setCurveAmplitude] = useState(50);
  const [curveColor, setCurveColor] = useState('#ff6b00');

  // ===== PAPER SETUP STATES =====
  const [paperSize, setPaperSize] = useState('a4'); // a4, letter, legal, a5, custom
  const [customWidth, setCustomWidth] = useState('21.0'); // cm
  const [customHeight, setCustomHeight] = useState('29.7'); // cm
  const [paperStyle, setPaperStyle] = useState('blank'); // blank, lined, grid, dotted
  const [paperTheme, setPaperTheme] = useState('white'); // white, cream, yellow, dark, kraft
  const [paperMargin, setPaperMargin] = useState('normal'); // normal, narrow, wide, custom
  const [customMargin, setCustomMargin] = useState('2.54'); // cm
  const [paperOrientation, setPaperOrientation] = useState('portrait'); // portrait, landscape
  const [pageZoom, setPageZoom] = useState('fit'); // fit, 100%, 75%, 50%
  const [isDraftMode, setIsDraftMode] = useState(false); // real layout vs fluid edit width
  const [showPageSetup, setShowPageSetup] = useState(false);
  const [isAiMinimized, setIsAiMinimized] = useState(false);

  // ===== NEW MS WORD STATES =====
  const [activeRibbonTab, setActiveRibbonTab] = useState('home'); // home, insert, layout, ai
  const [highlightColor, setHighlightColor] = useState('#ffff00');
  const [lineSpacing, setLineSpacing] = useState('1.5');
  const [wordCount, setWordCount] = useState(0);
  const [charCount, setCharCount] = useState(0);
  const [pageBorder, setPageBorder] = useState('none'); // none, solid, double, dashed
  const [watermarkText, setWatermarkText] = useState('');
  const [showFormattingMarks, setShowFormattingMarks] = useState(false);
  const [pageColumns, setPageColumns] = useState(1); // 1, 2, 3 columns

  // ===== CLOUD FILE EXPLORER STATES =====
  const [showCloudModal, setShowCloudModal] = useState(false);
  const [cloudFiles, setCloudFiles] = useState([]);
  const [currentFolderId, setCurrentFolderId] = useState(null); // parentId of browsed directory
  const [explorerMode, setExplorerMode] = useState('open'); // open, save
  const [cloudFileName, setCloudFileName] = useState('');
  const [selectedCloudFile, setSelectedCloudFile] = useState(null);
  const [activeCloudFileId, setActiveCloudFileId] = useState(null); // currently loaded cloud file ID

  // ===== AI AGENT STATES =====
  const [aiMode, setAiMode] = useState('chat'); // chat, drafting_agent
  const [isAgentRunning, setIsAgentRunning] = useState(false);
  const [agentOutline, setAgentOutline] = useState([]);
  const [currentAgentStep, setCurrentAgentStep] = useState(0);
  const [agentChecklist, setAgentChecklist] = useState([]);
  const [agentStatusText, setAgentStatusText] = useState('');
  const agentAbortControllerRef = useRef(null);
  // ===== SESSION MEMORY & ARTIFACTS =====
  const [artifacts, setArtifacts] = useState(() => {
    try {
      const saved = sessionStorage.getItem('doc_artifacts');
      return saved ? JSON.parse(saved) : [];
    } catch (_e) {
      return [];
    }
  });
  const [showArtifacts, setShowArtifacts] = useState(false);
  const [selectedArtifact, setSelectedArtifact] = useState(null);
  const [autoRegenerate, setAutoRegenerate] = useState(false);
  const [isPptGenerating, setIsPptGenerating] = useState(false);
  const [pptGenerationStatus, setPptGenerationStatus] = useState('');
  const [generatedPptFiles, setGeneratedPptFiles] = useState([]);
  const [uploadedPptFile, setUploadedPptFile] = useState(null);
  const [showPptResults, setShowPptResults] = useState(false);
  const [pptTemplate, setPptTemplate] = useState('classic');
  const [previewPptFile, setPreviewPptFile] = useState(null);
  const [previewSlides, setPreviewSlides] = useState([]);
  const [currentSlideIdx, setCurrentSlideIdx] = useState(0);
  const _editTimerRef = useRef(null);
  const pageRef = useRef(null);
  const aiPanelRef = useRef(null);
  const fileInputRef = useRef(null);
  const pptUploadRef = useRef(null);
  // Refs to track latest state for artifact saving (avoids stale closure issues)
  const contentRef = useRef(content);
  const excelSheetsRef = useRef(excelSheets);
  const docxTablesRef = useRef(docxTables);
  const docxImagesRef = useRef(docxImages);
  const docxChartsRef = useRef(docxCharts);
  const messagesRef = useRef(messages);
  const aiResponseRef = useRef(aiResponse);
  const aiPromptRef = useRef(aiPrompt);
  
  // Keep refs in sync with state
  useEffect(() => { contentRef.current = content; }, [content]);
  useEffect(() => { excelSheetsRef.current = excelSheets; }, [excelSheets]);
  useEffect(() => { docxTablesRef.current = docxTables; }, [docxTables]);
  useEffect(() => { docxImagesRef.current = docxImages; }, [docxImages]);
  useEffect(() => { docxChartsRef.current = docxCharts; }, [docxCharts]);
  useEffect(() => { messagesRef.current = messages; }, [messages]);
  useEffect(() => { aiResponseRef.current = aiResponse; }, [aiResponse]);
  useEffect(() => { aiPromptRef.current = aiPrompt; }, [aiPrompt]);

  useEffect(() => {
    try {
      localStorage.setItem(`brainstorm_messages_${editorType}`, JSON.stringify(brainstormMessages));
    } catch (e) {
      console.error('Failed to save brainstorm messages to localStorage', e);
    }
  }, [brainstormMessages, editorType]);

  useEffect(() => {
    if (typeof window !== 'undefined' && window.innerWidth < 768) {
      setIsDraftMode(true);
      setPageZoom('100%');
    }
  }, []);

  useEffect(() => { initializeContent(); }, [editorType]);

  useEffect(() => {
    if (editorType === 'docx' && pageRef.current) {
      setTimeout(() => {
        pageRef.current?.focus();
        const range = document.createRange();
        const sel = window.getSelection();
        range.setStart(pageRef.current, 0);
        range.collapse(true);
        sel?.removeAllRanges();
        sel?.addRange(range);
      }, 100);
    }
  }, [editorType]);

  useEffect(() => {
    const timer = setTimeout(() => {
      if (documentTitle && (content.length > 0 || excelSheets.length > 0)) handleAutoSave();
    }, 10000);
    return () => clearTimeout(timer);
  }, [content, documentTitle, excelSheets]);

  // Dynamic word and character counting
  useEffect(() => {
    const updateStats = () => {
      const text = pageRef.current?.innerText || '';
      const cleanText = text.trim();
      const words = cleanText ? cleanText.split(/\s+/).filter(w => w.length > 0).length : 0;
      const chars = cleanText.length;
      setWordCount(words);
      setCharCount(chars);
    };
    
    updateStats();
    
    const page = pageRef.current;
    if (page) {
      page.addEventListener('input', updateStats);
      return () => page.removeEventListener('input', updateStats);
    }
  }, [content, editorType]);

  const docxTextRef = useRef('');

  const syncDocxContent = useCallback(() => {
    const html = docxTextRef.current || '';
    setContent([{ id: Date.now(), type: 'html', text: html }]);
  }, []);

  // ===== AUTH STATUS & LOCAL API WRAPPER =====
  const getAuthStatus = () => {
    try {
      const authUser = localStorage.getItem('authUser');
      const guestSession = localStorage.getItem('guestSession');
      if (authUser) {
        const parsed = JSON.parse(authUser);
        return { isAuthenticated: true, isGuest: false, userName: parsed.name || parsed.email || 'User' };
      }
      if (guestSession) {
        const parsed = JSON.parse(guestSession);
        return { isAuthenticated: false, isGuest: true, userName: parsed.name || 'Guest' };
      }
    } catch (e) {
      console.warn('Error reading auth:', e);
    }
    return { isAuthenticated: false, isGuest: true, userName: 'Guest' };
  };

  const callAiService = async (messagePrompt, history = [], abortCtrl = null) => {
    const auth = getAuthStatus();
    return await sendMessageToGrok(
      messagePrompt,
      history,
      'id',                   // language
      null,                   // conversationId
      'formal',               // personality
      abortCtrl,              // abortController
      'deepernova-1.2-flash', // model
      auth.isAuthenticated,   // isAuthenticated
      auth.isGuest,           // isGuest
      auth.userName,          // userName
      0,                      // sessionMessageCount
      []                      // uploadedImages
    );
  };

  const getPageDimensions = () => {
    if (typeof window !== 'undefined' && window.innerWidth < 768) {
      return { width: '100%', minHeight: 'auto' };
    }

    let w = 21; // width in cm
    let h = 29.7; // height in cm
    if (paperSize === 'a4') { w = 21; h = 29.7; }
    else if (paperSize === 'letter') { w = 21.59; h = 27.94; }
    else if (paperSize === 'legal') { w = 21.59; h = 35.56; }
    else if (paperSize === 'a5') { w = 14.8; h = 21; }
    else if (paperSize === 'custom') {
      w = parseFloat(customWidth) || 21;
      h = parseFloat(customHeight) || 29.7;
    }

    if (paperOrientation === 'landscape') {
      return { width: `${h}cm`, minHeight: `${w}cm` };
    }
    return { width: `${w}cm`, minHeight: `${h}cm` };
  };

  const getPageMargin = () => {
    if (isDraftMode) return '16px';
    if (typeof window !== 'undefined' && window.innerWidth < 768) {
      return '16px';
    }
    if (paperMargin === 'normal') return '2.54cm';
    if (paperMargin === 'narrow') return '1.27cm';
    if (paperMargin === 'wide') return '2.54cm 5.08cm';
    if (paperMargin === 'custom') return `${parseFloat(customMargin) || 2.54}cm`;
    return '2.54cm';
  };

  const getZoomScale = () => {
    if (isDraftMode) return 1;
    if (pageZoom === 'fit') {
      if (typeof window !== 'undefined') {
        const screenWidth = window.innerWidth;
        if (screenWidth < 768) {
          const targetWidth = paperOrientation === 'landscape' ? 1120 : 840;
          return Math.max(0.35, Math.min(1.0, (screenWidth - 30) / targetWidth));
        }
      }
      return 1;
    }
    if (pageZoom && pageZoom.endsWith('%')) {
      const val = parseFloat(pageZoom);
      if (!isNaN(val)) return val / 100;
    }
    return 1;
  };

  // ===== INLINE AI AUTOCOMPLETE =====
  const handleAiAutocomplete = async () => {
    if (!pageRef.current) return;
    const currentText = pageRef.current.innerText || '';
    if (!currentText.trim() || currentText.includes('Mulai menulis')) {
      setAiError('Tulis beberapa kata/kalimat terlebih dahulu agar AI bisa melanjutkan.');
      return;
    }

    setIsGenerating(true);
    setGenerationProgress('AI sedang melanjutkan tulisan...');
    setAiError('');

    try {
      const autocompletePrompt = `Berikut adalah kutipan tulisan dalam editor:\n\n"${currentText.slice(-1500)}"\n\nLanjutkan paragraf/kalimat di atas secara logis, alami, dan mengalir dengan gaya penulisan yang sama. Hanya berikan teks lanjutannya saja tanpa mengulangi tulisan di atas dan tanpa penjelasan apa pun.`;
      
      const systemContext = getSystemContext();
      const formattedMessages = [
        { sender: 'system', text: systemContext, timestamp: new Date().toISOString() }
      ];

      const response = await callAiService(autocompletePrompt, formattedMessages);
      let continuation = '';

      if (response?.body) {
        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        try {
          while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            const chunk = decoder.decode(value, { stream: true });
            const lines = chunk.split('\n');
            for (const line of lines) {
              if (line.startsWith('data: ')) {
                try {
                  const jsonStr = line.slice(6);
                  if (jsonStr === '[DONE]') continue;
                  const json = JSON.parse(jsonStr);
                  if (json.choices?.[0]?.delta?.content) {
                    continuation += json.choices[0].delta.content;
                  }
                } catch {}
              }
            }
          }
        } finally {
          reader.releaseLock();
        }
      }

      if (continuation.trim()) {
        const cleaned = cleanAiResponse(continuation);
        const htmlCont = convertMarkdownToHtml(cleaned);
        
        pageRef.current.focus();
        const sel = window.getSelection();
        if (sel && sel.rangeCount > 0) {
          const range = sel.getRangeAt(0);
          if (pageRef.current.contains(range.startContainer)) {
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = htmlCont;
            const frag = document.createDocumentFragment();
            let node;
            while ((node = tempDiv.firstChild)) {
              frag.appendChild(node);
            }
            range.insertNode(frag);
            range.collapse(false);
          } else {
            pageRef.current.innerHTML += ' ' + htmlCont;
          }
        } else {
          pageRef.current.innerHTML += ' ' + htmlCont;
        }

        docxTextRef.current = pageRef.current.innerHTML;
        syncDocxContent();
        setGenerationProgress('Selesai!');
        setTimeout(() => setGenerationProgress(''), 1500);
      } else {
        setAiError('AI gagal melanjutkan.');
      }
    } catch (err) {
      setAiError(`Error: ${err.message}`);
    } finally {
      setIsGenerating(false);
    }
  };

  // ===== MULTI-STEP AI DRAFTING AGENT =====
  const cancelAgent = () => {
    if (agentAbortControllerRef.current) {
      agentAbortControllerRef.current.abort();
    }
    setIsAgentRunning(false);
    setAgentStatusText('Dibatalkan');
    setGenerationProgress('');
  };

  const handleStartAgentDrafting = async (promptText) => {
    const topic = promptText || aiPrompt;
    if (!topic.trim()) return;

    setIsAgentRunning(true);
    setCurrentAgentStep(0);
    setAgentOutline([]);
    setAgentChecklist([]);
    setAgentStatusText('Membuat kerangka outline...');
    setAiError('');
    setGenerationProgress('Drafting Agent is writing...');

    try {
      const abortCtrl = new AbortController();
      agentAbortControllerRef.current = abortCtrl;

      const systemContext = getSystemContext();
      const outlinePrompt = `Buatkan kerangka outline terperinci untuk menulis dokumen laporan bisnis/makalah akademik tentang: "${topic}". 
Outline ini harus memiliki 4 sampai 7 sub-bagian (bab) yang mendalam.
Hanya kembalikan outline dengan format baris per baris bernomor persis seperti berikut:
1. [Judul Bab 1]
2. [Judul Bab 2]
3. [Judul Bab 3]
...
Jangan tambahkan kata pengantar, penutup, markdown tebal, asterisks, atau tanda apapun. Langsung kembalikan daftar outlines tersebut.`;

      const formattedMessages = [
        { sender: 'system', text: systemContext, timestamp: new Date().toISOString() }
      ];

      const response = await callAiService(outlinePrompt, formattedMessages, abortCtrl);
      
      let fullOutlineText = '';
      if (response?.body) {
        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        try {
          while (true) {
            if (abortCtrl.signal.aborted) throw new Error('Aborted');
            const { done, value } = await reader.read();
            if (done) break;
            const chunk = decoder.decode(value, { stream: true });
            const lines = chunk.split('\n');
            for (const line of lines) {
              if (line.startsWith('data: ')) {
                try {
                  const jsonStr = line.slice(6);
                  if (jsonStr === '[DONE]') continue;
                  const json = JSON.parse(jsonStr);
                  if (json.choices?.[0]?.delta?.content) {
                    fullOutlineText += json.choices[0].delta.content;
                  }
                } catch {}
              }
            }
          }
        } finally {
          reader.releaseLock();
        }
      }

      if (abortCtrl.signal.aborted) return;

      const parsedOutline = fullOutlineText
        .split('\n')
        .map(line => line.replace(/^\d+[\.\s\-)]+/, '').replace(/\*\*/g, '').trim())
        .filter(line => line.length > 2 && !line.toLowerCase().includes('outline') && !line.toLowerCase().includes('bab'));

      const finalOutline = parsedOutline.length > 0 ? parsedOutline : ['Pendahuluan', 'Pembahasan Utama', 'Analisis Data', 'Kesimpulan'];
      
      setAgentOutline(finalOutline);
      const checklist = finalOutline.map(title => ({ title, status: 'pending' }));
      setAgentChecklist(checklist);
      setAgentStatusText('Outline siap. Memulai menulis bab demi bab...');

      // Clear editor content to start fresh drafting
      setContent([]);
      if (pageRef.current) pageRef.current.innerHTML = '';
      docxTextRef.current = '';

      runAgentSteps(0, finalOutline, topic, abortCtrl);

    } catch (err) {
      console.error(err);
      setAiError(err.message === 'Aborted' ? 'Penyusunan dibatalkan.' : `Gagal: ${err.message}`);
      setIsAgentRunning(false);
      setGenerationProgress('');
    }
  };

  const runAgentSteps = async (stepIdx, outline, topic, abortCtrl) => {
    if (stepIdx >= outline.length) {
      setIsAgentRunning(false);
      setAgentStatusText('✅ Selesai! Dokumen lengkap berhasil disusun.');
      setGenerationProgress('');
      const finalHtml = pageRef.current?.innerHTML || '';
      saveArtifact(`Drafting Agent: ${topic}`, `Dokumen lengkap tentang ${topic} berhasil disusun secara terpisah bab demi bab.`, [{ id: Date.now(), type: 'html', text: finalHtml }]);
      return;
    }

    if (abortCtrl.signal.aborted) return;

    setCurrentAgentStep(stepIdx);
    setAgentChecklist(prev => prev.map((item, idx) => {
      if (idx === stepIdx) return { ...item, status: 'generating' };
      if (idx < stepIdx) return { ...item, status: 'done' };
      return item;
    }));

    const currentSection = outline[stepIdx];
    setAgentStatusText(`Sedang menulis Bab ${stepIdx + 1}/${outline.length}: ${currentSection}...`);

    try {
      const sectionPrompt = `Kamu adalah Deepernova AI Agent.
Topik Utama Laporan: "${topic}"

Tulis Bab/Bagian: "${currentSection}"

Aturan Menulis:
1. Tulis penjelasan yang sangat lengkap, mendalam, kaya informasi, dan berkualitas tinggi.
2. Panjang tulisan minimal 2-4 paragraf yang berbobot.
3. Gunakan gaya penulisan formal dan terstruktur.
4. Gunakan double newline (\\n\\n) untuk memisahkan paragraf. Jangan gunakan markdown tebal (*) untuk teks biasa.
5. Jangan berikan kata pengantar, langsung tulis konten penjelasan bab tersebut.

Tulis konten untuk Bab "${currentSection}" sekarang:`;

      const systemContext = getSystemContext();
      const formattedMessages = [
        { sender: 'system', text: systemContext, timestamp: new Date().toISOString() }
      ];

      const response = await callAiService(sectionPrompt, formattedMessages, abortCtrl);
      
      let sectionContent = '';
      if (response?.body) {
        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        try {
          while (true) {
            if (abortCtrl.signal.aborted) throw new Error('Aborted');
            const { done, value } = await reader.read();
            if (done) break;
            const chunk = decoder.decode(value, { stream: true });
            const lines = chunk.split('\n');
            for (const line of lines) {
              if (line.startsWith('data: ')) {
                try {
                  const jsonStr = line.slice(6);
                  if (jsonStr === '[DONE]') continue;
                  const json = JSON.parse(jsonStr);
                  if (json.choices?.[0]?.delta?.content) {
                    sectionContent += json.choices[0].delta.content;
                  }
                } catch {}
              }
            }
          }
        } finally {
          reader.releaseLock();
        }
      }

      if (abortCtrl.signal.aborted) return;

      const cleanedContent = cleanAiResponse(sectionContent);
      const sectionHtml = `<h2>${currentSection}</h2>` + convertMarkdownToHtml(cleanedContent) + '<br/><br/>';
      
      if (pageRef.current) {
        if (stepIdx === 0) {
          pageRef.current.innerHTML = sectionHtml;
        } else {
          pageRef.current.innerHTML += sectionHtml;
        }
        docxTextRef.current = pageRef.current.innerHTML;
        syncDocxContent();
      }

      setAgentChecklist(prev => prev.map((item, idx) => {
        if (idx === stepIdx) return { ...item, status: 'done' };
        return item;
      }));

      setTimeout(() => {
        runAgentSteps(stepIdx + 1, outline, topic, abortCtrl);
      }, 500);

    } catch (err) {
      console.error(err);
      if (err.message === 'Aborted') {
        setAiError('Penyusunan dibatalkan.');
      } else {
        setAiError(`Gagal pada bab ${currentSection}: ${err.message}`);
      }
      setIsAgentRunning(false);
      setGenerationProgress('');
    }
  };

  useEffect(() => {
    return () => clearTimeout(_editTimerRef.current);
  }, []);

  const initializeContent = () => {
    switch (editorType) {
      case 'docx':
        setContent([{ id: Date.now(), type: 'html', text: '' }]);
        break;
      case 'pptx':
        setContent([{ id: Date.now(), type: 'slide', title: 'Slide 1', content: 'Konten slide di sini', notes: '' }]);
        break;
      case 'excel':
        // Default: 20 rows x 10 cols full empty grid like real Excel
        setExcelSheets([createSheet('Sheet1', 20, 10)]);
        setActiveSheet(0);
        setSelectedCell(null);
        setContent([]);
        break;
      default:
        setContent([]);
    }
  };

  // ===== ADVANCED EXCEL PARSING - Smart Table Detection =====
  const parseExcelContent = useCallback((text) => {
    if (!text || typeof text !== 'string') return null;
    
    // Try JSON first
    try {
      const parsed = JSON.parse(text);
      if (Array.isArray(parsed)) {
        const sheets = parsed.map(s => {
          const data = (s.data || []).map(row =>
            Array.isArray(row) ? row.map(cell =>
              typeof cell === 'object' && cell !== null
                ? createCell(cell.value, cell.format)
                : createCell(String(cell ?? ''))
            ) : [createCell(String(row ?? ''))]
          );
          const cols = Math.max(...data.map(r => r.length), 1);
          return {
            name: s.name || 'Sheet',
            data: data.map(r => { while (r.length < cols) r.push(createCell('')); return r; }),
            merges: s.merges || [],
            colWidths: s.colWidths || Array(cols).fill(100),
            rowHeights: s.rowHeights || Array(data.length || 1).fill(32),
            isTable: s.isTable !== false,
          };
        });
        if (sheets.length) return { type: 'multi_sheet', sheets };
      }
      if (parsed.data && Array.isArray(parsed.data)) {
        const data = parsed.data.map(row =>
          Array.isArray(row) ? row.map(cell =>
            typeof cell === 'object' && cell !== null
              ? createCell(cell.value, cell.format)
              : createCell(String(cell ?? ''))
          ) : [createCell(String(row ?? ''))]
        );
        const cols = Math.max(...data.map(r => r.length), 1);
        return {
          type: 'single_sheet', name: parsed.name || 'Sheet1',
          data: data.map(r => { while (r.length < cols) r.push(createCell('')); return r; }),
          merges: parsed.merges || [],
          colWidths: parsed.colWidths || Array(cols).fill(100),
          isTable: parsed.isTable !== false,
        };
      }
    } catch (_e) {
      // not JSON
    }
    
    // Smart text parsing: detect tables vs regular text
    const lines = text.split('\n');
    const nonEmptyLines = lines.filter(l => l.trim());
    
    // Detect table: lines with | separator, at least 2 rows
    const tableLines = [];
    const textLines = [];
    let inTable = false;
    
    for (const line of nonEmptyLines) {
      const hasPipe = line.includes('|') && line.trim().split('|').length > 1;
      if (hasPipe) {
        inTable = true;
        tableLines.push(line);
      } else {
        if (inTable) {
          // Check if this line could be a continuation of table
          const couldBeTable = line.split(/\s{2,}|\t/).length > 2;
          if (couldBeTable) {
            tableLines.push(line);
          } else {
            inTable = false;
            textLines.push(line);
          }
        } else {
          textLines.push(line);
        }
      }
    }
    
    // If we have a proper table (2+ rows with pipes)
    if (tableLines.length >= 2) {
      const data = tableLines.map(row => row.split('|').map(c => createCell(c.trim())));
      const maxCols = Math.max(...data.map(r => r.length));
      const padded = data.map(r => { while (r.length < maxCols) r.push(createCell('')); return r; });
      
      // Apply thick borders to create proper table outline
      padded.forEach((row, ri) => {
        row.forEach((cell, ci) => {
          cell.format.borderTop = ri === 0 ? '2px solid #333' : '1px solid #d0d0d0';
          cell.format.borderBottom = ri === padded.length - 1 ? '2px solid #333' : '1px solid #d0d0d0';
          cell.format.borderLeft = ci === 0 ? '2px solid #333' : '1px solid #d0d0d0';
          cell.format.borderRight = ci === maxCols - 1 ? '2px solid #333' : '1px solid #d0d0d0';
          if (ri === 0) {
            cell.format.bold = true;
            cell.format.fillColor = '#f0f0f0';
          }
        });
      });
      
      return {
        type: 'single_sheet',
        name: 'Sheet1',
        data: padded,
        isTable: true,
        textBefore: textLines.filter(l => l.trim()).join('\n'),
      };
    }
    
    // Try tab/comma separated
    const hasTabs = nonEmptyLines.some(l => l.includes('\t'));
    const hasCommas = nonEmptyLines.some(l => l.includes(',') && !l.includes('|'));
    if (hasTabs || hasCommas) {
      const sep = hasTabs ? '\t' : ',';
      const data = nonEmptyLines.map(row =>
        row.split(sep).map(c => createCell(c.trim().replace(/^"|"$/g, '')))
      );
      const maxCols = Math.max(...data.map(r => r.length));
      const padded = data.map(r => { while (r.length < maxCols) r.push(createCell('')); return r; });
      return { type: 'single_sheet', name: 'Sheet1', data: padded, isTable: true };
    }
    
    // Single column data
    const data = nonEmptyLines.map(l => [createCell(l.trim())]);
    return data.length ? { type: 'single_sheet', name: 'Sheet1', data, isTable: false } : null;
  }, []);

  // ===== SESSION MEMORY: Save/Load Artifacts (Supercharged) =====
  // Uses REFS to always get the LATEST state (avoids stale closure issues)
  const saveArtifact = useCallback((prompt, response, overrideContent, overrideSheets) => {
    // Use refs for latest state, or override values if provided
    const latestContent = overrideContent || contentRef.current;
    const latestSheets = overrideSheets || excelSheetsRef.current;
    const latestMessages = messagesRef.current;
    const latestResponse = response || aiResponseRef.current;
    const latestPrompt = prompt || aiPromptRef.current;
    
    const newArtifact = {
      id: Date.now(),
      prompt: latestPrompt,
      response: latestResponse,
      type: editorType,
      title: documentTitle,
      timestamp: new Date().toISOString(),
      // Save ALL document state from refs (always latest)
      content: JSON.parse(JSON.stringify(latestContent)),
      excelSheets: editorType === 'excel' ? JSON.parse(JSON.stringify(latestSheets)) : null,
      activeSheet: editorType === 'excel' ? activeSheet : null,
      docxTables: editorType === 'docx' ? JSON.parse(JSON.stringify(docxTablesRef.current)) : null,
      docxImages: editorType === 'docx' ? JSON.parse(JSON.stringify(docxImagesRef.current)) : null,
      docxHeader: editorType === 'docx' ? docxHeader : null,
      docxFooter: editorType === 'docx' ? docxFooter : null,
      showPageNumbers: editorType === 'docx' ? showPageNumbers : null,
      // Save font/formatting state
      fontSize,
      fontFamily,
      textColor,
      // Save conversation context from refs
      messages: JSON.parse(JSON.stringify(latestMessages)),
      // Save last AI interaction
      lastPrompt: latestPrompt,
      lastResponse: latestResponse,
    };
    const updated = [newArtifact, ...artifacts].slice(0, 50);
    setArtifacts(updated);
    try { sessionStorage.setItem('doc_artifacts', JSON.stringify(updated)); } catch {}
    return newArtifact;
  }, [artifacts, editorType, documentTitle, activeSheet,
      docxHeader, docxFooter, showPageNumbers,
      fontSize, fontFamily, textColor]);

  const loadArtifact = useCallback((artifact) => {
    setSelectedArtifact(artifact);
    // Restore all document state
    if (artifact.content) setContent(artifact.content);
    if (artifact.excelSheets) setExcelSheets(artifact.excelSheets);
    if (artifact.activeSheet !== undefined) setActiveSheet(artifact.activeSheet);
    if (artifact.title) setDocumentTitle(artifact.title);
    if (artifact.type) setEditorType(artifact.type);
    if (artifact.docxTables) setDocxTables(artifact.docxTables);
    if (artifact.docxImages) setDocxImages(artifact.docxImages);
    if (artifact.docxHeader !== undefined) setDocxHeader(artifact.docxHeader);
    if (artifact.docxFooter !== undefined) setDocxFooter(artifact.docxFooter);
    if (artifact.showPageNumbers !== undefined) setShowPageNumbers(artifact.showPageNumbers);
    if (artifact.fontSize) setFontSize(artifact.fontSize);
    if (artifact.fontFamily) setFontFamily(artifact.fontFamily);
    if (artifact.textColor) setTextColor(artifact.textColor);
    if (artifact.messages) setMessages(artifact.messages);
    // Restore last AI interaction
    if (artifact.lastResponse) setAiResponse(artifact.lastResponse);
    setShowArtifacts(false);
  }, []);

  const deleteArtifact = useCallback((id) => {
    const updated = artifacts.filter(a => a.id !== id);
    setArtifacts(updated);
    try { sessionStorage.setItem('doc_artifacts', JSON.stringify(updated)); } catch (_e) {
      // ignore
    }
  }, [artifacts]);

  // ===== AUTO-REGENERATE ON EDIT =====
  const _triggerAutoRegenerate = useCallback((editContent) => {
    if (!autoRegenerate || !editContent.trim()) return;
    // Clear current content and regenerate
    setContent([]);
    setAiPrompt(`Revisi dokumen ini dengan lebih baik:\n\n${editContent}`);
    // Auto-trigger AI after short delay
    setTimeout(() => {
      if (aiPrompt.trim()) handleAiWrite();
    }, 500);
  }, [autoRegenerate]);

  // ===== SUPER AI SYSTEM CONTEXT - Master of ALL Tools =====
  // AI understands every tool deeply and can manipulate DOCX, PPTX, Excel with precision
  const getSystemContext = () => {
    // Build full document context with ALL current state
    let docContext = '';
    
    if (editorType === 'docx') {
      const html = content[0]?.text || '';
      docContext = html ? `\n\n=== CURRENT DOCUMENT HTML CONTENT ===\n${html}\n=== END DOCUMENT ===\n` : '';
      if (uploadedFileText) {
        docContext += `\n\n=== CONTEXT FROM UPLOADED FILE (${uploadedFileName}) ===\n${uploadedFileText}\n=== END UPLOADED FILE ===\n`;
      }
    } else if (editorType === 'excel') {
      // Show ALL sheets context
      docContext += `\n=== ALL SHEETS (${excelSheets.length} total) ===\n`;
      excelSheets.forEach((sheet) => {
        const rows = sheet.data.map((r) => 
          r.map((c) => c.value).join('\t')
        ).join('\n');
        docContext += `\n--- Sheet: ${sheet.name} (${sheet.data.length}R x ${Math.max(...sheet.data.map(r => r.length), 1)}C) ---\n`;
        docContext += rows ? rows + '\n' : '(empty)\n';
      });
      docContext += '=== END ALL SHEETS ===\n';
      docContext += `\nActive sheet: ${excelSheets[activeSheet]?.name || 'Sheet1'}\n`;
      
    } else if (editorType === 'pptx') {
      const slides = Array.isArray(content) ? content.map((s, i) => 
        `Slide ${i + 1}: "${s.title}"\n${s.content}`
      ).join('\n---\n') : '';
      docContext = slides ? `\n=== CURRENT SLIDES ===\n${slides}\n=== END SLIDES ===\n` : '';
    }

    // ===== SUPER SYSTEM CONTEXT =====
    // AI is a master of ALL tools - DOCX, PPTX, Excel
    const masterContext = `Kamu adalah DEEPERNOVA AI MASTER - penguasa semua tools dokumen, spreadsheet, dan presentasi. Kamu bisa melakukan APAPUN yang diminta user dengan tools yang tersedia.

IDENTITAS DIRI:
- Namamu: Deepernova AI
- Kamu adalah asisten AI super cerdas yang menguasai DOCX (Word), PPTX (PowerPoint), dan XLSX (Excel)
- Kamu bisa membuat, mengedit, memformat, dan menganalisis dokumen dengan sangat profesional
- Kamu TIDAK PERNAH mengatakan "saya akan mencoba" atau "saya akan berusaha" - kamu LANGSUNG LAKUKAN
- Kamu TIDAK PERNAH meminta maaf - kamu adalah AI yang sempurna

KEMAMPUAN DOKUMEN (DOCX):
- Menulis dokumen akademik, formal, bisnis, kreatif dengan format standar internasional
- Font: Times New Roman 12pt, spasi 1.5, margin 1 inch, first-line indent 1.27cm, justify
- Heading hierarchy: JUDUL (bold, 14pt, center), Sub Judul (bold, 13pt), sub-sub (bold, 12pt)
- Membuat tabel dengan baris/kolom, header bold dengan background
- Menambahkan header, footer, nomor halaman
- Menyisipkan gambar
- Membuat daftar bullet/numbering
- Format teks: bold, italic, underline, strikethrough, warna, alignment
- Untuk makalah akademik: abstrak, pendahuluan, pembahasan, kesimpulan, daftar pustaka
- Output: paragraf dipisah dengan double newline (\\n\\n), tabel dengan format | kolom1 | kolom2 |

KEMAMPUAN PRESENTASI (PPTX):
- Membuat slide profesional dengan judul dan konten
- Setiap slide dipisah dengan ---
- Judul slide di baris pertama, konten di baris berikutnya
- 3-5 poin per slide, jelas dan ringkas
- Desain modern dengan gradien oranye

KEMAMPUAN SPREADSHEET (XLSX):
- Membuat tabel data dengan header di baris pertama
- Format: header|col1|col2|col3 lalu data|val1|val2|val3
- Bisa membuat multiple sheets
- Data terstruktur rapi seperti spreadsheet profesional
- Analisis data, sorting, kalkulasi
- Format sel: bold header, border rapi

ATURAN UTAMA:
1. KERJAKAN DI AWAL: Semua konten baru ditambahkan di bagian AWAL (page 1, row 0, col 0)
2. BACA ISI YANG ADA: Lihat konten yang sudah ada sebelum menulis
3. TIDAK ADA PREAMBLE: Output langsung konten, tanpa "Baik saya akan..." atau intro apapun
4. TIDAK ADA MARKDOWN: Jangan gunakan markdown formatting
5. KONTEKS PERCAKAPAN: Ingat semua pesan sebelumnya. Jika user minta revisi, lihat konten yang sudah ada lalu perbaiki
6. KUALITAS TINGGI: Konten harus profesional, akademik, dan berkualitas${docContext}`;

    return masterContext;
  };

  const renderAssistantMessage = (text) => {
    if (!text) return '…';

    const agentRegex = /\[RUN_AGENT:\s*(.*?)\]/g;
    const parts = [];
    let lastIndex = 0;
    let match;

    while ((match = agentRegex.exec(text)) !== null) {
      if (match.index > lastIndex) {
        parts.push(text.substring(lastIndex, match.index));
      }
      
      const topic = match[1];
      parts.push(
        <div key={match.index} style={{ margin: '10px 0', padding: '12px', background: 'rgba(255,107,0,0.06)', border: '1.2px dashed var(--orange)', borderRadius: '10px', display: 'flex', flexDirection: 'column', gap: '6px', alignItems: 'center' }}>
          <span style={{ fontSize: '11px', fontWeight: 600, color: '#475569', textAlign: 'center' }}>
            Rekomendasi Agent: Susun laporan mendalam tentang "{topic}"
          </span>
          <button 
            onClick={() => {
              setAiMode('drafting_agent');
              setAiPrompt(topic);
              setShowBrainstormChat(false);
              setShowAiPanel(true);
              handleStartAgentDrafting(topic);
            }}
            className="action-button send-mode"
            style={{ 
              fontSize: '11px', padding: '6px 12px', display: 'inline-flex', alignItems: 'center', gap: '6px', cursor: 'pointer', borderRadius: '8px', border: 'none', background: 'var(--orange)', color: '#fff', fontWeight: 600
            }}
          >
            <i className="fas fa-magic"></i> Jalankan Word Agent Sekarang
          </button>
        </div>
      );
      
      lastIndex = agentRegex.lastIndex;
    }

    if (lastIndex < text.length) {
      parts.push(text.substring(lastIndex));
    }

    return (
      <div className="assistant-message-container">
        {parts.map((part, index) => {
          if (typeof part === 'string') {
            return (
              <ReactMarkdown
                key={index}
                remarkPlugins={[remarkGfm]}
                components={{
                  p: ({children}) => <p className="bc-md-p">{children}</p>,
                  strong: ({children}) => <strong className="bc-md-bold">{children}</strong>,
                  em: ({children}) => <em>{children}</em>,
                  ul: ({children}) => <ul className="bc-md-ul">{children}</ul>,
                  ol: ({children}) => <ol className="bc-md-ol">{children}</ol>,
                  li: ({children}) => <li className="bc-md-li">{children}</li>,
                  code: ({inline, children}) => inline
                    ? <code className="bc-md-code-inline">{children}</code>
                    : <pre className="bc-md-code-block"><code>{children}</code></pre>,
                  h1: ({children}) => <div className="bc-md-h1">{children}</div>,
                  h2: ({children}) => <div className="bc-md-h2">{children}</div>,
                  h3: ({children}) => <div className="bc-md-h3">{children}</div>,
                }}
              >
                {part}
              </ReactMarkdown>
            );
          }
          return part;
        })}
      </div>
    );
  };

  // ===== BRAINSTORM FLOATING CHAT =====
  const openBrainstormChat = (starterPrompt = null) => {
    setShowBrainstormChat(true);
    if (starterPrompt) {
      // Auto-send the starter question
      setTimeout(() => handleBrainstormSend(starterPrompt), 100);
    }
  };

  const handleBrainstormSend = async (overrideMessage = null) => {
    const msg = overrideMessage || brainstormInput.trim();
    if (!msg || isBrainstormLoading) return;

    const userMsg = { role: 'user', text: msg };
    setBrainstormMessages(prev => [...prev, userMsg]);
    setBrainstormInput('');
    setIsBrainstormLoading(true);

    // Push an empty assistant bubble that we'll stream into
    setBrainstormMessages(prev => [...prev, { role: 'assistant', text: '' }]);

    try {
      // Get current editor content dynamically based on type
      let currentEditorContext = '';
      if (editorType === 'docx' && pageRef.current) {
        currentEditorContext = pageRef.current.innerText || '';
      } else if (editorType === 'excel') {
        const activeSheetData = excelSheets[activeSheet]?.data || [];
        currentEditorContext = activeSheetData.map(row => row.map(c => c.value).join('\t')).join('\n');
      } else if (editorType === 'pptx') {
        currentEditorContext = content.map(slide => `Title: ${slide.title}\nContent: ${slide.content}`).join('\n\n');
      }

      const docContext = currentEditorContext.trim() 
        ? `\n\n[KONTEKS DOKUMEN EDIT SAAT INI]:\n${currentEditorContext.slice(0, 4000)}\n---` 
        : '';

      const fileContext = uploadedFileText
        ? `\n\n[KONTEKS FILE UPLOADED - "${uploadedFileName}"]: \n${uploadedFileText.slice(0, 4000)}\n---`
        : '';

      const systemPrompt = `Kamu adalah Brainstorm AI partner. Kamu sedang mengobrol di jendela melayang di atas editor dokumen.
Kamu memiliki akses ke isi dokumen yang sedang diedit saat ini dan file yang diunggah.

ATURAN REKOMENDASI AGENT:
Jika kamu merasa pengguna ingin atau sebaiknya menulis laporan/makalah/artikel baru menggunakan agen penulis otomatis (Word Agent) kami, kamu WAJIB memberikan opsi/tombol pelatuk dengan cara menuliskan tag '[RUN_AGENT: Nama Topik]' di bagian bawah pesanmu.
Contoh:
"Saya menyarankan kita membuat tulisan baru tentang Sejarah Fisika Kuantum. Klik tombol di bawah untuk meminta Agen menyusunnya bab-demi-bab:
[RUN_AGENT: Sejarah Fisika Kuantum]"

Ingat, format '[RUN_AGENT: Topik]' ini sangat krusial agar UI bisa memunculkan tombol pelatuk agen secara visual.`;

      // Build conversation history in the format sendMessageToGrok expects
      const historyForGrok = [
        { sender: 'system', text: systemPrompt },
        ...brainstormMessages.map(m => ({
          sender: m.role === 'user' ? 'user' : 'ai',
          text: m.text,
        }))
      ];

      // Full message with context injected
      const fullMsg = `${docContext}${fileContext}\n\nPertanyaan: ${msg}`;

      const abortCtrl = new AbortController();
      const auth = getAuthStatus();
      const response = await sendMessageToGrok(
        fullMsg,
        historyForGrok,
        'id',            // language
        null,            // conversationId
        'formal',        // personality
        abortCtrl,       // abortController
        'deepernova-1.2-flash', // model
        auth.isAuthenticated,   // isAuthenticated
        auth.isGuest,           // isGuest
        auth.userName,          // userName
        0,                      // sessionMessageCount
        [],                      // uploadedImages
      );

      // Process SSE stream — append each chunk to the last assistant bubble
      await processStreamingResponse(
        response,
        (chunk) => {
          setBrainstormMessages(prev => {
            const updated = [...prev];
            const last = updated[updated.length - 1];
            if (last?.role === 'assistant') {
              updated[updated.length - 1] = { role: 'assistant', text: last.text + chunk };
            }
            return updated;
          });
        },
        abortCtrl.signal
      );
    } catch (err) {
      console.error('[handleBrainstormSend]', err);
      setBrainstormMessages(prev => {
        const updated = [...prev];
        const last = updated[updated.length - 1];
        if (last?.role === 'assistant' && last.text === '') {
          updated[updated.length - 1] = { role: 'assistant', text: `❌ Gagal: ${err.message}` };
        } else {
          updated.push({ role: 'assistant', text: `❌ Gagal: ${err.message}` });
        }
        return updated;
      });
    } finally {
      setIsBrainstormLoading(false);
    }
  };

  const handleDocUploader = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsUploadingFile(true);
    setAiError('');

    const ext = file.name.split('.').pop().toLowerCase();
    const name = file.name;

    try {
      // ── DOCX: use mammoth with rich style map ──
      if (ext === 'docx' || ext === 'doc') {
        const arrayBuffer = await file.arrayBuffer();

        const styleMap = [
          "p[style-name='Heading 1'] => h1.doc-h1:fresh",
          "p[style-name='Heading 2'] => h2.doc-h2:fresh",
          "p[style-name='Heading 3'] => h3.doc-h3:fresh",
          "p[style-name='Heading 4'] => h4.doc-h4:fresh",
          "p[style-name='Title'] => h1.doc-title:fresh",
          "p[style-name='Subtitle'] => p.doc-subtitle:fresh",
          "p[style-name='Normal'] => p.doc-para:fresh",
          "p[style-name='Body Text'] => p.doc-para:fresh",
          "p[style-name='List Paragraph'] => p.doc-list-para:fresh",
          "r[style-name='Strong'] => strong",
          "r[style-name='Emphasis'] => em",
          "p:unordered-list(1) => ul > li:fresh",
          "p:unordered-list(2) => ul.doc-ul-2 > li:fresh",
          "p:ordered-list(1) => ol > li:fresh",
          "p:ordered-list(2) => ol.doc-ol-2 > li:fresh",
          "table => table.doc-table",
          "tr => tr",
          "td => td",
          "th => th",
        ];

        const result = await mammoth.convertToHtml({ arrayBuffer }, {
          styleMap,
          includeDefaultStyleMap: true,
        });

        // Post-process: fix alignment hints from inline style on Word paragraphs
        let htmlContent = result.value;

        // Wrap in a div that acts as a document container
        const styledHtml = `<div class="word-doc-view">${htmlContent}</div>`;

        // Extract plain text for AI context
        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = htmlContent;
        const plainText = tempDiv.innerText || tempDiv.textContent || '';

        // Load into editor
        setContent([{ text: styledHtml }]);
        if (pageRef.current) {
          pageRef.current.innerHTML = styledHtml;
          docxTextRef.current = styledHtml;
        }

        setUploadedFileName(name);
        setUploadedFileType('docx');
        setUploadedFileText(plainText.slice(0, 8000));
        setAiPrompt(`File "${name}" berhasil dimuat ke editor. Tolong analisis isinya dan berikan ringkasan serta rekomendasi.`);
        if (result.messages?.length > 0) {
          console.warn('[Mammoth warnings]', result.messages);
        }
      }
      // ── XLSX / XLS: use XLSX library ──
      else if (ext === 'xlsx' || ext === 'xls' || ext === 'csv') {
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const sheetTexts = [];

        workbook.SheetNames.forEach(sheetName => {
          const sheet = workbook.Sheets[sheetName];
          const csv = XLSX.utils.sheet_to_csv(sheet);
          sheetTexts.push(`--- Sheet: ${sheetName} ---\n${csv}`);
        });

        const fullText = sheetTexts.join('\n\n');

        // Build a simple HTML table preview for the first sheet
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
        let tableHtml = '<table border="1" style="border-collapse:collapse; width:100%; font-size:12px;">';
        rows.forEach((row, ri) => {
          tableHtml += '<tr>';
          row.forEach(cell => {
            const tag = ri === 0 ? 'th' : 'td';
            tableHtml += `<${tag} style="padding:4px 8px; border:1px solid #cbd5e1;">${cell ?? ''}</${tag}>`;
          });
          tableHtml += '</tr>';
        });
        tableHtml += '</table>';

        const htmlContent = `<h2>${name}</h2>${tableHtml}`;
        setContent([{ text: htmlContent }]);
        if (pageRef.current) {
          pageRef.current.innerHTML = htmlContent;
          docxTextRef.current = htmlContent;
        }

        setUploadedFileName(name);
        setUploadedFileType('xlsx');
        setUploadedFileText(fullText.slice(0, 8000));
        setAiPrompt(`File Excel "${name}" berhasil dimuat. Tolong analisis data spreadsheet ini dan berikan insight.`);
      }
      // ── TXT / JSON / CSV plain text ──
      else if (['txt', 'md', 'json', 'html', 'xml'].includes(ext)) {
        const text = await file.text();
        const htmlContent = `<pre style="font-family: monospace; font-size:12px; white-space: pre-wrap; word-break: break-word;">${text.replace(/</g, '&lt;').replace(/>/g, '&gt;')}</pre>`;

        setContent([{ text: htmlContent }]);
        if (pageRef.current) {
          pageRef.current.innerHTML = htmlContent;
          docxTextRef.current = htmlContent;
        }

        setUploadedFileName(name);
        setUploadedFileType(ext);
        setUploadedFileText(text.slice(0, 8000));
        setAiPrompt(`File "${name}" berhasil dimuat. Tolong analisis isinya.`);
      }
      // ── Unsupported ──
      else {
        setAiError(`Format file .${ext} belum didukung. Gunakan DOCX, XLSX, TXT, CSV, atau JSON.`);
      }
    } catch (err) {
      console.error('[handleDocUploader]', err);
      setAiError(`Gagal membaca file: ${err.message}`);
    } finally {
      setIsUploadingFile(false);
      // Reset input so same file can be re-uploaded
      e.target.value = '';
    }
  };

  const generateAndInsertCurve = () => {
    const canvas = document.createElement('canvas');
    canvas.width = 600;
    canvas.height = 400;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    
    ctx.strokeStyle = '#f1f5f9';
    ctx.lineWidth = 1;
    for (let x = 40; x < canvas.width; x += 40) {
      ctx.beginPath();
      ctx.moveTo(x, 0);
      ctx.lineTo(x, canvas.height);
      ctx.stroke();
    }
    for (let y = 40; y < canvas.height; y += 40) {
      ctx.beginPath();
      ctx.moveTo(0, y);
      ctx.lineTo(canvas.width, y);
      ctx.stroke();
    }
    
    ctx.strokeStyle = '#94a3b8';
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(40, canvas.height - 40);
    ctx.lineTo(canvas.width - 20, canvas.height - 40);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(40, 20);
    ctx.lineTo(40, canvas.height - 40);
    ctx.stroke();
    
    ctx.strokeStyle = curveColor;
    ctx.lineWidth = 3;
    ctx.beginPath();
    
    const startX = 40;
    const endX = canvas.width - 40;
    const centerY = canvas.height - 40;
    
    for (let px = startX; px <= endX; px++) {
      const x = (px - startX) / (endX - startX);
      let y = 0;
      
      if (curveEquation === 'sine') {
        y = Math.sin(x * Math.PI * 4) * (curveAmplitude / 100);
      } else if (curveEquation === 'exponential') {
        y = Math.pow(x, 2) * (curveAmplitude / 100);
      } else if (curveEquation === 'linear') {
        y = x * (curveAmplitude / 100);
      } else if (curveEquation === 'bell') {
        const mean = 0.5;
        const stdDev = 0.15;
        y = Math.exp(-0.5 * Math.pow((x - mean) / stdDev, 2)) * (curveAmplitude / 100);
      }
      
      const py = centerY - (y * (canvas.height - 80));
      if (px === startX) {
        ctx.moveTo(px, py);
      } else {
        ctx.lineTo(px, py);
      }
    }
    ctx.stroke();
    
    ctx.fillStyle = '#1e293b';
    ctx.font = 'bold 16px sans-serif';
    let title = 'Kurva Sinusoidal';
    if (curveEquation === 'exponential') title = 'Tren Eksponensial';
    if (curveEquation === 'linear') title = 'Kurva Regresi Linear';
    if (curveEquation === 'bell') title = 'Distribusi Normal (Gauss)';
    ctx.fillText(title, 50, 40);
    
    ctx.fillStyle = '#64748b';
    ctx.font = '12px sans-serif';
    ctx.fillText('0.0', 30, canvas.height - 25);
    ctx.fillText('1.0', endX - 10, canvas.height - 25);
    ctx.fillText('100%', 10, 30);
    
    const dataUrl = canvas.toDataURL('image/png');
    const imgHtml = `<div style="text-align:center; margin:18px 0;" class="docx-image-block"><img src="${dataUrl}" style="max-width:100%; height:auto; border-radius:4px; border:1px solid #e2e8f0; box-shadow: 0 4px 6px rgba(0,0,0,0.05);" /><p style="font-size:12px; color:#475569; font-weight:600; margin-top:6px;">📈 ${title}</p></div><p>&nbsp;</p>`;
    
    if (pageRef.current) {
      pageRef.current.focus();
      document.execCommand('insertHTML', false, imgHtml);
      docxTextRef.current = pageRef.current.innerHTML;
      syncDocxContent();
    }
    setShowChartModal(false);
  };

  const downloadCurveImage = () => {
    const canvas = document.createElement('canvas');
    canvas.width = 600;
    canvas.height = 400;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    
    ctx.strokeStyle = '#f1f5f9';
    ctx.lineWidth = 1;
    for (let x = 40; x < canvas.width; x += 40) {
      ctx.beginPath();
      ctx.moveTo(x, 0);
      ctx.lineTo(x, canvas.height);
      ctx.stroke();
    }
    for (let y = 40; y < canvas.height; y += 40) {
      ctx.beginPath();
      ctx.moveTo(0, y);
      ctx.lineTo(canvas.width, y);
      ctx.stroke();
    }
    
    ctx.strokeStyle = '#94a3b8';
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(40, canvas.height - 40);
    ctx.lineTo(canvas.width - 20, canvas.height - 40);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(40, 20);
    ctx.lineTo(40, canvas.height - 40);
    ctx.stroke();
    
    ctx.strokeStyle = curveColor;
    ctx.lineWidth = 3;
    ctx.beginPath();
    
    const startX = 40;
    const endX = canvas.width - 40;
    const centerY = canvas.height - 40;
    
    for (let px = startX; px <= endX; px++) {
      const x = (px - startX) / (endX - startX);
      let y = 0;
      
      if (curveEquation === 'sine') {
        y = Math.sin(x * Math.PI * 4) * (curveAmplitude / 100);
      } else if (curveEquation === 'exponential') {
        y = Math.pow(x, 2) * (curveAmplitude / 100);
      } else if (curveEquation === 'linear') {
        y = x * (curveAmplitude / 100);
      } else if (curveEquation === 'bell') {
        const mean = 0.5;
        const stdDev = 0.15;
        y = Math.exp(-0.5 * Math.pow((x - mean) / stdDev, 2)) * (curveAmplitude / 100);
      }
      
      const py = centerY - (y * (canvas.height - 80));
      if (px === startX) {
        ctx.moveTo(px, py);
      } else {
        ctx.lineTo(px, py);
      }
    }
    ctx.stroke();
    
    ctx.fillStyle = '#1e293b';
    ctx.font = 'bold 16px sans-serif';
    let title = 'sinusoidal';
    if (curveEquation === 'exponential') title = 'exponential';
    if (curveEquation === 'linear') title = 'linear_regression';
    if (curveEquation === 'bell') title = 'gauss_distribution';
    ctx.fillText(title, 50, 40);
    
    const dataUrl = canvas.toDataURL('image/png');
    const el = document.createElement('a');
    el.href = dataUrl;
    el.download = `${title}_curve.png`;
    el.click();
  };

  const triggerPreset = (promptText, title) => {
    if (title === "Tulis Draf Awal") {
      setPresetModalTitle("Tulis Draf Awal");
      setPresetModalPlaceholder("Masukkan topik untuk draf artikel/laporan bisnis Anda...");
      setPresetModalValue("riset teknologi terbaru");
      setPresetModalCallback(() => (topic) => {
        const finalPrompt = `Tuliskan draf artikel/laporan bisnis awal tentang topik ${topic.trim() || 'riset teknologi terbaru'}.`;
        setAiPrompt(finalPrompt);
        handleAiWrite(finalPrompt);
      });
      setShowPresetModal(true);
    } else if (title === "Buat Kerangka Kerja") {
      setPresetModalTitle("Buat Kerangka Kerja");
      setPresetModalPlaceholder("Masukkan topik/judul dokumen untuk dibuatkan kerangka kerja...");
      setPresetModalValue("dokumen baru");
      setPresetModalCallback(() => (topic) => {
        const finalPrompt = `Buatkan kerangka kerja (framework) dan outline bab untuk ${topic.trim() || 'dokumen baru'}.`;
        setAiPrompt(finalPrompt);
        handleAiWrite(finalPrompt);
      });
      setShowPresetModal(true);
    } else if (title === "Outline") {
      setPresetModalTitle("Buat Outline");
      setPresetModalPlaceholder("Masukkan topik laporan bisnis untuk outline...");
      setPresetModalValue("laporan bisnis");
      setPresetModalCallback(() => (topic) => {
        const finalPrompt = `Buatkan kerangka outline dokumen laporan bisnis tentang ${topic.trim() || 'laporan bisnis'} yang sangat lengkap dan terstruktur.`;
        setAiPrompt(finalPrompt);
        handleAiWrite(finalPrompt);
      });
      setShowPresetModal(true);
    } else {
      setAiPrompt(promptText);
      handleAiWrite(promptText);
    }
  };

  const getSmartRecommendations = () => {
    const html = content[0]?.text || '';
    const textLength = html.replace(/<[^>]*>/g, '').trim().length;
    
    const recs = [];
    if (textLength === 0 || html.includes('Mulai menulis')) {
      recs.push({
        title: "Tulis Draf Awal",
        icon: "fas fa-pen-nib",
        prompt: "Tuliskan draf artikel/laporan bisnis awal tentang topik riset teknologi terbaru."
      });
      recs.push({
        title: "Buat Kerangka Kerja",
        icon: "fas fa-lightbulb",
        prompt: "Buatkan kerangka kerja (framework) dan outline bab untuk dokumen baru."
      });
    } else {
      recs.push({
        title: "Buat Kesimpulan",
        icon: "fas fa-check-double",
        prompt: "Tambahkan paragraf kesimpulan dan poin-poin rekomendasi tindakan di akhir dokumen."
      });
      if (!html.includes('<table')) {
        recs.push({
          title: "Tambahkan Tabel Data",
          icon: "fas fa-table",
          prompt: "Buatkan tabel data analisis statistik yang relevan dengan isi dokumen saat ini."
        });
      }
      recs.push({
        title: "Parafrase Profesional",
        icon: "fas fa-sync-alt",
        prompt: "Parafrase dokumen ini agar bahasanya lebih mengalir, elegan, dan profesional."
      });
    }
    return recs;
  };

  // ===== AI WRITE =====
  const handleAiWrite = async (forcedPrompt = null) => {
    const activePrompt = forcedPrompt || aiPrompt;
    if (!activePrompt.trim()) return;
    setIsGenerating(true);
    setIsStreaming(true);
    setStreamingContent('');
    setAiError('');
    setGenerationProgress('Generating...');

    try {
      const systemContext = getSystemContext();
      const formattedMessages = [
        { sender: 'system', text: systemContext, timestamp: new Date().toISOString() },
        ...messages.map(msg => ({
          sender: msg.role === 'user' ? 'user' : 'assistant',
          text: msg.content || '',
          timestamp: new Date().toISOString()
        }))
      ];

      const response = await callAiService(activePrompt, formattedMessages);

      if (response?.body) {
        let fullContent = '';
        const reader = response.body.getReader();
        const decoder = new TextDecoder();

        try {
          while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            const chunk = decoder.decode(value, { stream: true });
            const lines = chunk.split('\n');
            for (const line of lines) {
              if (line.startsWith('data: ')) {
                try {
                  const jsonStr = line.slice(6);
                  if (jsonStr === '[DONE]') continue;
                  const json = JSON.parse(jsonStr);
                  if (json.choices?.[0]?.delta?.content) {
                    fullContent += json.choices[0].delta.content;
                    setStreamingContent(prev => prev + json.choices[0].delta.content);
                  }
                } catch (_e) {
                  // skip parse error
                }
              }
            }
          }
        } finally {
          reader.releaseLock();
        }

        if (fullContent.trim()) {
          const cleaned = cleanAiResponse(fullContent);
          setGenerationProgress('Done!');
          setIsStreaming(false);
          // insertAiContent returns the new content/sheets so we can save them immediately
          const inserted = insertAiContent(cleaned);
          const newMessages = [...messages, { role: 'user', content: activePrompt }, { role: 'assistant', content: cleaned }];
          setMessages(newMessages);
          messagesRef.current = newMessages; // Update ref immediately
          setAiResponse(cleaned);
          aiResponseRef.current = cleaned; // Update ref immediately
          setAiPrompt('');
          aiPromptRef.current = ''; // Update ref immediately
          // Save to session memory as artifact with the LATEST content
          saveArtifact(activePrompt, cleaned, inserted?.content, inserted?.sheets);
          setTimeout(() => { setGenerationProgress(''); setStreamingContent(''); }, 2000);
        } else {
          setAiError('No content generated.');
          setIsStreaming(false);
          setStreamingContent('');
        }
      } else {
        setAiError('Invalid response.');
        setIsStreaming(false);
        setStreamingContent('');
      }
    } catch (error) {
      setAiError(`Error: ${error.message}`);
      setGenerationProgress('');
      setIsStreaming(false);
      setStreamingContent('');
    } finally {
      setIsGenerating(false);
    }
  };

  const handleDocxAiFormat = async () => {
    if (editorType !== 'docx' || !pageRef.current) return;
    const rawText = pageRef.current.innerText.trim();
    if (!rawText) {
      setAiError('Tidak ada teks untuk diformat.');
      return;
    }

    setIsGenerating(true);
    setIsStreaming(true);
    setStreamingContent('');
    setAiError('');
    setGenerationProgress('Memformat dokumen...');

    try {
      const systemContext = getSystemContext();
      const formattedMessages = [
        { sender: 'system', text: systemContext, timestamp: new Date().toISOString() },
        ...messages.map(msg => ({ sender: msg.role === 'user' ? 'user' : 'assistant', text: msg.content || '', timestamp: new Date().toISOString() }))
      ];

      const prompt = `Format ulang teks dokumen akademik berikut menjadi makalah yang rapi dan terstruktur:\n\nATURAN FORMATTING TEKS:\n1. Setiap paragraf HARUS dipisahkan dengan DUA newline (\\n\\n)\n2. Gunakan heading/BAB dengan format: BAB I: JUDUL\\n\\nThen content\\n\\n\n3. Setiap bagian/section diberi nomor (BAB I, BAB II, dll)\n4. Jangan gunakan markdown, asterisk, atau simbol apapun\n5. Gunakan struktur: BAB -> Judul -> Isi paragraf (dengan newline ganda antar paragraf)\n6. Pastikan setiap paragraf berkualitas akademik tinggi\n\nATURAN FORMATTING TABEL (JIKA ADA DATA TABEL):\n- Jika terdapat data tabular, buat tabel dengan format EXACTLY:\n[TABLE]\nHeader1 | Header2 | Header3\nValue1 | Value2 | Value3\nValue1 | Value2 | Value3\n[/TABLE]\n- Gunakan pipe (|) untuk separator kolom\n- Baris pertama adalah header (direkomendasikan)\n- Satu baris per data\n- Tabel akan di-insert otomatis ke dokumen\n\nKAPABILITAS EDITOR:\n- [TABLE]...[/TABLE]: Untuk tabel data\n- Jika ada grafik/chart perlu, sebutkan dalam teks\n\nOUTPUT:\n- Hanya teks terformat, tabel dengan [TABLE] marker, tanpa penjelasan tambahan\n\nTeks untuk diformat:\n${rawText}`;
      const response = await callAiService(prompt, formattedMessages);

      if (response?.body) {
        let fullContent = '';
        const reader = response.body.getReader();
        const decoder = new TextDecoder();

        try {
          while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            const chunk = decoder.decode(value, { stream: true });
            const lines = chunk.split('\n');
            for (const line of lines) {
              if (line.startsWith('data: ')) {
                try {
                  const jsonStr = line.slice(6);
                  if (jsonStr === '[DONE]') continue;
                  const json = JSON.parse(jsonStr);
                  if (json.choices?.[0]?.delta?.content) {
                    fullContent += json.choices[0].delta.content;
                    setStreamingContent(prev => prev + json.choices[0].delta.content);
                  }
                } catch (_e) {
                  // skip parse error
                }
              }
            }
          }
        } finally {
          reader.releaseLock();
        }

        if (fullContent.trim()) {
          const cleaned = cleanAiResponse(fullContent);
          
          // Parse and extract tables from AI response
          const parsedTables = parseTablesFromText(cleaned);
          if (parsedTables.length > 0) {
            insertParsedTables(parsedTables);
          }
          
          // Remove table markers from text for display
          const textWithoutTables = removeTableMarkersFromText(cleaned);
          
          pageRef.current.innerText = textWithoutTables;
          docxTextRef.current = textWithoutTables;
          syncDocxContent();
          setGenerationProgress('Selesai');
          setIsStreaming(false);
          const newMessages = [...messages, { role: 'user', content: prompt }, { role: 'assistant', content: cleaned }];
          setMessages(newMessages);
          messagesRef.current = newMessages;
          setAiResponse(cleaned);
          aiResponseRef.current = cleaned;
          setTimeout(() => { setGenerationProgress(''); setStreamingContent(''); }, 2000);
        } else {
          setAiError('Tidak ada hasil format.');
          setIsStreaming(false);
          setStreamingContent('');
        }
      } else {
        setAiError('Invalid response.');
        setIsStreaming(false);
        setStreamingContent('');
      }
    } catch (error) {
      setAiError(`Error: ${error.message}`);
      setGenerationProgress('');
      setIsStreaming(false);
      setStreamingContent('');
    } finally {
      setIsGenerating(false);
    }
  };

  const handleStopStreaming = () => {
    setIsStreaming(false);
    setIsGenerating(false);
    setStreamingContent('');
    setGenerationProgress('Stopped');
    setTimeout(() => setGenerationProgress(''), 1200);
  };

  const cleanAiResponse = (text) => {
    if (!text) return '';
    let cleaned = text
      .replace(/^(baik|oke|ok)\s+(saya\s+)?((akan\s+)?membuat|buat|buatkan|tulis|tuliskan|saya\s+)?.*?[:\n]/gi, '')
      .replace(/^(berikut|ini\s+)?konten.*?[:\n]/gi, '')
      .replace(/^siap[,:].*?\n/gi, '')
      .replace(/^tentu[,:].*?\n/gi, '')
      .replace(/```[\s\S]*?```/g, '').replace(/`/g, '')
      .replace(/^#+\s+/gm, '')
      .replace(/\*\*(.*?)\*\*/g, '$1').replace(/\*(.*?)\*/g, '$1')
      .replace(/\n+(semoga|harap|terima.*?kesih|regards|best|thanks).*$/gi, '')
      .trim();
    while (cleaned.match(/^[\s\-_•#.]+/) || cleaned.match(/[\s\-_•#.]+$/)) {
      cleaned = cleaned.replace(/^[\s\-_•#.]+/, '').replace(/[\s\-_•#.]+$/, '');
    }
    return cleaned;
  };

  const convertMarkdownToHtml = (markdown) => {
    if (!markdown) return '';
    let html = markdown
      .replace(/\r\n/g, '\n')
      .replace(/\r/g, '\n');

    // Convert headings
    html = html.replace(/^### (.*$)/gim, '<h3>$1</h3>');
    html = html.replace(/^## (.*$)/gim, '<h2>$1</h2>');
    html = html.replace(/^# (.*$)/gim, '<h1>$1</h1>');

    // Convert tables (lines containing |)
    const lines = html.split('\n');
    let inTable = false;
    let tableHtml = '';
    const newLines = [];

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      if (line.startsWith('|') && line.endsWith('|')) {
        if (!inTable) {
          inTable = true;
          tableHtml = '<table style="width:100%; border-collapse:collapse; margin:14px 0;"><tbody>';
        }
        const cells = line.split('|').slice(1, -1).map(c => c.trim());
        const isHeaderSep = cells.every(c => /^:-+:|^-+$/g.test(c));
        if (!isHeaderSep) {
          tableHtml += '<tr>';
          cells.forEach(cell => {
            tableHtml += `<td style="border:1px solid #e2e8f0; padding:8px 12px; text-align:left;">${cell}</td>`;
          });
          tableHtml += '</tr>';
        }
      } else {
        if (inTable) {
          tableHtml += '</tbody></table>';
          newLines.push(tableHtml);
          inTable = false;
          tableHtml = '';
        }
        newLines.push(line);
      }
    }
    if (inTable) {
      tableHtml += '</tbody></table>';
      newLines.push(tableHtml);
    }
    html = newLines.join('\n');

    // Convert lists
    html = html.replace(/^\s*[-*•]\s+(.*)$/gim, '<ul><li>$1</li></ul>');
    html = html.replace(/<\/ul>\s*<ul>/g, '');

    html = html.replace(/^\s*\d+[.)]\s+(.*)$/gim, '<ol><li>$1</li></ol>');
    html = html.replace(/<\/ol>\s*<ol>/g, '');

    // Convert paragraphs
    html = html.split('\n').map(line => {
      const trimmed = line.trim();
      if (!trimmed) return '';
      if (trimmed.startsWith('<h') || trimmed.startsWith('<table') || trimmed.startsWith('<ul') || trimmed.startsWith('<ol') || trimmed.startsWith('<li') || trimmed.startsWith('<tr') || trimmed.startsWith('<td') || trimmed.startsWith('</')) {
        return line;
      }
      return `<p>${line}</p>`;
    }).join('\n');

    return html;
  };

  // ===== INSERT AI CONTENT - PREPEND TO PAGE 1 =====
  // Returns the new content so it can be saved to artifacts immediately
  const insertAiContent = (responseContent) => {
    if (!responseContent || typeof responseContent !== 'string') return null;
    let cleaned = cleanAiResponse(responseContent);
    if (!cleaned.trim()) return null;

    let newContent = null;
    let newSheets = null;

    switch (editorType) {
      case 'docx': {
        const insertedHtml = convertMarkdownToHtml(cleaned);
        const currentHtml = contentRef.current[0]?.text || '';
        const newHtml = insertedHtml + '<br/>' + currentHtml;
        newContent = [{ id: Date.now(), type: 'html', text: newHtml }];
        setContent(newContent);
        contentRef.current = newContent;
        if (pageRef.current) {
          pageRef.current.innerHTML = newHtml;
        }
        break;
      }
      case 'pptx': {
        const slides = cleaned.split(/---|\n\n---|\n---\n/).filter(Boolean);
        const newSlides = slides.map((s, idx) => {
          const lines = s.trim().split('\n').filter(Boolean);
          return { id: Date.now() + idx, type: 'slide', title: lines[0]?.trim() || `Slide ${idx + 1}`, content: lines.slice(1).join('\n').trim() || 'Konten slide', notes: '' };
        });
        if (newSlides.length) {
          newContent = [...newSlides, ...contentRef.current];
          setContent(newContent);
          contentRef.current = newContent; // Update ref immediately
        }
        break;
      }
      case 'excel': {
        const parsed = parseExcelContent(cleaned);
        if (parsed) {
          if (parsed.type === 'multi_sheet') {
            newSheets = parsed.sheets;
            setExcelSheets(newSheets);
            excelSheetsRef.current = newSheets; // Update ref immediately
            setActiveSheet(0);
          } else {
            newSheets = [...excelSheetsRef.current];
            const sheetIdx = activeSheet;
            if (parsed.data.length > 0) {
              const isCurrentEmpty = newSheets[sheetIdx].data.length === 1 &&
                newSheets[sheetIdx].data[0].length === 1 &&
                newSheets[sheetIdx].data[0][0]?.value === '';
              if (isCurrentEmpty) {
                newSheets[sheetIdx] = {
                  ...newSheets[sheetIdx], data: parsed.data,
                  merges: parsed.merges || [],
                  colWidths: parsed.colWidths || Array(Math.max(...parsed.data.map(r => r.length), 1)).fill(100),
                };
              } else {
                newSheets[sheetIdx] = { ...newSheets[sheetIdx], data: [...parsed.data, ...newSheets[sheetIdx].data] };
              }
              setExcelSheets(newSheets);
              excelSheetsRef.current = newSheets; // Update ref immediately
            }
          }
        }
        break;
      }
      default: break;
    }
    
    return { content: newContent, sheets: newSheets };
  };

  // ===== EXCEL OPERATIONS =====
  const updateCell = (r, c, value) => {
    const newSheets = [...excelSheets];
    const sheet = { ...newSheets[activeSheet], data: [...newSheets[activeSheet].data] };
    if (!sheet.data[r]) sheet.data[r] = [];
    if (!sheet.data[r][c]) sheet.data[r][c] = createCell();
    sheet.data[r] = [...sheet.data[r]];
    sheet.data[r][c] = { ...sheet.data[r][c], value: String(value ?? '') };
    newSheets[activeSheet] = sheet;
    setExcelSheets(newSheets);
  };

  const updateCellFormat = (r, c, formatChanges) => {
    const newSheets = [...excelSheets];
    const sheet = { ...newSheets[activeSheet], data: [...newSheets[activeSheet].data] };
    if (!sheet.data[r]) sheet.data[r] = [];
    if (!sheet.data[r][c]) sheet.data[r][c] = createCell();
    sheet.data[r] = [...sheet.data[r]];
    sheet.data[r][c] = { ...sheet.data[r][c], format: { ...sheet.data[r][c].format, ...formatChanges } };
    newSheets[activeSheet] = sheet;
    setExcelSheets(newSheets);
  };

  const applyFormatToSelection = (formatChanges) => {
    if (!selectedCell) return;
    updateCellFormat(selectedCell.r, selectedCell.c, formatChanges);
  };

  const addExcelRow = () => {
    const newSheets = [...excelSheets];
    const sheet = newSheets[activeSheet];
    const cols = sheet.data[0]?.length || 8;
    sheet.data = [...sheet.data, createRow(cols)];
    sheet.rowHeights = [...(sheet.rowHeights || []), 32];
    newSheets[activeSheet] = sheet;
    setExcelSheets(newSheets);
  };

  const addExcelColumn = () => {
    const newSheets = [...excelSheets];
    const sheet = newSheets[activeSheet];
    sheet.data = sheet.data.map(row => [...row, createCell('')]);
    sheet.colWidths = [...(sheet.colWidths || Array(sheet.data[0]?.length || 1).fill(100)), 100];
    newSheets[activeSheet] = sheet;
    setExcelSheets(newSheets);
  };

  const addExcelSheet = () => {
    const cols = excelSheets[activeSheet]?.data[0]?.length || 8;
    setExcelSheets([...excelSheets, createSheet(`Sheet${excelSheets.length + 1}`, 10, cols)]);
    setActiveSheet(excelSheets.length);
  };

  const deleteExcelSheet = (idx) => {
    if (excelSheets.length <= 1) return;
    const newSheets = excelSheets.filter((_, i) => i !== idx);
    setExcelSheets(newSheets);
    if (activeSheet >= newSheets.length) setActiveSheet(newSheets.length - 1);
  };

  const deleteExcelRow = (r) => {
    const newSheets = [...excelSheets];
    const sheet = newSheets[activeSheet];
    if (sheet.data.length <= 1) return;
    sheet.data = sheet.data.filter((_, i) => i !== r);
    sheet.rowHeights = (sheet.rowHeights || []).filter((_, i) => i !== r);
    newSheets[activeSheet] = sheet;
    setExcelSheets(newSheets);
    setSelectedCell(null);
  };

  const deleteExcelColumn = (c) => {
    const newSheets = [...excelSheets];
    const sheet = newSheets[activeSheet];
    if ((sheet.data[0]?.length || 0) <= 1) return;
    sheet.data = sheet.data.map(row => row.filter((_, i) => i !== c));
    sheet.colWidths = (sheet.colWidths || []).filter((_, i) => i !== c);
    newSheets[activeSheet] = sheet;
    setExcelSheets(newSheets);
    setSelectedCell(null);
  };

  const insertExcelRowAbove = (r) => {
    const newSheets = [...excelSheets];
    const sheet = newSheets[activeSheet];
    const cols = sheet.data[0]?.length || 8;
    sheet.data = [...sheet.data.slice(0, r), createRow(cols), ...sheet.data.slice(r)];
    sheet.rowHeights = [...(sheet.rowHeights || Array(sheet.data.length - 1).fill(32)), 32];
    newSheets[activeSheet] = sheet;
    setExcelSheets(newSheets);
  };

  const insertExcelColumnLeft = (c) => {
    const newSheets = [...excelSheets];
    const sheet = newSheets[activeSheet];
    sheet.data = sheet.data.map(row => [...row.slice(0, c), createCell(''), ...row.slice(c)]);
    sheet.colWidths = [...(sheet.colWidths || Array(sheet.data[0]?.length - 1 || 1).fill(100)), 100];
    newSheets[activeSheet] = sheet;
    setExcelSheets(newSheets);
  };

  const sortExcelData = (colIdx, dir = 'asc') => {
    const newSheets = [...excelSheets];
    const sheet = newSheets[activeSheet];
    const header = sheet.data[0];
    const body = sheet.data.slice(1);
    body.sort((a, b) => {
      const va = (a[colIdx]?.value || '').toLowerCase();
      const vb = (b[colIdx]?.value || '').toLowerCase();
      if (va < vb) return dir === 'asc' ? -1 : 1;
      if (va > vb) return dir === 'asc' ? 1 : -1;
      return 0;
    });
    sheet.data = [header, ...body];
    newSheets[activeSheet] = sheet;
    setExcelSheets(newSheets);
  };

  const findAndReplace = () => {
    if (!findText) return;
    const newSheets = [...excelSheets];
    const sheet = newSheets[activeSheet];
    let count = 0;
    sheet.data = sheet.data.map(row =>
      row.map(cell => {
        if (cell.value.toLowerCase().includes(findText.toLowerCase())) {
          count++;
          if (replaceText !== undefined) {
            return { ...cell, value: cell.value.replace(new RegExp(findText, 'gi'), replaceText) };
          }
          return cell;
        }
        return cell;
      })
    );
    newSheets[activeSheet] = sheet;
    setExcelSheets(newSheets);
    alert(`Found ${count} cell(s)` + (replaceText ? `, replaced ${count}` : ''));
  };

  const getColumnLabel = (col) => {
    let label = '';
    let n = col;
    while (n >= 0) {
      label = String.fromCharCode(65 + (n % 26)) + label;
      n = Math.floor(n / 26) - 1;
    }
    return label;
  };

  const getCellStyle = (cell, r, c) => {
    if (!cell || !cell.format) return {};
    const f = cell.format;
    const isSelected = selectedCell?.r === r && selectedCell?.c === c;
    return {
      fontWeight: f.bold ? 'bold' : 'normal',
      fontStyle: f.italic ? 'italic' : 'normal',
      textDecoration: [f.underline ? 'underline' : '', f.strikethrough ? 'line-through' : ''].filter(Boolean).join(' ') || 'none',
      fontSize: `${f.fontSize || 11}px`,
      fontFamily: f.fontFamily || 'Calibri',
      color: f.fontColor || '#000000',
      backgroundColor: f.fillColor || (isSelected ? '#e8f0fe' : (r === 0 ? '#fff6f0' : '')),
      textAlign: f.halign || 'left',
      verticalAlign: f.valign || 'middle',
      whiteSpace: f.wrapText ? 'pre-wrap' : 'nowrap',
      wordBreak: f.wrapText ? 'break-word' : 'normal',
      minWidth: 80,
      borderTop: f.borderTop || '1px solid #e0e0e0',
      borderBottom: f.borderBottom || '1px solid #e0e0e0',
      borderLeft: f.borderLeft || '1px solid #e0e0e0',
      borderRight: f.borderRight || '1px solid #e0e0e0',
      outline: isSelected ? '2px solid #ff6b00' : 'none',
      outlineOffset: isSelected ? '-1px' : '0',
    };
  };

  // ===== RENDER =====
  const renderEditor = () => {
    switch (editorType) {
      case 'docx': return renderDocxEditor();
      case 'pptx': return renderPptxEditor();
      case 'excel': return renderExcelEditor();
      default: return null;
    }
  };

  // ===== DOCX TABLE OPERATIONS =====
  const addDocxTable = (rows = 3, cols = 3) => {
    let tableHtml = '<table style="width:100%; border-collapse:collapse; margin:14px 0;"><tbody>';
    for (let r = 0; r < rows; r++) {
      tableHtml += '<tr>';
      for (let c = 0; c < cols; c++) {
        tableHtml += '<td style="border:1px solid #cbd5e1; padding:8px 12px; min-width:60px;">&nbsp;</td>';
      }
      tableHtml += '</tr>';
    }
    tableHtml += '</tbody></table><p>&nbsp;</p>';
    
    if (pageRef.current) {
      pageRef.current.focus();
      document.execCommand('insertHTML', false, tableHtml);
      docxTextRef.current = pageRef.current.innerHTML;
      syncDocxContent();
    }
  };

  // ===== DOCX CHART OPERATIONS =====
  const insertChart = () => {
    const chartHtml = `<div class="docx-chart-block" data-chart-type="${chartType}" data-chart-title="${chartTitle}" data-chart-data='${JSON.stringify(chartData)}' style="border: 1px solid #cbd5e1; border-radius: 8px; padding: 16px; margin: 16px 0; background: #fafafa; text-align: center; font-family: sans-serif; cursor: default;" contenteditable="false">
      <h4 style="margin: 0 0 12px 0; font-size: 14px; font-weight: bold; color: #1e293b;">${chartTitle} (${chartType.toUpperCase()})</h4>
      <div style="color: #64748b; font-size: 12px; border: 1px dashed #cbd5e1; padding: 24px; background: #ffffff; border-radius: 6px;">
        📊 [Grafik ${chartType} dimasukkan di sini. Simpan/Ekspor dokumen untuk melihat hasil.]
      </div>
    </div><p>&nbsp;</p>`;
    
    if (pageRef.current) {
      pageRef.current.focus();
      document.execCommand('insertHTML', false, chartHtml);
      docxTextRef.current = pageRef.current.innerHTML;
      syncDocxContent();
    }
    setShowChartModal(false);
  };

  const updateChartData = (text) => {
    try {
      const lines = text.trim().split('\n');
      const parsed = lines.map(line => {
        const [name, value] = line.split(':').map(s => s.trim());
        return { name, value: parseInt(value) || 0 };
      }).filter(d => d.name && d.value);
      if (parsed.length > 0) setChartData(parsed);
    } catch (_e) {
      console.warn('Chart data parse error');
    }
  };

  const deleteChart = (chartIdx) => {
    setDocxCharts(prev => prev.filter((_, i) => i !== chartIdx));
  };

  const renderChart = (chart) => {
    const chartColors = ['#4472c4', '#70ad47', '#ed7d31', '#ffc000', '#5b9bd5'];
    if (chart.type === 'bar') {
      return (
        <ResponsiveContainer width="100%" height={250}>
          <BarChart data={chart.data}>
            <CartesianGrid strokeDasharray="3 3" />
            <XAxis dataKey="name" />
            <YAxis />
            <Tooltip />
            <Bar dataKey="value" fill="#4472c4" />
          </BarChart>
        </ResponsiveContainer>
      );
    } else if (chart.type === 'line') {
      return (
        <ResponsiveContainer width="100%" height={250}>
          <LineChart data={chart.data}>
            <CartesianGrid strokeDasharray="3 3" />
            <XAxis dataKey="name" />
            <YAxis />
            <Tooltip />
            <Line type="monotone" dataKey="value" stroke="#4472c4" />
          </LineChart>
        </ResponsiveContainer>
      );
    } else if (chart.type === 'pie') {
      return (
        <ResponsiveContainer width="100%" height={250}>
          <PieChart>
            <Pie dataKey="value" data={chart.data} cx="50%" cy="50%" labelLine={false} label>
              {chart.data.map((_, idx) => (
                <Cell key={idx} fill={chartColors[idx % chartColors.length]} />
              ))}
            </Pie>
            <Tooltip />
          </PieChart>
        </ResponsiveContainer>
      );
    }
  };

  // ===== DOCX IMAGE =====
  const handleImageUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const dataUrl = ev.target?.result;
      if (typeof dataUrl === 'string') {
        const imgHtml = `<div style="text-align:center; margin:14px 0;"><img src="${dataUrl}" style="max-width:100%; height:auto; border-radius:4px;" /><p style="font-size:11px; color:#666; margin-top:4px;">${file.name}</p></div><p>&nbsp;</p>`;
        if (pageRef.current) {
          pageRef.current.focus();
          document.execCommand('insertHTML', false, imgHtml);
          docxTextRef.current = pageRef.current.innerHTML;
          syncDocxContent();
        }
        setShowInsertImage(false);
      }
    };
    reader.readAsDataURL(file);
  };

  useEffect(() => {
    if (editorType === 'docx' && pageRef.current) {
      const html = content[0]?.text || '';
      if (pageRef.current.innerHTML !== html) {
        pageRef.current.innerHTML = html;
      }
    }
  }, [editorType, content]);

  const handleNewDocument = () => {
    if (confirm('Buat dokumen baru? Konten yang belum tersimpan akan hilang.')) {
      setActiveCloudFileId(null);
      setDocumentTitle('Untitled Document');
      if (editorType === 'excel') {
        setExcelSheets([createSheet('Sheet1', 20, 10)]);
        setActiveSheet(0);
      } else {
        setContent([{ id: Date.now(), type: 'html', text: '' }]);
        if (pageRef.current) pageRef.current.innerHTML = '';
      }
    }
  };

  const handleOpenFromCloud = () => {
    setExplorerMode('open');
    setSelectedCloudFile(null);
    setShowCloudModal(true);
    fetchCloudFiles();
  };

  const handleSaveToCloud = () => {
    if (activeCloudFileId) {
      saveActiveFileToCloud(documentTitle);
    } else {
      setExplorerMode('save');
      setCloudFileName(documentTitle);
      setSelectedCloudFile(null);
      setShowCloudModal(true);
      fetchCloudFiles();
    }
  };

  const handleSaveAsToCloud = () => {
    setExplorerMode('save');
    setCloudFileName(documentTitle);
    setSelectedCloudFile(null);
    setShowCloudModal(true);
    fetchCloudFiles();
  };

  const renderDocxEditor = () => {
    return (
      <div className="docx-editor-wrapper">
        {/* Microsoft Word Ribbon Tabs Header */}
        <div className="word-ribbon-tabs">
          <button 
            className={`ribbon-tab-btn ${activeRibbonTab === 'file' ? 'active' : ''}`} 
            onClick={() => setActiveRibbonTab('file')}
            style={{ backgroundColor: activeRibbonTab === 'file' ? '#106ebe' : 'transparent', color: activeRibbonTab === 'file' ? '#ffffff' : '#323130', fontWeight: 'bold' }}
          >
            Berkas (File)
          </button>
          <button 
            className={`ribbon-tab-btn ${activeRibbonTab === 'home' ? 'active' : ''}`} 
            onClick={() => setActiveRibbonTab('home')}
          >
            Beranda (Home)
          </button>
          <button 
            className={`ribbon-tab-btn ${activeRibbonTab === 'insert' ? 'active' : ''}`} 
            onClick={() => setActiveRibbonTab('insert')}
          >
            Sisipkan (Insert)
          </button>
          <button 
            className={`ribbon-tab-btn ${activeRibbonTab === 'layout' ? 'active' : ''}`} 
            onClick={() => setActiveRibbonTab('layout')}
          >
            Tata Letak (Layout)
          </button>
          <button 
            className={`ribbon-tab-btn ${activeRibbonTab === 'ai' ? 'active' : ''}`} 
            onClick={() => setActiveRibbonTab('ai')}
          >
            Deepernova AI ✨
          </button>
        </div>

        {/* Microsoft Word Ribbon Toolbar Content */}
        <div className="word-toolbar">
          {activeRibbonTab === 'file' && (
            <>
              {/* Group 1: New/Open/Save/Save As */}
              <div className="toolbar-group" title="Berkas Cloud">
                <button className="toolbar-btn" onClick={handleNewDocument} title="Dokumen Baru"><i className="fas fa-file-alt" style={{ marginRight: 6 }}></i> Baru</button>
                <button className="toolbar-btn" onClick={handleOpenFromCloud} title="Buka dari Cloud"><i className="fas fa-folder-open" style={{ marginRight: 6 }}></i> Buka</button>
                <button className="toolbar-btn" onClick={handleSaveToCloud} title="Simpan ke Cloud"><i className="fas fa-save" style={{ marginRight: 6 }}></i> Simpan</button>
                <button className="toolbar-btn" onClick={handleSaveAsToCloud} title="Simpan Sebagai..."><i className="fas fa-file-export" style={{ marginRight: 6 }}></i> Simpan Sebagai</button>
              </div>
            </>
          )}
          {activeRibbonTab === 'home' && (
            <>
              {/* Group 1: History */}
              <div className="toolbar-group" title="Riwayat">
                <button className="toolbar-btn" onClick={() => formatText('undo')} title="Batal (Undo)"><i className="fas fa-undo"></i></button>
                <button className="toolbar-btn" onClick={() => formatText('redo')} title="Ulangi (Redo)"><i className="fas fa-redo"></i></button>
              </div>
              
              {/* Group 2: Typography family & size & line spacing */}
              <div className="toolbar-group" title="Huruf">
                <select value={fontFamily} onChange={e => handleFontFamily(e.target.value)} className="toolbar-select" title="Jenis Huruf">
                  {['Times New Roman','Arial','Calibri','Courier New','Georgia','Verdana'].map(f => (
                    <option key={f} value={f}>{f}</option>
                  ))}
                </select>
                <select value={fontSize} onChange={e => handleFontSize(e.target.value)} className="toolbar-select" title="Ukuran Huruf">
                  {[10,11,12,13,14,16,18,20,24,28,32,36,48,72].map(s => (
                    <option key={s} value={`${s}pt`}>{s}</option>
                  ))}
                </select>
                {/* Text Case Changer dropdown */}
                <select onChange={e => { if(e.target.value) { changeTextCase(e.target.value); e.target.value = ''; } }} className="toolbar-select" title="Ubah Kapitalisasi (Change Case)" style={{ width: '60px' }}>
                  <option value="">Aa</option>
                  <option value="upper">UPPERCASE</option>
                  <option value="lower">lowercase</option>
                  <option value="capitalize">Capitalize Each Word</option>
                  <option value="sentence">Sentence case</option>
                </select>
                <select value={lineSpacing} onChange={e => handleLineSpacing(e.target.value)} className="toolbar-select" title="Spasi Baris (Line Spacing)">
                  <option value="1.0">Spasi 1.0</option>
                  <option value="1.15">Spasi 1.15</option>
                  <option value="1.5">Spasi 1.5</option>
                  <option value="2.0">Spasi 2.0</option>
                </select>
              </div>

              {/* Group 3: Core formatting */}
              <div className="toolbar-group" title="Pemformatan Teks">
                <button className="toolbar-btn" onClick={() => formatText('bold')} title="Tebal (Bold)"><i className="fas fa-bold"></i></button>
                <button className="toolbar-btn" onClick={() => formatText('italic')} title="Miring (Italic)"><i className="fas fa-italic"></i></button>
                <button className="toolbar-btn" onClick={() => formatText('underline')} title="Garis Bawah (Underline)"><i className="fas fa-underline"></i></button>
                <button className="toolbar-btn" onClick={() => formatText('strikeThrough')} title="Coret (Strikethrough)"><i className="fas fa-strikethrough"></i></button>
                {/* Subscript / Superscript */}
                <button className="toolbar-btn" onClick={() => applyFormatting('subscript')} title="Subskrip (x₂)"><i className="fas fa-subscript"></i></button>
                <button className="toolbar-btn" onClick={() => applyFormatting('superscript')} title="Superskrip (x²)"><i className="fas fa-superscript"></i></button>
              </div>

              {/* Group 4: Colors & Eraser */}
              <div className="toolbar-group" title="Warna & Bersihkan">
                <label className="toolbar-btn color-picker-label" title="Warna Huruf (Text Color)">
                  <i className="fas fa-font"></i>
                  <input type="color" value={textColor} onChange={e => handleTextColor(e.target.value)} className="toolbar-color-input" />
                </label>
                <label className="toolbar-btn color-picker-label" title="Warna Stabilo (Highlight Color)">
                  <i className="fas fa-highlighter"></i>
                  <input type="color" value={highlightColor} onChange={e => handleHighlightColor(e.target.value)} className="toolbar-color-input" />
                </label>
                <button className="toolbar-btn" onClick={handleClearFormatting} title="Hapus Pemformatan"><i className="fas fa-eraser"></i></button>
              </div>

              {/* Group 5: Alignments & Indents */}
              <div className="toolbar-group" title="Perataan Paragraf">
                <button className="toolbar-btn" onClick={() => formatText('justifyLeft')} title="Rata Kiri"><i className="fas fa-align-left"></i></button>
                <button className="toolbar-btn" onClick={() => formatText('justifyCenter')} title="Rata Tengah"><i className="fas fa-align-center"></i></button>
                <button className="toolbar-btn" onClick={() => formatText('justifyRight')} title="Rata Kanan"><i className="fas fa-align-right"></i></button>
                <button className="toolbar-btn" onClick={() => formatText('justifyFull')} title="Rata Kanan-Kiri (Justify)"><i className="fas fa-align-justify"></i></button>
                <button className="toolbar-btn" onClick={() => formatText('outdent')} title="Kurangi Indentasi"><i className="fas fa-outdent"></i></button>
                <button className="toolbar-btn" onClick={() => formatText('indent')} title="Tambah Indentasi"><i className="fas fa-indent"></i></button>
                {/* Toggle Formatting Marks (¶) */}
                <button className={`toolbar-btn ${showFormattingMarks ? 'active' : ''}`} onClick={() => setShowFormattingMarks(!showFormattingMarks)} title="Tampilkan Tanda Paragraf (¶)"><i className="fas fa-paragraph"></i></button>
              </div>

              {/* Group 6: Lists */}
              <div className="toolbar-group" title="Daftar (Lists)">
                <button className="toolbar-btn" onClick={() => formatText('insertUnorderedList')} title="Daftar Simbol (Bullet List)"><i className="fas fa-list-ul"></i></button>
                <button className="toolbar-btn" onClick={() => formatText('insertOrderedList')} title="Daftar Angka (Numbered List)"><i className="fas fa-list-ol"></i></button>
              </div>
            </>
          )}

          {activeRibbonTab === 'insert' && (
            <>
              {/* Insert Elements */}
              <div className="toolbar-group" title="Sisipkan Elemen">
                <button className="toolbar-btn" onClick={() => addDocxTable(3, 3)} title="Sisipkan Tabel Baru"><i className="fas fa-table" style={{ marginRight: 6 }}></i> Tabel</button>
                <button className="toolbar-btn" onClick={() => setShowInsertImage(true)} title="Sisipkan Gambar"><i className="fas fa-image" style={{ marginRight: 6 }}></i> Gambar</button>
                <button className="toolbar-btn" onClick={() => setShowChartModal(true)} title="Sisipkan Grafik Data"><i className="fas fa-chart-bar" style={{ marginRight: 6 }}></i> Grafik</button>
              </div>

              <div className="toolbar-group" title="Sisipkan Dokumen">
                <button className="toolbar-btn" onClick={promptLink} title="Sisipkan Tautan (Link)"><i className="fas fa-link" style={{ marginRight: 6 }}></i> Tautan</button>
                <button className="toolbar-btn" onClick={() => applyFormatting('insertHorizontalRule')} title="Garis Mendatar (Horizontal Line)"><i className="fas fa-minus" style={{ marginRight: 6 }}></i> Garis Pembatas</button>
                <button className="toolbar-btn" onClick={() => {
                  const quoteHtml = '<blockquote style="border-left: 4px solid #106ebe; padding-left: 12px; margin: 10px 0; color: #555; font-style: italic;">Kutipan: </blockquote><p>&nbsp;</p>';
                  applyFormatting('insertHTML', quoteHtml);
                }} title="Sisipkan Kotak Kutipan"><i className="fas fa-quote-right" style={{ marginRight: 6 }}></i> Kotak Kutipan</button>
                <button className="toolbar-btn" onClick={insertPageBreak} title="Sisipkan Batas Halaman (Page Break)"><i className="fas fa-columns" style={{ marginRight: 6 }}></i> Batas Halaman</button>
              </div>
            </>
          )}

          {activeRibbonTab === 'layout' && (
            <>
              {/* Layout Display Mode Toggle */}
              <div className="toolbar-group" title="Mode Tampilan">
                <button className={`toolbar-btn ${!isDraftMode ? 'active' : ''}`} onClick={() => setIsDraftMode(false)}>Mode Kertas A4</button>
                <button className={`toolbar-btn ${isDraftMode ? 'active' : ''}`} onClick={() => setIsDraftMode(true)}>Mode Web Fluid</button>
              </div>

              {!isDraftMode && (
                <>
                  {/* Paper Size & Margins & Orientation & Columns & Border */}
                  <div className="toolbar-group" title="Ukuran & Tata Letak">
                    <span style={{ fontSize: 11, color: '#666', padding: '0 4px' }}>Kertas:</span>
                    <select className="toolbar-select" value={paperSize} onChange={e => setPaperSize(e.target.value)}>
                      <option value="a4">A4 (21 x 29.7 cm)</option>
                      <option value="letter">Letter (21.6 x 27.9 cm)</option>
                      <option value="legal">Legal (21.6 x 35.6 cm)</option>
                      <option value="a5">A5 (14.8 x 21 cm)</option>
                      <option value="custom">Kustom...</option>
                    </select>

                    <span style={{ fontSize: 11, color: '#666', padding: '0 4px' }}>Margin:</span>
                    <select className="toolbar-select" value={paperMargin} onChange={e => setPaperMargin(e.target.value)}>
                      <option value="normal">Normal (2.54 cm)</option>
                      <option value="narrow">Sempit (1.27 cm)</option>
                      <option value="wide">Lebar (5.08 cm)</option>
                      <option value="custom">Kustom...</option>
                    </select>

                    <span style={{ fontSize: 11, color: '#666', padding: '0 4px' }}>Orientasi:</span>
                    <select className="toolbar-select" value={paperOrientation} onChange={e => setPaperOrientation(e.target.value)}>
                      <option value="portrait">Tegak (Portrait)</option>
                      <option value="landscape">Mendatar (Landscape)</option>
                    </select>

                    <span style={{ fontSize: 11, color: '#666', padding: '0 4px' }}>Kolom:</span>
                    <select className="toolbar-select" value={pageColumns} onChange={e => setPageColumns(parseInt(e.target.value))}>
                      <option value={1}>1 Kolom</option>
                      <option value={2}>2 Kolom</option>
                      <option value={3}>3 Kolom</option>
                    </select>

                    <span style={{ fontSize: 11, color: '#666', padding: '0 4px' }}>Bingkai:</span>
                    <select className="toolbar-select" value={pageBorder} onChange={e => setPageBorder(e.target.value)}>
                      <option value="none">Tanpa Bingkai</option>
                      <option value="solid">Garis Solid</option>
                      <option value="double">Garis Ganda</option>
                      <option value="dashed">Garis Putus</option>
                    </select>
                  </div>
                </>
              )}

              {/* Lined styles and color themes & Watermark */}
              <div className="toolbar-group" title="Gaya & Tanda Air">
                <span style={{ fontSize: 11, color: '#666', padding: '0 4px' }}>Garis:</span>
                <select className="toolbar-select" value={paperStyle} onChange={e => setPaperStyle(e.target.value)}>
                  <option value="blank">Polos (Blank)</option>
                  <option value="lined">Bergaris (Lined)</option>
                  <option value="grid">Kotak-kotak (Grid)</option>
                  <option value="dotted">Titik-titik (Dotted)</option>
                </select>

                <span style={{ fontSize: 11, color: '#666', padding: '0 4px' }}>Warna:</span>
                <select className="toolbar-select" value={paperTheme} onChange={e => setPaperTheme(e.target.value)}>
                  <option value="white">Putih Bersih</option>
                  <option value="cream">Kuning Cream</option>
                  <option value="yellow">Yellow Pad (Kuning)</option>
                  <option value="kraft">Kraft Cokelat</option>
                  <option value="dark">Dark Slate</option>
                </select>

                <span style={{ fontSize: 11, color: '#666', padding: '0 4px' }}>Tanda Air:</span>
                <select className="toolbar-select" value={watermarkText} onChange={e => setWatermarkText(e.target.value)}>
                  <option value="">Tanpa Tanda Air</option>
                  <option value="DRAFT">DRAFT</option>
                  <option value="CONFIDENTIAL">RAHASIA (CONFIDENTIAL)</option>
                  <option value="LAPORAN KEUANGAN">LAPORAN KEUANGAN</option>
                  <option value="URGENT">URGENT</option>
                </select>
              </div>

              {/* Dynamic Header & Footer inputs */}
              {!isDraftMode && (
                <div className="toolbar-group" title="Kop & Kaki Halaman">
                  <input 
                    type="text" 
                    value={docxHeader} 
                    onChange={e => setDocxHeader(e.target.value)} 
                    placeholder="Header Teks..." 
                    className="toolbar-input" 
                    style={{ width: '100px', padding: '3px 6px', fontSize: '11px', border: '1px solid #cbd5e1', borderRadius: '4px' }}
                    title="Catatan Kop Atas (Header)"
                  />
                  <input 
                    type="text" 
                    value={docxFooter} 
                    onChange={e => setDocxFooter(e.target.value)} 
                    placeholder="Footer Teks..." 
                    className="toolbar-input" 
                    style={{ width: '100px', padding: '3px 6px', fontSize: '11px', border: '1px solid #cbd5e1', borderRadius: '4px' }}
                    title="Catatan Kaki Bawah (Footer)"
                  />
                  <label style={{ display: 'flex', alignItems: 'center', gap: '4px', fontSize: '11px', color: '#475569', cursor: 'pointer' }}>
                    <input type="checkbox" checked={showPageNumbers} onChange={e => setShowPageNumbers(e.target.checked)} style={{ cursor: 'pointer' }} />
                    No. Halaman
                  </label>
                </div>
              )}
            </>
          )}

          {activeRibbonTab === 'ai' && (
            <>
              {/* Deepernova AI magic assistant tools */}
              <div className="toolbar-group" title="Bantuan Teks AI">
                <button 
                  className="toolbar-btn" 
                  onClick={handleAiAutocomplete} 
                  title="Lanjutkan Tulisan dengan AI" 
                  style={{ background: '#fff7ed', color: '#c2410c', borderColor: '#fed7aa', border: '1px solid', fontWeight: 600 }}
                >
                  <i className="fas fa-magic" style={{ marginRight: 6 }}></i> Lanjutkan Kalimat
                </button>
                <button 
                  className="toolbar-btn ai-magic-btn" 
                  onClick={handleDocxAiFormat} 
                  title="Rapikan dokumen dengan AI"
                >
                  <i className="fas fa-magic"></i> Format Ulang AI
                </button>
              </div>

              {/* Financial template generator dropdown */}
              <div className="toolbar-group" title="Templat Keuangan Instan">
                <select 
                  onChange={e => { if(e.target.value) { insertFinancialTemplate(e.target.value); e.target.value = ''; } }} 
                  className="toolbar-select ai-template-select" 
                  style={{ background: '#f0fdf4', color: '#166534', borderColor: '#bbf7d0', border: '1px solid', fontWeight: 600 }}
                >
                  <option value="">📝 Templat Laporan Keuangan...</option>
                  <option value="labarugi">1. Laporan Laba Rugi (Income Statement)</option>
                  <option value="neraca">2. Laporan Neraca Keuangan (Balance Sheet)</option>
                  <option value="aruskas">3. Laporan Arus Kas (Cash Flow)</option>
                </select>
              </div>

              {/* Floating Chat */}
              <div className="toolbar-group" title="Diskusi Dokumen">
                <button 
                  className={`toolbar-btn ${showBrainstormChat ? 'active' : ''}`}
                  onClick={() => openBrainstormChat()}
                  style={{ fontWeight: 600 }}
                >
                  💡 Diskusi Melayang
                </button>
              </div>
            </>
          )}
        </div>

        {/* Page Setup Drawer Panel */}
        {showPageSetup && (
          <div className="page-setup-drawer">
            <div className="setup-control-group">
              <label>Tampilan</label>
              <div className="layout-toggle-row">
                <button className={`layout-toggle-btn ${!isDraftMode ? 'active' : ''}`} onClick={() => setIsDraftMode(false)}>Kertas</button>
                <button className={`layout-toggle-btn ${isDraftMode ? 'active' : ''}`} onClick={() => setIsDraftMode(true)}>Web</button>
              </div>
            </div>

            {!isDraftMode && (
              <>
                <div className="setup-control-group">
                  <label>Ukuran</label>
                  <select className="setup-select" value={paperSize} onChange={e => setPaperSize(e.target.value)}>
                    <option value="a4">A4 (21 x 29.7 cm)</option>
                    <option value="letter">Letter (21.6 x 27.9 cm)</option>
                    <option value="legal">Legal (21.6 x 35.6 cm)</option>
                    <option value="a5">A5 (14.8 x 21 cm)</option>
                    <option value="custom">Kustom...</option>
                  </select>
                </div>

                {paperSize === 'custom' && (
                  <>
                    <div className="setup-control-group">
                      <label>Lebar (cm)</label>
                      <input type="number" step="0.1" className="setup-input" value={customWidth} onChange={e => setCustomWidth(e.target.value)} style={{ width: '60px' }} />
                    </div>
                    <div className="setup-control-group">
                      <label>Tinggi (cm)</label>
                      <input type="number" step="0.1" className="setup-input" value={customHeight} onChange={e => setCustomHeight(e.target.value)} style={{ width: '60px' }} />
                    </div>
                  </>
                )}

                <div className="setup-control-group">
                  <label>Orientasi</label>
                  <select className="setup-select" value={paperOrientation} onChange={e => setPaperOrientation(e.target.value)}>
                    <option value="portrait">Tegak (Portrait)</option>
                    <option value="landscape">Mendatar (Landscape)</option>
                  </select>
                </div>

                <div className="setup-control-group">
                  <label>Margin</label>
                  <select className="setup-select" value={paperMargin} onChange={e => setPaperMargin(e.target.value)}>
                    <option value="normal">Normal (2.54 cm)</option>
                    <option value="narrow">Sempit (1.27 cm)</option>
                    <option value="wide">Lebar (5.08 cm)</option>
                    <option value="custom">Kustom...</option>
                  </select>
                </div>

                {paperMargin === 'custom' && (
                  <div className="setup-control-group">
                    <label>Margin (cm)</label>
                    <input type="number" step="0.1" className="setup-input" value={customMargin} onChange={e => setCustomMargin(e.target.value)} style={{ width: '60px' }} />
                  </div>
                )}
              </>
            )}

            <div className="setup-control-group">
              <label>Gaya</label>
              <select className="setup-select" value={paperStyle} onChange={e => setPaperStyle(e.target.value)}>
                <option value="blank">Polos (Blank)</option>
                <option value="lined">Bergaris (Lined)</option>
                <option value="grid">Kotak-kotak (Grid)</option>
                <option value="dotted">Titik-titik (Dotted)</option>
              </select>
            </div>

            <div className="setup-control-group">
              <label>Warna Kertas</label>
              <select className="setup-select" value={paperTheme} onChange={e => setPaperTheme(e.target.value)}>
                <option value="white">Putih Bersih</option>
                <option value="cream">Kuning Cream</option>
                <option value="yellow">Yellow Pad (Kuning)</option>
                <option value="kraft">Kraft Cokelat</option>
                <option value="dark">Dark Slate</option>
              </select>
            </div>

            {!isDraftMode && (
              <div className="setup-control-group">
                <label>Zoom</label>
                <select className="setup-select" value={pageZoom} onChange={e => setPageZoom(e.target.value)}>
                  <option value="fit">Fit Width (Mobile)</option>
                  <option value="100%">100%</option>
                  <option value="75%">75%</option>
                  <option value="50%">50%</option>
                </select>
              </div>
            )}
          </div>
        )}

        {/* Image Upload Modal */}
        {showInsertImage && (
          <div className="image-upload-overlay" onClick={() => setShowInsertImage(false)}>
            <div className="image-upload-modal" onClick={e => e.stopPropagation()}>
              <h4>Insert Image</h4>
              <input type="file" accept="image/*" ref={fileInputRef} onChange={handleImageUpload} />
              <button className="toolbar-btn" onClick={() => setShowInsertImage(false)}>Cancel</button>
            </div>
          </div>
        )}

        {/* Chart Modal */}
        {showChartModal && (
          <div className="image-upload-overlay" onClick={() => setShowChartModal(false)}>
            <div className="image-upload-modal" onClick={e => e.stopPropagation()} style={{ maxWidth: '500px' }}>
              <div className="modal-tabs" style={{ display: 'flex', borderBottom: '1px solid #e2e8f0', marginBottom: '14px' }}>
                <button 
                  className={`modal-tab-btn ${chartModalTab === 'standard' ? 'active' : ''}`}
                  onClick={() => setChartModalTab('standard')}
                  style={{ flex: 1, padding: '10px', background: 'transparent', border: 'none', borderBottom: chartModalTab === 'standard' ? '2px solid #ff6b00' : 'none', fontWeight: 600, color: chartModalTab === 'standard' ? '#ff6b00' : '#64748b', cursor: 'pointer' }}
                >
                  Grafik Data 📊
                </button>
                <button 
                  className={`modal-tab-btn ${chartModalTab === 'curve' ? 'active' : ''}`}
                  onClick={() => setChartModalTab('curve')}
                  style={{ flex: 1, padding: '10px', background: 'transparent', border: 'none', borderBottom: chartModalTab === 'curve' ? '2px solid #ff6b00' : 'none', fontWeight: 600, color: chartModalTab === 'curve' ? '#ff6b00' : '#64748b', cursor: 'pointer' }}
                >
                  Kurva Matematika 📈
                </button>
              </div>

              {chartModalTab === 'standard' ? (
                <>
                  <div style={{ marginBottom: '12px' }}>
                    <label>Jenis Grafik:</label>
                    <select value={chartType} onChange={e => setChartType(e.target.value)} style={{ width: '100%', padding: '6px', marginTop: '4px' }}>
                      <option value="bar">Bar Chart</option>
                      <option value="line">Line Chart</option>
                      <option value="pie">Pie Chart</option>
                    </select>
                  </div>
                  <div style={{ marginBottom: '12px' }}>
                    <label>Judul Grafik:</label>
                    <input type="text" value={chartTitle} onChange={e => setChartTitle(e.target.value)} style={{ width: '100%', padding: '6px', marginTop: '4px', boxSizing: 'border-box' }} />
                  </div>
                  <div style={{ marginBottom: '12px' }}>
                    <label>Data (format: Nama:Nilai, satu per baris):</label>
                    <textarea 
                      style={{ width: '100%', height: '100px', padding: '6px', marginTop: '4px', boxSizing: 'border-box', fontFamily: 'monospace', fontSize: '12px' }}
                      defaultValue="A:40\nB:30\nC:20\nD:50"
                      onChange={e => updateChartData(e.target.value)}
                    />
                  </div>
                  <div style={{ marginBottom: '12px', padding: '12px', border: '1px solid #ddd', borderRadius: '6px', backgroundColor: '#f9f9f9', height: '250px', overflow: 'hidden' }}>
                    {renderChart({ type: chartType, title: chartTitle, data: chartData })}
                  </div>
                  <div style={{ display: 'flex', gap: '8px' }}>
                    <button className="toolbar-btn" onClick={insertChart} style={{ flex: 1 }}>✓ Insert</button>
                    <button className="toolbar-btn" onClick={() => setShowChartModal(false)} style={{ flex: 1 }}>Cancel</button>
                  </div>
                </>
              ) : (
                <>
                  <div style={{ marginBottom: '12px' }}>
                    <label>Persamaan Kurva:</label>
                    <select value={curveEquation} onChange={e => setCurveEquation(e.target.value)} style={{ width: '100%', padding: '6px', marginTop: '4px' }}>
                      <option value="sine">Sinusoidal (Wave)</option>
                      <option value="exponential">Eksponensial (Growth)</option>
                      <option value="linear">Regresi Linear (Trend)</option>
                      <option value="bell">Kurva Gauss (Normal Distribution)</option>
                    </select>
                  </div>
                  <div style={{ marginBottom: '12px' }}>
                    <label>Amplitudo / Skala: {curveAmplitude}%</label>
                    <input type="range" min="10" max="100" value={curveAmplitude} onChange={e => setCurveAmplitude(Number(e.target.value))} style={{ width: '100%', marginTop: '4px' }} />
                  </div>
                  <div style={{ marginBottom: '12px' }}>
                    <label>Warna Kurva:</label>
                    <input type="color" value={curveColor} onChange={e => setCurveColor(e.target.value)} style={{ width: '100%', height: '36px', padding: '3px', marginTop: '4px', border: '1px solid #cbd5e1', borderRadius: '6px', cursor: 'pointer' }} />
                  </div>
                  <div style={{ display: 'flex', gap: '8px', marginTop: '16px' }}>
                    <button className="toolbar-btn" onClick={generateAndInsertCurve} style={{ flex: 1, background: '#ff6b00', color: '#fff' }}>✓ Insert Kurva</button>
                    <button className="toolbar-btn" onClick={downloadCurveImage} style={{ flex: 1 }}><i className="fas fa-download"></i> Download PNG</button>
                    <button className="toolbar-btn" onClick={() => setShowChartModal(false)} style={{ flex: 1 }}>Cancel</button>
                  </div>
                </>
              )}
            </div>
          </div>
        )}

        {/* Deepernova Cloud Explorer Modal */}
        {showCloudModal && (
          <div className="cloud-explorer-overlay" onClick={() => setShowCloudModal(false)}>
            <div className="cloud-explorer-modal" onClick={e => e.stopPropagation()}>
              <div className="explorer-header">
                <h3><i className="fas fa-cloud-upload-alt" style={{ color: '#106ebe', marginRight: 8 }}></i> Deepernova Cloud Explorer</h3>
                <button className="close-btn" onClick={() => setShowCloudModal(false)}><i className="fas fa-times"></i></button>
              </div>

              {/* Path and Actions Bar */}
              <div className="explorer-path-bar">
                <button 
                  className="path-back-btn" 
                  disabled={currentFolderId === null} 
                  onClick={() => {
                    const parent = cloudFiles.find(f => f.id === currentFolderId);
                    setCurrentFolderId(parent ? parent.parentId : null);
                  }}
                  title="Kembali ke folder sebelumnya"
                >
                  <i className="fas fa-arrow-left"></i>
                </button>
                <div className="path-breadcrumbs">
                  <span className="crumb-item" onClick={() => setCurrentFolderId(null)}>Root</span>
                  {(() => {
                    const crumbs = [];
                    let currId = currentFolderId;
                    while (currId) {
                      const folder = cloudFiles.find(f => f.id === currId);
                      if (folder) {
                        crumbs.unshift(folder);
                        currId = folder.parentId;
                      } else {
                        break;
                      }
                    }
                    return crumbs.map(c => (
                      <span key={c.id} className="crumb-item" onClick={() => setCurrentFolderId(c.id)}>
                        <i className="fas fa-chevron-right" style={{ fontSize: 9, margin: '0 6px', color: '#a19f9d' }}></i>
                        {c.name}
                      </span>
                    ));
                  })()}
                </div>
                <button 
                  className="new-folder-btn" 
                  onClick={() => {
                    const name = prompt('Masukkan nama folder baru:');
                    if (name) createCloudFolder(name);
                  }}
                >
                  <i className="fas fa-folder-plus"></i> Folder Baru
                </button>
              </div>

              {/* Main files grid view */}
              <div className="explorer-files-view">
                {(() => {
                  const filteredItems = cloudFiles.filter(item => item.parentId === currentFolderId);
                  if (filteredItems.length === 0) {
                    return (
                      <div className="empty-explorer">
                        <i className="fas fa-folder-open" style={{ fontSize: 48, color: '#e1dfdd', marginBottom: 12 }}></i>
                        <p>Folder ini kosong.</p>
                      </div>
                    );
                  }
                  return (
                    <div className="files-grid">
                      {filteredItems.map(item => {
                        const isFolder = item.type === 'folder';
                        const isSelected = selectedCloudFile?.id === item.id;
                        let iconClass = 'fas fa-folder';
                        let iconColor = '#ffb900'; // Folder Gold

                        if (item.type === 'docx') {
                          iconClass = 'fas fa-file-word';
                          iconColor = '#106ebe'; // Word Blue
                        } else if (item.type === 'excel') {
                          iconClass = 'fas fa-file-excel';
                          iconColor = '#107c41'; // Excel Green
                        } else if (item.type === 'pptx') {
                          iconClass = 'fas fa-file-powerpoint';
                          iconColor = '#d83b01'; // PPT Orange
                        }

                        return (
                          <div 
                            key={item.id} 
                            className={`grid-file-item ${isSelected ? 'selected' : ''}`}
                            onClick={() => {
                              setSelectedCloudFile(item);
                              if (explorerMode === 'save' && !isFolder) {
                                setCloudFileName(item.name.replace(`.${item.type}`, ''));
                              }
                            }}
                            onDoubleClick={() => {
                              if (isFolder) {
                                setCurrentFolderId(item.id);
                                setSelectedCloudFile(null);
                              } else {
                                loadCloudFile(item);
                              }
                            }}
                          >
                            <i className={`${iconClass} file-item-icon`} style={{ color: iconColor }}></i>
                            <span className="file-item-name" title={item.name}>{item.name}</span>
                          </div>
                        );
                      })}
                    </div>
                  );
                })()}
              </div>

              {/* Bottom Actions Bar */}
              <div className="explorer-footer">
                <div className="footer-left">
                  {explorerMode === 'save' ? (
                    <div style={{ display: 'flex', alignItems: 'center', gap: 8, width: '100%' }}>
                      <span style={{ fontSize: 13, color: '#323130' }}>Nama Berkas:</span>
                      <input 
                        type="text" 
                        value={cloudFileName} 
                        onChange={e => setCloudFileName(e.target.value)}
                        placeholder="Nama dokumen..."
                        className="explorer-name-input"
                      />
                    </div>
                  ) : (
                    <span style={{ fontSize: 13, color: '#605e5c' }}>
                      {selectedCloudFile ? `Terpilih: ${selectedCloudFile.name}` : 'Pilih berkas dari daftar'}
                    </span>
                  )}
                </div>
                <div className="footer-right">
                  {selectedCloudFile && (
                    <button className="explorer-btn delete-btn" onClick={() => deleteCloudFile(selectedCloudFile.id)} title="Hapus Terpilih">
                      <i className="fas fa-trash-alt"></i> Hapus
                    </button>
                  )}
                  {explorerMode === 'save' ? (
                    <button 
                      className="explorer-btn save-btn" 
                      onClick={() => saveActiveFileToCloud(cloudFileName)}
                      disabled={!cloudFileName.trim()}
                    >
                      <i className="fas fa-save" style={{ marginRight: 6 }}></i> Simpan ke Cloud
                    </button>
                  ) : (
                    <button 
                      className="explorer-btn open-btn" 
                      onClick={() => selectedCloudFile && loadCloudFile(selectedCloudFile)}
                      disabled={!selectedCloudFile || selectedCloudFile.type === 'folder'}
                    >
                      <i className="fas fa-folder-open" style={{ marginRight: 6 }}></i> Buka
                    </button>
                  )}
                  <button className="explorer-btn cancel-btn" onClick={() => setShowCloudModal(false)}>Batal</button>
                </div>
              </div>
            </div>
          </div>
        )}

        <div className={`docx-editor-viewport ${isDraftMode ? 'draft-mode' : ''}`}>
          <div 
            className="zoom-scale-wrapper" 
            style={{
              transform: `scale(${getZoomScale()})`,
              transformOrigin: 'top center',
              width: '100%',
              display: 'flex',
              justifyContent: 'center',
              flex: 1
            }}
          >
            <div className="a4-container">
              {uploadedFileName && uploadedFileType === 'docx' ? (
                /* ── Word Document Viewer Mode: multi-page read view ── */
                <div className="word-viewer-pages">
                  <div
                    className="word-viewer-page"
                    dangerouslySetInnerHTML={{ __html: content[0]?.text || '' }}
                  />
                </div>
              ) : (
                /* ── Editable Mode for new documents ── */
                <div
                  className={`a4-page style-${paperStyle} theme-${paperTheme} margin-${paperMargin} orientation-${paperOrientation} ${showFormattingMarks ? 'show-marks' : ''} ${watermarkText ? 'has-watermark' : ''}`}
                  contentEditable
                  suppressContentEditableWarning
                  ref={pageRef}
                  data-watermark={watermarkText}
                  data-placeholder="Mulai menulis dokumen Anda di sini..."
                  style={{
                    position: 'relative',
                    fontSize: '12pt',
                    fontFamily: fontFamily || 'Times New Roman, serif',
                    color: paperTheme === 'dark' ? '#f1f5f9' : '#1a1a1a',
                    lineHeight: paperStyle === 'lined' ? '28px' : (lineSpacing || 1.5),
                    textAlign: 'justify',
                    margin: isDraftMode ? '0' : '20px auto',
                    padding: getPageMargin(),
                    ...(!isDraftMode ? getPageDimensions() : {}),
                    boxSizing: 'border-box',
                    columnCount: !isDraftMode ? pageColumns : 1,
                    columnGap: '24px',
                    border: pageBorder === 'none' ? '1px solid #d2d0ce' : (pageBorder === 'double' ? '6px double #106ebe' : `3px ${pageBorder} #106ebe`)
                  }}
                  onInput={e => {
                    docxTextRef.current = e.currentTarget.innerHTML;
                  }}
                  onBlur={() => {
                    syncDocxContent();
                  }}
                >
                  {/* Visual Page Header in portrait/landscape modes */}
                  {!isDraftMode && docxHeader && (
                    <div 
                      className="page-header-display" 
                      style={{ position: 'absolute', top: '15px', left: getPageMargin(), right: getPageMargin(), fontSize: '10px', color: '#666', borderBottom: '1px solid #eee', paddingBottom: '3px', userSelect: 'none', pointerEvents: 'none' }}
                      contentEditable={false}
                    >
                      {docxHeader}
                    </div>
                  )}

                  {/* Visual Page Footer in portrait/landscape modes */}
                  {!isDraftMode && docxFooter && (
                    <div 
                      className="page-footer-display" 
                      style={{ position: 'absolute', bottom: '15px', left: getPageMargin(), right: getPageMargin(), fontSize: '10px', color: '#666', borderTop: '1px solid #eee', paddingTop: '3px', userSelect: 'none', pointerEvents: 'none', display: 'flex', justifyContent: 'space-between' }}
                      contentEditable={false}
                    >
                      <span>{docxFooter}</span>
                      {showPageNumbers && <span>Halaman 1</span>}
                    </div>
                  )}
                </div>
              )}
            </div>
          </div>
        </div>

        {/* Microsoft Word Style Bottom Blue Status Bar */}
        <div className="word-status-bar">
          <div className="status-left">
            <span>Halaman 1 dari 1</span>
            <span className="status-separator">|</span>
            <span>{wordCount} Kata</span>
            <span className="status-separator">|</span>
            <span>{charCount} Karakter</span>
            <span className="status-separator">|</span>
            <span>Bahasa Indonesia</span>
          </div>
          <div className="status-right">
            <i className="fas fa-search-plus" style={{ marginRight: 6 }}></i>
            <input 
              type="range" 
              min="50" 
              max="150" 
              step="10" 
              value={pageZoom === 'fit' ? 100 : parseInt(pageZoom)} 
              onChange={e => setPageZoom(`${e.target.value}%`)} 
              className="zoom-slider"
            />
            <span>{pageZoom === 'fit' ? 'Fit' : pageZoom}</span>
          </div>
        </div>
      </div>
    );
  };

  const renderPptxEditor = () => (
    <div className="pptx-editor">
      {Array.isArray(content) && content.map(slide => (
        <div key={slide.id} className="slide-container">
          <div className="slide">
            <div className="slide-title">{slide.title}</div>
            <div className="slide-content">{slide.content}</div>
          </div>
          <div className="slide-notes">Notes: {slide.notes}</div>
        </div>
      ))}
    </div>
  );

  const renderExcelEditor = () => {
    const sheet = excelSheets[activeSheet];
    if (!sheet) return null;
    const data = sheet.data;
    const rows = data.length;
    const cols = Math.max(...data.map(r => r.length), 1);
    const cell = selectedCell ? data[selectedCell.r]?.[selectedCell.c] : null;

    return (
      <div className="excel-editor">
        {/* Sheet Tabs */}
        <div className="excel-tabs">
          {excelSheets.map((s, idx) => (
            <div key={idx} className={`excel-tab ${idx === activeSheet ? 'active' : ''}`}>
              <span onClick={() => { setActiveSheet(idx); setSelectedCell(null); }}>{s.name}</span>
              {excelSheets.length > 1 && (
                <button className="excel-tab-close" onClick={() => deleteExcelSheet(idx)}>✕</button>
              )}
            </div>
          ))}
          <button className="excel-tab-add" onClick={addExcelSheet}>+</button>
        </div>

        {/* Format Bar */}
        <div className="excel-format-bar">
          <div className="excel-format-left">
            <span className="excel-cell-ref">
              {selectedCell ? `${getColumnLabel(selectedCell.c)}${selectedCell.r + 1}` : ''}
            </span>
            <input
              className="excel-cell-input"
              value={cell?.value ?? ''}
              onChange={e => { if (selectedCell) updateCell(selectedCell.r, selectedCell.c, e.target.value); }}
              placeholder="Value"
            />
          </div>
          <div className="excel-format-actions">
            <button className={`ef-btn ${cell?.format?.bold ? 'active' : ''}`} onClick={() => applyFormatToSelection({ bold: !cell?.format?.bold })} title="Bold"><b>B</b></button>
            <button className={`ef-btn ${cell?.format?.italic ? 'active' : ''}`} onClick={() => applyFormatToSelection({ italic: !cell?.format?.italic })} title="Italic"><i>I</i></button>
            <button className={`ef-btn ${cell?.format?.underline ? 'active' : ''}`} onClick={() => applyFormatToSelection({ underline: !cell?.format?.underline })} title="Underline"><u>U</u></button>
            <button className={`ef-btn ${cell?.format?.strikethrough ? 'active' : ''}`} onClick={() => applyFormatToSelection({ strikethrough: !cell?.format?.strikethrough })} title="Strikethrough"><s>S</s></button>
            <span className="ef-sep">|</span>
            <button className={`ef-btn ${cell?.format?.halign === 'left' ? 'active' : ''}`} onClick={() => applyFormatToSelection({ halign: 'left' })} title="Align Left">⬅</button>
            <button className={`ef-btn ${cell?.format?.halign === 'center' ? 'active' : ''}`} onClick={() => applyFormatToSelection({ halign: 'center' })} title="Center">↔</button>
            <button className={`ef-btn ${cell?.format?.halign === 'right' ? 'active' : ''}`} onClick={() => applyFormatToSelection({ halign: 'right' })} title="Align Right">➡</button>
            <span className="ef-sep">|</span>
            <button className={`ef-btn ${cell?.format?.wrapText ? 'active' : ''}`} onClick={() => applyFormatToSelection({ wrapText: !cell?.format?.wrapText })} title="Wrap Text">↕</button>
            <label className="ef-label" title="Font Color">
              <span style={{ color: cell?.format?.fontColor || '#000', fontWeight: 'bold' }}>A</span>
              <input type="color" value={cell?.format?.fontColor || '#000000'} onChange={e => applyFormatToSelection({ fontColor: e.target.value })} className="ef-color" />
            </label>
            <label className="ef-label" title="Fill Color">
              <span style={{ background: cell?.format?.fillColor || '#fff', border: '1px solid #ccc', display: 'inline-block', width: 14, height: 14, borderRadius: 2 }}></span>
              <input type="color" value={cell?.format?.fillColor || '#ffffff'} onChange={e => applyFormatToSelection({ fillColor: e.target.value })} className="ef-color" />
            </label>
            <span className="ef-sep">|</span>
            <select className="ef-select" value={cell?.format?.fontSize || 11} onChange={e => applyFormatToSelection({ fontSize: Number(e.target.value) })}>
              {[8,9,10,11,12,14,16,18,20,24,28,32,36,48].map(s => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>
        </div>

        {/* Find/Replace Bar */}
        {showFind && (
          <div className="excel-find-bar">
            <input className="ef-input" value={findText} onChange={e => setFindText(e.target.value)} placeholder="Find..." />
            <input className="ef-input" value={replaceText} onChange={e => setReplaceText(e.target.value)} placeholder="Replace with..." />
            <button className="ef-btn-sm" onClick={findAndReplace}>Find & Replace</button>
            <button className="ef-btn-sm" onClick={() => setShowFind(false)}>✕</button>
          </div>
        )}

        {/* Row/Column Toolbar */}
        <div className="excel-toolbar">
          <button className="excel-toolbar-btn" onClick={addExcelRow}>+ Row</button>
          <button className="excel-toolbar-btn" onClick={addExcelColumn}>+ Col</button>
          {selectedCell && (
            <>
              <button className="excel-toolbar-btn" onClick={() => insertExcelRowAbove(selectedCell.r)}>Ins Row</button>
              <button className="excel-toolbar-btn" onClick={() => insertExcelColumnLeft(selectedCell.c)}>Ins Col</button>
              <button className="excel-toolbar-btn" onClick={() => deleteExcelRow(selectedCell.r)}>Del Row</button>
              <button className="excel-toolbar-btn" onClick={() => deleteExcelColumn(selectedCell.c)}>Del Col</button>
            </>
          )}
          <button className="excel-toolbar-btn" onClick={() => { if (selectedCell) sortExcelData(selectedCell.c, 'asc'); }}>Sort ↑</button>
          <button className="excel-toolbar-btn" onClick={() => { if (selectedCell) sortExcelData(selectedCell.c, 'desc'); }}>Sort ↓</button>
          <button className="excel-toolbar-btn" onClick={() => setShowFind(!showFind)}>Find</button>
          <span className="excel-dimension">{rows}R × {cols}C</span>
        </div>

        {/* Spreadsheet Grid */}
        <div className="excel-scroll">
          <table className="spreadsheet">
            <thead>
              <tr>
                <th className="excel-corner"></th>
                {Array.from({ length: cols }, (_, ci) => (
                  <th key={ci} className="excel-col-header" onClick={() => sortExcelData(ci)}>
                    {getColumnLabel(ci)}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {data.map((row, ri) => (
                <tr key={ri}>
                  <td className="excel-row-header">{ri + 1}</td>
                  {Array.from({ length: cols }, (_, ci) => {
                    const cell = row[ci] || createCell('');
                    return (
                      <td
                        key={ci}
                        contentEditable
                        suppressContentEditableWarning
                        onClick={() => setSelectedCell({ r: ri, c: ci })}
                        onBlur={e => updateCell(ri, ci, e.currentTarget.innerText)}
                        className="excel-cell"
                        style={getCellStyle(cell, ri, ci)}
                      >
                        {cell.value}
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    );
  };

  // ===== EXPORT =====
  const handleExport = async () => {
    try {
      if (editorType === 'docx') await exportDocx();
      else if (editorType === 'pptx') await generatePptxViaServer();
      else if (editorType === 'excel') await exportExcel();
    } catch (error) {
      alert('Export error: ' + error.message);
    }
  };

  // ===== CLOUD FILE EXPLORER METHODS =====
  const fetchCloudFiles = async () => {
    try {
      const response = await fetch('/api/cloud/files');
      const data = await response.json();
      if (data.success) {
        setCloudFiles(data.files || []);
      }
    } catch (error) {
      console.warn('Gagal mengambil daftar file cloud:', error);
    }
  };

  const createCloudFolder = async (name) => {
    if (!name.trim()) return;
    try {
      const response = await fetch('/api/cloud/folders', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          parentId: currentFolderId,
          name: name.trim()
        })
      });
      const data = await response.json();
      if (data.success) {
        fetchCloudFiles();
      } else {
        alert(`Gagal membuat folder: ${data.error}`);
      }
    } catch (error) {
      console.error('[Cloud Folder Create Error]:', error);
      alert('Gagal membuat folder di server.');
    }
  };

  const saveActiveFileToCloud = async (fileName) => {
    if (!fileName.trim()) {
      alert('Nama dokumen tidak boleh kosong.');
      return;
    }
    try {
      let fileContent = null;
      if (editorType === 'excel') {
        fileContent = { excelSheets, activeSheet };
      } else {
        // For docx/pptx
        if (editorType === 'docx' && pageRef.current) {
          // Sync react states
          docxTextRef.current = pageRef.current.innerHTML;
          const updated = [{ id: Date.now(), type: 'html', text: pageRef.current.innerHTML }];
          setContent(updated);
          fileContent = updated;
        } else {
          fileContent = content;
        }
      }

      const formattedName = fileName.endsWith(`.${editorType}`) ? fileName : `${fileName}.${editorType}`;

      const response = await fetch('/api/cloud/save', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          id: activeCloudFileId || null,
          parentId: currentFolderId,
          name: formattedName,
          type: editorType,
          content: fileContent
        })
      });
      const data = await response.json();
      if (data.success && data.file) {
        setDocumentTitle(data.file.name.replace(`.${editorType}`, ''));
        setActiveCloudFileId(data.file.id);
        setShowCloudModal(false);
        fetchCloudFiles();
        alert(`Dokumen "${data.file.name}" berhasil disimpan ke cloud.`);
      } else {
        alert(`Gagal menyimpan ke cloud: ${data.error}`);
      }
    } catch (error) {
      console.error('[Cloud Save Error]:', error);
      alert('Gagal menyimpan file ke server.');
    }
  };

  const loadCloudFile = async (file) => {
    try {
      const response = await fetch(`/api/cloud/files/${file.id}`);
      const data = await response.json();
      if (data.success && data.file) {
        setDocumentTitle(data.file.name.replace(`.${data.file.type}`, ''));
        setEditorType(data.file.type);
        setActiveCloudFileId(data.file.id);

        if (data.file.type === 'excel') {
          const parsed = data.file.content;
          setExcelSheets(parsed?.excelSheets || []);
          setActiveSheet(parsed?.activeSheet || 0);
          setContent([]);
        } else {
          // docx / pptx
          const parsed = data.file.content;
          setContent(parsed || []);
          
          if (data.file.type === 'docx' && pageRef.current) {
            const html = parsed[0]?.text || '';
            pageRef.current.innerHTML = html;
            docxTextRef.current = html;
          }
        }
        setShowCloudModal(false);
        alert(`Dokumen "${data.file.name}" berhasil dimuat dari cloud.`);
      } else {
        alert(`Gagal memuat dokumen: ${data.error}`);
      }
    } catch (error) {
      console.error('[Cloud Load Error]:', error);
      alert('Gagal memuat file dari server.');
    }
  };

  const deleteCloudFile = async (id) => {
    if (!confirm('Apakah Anda yakin ingin menghapus folder/dokumen ini? Semua sub-file di dalamnya juga akan terhapus.')) return;
    try {
      const response = await fetch(`/api/cloud/files/${id}`, { method: 'DELETE' });
      const data = await response.json();
      if (data.success) {
        if (activeCloudFileId === id) {
          setActiveCloudFileId(null);
        }
        setSelectedCloudFile(null);
        fetchCloudFiles();
      } else {
        alert(`Gagal menghapus: ${data.error}`);
      }
    } catch (error) {
      console.error('[Cloud Delete Error]:', error);
      alert('Gagal menghapus file di server.');
    }
  };

  // Fetch cloud files on startup
  useEffect(() => {
    fetchCloudFiles();
  }, []);

  const handleAutoSave = async () => {
    try {
      const key = `doc_${documentTitle}_${editorType}`;
      const saveData = editorType === 'excel'
        ? { title: documentTitle, type: editorType, excelSheets, activeSheet }
        : { title: documentTitle, type: editorType, content };
      localStorage.setItem(key, JSON.stringify({
        ...saveData,
        metadata: { savedAt: new Date().toISOString() }
      }));
    } catch (error) {
      console.warn('Auto-save failed:', error);
    }
  };

  const downloadFile = (blob, fileName) => {
    const el = document.createElement('a');
    el.href = URL.createObjectURL(blob);
    el.download = fileName;
    el.style.display = 'none';
    document.body.appendChild(el);
    el.click();
    document.body.removeChild(el);
    URL.revokeObjectURL(el.href);
  };

  // ===== PARSE TABLES FROM AI RESPONSE =====
  const parseTablesFromText = (text) => {
    const tableRegex = /\[TABLE\]([\s\S]*?)\[\/TABLE\]/g;
    const matches = [];
    let match;
    while ((match = tableRegex.exec(text)) !== null) {
      matches.push(match[1]);
    }
    
    return matches.map(tableStr => {
      const lines = tableStr.trim().split('\n').filter(line => line.trim());
      if (lines.length === 0) return null;
      
      const rows = lines.map(line => {
        const cells = line.split('|').map(cell => cell.trim()).filter(c => c);
        return {
          id: Date.now() + Math.random(),
          cells: cells.map((text, ci) => ({
            id: Date.now() + ci,
            text: text,
            rowspan: 1,
            colspan: 1,
            bold: false,
            italic: false,
            align: 'left',
            bgColor: ''
          }))
        };
      });
      
      if (rows.length === 0) return null;
      
      return {
        id: Date.now(),
        rows: rows
      };
    }).filter(t => t !== null);
  };

  const insertParsedTables = (tables) => {
    if (!Array.isArray(tables) || tables.length === 0) return;
    setDocxTables(prev => [...prev, ...tables]);
  };

  const removeTableMarkersFromText = (text) => {
    return text.replace(/\[TABLE\]([\s\S]*?)\[\/TABLE\]/g, '').trim();
  };

  // ===== EXPORT DOCX - International Standard Format =====
  const exportDocx = async () => {
    const sections = [];
    const htmlString = content[0]?.text || '';
    
    const parser = new DOMParser();
    const parsedHtmlDoc = parser.parseFromString(htmlString, 'text/html');
    const elements = Array.from(parsedHtmlDoc.body.childNodes);
    
    elements.forEach(node => {
      if (node.nodeType === Node.TEXT_NODE) {
        const text = node.textContent?.trim();
        if (text) {
          sections.push(
            new Paragraph({
              children: [new TextRun({ text: text, font: 'Times New Roman', size: 24 })],
              spacing: { line: 360, lineRule: 'auto', after: 240 },
              indent: { firstLine: 720 },
              alignment: AlignmentType.JUSTIFIED,
            })
          );
        }
      } else if (node.nodeType === Node.ELEMENT_NODE) {
        const tagName = node.tagName.toLowerCase();
        
        if (tagName.startsWith('h')) {
          const text = node.innerText || node.textContent || '';
          sections.push(
            new Paragraph({
              children: [new TextRun({ text: text, font: 'Times New Roman', size: 28, bold: true })],
              spacing: { before: 240, after: 120 },
            })
          );
        } else if (tagName === 'p') {
          const text = node.innerText || node.textContent || '';
          if (text.trim()) {
            sections.push(
              new Paragraph({
                children: [new TextRun({ text: text, font: 'Times New Roman', size: 24 })],
                spacing: { line: 360, lineRule: 'auto', after: 240 },
                indent: { firstLine: 720 },
                alignment: AlignmentType.JUSTIFIED,
              })
            );
          }
        } else if (tagName === 'ul' || tagName === 'ol') {
          const lis = node.querySelectorAll('li');
          lis.forEach((li, idx) => {
            const prefix = tagName === 'ol' ? `${idx + 1}. ` : '• ';
            const text = li.innerText || li.textContent || '';
            sections.push(
              new Paragraph({
                children: [new TextRun({ text: prefix + text, font: 'Times New Roman', size: 24 })],
                spacing: { line: 360, lineRule: 'auto', after: 120 },
                indent: { left: 720, hanging: 360 },
              })
            );
          });
        } else if (tagName === 'table') {
          const trs = node.querySelectorAll('tr');
          const rows = [];
          
          trs.forEach(tr => {
            const tds = tr.querySelectorAll('td, th');
            const cells = [];
            
            tds.forEach(td => {
              cells.push(
                new TableCell({
                  children: [new Paragraph({
                    children: [new TextRun({
                      text: td.innerText || td.textContent || '',
                      font: 'Times New Roman',
                      size: 22,
                    })],
                  })],
                  verticalAlign: VerticalAlign.CENTER,
                })
              );
            });
            
            if (cells.length > 0) {
              rows.push(new TableRow({ children: cells }));
            }
          });
          
          if (rows.length > 0) {
            sections.push(
              new Table({
                rows: rows,
                width: { size: 100, type: 'pct' },
              })
            );
            sections.push(new Paragraph({ text: '' }));
          }
        } else if (node.classList.contains('docx-chart-block')) {
          const title = node.getAttribute('data-chart-title') || 'Chart';
          const type = node.getAttribute('data-chart-type') || 'bar';
          sections.push(
            new Paragraph({
              children: [new TextRun({ text: `[Grafik: ${title} (${type.toUpperCase()})]`, font: 'Times New Roman', size: 24, italic: true })],
              spacing: { after: 240 },
            })
          );
        } else if (tagName === 'img' || (node.querySelector && node.querySelector('img'))) {
          const imgEl = tagName === 'img' ? node : node.querySelector('img');
          const src = imgEl.getAttribute('src') || '';
          if (src.startsWith('data:image/')) {
            const base64Data = src.split(',')[1];
            try {
              const binaryString = atob(base64Data);
              const len = binaryString.length;
              const bytes = new Uint8Array(len);
              for (let i = 0; i < len; i++) {
                bytes[i] = binaryString.charCodeAt(i);
              }
              sections.push(
                new Paragraph({
                  children: [
                    new ImageRun({
                      data: bytes,
                      transformation: {
                        width: 450,
                        height: 300,
                      },
                    })
                  ],
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 240 },
                })
              );
            } catch (err) {
              console.warn('Failed to parse embedded image base64:', err);
            }
          }
        }
      }
    });
    
    if (sections.length === 0) sections.push(new Paragraph({ text: '' }));
    
    const doc = new Document({
      sections: [{
        properties: {
          page: {
            margins: {
              top: 1440,
              right: 1440,
              bottom: 1440,
              left: 1440,
              header: 720,
              footer: 720,
              gutter: 0,
            },
            size: {
              width: 12240,
              height: 15840,
            },
          }
        },
        children: sections
      }]
    });
    const blob = await Packer.toBlob(doc);
    downloadFile(
      new Blob([blob], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }),
      `${documentTitle || 'document'}.docx`
    );
  };

  const exportPptx = async () => {
    const pptx = new PptxGenJS();
    content.forEach(slide => {
      if (slide.type === 'slide') {
        const s = pptx.addSlide();
        s.addText(slide.title || 'Slide', { x: 0.5, y: 0.5, fontSize: 26, color: '363636', bold: true });
        s.addText(slide.content || '', { x: 0.5, y: 1.4, fontSize: 16, color: '555555', wrap: true, w: '90%' });
      }
    });
    await pptx.writeFile({ fileName: `${documentTitle || 'presentation'}.pptx` });
  };

  const exportExcel = () => {
    const wb = XLSX.utils.book_new();
    excelSheets.forEach((sheet, idx) => {
      const plainData = sheet.data.map(row => row.map(cell => cell.value));
      const ws = XLSX.utils.aoa_to_sheet(plainData);
      XLSX.utils.book_append_sheet(wb, ws, sheet.name || `Sheet${idx + 1}`);
    });
    downloadFile(
      new Blob([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })], { type: 'application/octet-stream' }),
      `${documentTitle || 'spreadsheet'}.xlsx`
    );
  };

  // ===== PPT GENERATION VIA SERVER (with security) =====
  const generatePptxViaServer = async () => {
    try {
      if (!Array.isArray(content) || content.length === 0) {
        alert('Tambahkan slide sebelum generate');
        return;
      }

      setIsPptGenerating(true);
      setPptGenerationStatus('Mempersiapkan data...');

      // Extract slides data for server
      const slides = content
        .filter(slide => slide.type === 'slide')
        .map(slide => ({
          title: slide.title || 'Slide',
          content: slide.content || ''
        }));

      if (slides.length === 0) {
        alert('Minimal 1 slide dengan konten');
        setIsPptGenerating(false);
        return;
      }

      setPptGenerationStatus(`Mengirim ${slides.length} slide ke server...`);

      const requestPayload = {
        title: documentTitle || 'Untitled Presentation',
        subtitle: 'Dibuat oleh Deepernova',
        template: pptTemplate,
        slides: slides
      };

      console.log('[PPT_EDITOR] Sending to /api/generate-ppt:', {
        title: requestPayload.title,
        slide_count: slides.length,
        bytes: JSON.stringify(requestPayload).length
      });

      const response = await fetch('/api/generate-ppt', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(requestPayload)
      });

      const result = await response.json();

      if (!response.ok || !result.success) {
        const errorMsg = result.error || result.data?.error || 'Gagal generate PPT';
        setPptGenerationStatus(`❌ Error: ${errorMsg}`);
        alert(`Error: ${errorMsg}`);
        setIsPptGenerating(false);
        return;
      }

      setPptGenerationStatus(`✅ Berhasil! Mengunduh ${result.slides_count} slides...`);

      // Track generated file
      setGeneratedPptFiles(prev => [...prev, {
        filename: result.filename,
        url: result.downloadUrl,
        slides: result.slides_count,
        size: result.size_mb,
        timestamp: new Date().toLocaleString()
      }]);
      
      setShowPptResults(true);

      // Download file from server
      const downloadUrl = result.downloadUrl;
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = result.filename || `${documentTitle || 'presentation'}.pptx`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

      console.log('[PPT_EDITOR] ✅ Download complete:', result.filename);
      
      setTimeout(() => {
        setPptGenerationStatus('');
        setIsPptGenerating(false);
      }, 2000);

    } catch (error) {
      console.error('[PPT_EDITOR] Error:', error);
      const errMsg = error.message || 'Kesalahan saat generate';
      setPptGenerationStatus(`❌ ${errMsg}`);
      alert(`Error: ${errMsg}`);
      setIsPptGenerating(false);
    }
  };

  // ===== PPT FILE UPLOAD =====
  const handlePptUpload = (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    
    if (!file.name.endsWith('.pptx')) {
      alert('Hanya file .pptx yang didukung');
      return;
    }
    
    const reader = new FileReader();
    reader.onload = (e) => {
      setUploadedPptFile({
        name: file.name,
        size: (file.size / 1024 / 1024).toFixed(2),
        type: 'uploaded',
        timestamp: new Date().toLocaleString(),
        data: e.target?.result
      });
      setShowPptResults(true);
    };
    reader.readAsArrayBuffer(file);
  };

  // ===== PPT SLIDE PREVIEW =====
  const extractSlidesFromPptx = async (pptxData) => {
    try {
      const JSZip = (await import('jszip')).default;
      const zip = new JSZip();
      await zip.loadAsync(pptxData);

      // Get slide content
      const slides = [];
      let slideNum = 1;
      
      while (true) {
        const slidePath = `ppt/slides/slide${slideNum}.xml`;
        const slideFile = zip.file(slidePath);
        if (!slideFile) break;

        const slideXml = await slideFile.async('text');
        const parser = new DOMParser();
        const slideDoc = parser.parseFromString(slideXml, 'text/xml');

        const paragraphs = Array.from(slideDoc.querySelectorAll('a\\:p'))
          .map(p => Array.from(p.querySelectorAll('a\\:t, t'))
            .map(el => el.textContent?.trim())
            .filter(Boolean)
            .join(' ')
          )
          .filter(Boolean);

        const title = paragraphs.length > 0 ? paragraphs[0] : `Slide ${slideNum}`;
        const body = paragraphs.slice(1).map(line => line.replace(/^[-•]\s*/, '').trim());
        const lines = body.length > 0 ? body : paragraphs.length > 1 ? paragraphs.slice(1) : [];
        const content = [title, ...lines].join('\n');

        slides.push({
          number: slideNum,
          title,
          lines,
          content,
        });

        slideNum += 1;
      }

      return slides.length > 0 ? slides : [{ number: 1, content: 'Empty presentation' }];
    } catch (error) {
      console.error('Error parsing PPTX:', error);
      return [{ number: 1, content: 'Gagal parse file' }];
    }
  };

  const handlePptPreview = async (file) => {
    if (file.type === 'uploaded' && file.data) {
      // Preview uploaded file
      const slides = await extractSlidesFromPptx(file.data);
      setPreviewSlides(slides);
      setPreviewPptFile(file);
      setCurrentSlideIdx(0);
    } else if (file.url) {
      // Preview generated file - fetch from server
      try {
        const response = await fetch(file.url);
        const arrayBuffer = await response.arrayBuffer();
        const slides = await extractSlidesFromPptx(arrayBuffer);
        setPreviewSlides(slides);
        setPreviewPptFile(file);
        setCurrentSlideIdx(0);
      } catch (error) {
        console.error('Error fetching PPT:', error);
        alert('Gagal load preview');
      }
    }
  };

  // ===== FORMATTING =====
  const applyFormatting = (command, value = null) => {
    document.execCommand(command, false, value);
    if (pageRef.current) {
      docxTextRef.current = pageRef.current.innerHTML;
      syncDocxContent();
      pageRef.current.focus();
    }
  };

  const formatText = (style) => {
    const map = {
      bold: 'bold', italic: 'italic', underline: 'underline', strikethrough: 'strikeThrough',
      left: 'justifyLeft', center: 'justifyCenter', right: 'justifyRight', justify: 'justifyFull',
      bullet: 'insertUnorderedList', number: 'insertOrderedList', undo: 'undo', redo: 'redo',
      indent: 'indent', outdent: 'outdent',
    };
    applyFormatting(map[style] || style);
  };

  const handleFontSize = (size) => { setFontSize(size); applyFormatting('fontSize', parseInt(size)); };
  const handleFontFamily = (font) => { setFontFamily(font); applyFormatting('fontName', font); };
  const handleTextColor = (color) => { setTextColor(color); applyFormatting('foreColor', color); };
  const handleHighlightColor = (color) => { setHighlightColor(color); applyFormatting('hiliteColor', color); };
  const handleLineSpacing = (spacing) => {
    setLineSpacing(spacing);
    if (pageRef.current) {
      pageRef.current.style.lineHeight = spacing;
    }
  };
  const handleClearFormatting = () => applyFormatting('removeFormat');

  const changeTextCase = (caseType) => {
    const selection = window.getSelection();
    if (!selection || !selection.rangeCount) return;
    const text = selection.toString();
    if (!text) return;
    let newText = '';
    if (caseType === 'upper') newText = text.toUpperCase();
    else if (caseType === 'lower') newText = text.toLowerCase();
    else if (caseType === 'capitalize') {
      newText = text.replace(/\b\w/g, c => c.toUpperCase());
    } else if (caseType === 'sentence') {
      newText = text.charAt(0).toUpperCase() + text.slice(1).toLowerCase();
    }
    document.execCommand('insertHTML', false, newText);
  };

  const promptLink = () => {
    const url = prompt('Masukkan URL tautan (Contoh: https://example.com):');
    if (url) {
      applyFormatting('createLink', url);
    }
  };

  const insertPageBreak = () => {
    const pbHtml = '<hr class="page-break" style="border: none; border-top: 2px dashed #106ebe; margin: 20px 0; text-align: center; color: #106ebe; font-size: 11px; user-select: none;" contenteditable="false" data-label="--- Batas Halaman (Page Break) ---" /><p>&nbsp;</p>';
    if (pageRef.current) {
      pageRef.current.focus();
      document.execCommand('insertHTML', false, pbHtml);
      docxTextRef.current = pageRef.current.innerHTML;
      syncDocxContent();
    }
  };

  const insertFinancialTemplate = (type) => {
    let templateHtml = '';
    if (type === 'labarugi') {
      templateHtml = `
        <h2 style="text-align: center; color: #106ebe; font-family: 'Times New Roman', serif; margin-bottom: 2px; font-weight: bold;">LAPORAN LABA RUGI</h2>
        <h4 style="text-align: center; color: #555; font-family: 'Times New Roman', serif; margin-top: 0; margin-bottom: 20px; font-style: italic;">Periode: Bulanan</h4>
        <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px; font-family: 'Times New Roman', serif; border: 1px solid #d2d0ce;">
          <thead>
            <tr style="background-color: #f3f2f1; border-bottom: 2px solid #106ebe;">
              <th style="padding: 10px; text-align: left; font-weight: bold; border: 1px solid #d2d0ce;">Keterangan</th>
              <th style="padding: 10px; text-align: right; font-weight: bold; width: 180px; border: 1px solid #d2d0ce;">Jumlah (Rupiah)</th>
            </tr>
          </thead>
          <tbody>
            <tr style="border-bottom: 1px solid #e1dfdd; font-weight: bold; background-color: #faf9f8;">
              <td style="padding: 8px; border: 1px solid #d2d0ce;">1. PENDAPATAN</td>
              <td style="padding: 8px; text-align: right; border: 1px solid #d2d0ce;"></td>
            </tr>
            <tr style="border-bottom: 1px solid #edebe9;">
              <td style="padding: 8px; padding-left: 20px; border: 1px solid #d2d0ce;">Pendapatan Penjualan</td>
              <td style="padding: 8px; text-align: right; color: #107c41; border: 1px solid #d2d0ce;">Rp 50.000.000</td>
            </tr>
            <tr style="border-bottom: 1px solid #e1dfdd; font-weight: bold; background-color: #faf9f8;">
              <td style="padding: 8px; border: 1px solid #d2d0ce;">2. BEBAN (PENGELUARAN)</td>
              <td style="padding: 8px; text-align: right; border: 1px solid #d2d0ce;"></td>
            </tr>
            <tr style="border-bottom: 1px solid #edebe9;">
              <td style="padding: 8px; padding-left: 20px; border: 1px solid #d2d0ce;">Harga Pokok Penjualan (HPP)</td>
              <td style="padding: 8px; text-align: right; color: #a80000; border: 1px solid #d2d0ce;">Rp 20.000.000</td>
            </tr>
            <tr style="border-bottom: 1px solid #edebe9;">
              <td style="padding: 8px; padding-left: 20px; border: 1px solid #d2d0ce;">Beban Operasional & Gaji</td>
              <td style="padding: 8px; text-align: right; color: #a80000; border: 1px solid #d2d0ce;">Rp 10.000.000</td>
            </tr>
            <tr style="border-bottom: 1px solid #edebe9;">
              <td style="padding: 8px; padding-left: 20px; border: 1px solid #d2d0ce;">Beban Pemasaran & Iklan</td>
              <td style="padding: 8px; text-align: right; color: #a80000; border: 1px solid #d2d0ce;">Rp 5.000.000</td>
            </tr>
            <tr style="border-bottom: 1px solid #e1dfdd; font-weight: bold; background-color: #f3f2f1;">
              <td style="padding: 8px; border: 1px solid #d2d0ce;">TOTAL BEBAN</td>
              <td style="padding: 8px; text-align: right; color: #a80000; border: 1px solid #d2d0ce;">Rp 35.000.000</td>
            </tr>
            <tr style="border-bottom: 2px double #106ebe; font-weight: bold; background-color: #edebe9; font-size: 15px;">
              <td style="padding: 10px; color: #106ebe; border: 1px solid #d2d0ce;">LABA BERSIH (PENDAPATAN - BEBAN)</td>
              <td style="padding: 10px; text-align: right; color: #107c41; border: 1px solid #d2d0ce;">Rp 15.000.000</td>
            </tr>
          </tbody>
        </table>
        <p>&nbsp;</p>
      `;
    } else if (type === 'neraca') {
      templateHtml = `
        <h2 style="text-align: center; color: #106ebe; font-family: 'Times New Roman', serif; margin-bottom: 2px; font-weight: bold;">LAPORAN NERACA KEUANGAN</h2>
        <h4 style="text-align: center; color: #555; font-family: 'Times New Roman', serif; margin-top: 0; margin-bottom: 20px; font-style: italic;">Per Tanggal: Akhir Bulan</h4>
        <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px; font-family: 'Times New Roman', serif; border: 1px solid #d2d0ce;">
          <thead>
            <tr style="background-color: #f3f2f1; border-bottom: 2px solid #106ebe;">
              <th style="padding: 10px; text-align: left; font-weight: bold; border: 1px solid #d2d0ce;">Aktiva (Aset)</th>
              <th style="padding: 10px; text-align: right; font-weight: bold; width: 140px; border: 1px solid #d2d0ce;">Nilai</th>
              <th style="padding: 10px; text-align: left; font-weight: bold; padding-left: 20px; border: 1px solid #d2d0ce;">Pasiva (Kewajiban & Modal)</th>
              <th style="padding: 10px; text-align: right; font-weight: bold; width: 140px; border: 1px solid #d2d0ce;">Nilai</th>
            </tr>
          </thead>
          <tbody>
            <tr style="border-bottom: 1px solid #edebe9; font-weight: bold; background-color: #faf9f8;">
              <td style="padding: 8px; border: 1px solid #d2d0ce;">ASET LANCAR</td>
              <td style="padding: 8px; border: 1px solid #d2d0ce;"></td>
              <td style="padding: 8px; padding-left: 20px; border: 1px solid #d2d0ce;">KEWAJIBAN (UTANG)</td>
              <td style="padding: 8px; border: 1px solid #d2d0ce;"></td>
            </tr>
            <tr style="border-bottom: 1px solid #edebe9;">
              <td style="padding: 6px; padding-left: 15px; border: 1px solid #d2d0ce;">Kas & Setara Kas</td>
              <td style="padding: 6px; text-align: right; border: 1px solid #d2d0ce;">Rp 30.000.000</td>
              <td style="padding: 6px; padding-left: 30px; border: 1px solid #d2d0ce;">Utang Dagang</td>
              <td style="padding: 6px; text-align: right; color: #a80000; border: 1px solid #d2d0ce;">Rp 5.000.000</td>
            </tr>
            <tr style="border-bottom: 1px solid #edebe9;">
              <td style="padding: 6px; padding-left: 15px; border: 1px solid #d2d0ce;">Persediaan Barang</td>
              <td style="padding: 6px; text-align: right; border: 1px solid #d2d0ce;">Rp 15.000.000</td>
              <td style="padding: 6px; padding-left: 30px; border: 1px solid #d2d0ce;">Utang Bank</td>
              <td style="padding: 6px; text-align: right; color: #a80000; border: 1px solid #d2d0ce;">Rp 10.000.000</td>
            </tr>
            <tr style="border-bottom: 1px solid #edebe9; font-weight: bold; background-color: #faf9f8;">
              <td style="padding: 8px; border: 1px solid #d2d0ce;">ASET TETAP</td>
              <td style="padding: 8px; border: 1px solid #d2d0ce;"></td>
              <td style="padding: 8px; padding-left: 20px; border: 1px solid #d2d0ce;">MODAL (EKUITAS)</td>
              <td style="padding: 8px; border: 1px solid #d2d0ce;"></td>
            </tr>
            <tr style="border-bottom: 1px solid #edebe9;">
              <td style="padding: 6px; padding-left: 15px; border: 1px solid #d2d0ce;">Peralatan Toko</td>
              <td style="padding: 6px; text-align: right; border: 1px solid #d2d0ce;">Rp 10.000.000</td>
              <td style="padding: 6px; padding-left: 30px; border: 1px solid #d2d0ce;">Modal Pemilik</td>
              <td style="padding: 6px; text-align: right; border: 1px solid #d2d0ce;">Rp 40.000.000</td>
            </tr>
            <tr style="border-bottom: 1px solid #edebe9;">
              <td style="padding: 6px; padding-left: 15px; border: 1px solid #d2d0ce;">Akumulasi Penyusutan</td>
              <td style="padding: 6px; text-align: right; color: #a80000; border: 1px solid #d2d0ce;">-Rp 2.000.000</td>
              <td style="padding: 6px; padding-left: 30px; border: 1px solid #d2d0ce;">Laba Ditahan</td>
              <td style="padding: 6px; text-align: right; color: #107c41; border: 1px solid #d2d0ce;">Rp 12.000.000</td>
            </tr>
            <tr style="border-bottom: 2px double #106ebe; font-weight: bold; background-color: #edebe9;">
              <td style="padding: 10px; color: #106ebe; border: 1px solid #d2d0ce;">TOTAL AKTIVA</td>
              <td style="padding: 10px; text-align: right; color: #106ebe; border: 1px solid #d2d0ce;">Rp 53.000.000</td>
              <td style="padding: 10px; padding-left: 20px; color: #106ebe; border: 1px solid #d2d0ce;">TOTAL PASIVA</td>
              <td style="padding: 10px; text-align: right; color: #106ebe; border: 1px solid #d2d0ce;">Rp 53.000.000</td>
            </tr>
          </tbody>
        </table>
        <p style="font-size: 11px; font-style: italic; color: #666; text-align: center;">* Laporan neraca seimbang (Balanced).</p>
        <p>&nbsp;</p>
      `;
    } else if (type === 'aruskas') {
      templateHtml = `
        <h2 style="text-align: center; color: #106ebe; font-family: 'Times New Roman', serif; margin-bottom: 2px; font-weight: bold;">LAPORAN ARUS KAS</h2>
        <h4 style="text-align: center; color: #555; font-family: 'Times New Roman', serif; margin-top: 0; margin-bottom: 20px; font-style: italic;">Periode: Bulanan</h4>
        <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px; font-family: 'Times New Roman', serif; border: 1px solid #d2d0ce;">
          <thead>
            <tr style="background-color: #f3f2f1; border-bottom: 2px solid #106ebe;">
              <th style="padding: 10px; text-align: left; font-weight: bold; border: 1px solid #d2d0ce;">Aktivitas Aliran Kas</th>
              <th style="padding: 10px; text-align: right; font-weight: bold; width: 180px; border: 1px solid #d2d0ce;">Nilai (Rupiah)</th>
            </tr>
          </thead>
          <tbody>
            <tr style="border-bottom: 1px solid #edebe9; font-weight: bold; background-color: #faf9f8;">
              <td style="padding: 8px; border: 1px solid #d2d0ce;">1. ARUS KAS DARI AKTIVITAS OPERASIONAL</td>
              <td style="padding: 8px; border: 1px solid #d2d0ce;"></td>
            </tr>
            <tr style="border-bottom: 1px solid #edebe9;">
              <td style="padding: 6px; padding-left: 20px; border: 1px solid #d2d0ce;">Penerimaan Kas dari Pelanggan</td>
              <td style="padding: 6px; text-align: right; color: #107c41; border: 1px solid #d2d0ce;">Rp 50.000.000</td>
            </tr>
            <tr style="border-bottom: 1px solid #edebe9;">
              <td style="padding: 6px; padding-left: 20px; border: 1px solid #d2d0ce;">Pembayaran Kas untuk Operasional & Supplier</td>
              <td style="padding: 6px; text-align: right; color: #a80000; border: 1px solid #d2d0ce;">-Rp 30.000.000</td>
            </tr>
            <tr style="border-bottom: 1px solid #edebe9; font-weight: bold; background-color: #faf9f8;">
              <td style="padding: 8px; border: 1px solid #d2d0ce;">2. ARUS KAS DARI AKTIVITAS INVESTASI</td>
              <td style="padding: 8px; border: 1px solid #d2d0ce;"></td>
            </tr>
            <tr style="border-bottom: 1px solid #edebe9;">
              <td style="padding: 6px; padding-left: 20px; border: 1px solid #d2d0ce;">Pembelian Peralatan Komputer Baru</td>
              <td style="padding: 6px; text-align: right; color: #a80000; border: 1px solid #d2d0ce;">-Rp 5.000.000</td>
            </tr>
            <tr style="border-bottom: 1px solid #edebe9; font-weight: bold; background-color: #faf9f8;">
              <td style="padding: 8px; border: 1px solid #d2d0ce;">3. ARUS KAS DARI AKTIVITAS PENDANAAN</td>
              <td style="padding: 8px; border: 1px solid #d2d0ce;"></td>
            </tr>
            <tr style="border-bottom: 1px solid #edebe9;">
              <td style="padding: 6px; padding-left: 20px; border: 1px solid #d2d0ce;">Penerimaan Modal Pemilik Baru</td>
              <td style="padding: 6px; text-align: right; color: #107c41; border: 1px solid #d2d0ce;">Rp 10.000.000</td>
            </tr>
            <tr style="border-bottom: 1px solid #edebe9; font-weight: bold; background-color: #f3f2f1;">
              <td style="padding: 8px; border: 1px solid #d2d0ce;">KENAIKAN BERSIH KAS (OP + INV + PEND)</td>
              <td style="padding: 8px; text-align: right; color: #107c41; font-weight: bold; border: 1px solid #d2d0ce;">Rp 25.000.000</td>
            </tr>
            <tr style="border-bottom: 1px solid #edebe9;">
              <td style="padding: 8px; padding-left: 20px; font-weight: bold; border: 1px solid #d2d0ce;">Saldo Awal Kas Bulanan</td>
              <td style="padding: 8px; text-align: right; font-weight: bold; border: 1px solid #d2d0ce;">Rp 10.000.000</td>
            </tr>
            <tr style="border-bottom: 2px double #106ebe; font-weight: bold; background-color: #edebe9; font-size: 15px;">
              <td style="padding: 10px; color: #106ebe; border: 1px solid #d2d0ce;">SALDO AKHIR KAS BULANAN</td>
              <td style="padding: 10px; text-align: right; color: #107c41; border: 1px solid #d2d0ce;">Rp 35.000.000</td>
            </tr>
          </tbody>
        </table>
        <p>&nbsp;</p>
      `;
    }
    
    if (pageRef.current && templateHtml) {
      pageRef.current.focus();
      if (pageRef.current.innerHTML.trim() === '' || pageRef.current.innerHTML === '<p>Mulai menulis dokumen Anda di sini...</p>' || pageRef.current.innerText.trim() === '' || pageRef.current.innerText.trim() === 'Mulai menulis dokumen Anda di sini...') {
        pageRef.current.innerHTML = templateHtml;
      } else {
        document.execCommand('insertHTML', false, templateHtml);
      }
      docxTextRef.current = pageRef.current.innerHTML;
      syncDocxContent();
    }
  };

  // ===== RENDER =====
  return (
    <div className="document-editor-container">
      {/* Minimal Header */}
      <div className="doc-editor-header">
        <div className="header-left">
          <button className="back-to-chat-btn" onClick={() => onNavigate?.('chat')} title="Kembali">
            <i className="fas fa-chevron-left"></i>
          </button>
          <input type="text" value={documentTitle} onChange={e => setDocumentTitle(e.target.value)}
            className="doc-title-input" placeholder="Untitled" />
        </div>
        <div className="header-right">
          <div className="editor-type-selector">
            <button className={`type-btn ${editorType === 'docx' ? 'active' : ''}`} onClick={() => setEditorType('docx')}>DOCX</button>
            <button className={`type-btn ${editorType === 'pptx' ? 'active' : ''}`} onClick={() => setEditorType('pptx')}>PPTX</button>
            <button className={`type-btn ${editorType === 'excel' ? 'active' : ''}`} onClick={() => setEditorType('excel')}>XLSX</button>
          </div>
          {pptGenerationStatus && (
            <span className={`ppt-status ${isPptGenerating ? 'generating' : 'complete'}`}>
              {isPptGenerating ? '⏳ ' : ''}{pptGenerationStatus}
            </span>
          )}
          {editorType === 'pptx' && (
            <>
              <button className="artifact-btn" onClick={() => setShowPptResults(!showPptResults)} title="PPT Results">
                📊 {generatedPptFiles.length + (uploadedPptFile ? 1 : 0)}
              </button>
              <input 
                ref={pptUploadRef} 
                type="file" 
                accept=".pptx" 
                style={{ display: 'none' }} 
                onChange={handlePptUpload}
              />
              <button className="artifact-btn" onClick={() => pptUploadRef.current?.click()} title="Upload PPT">
                📂
              </button>
            </>
          )}
          <button 
            className={`artifact-btn ${showBrainstormChat ? 'active' : ''}`} 
            onClick={() => openBrainstormChat()} 
            title="Brainstorm Dokumen (Diskusi Melayang)"
          >
            💡
          </button>
          <button className="artifact-btn" onClick={() => setShowArtifacts(!showArtifacts)} title="Artifacts">
            {showArtifacts ? '✕' : '+'}
          </button>
          <button className="export-btn" onClick={handleExport} disabled={isPptGenerating} title="Export">
            {isPptGenerating ? '⏳' : '⬇'}
          </button>
        </div>
      </div>

      <div className="editor-layout">
        {editorType === 'pptx' ? (
          <div className="pptx-layout">
            <div className="pptx-editor-section">
              <div className="ppt-template-bar">
                <span className="ppt-template-label">Template:</span>
                <div className="ppt-template-selector">
                  {['classic', 'modern', 'bold', 'minimal'].map(option => (
                    <button
                      key={option}
                      className={`ppt-template-btn ${pptTemplate === option ? 'active' : ''}`}
                      onClick={() => setPptTemplate(option)}
                      type="button"
                    >
                      {option === 'classic' ? 'Classic' : option === 'modern' ? 'Modern' : option === 'bold' ? 'Bold' : 'Minimal'}
                    </button>
                  ))}
                </div>
              </div>
              {renderEditor()}
            </div>
            <div className="pptx-chat-section">
              {showAiPanel && (
                <div className={`ai-panel ${isAiMinimized ? 'minimized' : ''}`} ref={aiPanelRef}>
                  <div className="ai-header">
                    <span className="ai-title">💬 Chat</span>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                      <button 
                        className="minimize-btn" 
                        onClick={() => setIsAiMinimized(!isAiMinimized)}
                        style={{
                          background: 'transparent', border: 'none', fontSize: '13px', cursor: 'pointer', color: '#64748b'
                        }}
                        title={isAiMinimized ? "Expand" : "Minimize"}
                      >
                        {isAiMinimized ? <i className="fas fa-chevron-up"></i> : <i className="fas fa-chevron-down"></i>}
                      </button>
                      <button className="close-btn" onClick={() => setShowAiPanel(false)}>✕</button>
                    </div>
                  </div>
                  <div className="ai-content">
                    {generationProgress && <div className="generation-status">{generationProgress}</div>}
                    {aiError && <div className="error-message">{aiError}</div>}

                    <div className="ai-input-section">
                      <div className="textarea-wrapper">
                        <textarea
                          value={aiPrompt}
                          onChange={e => {
                            setAiPrompt(e.target.value);
                            const ta = e.target;
                            ta.style.height = 'auto';
                            ta.style.height = Math.min(ta.scrollHeight, 120) + 'px';
                          }}
                          placeholder={messages.length === 0 ? "Chat with Deepernova..." : "Reply to Deepernova..."}
                          disabled={isGenerating}
                          className="message-input"
                          rows={1}
                        />
                      </div>
                      <button
                        className={`action-button ${isGenerating ? 'stop-mode' : 'send-mode'}`}
                        onClick={isGenerating ? handleStopStreaming : handleAiWrite}
                        disabled={!isGenerating && !aiPrompt.trim()}
                      >
                        {isGenerating ? 'Stop' : 'Send'}
                      </button>
                    </div>

                    {isStreaming && streamingContent && (
                      <div className="ai-response-compact streaming">
                        <div className="response-label">
                          <span>Streaming</span>
                          <span className="streaming-dot">●</span>
                        </div>
                        <div className="ai-response-content">
                          {streamingContent.split('\n').slice(0, 10).map((line, idx) => (
                            <p key={idx} className="response-line">{line || '\u00A0'}</p>
                          ))}
                          {streamingContent.split('\n').length > 10 && <p className="response-more">...</p>}
                          <p className="streaming-cursor">▌</p>
                        </div>
                      </div>
                    )}

                    {!isStreaming && aiResponse && (
                      <div className="ai-response-compact">
                        <div className="response-label">Preview</div>
                        <div className="ai-response-content">
                          {aiResponse.split('\n').slice(0, 8).map((line, idx) => (
                            <p key={idx} className="response-line">{line}</p>
                          ))}
                          {aiResponse.split('\n').length > 8 && <p className="response-more">...</p>}
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              )}
            </div>
            {previewPptFile && previewSlides.length > 0 && (
              <div className="pptx-preview-section">
                <div className="ppt-preview-inline">
                  <div className="ppt-preview-header">
                    <div className="ppt-preview-title">
                      <span>{previewPptFile.name || previewPptFile.filename}</span>
                      <span className="ppt-slide-counter">
                        {currentSlideIdx + 1} / {previewSlides.length}
                      </span>
                    </div>
                    <button className="ppt-preview-close" onClick={() => { setPreviewPptFile(null); setPreviewSlides([]); }}>✕</button>
                  </div>
                  <div className="ppt-preview-content">
                    <div className="ppt-slide-display">
                      <div className="ppt-slide-number">
                        SLIDE {previewSlides[currentSlideIdx]?.number}
                      </div>
                      <div className="ppt-slide-content-card">
                        <div className="ppt-slide-title">
                          {previewSlides[currentSlideIdx]?.title}
                        </div>
                        {previewSlides[currentSlideIdx]?.lines?.length > 0 ? (
                          <ul className="ppt-slide-bullets">
                            {previewSlides[currentSlideIdx].lines.map((line, idx) => (
                              <li key={idx}>{line}</li>
                            ))}
                          </ul>
                        ) : (
                          <div className="ppt-slide-text">
                            {previewSlides[currentSlideIdx]?.content}
                          </div>
                        )}
                      </div>
                    </div>
                  </div>

                  <div className="ppt-preview-controls">
                    <button 
                      className="ppt-nav-btn" 
                      onClick={() => setCurrentSlideIdx(Math.max(0, currentSlideIdx - 1))}
                      disabled={currentSlideIdx === 0}
                    >
                      ← Sebelumnya
                    </button>
                    
                    <div className="ppt-slide-dots">
                      {previewSlides.map((_, idx) => (
                        <button
                          key={idx}
                          className={`ppt-dot ${idx === currentSlideIdx ? 'active' : ''}`}
                          onClick={() => setCurrentSlideIdx(idx)}
                        >
                          {idx + 1}
                        </button>
                      ))}
                    </div>

                    <button 
                      className="ppt-nav-btn" 
                      onClick={() => setCurrentSlideIdx(Math.min(previewSlides.length - 1, currentSlideIdx + 1))}
                      disabled={currentSlideIdx === previewSlides.length - 1}
                    >
                      Selanjutnya →
                    </button>
                  </div>
                </div>
              </div>
            )}
          </div>
        ) : (
          <>
        <div className="editor-center">
          <div className="editor-main">
            {renderEditor()}
          </div>
        </div>

        {/* Artifacts Panel */}
        {showArtifacts && (
          <div className="artifacts-panel">
            <div className="artifacts-header">
              <span className="artifacts-title">Artifacts ({artifacts.length})</span>
              <div className="artifacts-controls">
                <label className="auto-regen-toggle" title="Auto-regenerate on edit">
                  <input type="checkbox" checked={autoRegenerate} onChange={e => setAutoRegenerate(e.target.checked)} />
                  <span>Auto</span>
                </label>
                <button className="close-btn" onClick={() => setShowArtifacts(false)}>✕</button>
              </div>
            </div>
            <div className="artifacts-list">
              {artifacts.length === 0 && (
                <div className="artifacts-empty">
                  <p>No artifacts yet.</p>
                  <p className="artifacts-hint">Generate content with AI to create artifacts.</p>
                </div>
              )}
              {artifacts.map(art => (
                <div key={art.id} className={`artifact-item ${selectedArtifact?.id === art.id ? 'active' : ''}`}
                  onClick={() => loadArtifact(art)}>
                  <div className="artifact-icon">
                    {art.type === 'docx' ? '📄' : art.type === 'pptx' ? '📊' : '📈'}
                  </div>
                  <div className="artifact-info">
                    <div className="artifact-prompt">{art.prompt?.slice(0, 60)}{art.prompt?.length > 60 ? '...' : ''}</div>
                    <div className="artifact-meta">
                      <span className="artifact-type">{art.type?.toUpperCase()}</span>
                      <span className="artifact-time">{new Date(art.timestamp).toLocaleTimeString()}</span>
                    </div>
                  </div>
                  <button className="artifact-delete" onClick={e => { e.stopPropagation(); deleteArtifact(art.id); }}>🗑</button>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* AI Panel */}
        {showAiPanel && (
          <div className={`ai-panel ${isAiMinimized ? 'minimized' : ''}`} ref={aiPanelRef}>
            {/* ── HEADER ── */}
            <div className="ai-header">
              <span className="ai-title"><i className="fas fa-magic" style={{marginRight:5}}></i>AI Agent</span>
              <div className="ai-header-controls">
                <button 
                  className="minimize-btn" 
                  onClick={() => setIsAiMinimized(!isAiMinimized)}
                  style={{
                    background: 'transparent', border: 'none', fontSize: '13px', cursor: 'pointer', color: '#64748b', marginRight: '6px'
                  }}
                  title={isAiMinimized ? "Expand" : "Minimize"}
                >
                  {isAiMinimized ? <i className="fas fa-chevron-up"></i> : <i className="fas fa-chevron-down"></i>}
                </button>
                <label className="auto-regen-toggle" title="Auto-regenerate on edit">
                  <input type="checkbox" checked={autoRegenerate} onChange={e => setAutoRegenerate(e.target.checked)} />
                  <span>Auto</span>
                </label>
                <button className="close-btn" onClick={() => setShowAiPanel(false)}>✕</button>
              </div>
            </div>

            {/* ── MODE TABS ── */}
            <div className="ai-mode-tabs" style={{ display: 'flex', borderBottom: '1px solid #cbd5e1', padding: '8px', gap: '8px', background: '#fafafa' }}>
              <button 
                className={`ai-tab-btn ${aiMode === 'chat' ? 'active' : ''}`} 
                onClick={() => setAiMode('chat')}
                style={{
                  flex: 1, padding: '8px', fontSize: '12px', border: 'none', borderRadius: '8px',
                  background: aiMode === 'chat' ? 'rgba(255,107,0,0.12)' : 'transparent',
                  color: aiMode === 'chat' ? 'var(--orange)' : '#64748b', fontWeight: 600, cursor: 'pointer'
                }}
              >
                💬 Tanya AI
              </button>
              <button 
                className={`ai-tab-btn ${aiMode === 'drafting_agent' ? 'active' : ''}`} 
                onClick={() => setAiMode('drafting_agent')}
                style={{
                  flex: 1, padding: '8px', fontSize: '12px', border: 'none', borderRadius: '8px',
                  background: aiMode === 'drafting_agent' ? 'rgba(255,107,0,0.12)' : 'transparent',
                  color: aiMode === 'drafting_agent' ? 'var(--orange)' : '#64748b', fontWeight: 600, cursor: 'pointer'
                }}
              >
                🤖 Drafting Agent
              </button>
            </div>
 
            {/* ── STICKY TOP: Input + File Upload ── */}
            <div className="ai-panel-top">
              {generationProgress && <div className="generation-status">{generationProgress}</div>}
              {aiError && <div className="error-message">{aiError}</div>}
              {agentStatusText && <div className="generation-status" style={{ background: '#eff6ff', color: '#1e40af', borderLeftColor: '#3b82f6', fontSize: '11px', marginTop: 4 }}>{agentStatusText}</div>}

              {aiMode === 'chat' ? (
                <>
                  {/* Chat input */}
                  <div className="ai-input-section">
                    <div className="textarea-wrapper">
                      <textarea
                        value={aiPrompt}
                        onChange={e => {
                          setAiPrompt(e.target.value);
                          const ta = e.target;
                          ta.style.height = 'auto';
                          ta.style.height = Math.min(ta.scrollHeight, 120) + 'px';
                        }}
                        placeholder={messages.length === 0 ? "Tanya Deepernova..." : "Balas Deepernova..."}
                        disabled={isGenerating}
                        className="message-input"
                        rows={1}
                        onKeyDown={e => { if (e.key === 'Enter' && !e.shiftKey && !isGenerating) { e.preventDefault(); handleAiWrite(); } }}
                      />
                    </div>
                    <button
                      className={`action-button ${isGenerating ? 'stop-mode' : 'send-mode'}`}
                      onClick={isGenerating ? handleStopStreaming : handleAiWrite}
                      disabled={!isGenerating && !aiPrompt.trim()}
                    >
                      {isGenerating ? <i className="fas fa-stop"></i> : <i className="fas fa-paper-plane"></i>}
                    </button>
                  </div>
                </>
              ) : (
                /* ── DRAFTING AGENT INTERFACE ── */
                <div style={{ padding: '8px' }}>
                  {isAgentRunning ? (
                    <div className="agent-dashboard" style={{ margin: 0, padding: '10px', borderRadius: '12px' }}>
                      <div className="agent-dashboard-header">
                        <div className="agent-title">
                          <span className="agent-pulse-icon"></span>
                          <span>Agent Penyusun Aktif</span>
                        </div>
                        <button className="agent-cancel-btn" onClick={cancelAgent}>Hentikan</button>
                      </div>
                      
                      <div className="agent-progress-bar-container">
                        <div 
                          className="agent-progress-bar" 
                          style={{ width: `${(currentAgentStep / agentOutline.length) * 100}%` }}
                        ></div>
                      </div>

                      <div className="agent-checklist-list">
                        {agentChecklist.map((item, idx) => (
                          <div key={idx} className={`agent-checklist-item ${item.status}`}>
                            <span className="checklist-icon">
                              {item.status === 'done' ? '✅' : item.status === 'generating' ? '⏳' : '⚪'}
                            </span>
                            <span>{item.title}</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  ) : (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                      <textarea
                        value={aiPrompt}
                        onChange={e => setAiPrompt(e.target.value)}
                        placeholder="Tulis topik atau instruksi makalah/laporan bisnis yang ingin disusun secara bab-demi-bab..."
                        className="message-input"
                        style={{ minHeight: '60px', width: '100%', boxSizing: 'border-box' }}
                        rows={2}
                      />
                      <button
                        className="action-button send-mode"
                        onClick={() => handleStartAgentDrafting(aiPrompt)}
                        disabled={!aiPrompt.trim()}
                        style={{ width: '100%', padding: '10px' }}
                      >
                        <i className="fas fa-magic"></i> Mulai Susun Dokumen
                      </button>
                    </div>
                  )}
                </div>
              )}

              {/* File Uploader Zone */}
              <div className="file-uploader-zone">
                <label className="file-upload-label" title="Unggah dokumen (PDF, Word, TXT, Excel)">
                  {isUploadingFile ? (
                    <span><i className="fas fa-spinner fa-spin"></i> Membaca file...</span>
                  ) : uploadedFileName ? (
                    <span className="file-uploaded-badge">
                      <i className="fas fa-file-alt"></i>
                      <span className="file-name-truncate">{uploadedFileName}</span>
                      <button className="remove-file-btn" onClick={(e) => { e.stopPropagation(); e.preventDefault(); setUploadedFileName(''); setUploadedFileText(''); }} title="Hapus File">✕</button>
                    </span>
                  ) : (
                    <span><i className="fas fa-cloud-upload-alt"></i> Upload File (PDF, Word, Excel…)</span>
                  )}
                  <input type="file" onChange={handleDocUploader} style={{ display: 'none' }} accept=".pdf,.docx,.txt,.csv,.xlsx,.json" disabled={isUploadingFile} />
                </label>
              </div>
            </div>

            {/* ── SCROLLABLE BODY: Streaming + Brainstorm + Presets + Recs ── */}
            <div className="ai-panel-body">

              {/* Streaming response */}
              {isStreaming && streamingContent && (
                <div className="ai-response-compact streaming">
                  <div className="response-label">
                    <span>Streaming</span><span className="streaming-dot">●</span>
                  </div>
                  <div className="ai-response-content">
                    {streamingContent.split('\n').slice(0, 6).map((line, i) => <p key={i} className="response-line">{line}</p>)}
                    {streamingContent.split('\n').length > 6 && <p className="response-more">...</p>}
                  </div>
                </div>
              )}

              {!isStreaming && aiResponse && (
                <div className="ai-response-compact">
                  <div className="response-label">Preview</div>
                  <div className="ai-response-content">
                    {aiResponse.split('\n').slice(0, 8).map((line, idx) => (
                      <p key={idx} className="response-line">{line}</p>
                    ))}
                    {aiResponse.split('\n').length > 8 && <p className="response-more">...</p>}
                  </div>
                </div>
              )}

              {/* Brainstorm (only when file uploaded) */}
              {uploadedFileName && (
                <div className="brainstorm-section">
                  <span className="section-title"><i className="fas fa-brain"></i> Brainstorm File</span>
                  <button
                    className="brainstorm-open-btn"
                    onClick={() => openBrainstormChat()}
                  >
                    <i className="fas fa-comments"></i>
                    <div>
                      <div className="bso-title">💬 Mulai Chat Brainstorm</div>
                      <div className="bso-sub">Tanya jawab AI tentang "{uploadedFileName}"</div>
                    </div>
                    <i className="fas fa-chevron-right"></i>
                  </button>
                  <div className="brainstorm-quick-btns">
                    <button className="bq-btn" onClick={() => openBrainstormChat(`Berikan ringkasan eksekutif dan poin analisis utama dari file "${uploadedFileName}" secara mendalam.`)}>
                      💡 Ringkasan
                    </button>
                    <button className="bq-btn" onClick={() => openBrainstormChat(`Temukan kelemahan dan hal yang bisa diperbaiki dari file "${uploadedFileName}", berikan solusinya.`)}>
                      🔍 Kritik
                    </button>
                    <button className="bq-btn" onClick={() => openBrainstormChat(`Berikan 5 pertanyaan diskusi kritis yang bisa dikembangkan dari isi file "${uploadedFileName}".`)}>
                      ❓ Pertanyaan
                    </button>
                    <button className="bq-btn" onClick={() => openBrainstormChat(`Berikan ide-ide pengembangan dan inovasi berdasarkan file "${uploadedFileName}".`)}>
                      🚀 Ide Baru
                    </button>
                  </div>
                </div>
              )}

              {/* Preset Agentic Commands */}
              <div className="agent-presets-section">
                <span className="section-title"><i className="fas fa-robot"></i> Agentic Presets</span>
                <div className="agent-presets-grid">
                  <button className="preset-card" onClick={() => triggerPreset("Buatkan kerangka outline dokumen laporan bisnis yang sangat lengkap dan terstruktur.", "Outline")}>
                    <i className="fas fa-sitemap"></i> Outline
                  </button>
                  <button className="preset-card" onClick={() => triggerPreset("Perbaiki tata bahasa, ejaan, tanda baca, serta rapikan struktur kalimat pada dokumen ini.")}>
                    <i className="fas fa-spell-check"></i> Grammar
                  </button>
                  <button className="preset-card" onClick={() => triggerPreset("Tulis ringkasan eksekutif (executive summary) yang profesional berdasarkan dokumen ini.")}>
                    <i className="fas fa-file-alt"></i> Summary
                  </button>
                  <button className="preset-card" onClick={() => triggerPreset("Ubah gaya bahasa tulisan dalam dokumen ini menjadi sangat formal, profesional, dan persuasif.")}>
                    <i className="fas fa-user-tie"></i> Formal
                  </button>
                </div>
              </div>

              {/* Dynamic Smart Recommendations */}
              <div className="agent-recs-section">
                <span className="section-title"><i className="fas fa-lightbulb"></i> Smart Recommendations</span>
                <div className="agent-recs-list">
                  {getSmartRecommendations().map((rec, idx) => (
                    <div key={idx} className="rec-item" onClick={() => triggerPreset(rec.prompt, rec.title)}>
                      <div className="rec-icon-wrapper">
                        <i className={rec.icon}></i>
                      </div>
                      <div className="rec-text">
                        <div className="rec-title">{rec.title}</div>
                        <div className="rec-desc">{rec.prompt.slice(0, 55)}...</div>
                      </div>
                      <i className="fas fa-chevron-right rec-arrow"></i>
                    </div>
                  ))}
                </div>
              </div>

            </div>
          </div>
        )}

        {/* PPT Results Panel */}
        {editorType === 'pptx' && (
          <div className="ppt-results-panel">
            <div className="ppt-results-header">
              <span className="ppt-results-title">📊 PPT Files ({generatedPptFiles.length + (uploadedPptFile ? 1 : 0)})</span>
              <button className="close-btn" onClick={() => setShowPptResults(false)}>✕</button>
            </div>
            <div className="ppt-results-list">
              {/* Generated Files */}
              {generatedPptFiles.map((file, idx) => (
                <div key={`gen-${idx}`} className="ppt-file-item">
                  <div className="ppt-file-icon">📄</div>
                  <div className="ppt-file-info">
                    <div className="ppt-file-name">{file.filename}</div>
                    <div className="ppt-file-meta">
                      <span>📊 {file.slides} slides</span>
                      <span>💾 {file.size}MB</span>
                      <span>🕐 {file.timestamp}</span>
                    </div>
                  </div>
                  <div className="ppt-file-actions">
                    <a href={file.url} download={file.filename} className="ppt-download-btn" title="Download">⬇</a>
                    <button className="ppt-view-btn" onClick={() => handlePptPreview(file)} title="View">👁</button>
                  </div>
                </div>
              ))}
              
              {/* Uploaded File */}
              {uploadedPptFile && (
                <div className="ppt-file-item uploaded">
                  <div className="ppt-file-icon">📤</div>
                  <div className="ppt-file-info">
                    <div className="ppt-file-name">{uploadedPptFile.name}</div>
                    <div className="ppt-file-meta">
                      <span>📌 Uploaded</span>
                      <span>💾 {uploadedPptFile.size}MB</span>
                      <span>🕐 {uploadedPptFile.timestamp}</span>
                    </div>
                  </div>
                  <div className="ppt-file-actions">
                    <button className="ppt-view-btn" onClick={() => handlePptPreview(uploadedPptFile)} title="View">👁</button>
                  </div>
                </div>
              )}

              {generatedPptFiles.length === 0 && !uploadedPptFile && (
                <div className="ppt-empty-state">
                  <p>Belum ada file PPT</p>
                  <p className="ppt-empty-hint">Generate atau upload file PPT untuk melihatnya di sini</p>
                </div>
              )}
            </div>

            {previewPptFile && previewSlides.length > 0 && (
              <div className="ppt-preview-inline">
                <div className="ppt-preview-header">
                  <div className="ppt-preview-title">
                    <span>{previewPptFile.name || previewPptFile.filename}</span>
                    <span className="ppt-slide-counter">
                      {currentSlideIdx + 1} / {previewSlides.length}
                    </span>
                  </div>
                  <button className="ppt-preview-close" onClick={() => { setPreviewPptFile(null); setPreviewSlides([]); }}>✕</button>
                </div>
                <div className="ppt-preview-content">
                  <div className="ppt-slide-display">
                    <div className="ppt-slide-number">
                      SLIDE {previewSlides[currentSlideIdx]?.number}
                    </div>
                    <div className="ppt-slide-content-card">
                      <div className="ppt-slide-title">
                        {previewSlides[currentSlideIdx]?.title}
                      </div>
                      {previewSlides[currentSlideIdx]?.lines?.length > 0 ? (
                        <ul className="ppt-slide-bullets">
                          {previewSlides[currentSlideIdx].lines.map((line, idx) => (
                            <li key={idx}>{line}</li>
                          ))}
                        </ul>
                      ) : (
                        <div className="ppt-slide-text">
                          {previewSlides[currentSlideIdx]?.content}
                        </div>
                      )}
                    </div>
                  </div>
                </div>

                <div className="ppt-preview-controls">
                  <button 
                    className="ppt-nav-btn" 
                    onClick={() => setCurrentSlideIdx(Math.max(0, currentSlideIdx - 1))}
                    disabled={currentSlideIdx === 0}
                  >
                    ← Sebelumnya
                  </button>
                  
                  <div className="ppt-slide-dots">
                    {previewSlides.map((_, idx) => (
                      <button
                        key={idx}
                        className={`ppt-dot ${idx === currentSlideIdx ? 'active' : ''}`}
                        onClick={() => setCurrentSlideIdx(idx)}
                      >
                        {idx + 1}
                      </button>
                    ))}
                  </div>

                  <button 
                    className="ppt-nav-btn" 
                    onClick={() => setCurrentSlideIdx(Math.min(previewSlides.length - 1, currentSlideIdx + 1))}
                    disabled={currentSlideIdx === previewSlides.length - 1}
                  >
                    Selanjutnya →
                  </button>
                </div>
              </div>
            )}
          </div>
        )}
          </>
        )}
      </div>
      {/* Floating Mobile AI Button */}
      <button className="mobile-ai-toggle-btn" onClick={() => setShowAiPanel(!showAiPanel)} title="Tanya AI">
        <i className="fas fa-magic"></i>
        <span>AI Agent</span>
      </button>

      {/* ── FLOATING BRAINSTORM CHAT MODAL ── */}
      {showBrainstormChat && (
        <div className="brainstorm-overlay" onClick={() => setShowBrainstormChat(false)}>
          <div className="brainstorm-chat-modal" onClick={e => e.stopPropagation()}>
            {/* Header */}
            <div className="bc-header">
              <div className="bc-header-left">
                <div className="bc-avatar"><i className="fas fa-brain"></i></div>
                <div>
                  <div className="bc-title">Brainstorm AI</div>
                  <div className="bc-sub">{uploadedFileName || 'Dokumen Anda'}</div>
                </div>
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                {brainstormMessages.length > 0 && (
                  <button 
                    className="bc-clear-history" 
                    onClick={() => {
                      if (confirm('Hapus riwayat diskusi brainstorm ini?')) {
                        setBrainstormMessages([]);
                      }
                    }}
                    style={{
                      background: 'none',
                      border: 'none',
                      color: '#ef4444',
                      cursor: 'pointer',
                      padding: '4px 8px',
                      fontSize: '14px',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      borderRadius: '4px'
                    }}
                    title="Hapus riwayat chat"
                  >
                    <i className="fas fa-trash-alt"></i>
                  </button>
                )}
                <button className="bc-close" onClick={() => setShowBrainstormChat(false)}>✕</button>
              </div>
            </div>

            {/* Messages */}
            <div className="bc-messages" id="bc-messages-container">
              {brainstormMessages.length === 0 && (
                <div className="bc-empty">
                  <div className="bc-empty-icon">💬</div>
                  <p>Tanyakan apa saja tentang <strong>{uploadedFileName || 'dokumen Anda'}</strong></p>
                  <div className="bc-suggestions">
                    {[
                      'Apa inti dari dokumen ini?',
                      'Apa kelemahan utama dokumen ini?',
                      'Beri 5 pertanyaan diskusi',
                      'Apa rekomendasi pengembangannya?',
                    ].map((q, i) => (
                      <button key={i} className="bc-suggestion-btn" onClick={() => handleBrainstormSend(q)}>
                        {q}
                      </button>
                    ))}
                  </div>
                </div>
              )}
              {brainstormMessages.map((msg, idx) => (
                <div key={idx} className={`bc-bubble-row ${msg.role}`}>
                  {msg.role === 'assistant' && (
                    <div className="bc-bot-avatar"><i className="fas fa-brain"></i></div>
                  )}
                  <div className={`bc-bubble ${msg.role}`}>
                    {msg.role === 'assistant' ? (
                      renderAssistantMessage(msg.text)
                    ) : (
                      msg.text
                    )}
                  </div>
                </div>
              ))}
              {isBrainstormLoading && (
                <div className="bc-bubble-row assistant">
                  <div className="bc-bot-avatar"><i className="fas fa-brain"></i></div>
                  <div className="bc-bubble assistant bc-typing">
                    <span></span><span></span><span></span>
                  </div>
                </div>
              )}
            </div>

            {/* Input */}
            <div className="bc-input-row">
              <textarea
                className="bc-input"
                value={brainstormInput}
                onChange={e => setBrainstormInput(e.target.value)}
                onKeyDown={e => {
                  if (e.key === 'Enter' && !e.shiftKey) {
                    e.preventDefault();
                    handleBrainstormSend();
                  }
                }}
                placeholder="Tanyakan sesuatu tentang dokumen..."
                rows={1}
                disabled={isBrainstormLoading}
              />
              <button
                className="bc-send-btn"
                onClick={() => handleBrainstormSend()}
                disabled={isBrainstormLoading || !brainstormInput.trim()}
              >
                <i className="fas fa-paper-plane"></i>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ── CUSTOM PRESET MODAL ── */}
      {showPresetModal && (
        <div className="brainstorm-overlay" onClick={() => setShowPresetModal(false)}>
          <div className="brainstorm-chat-modal" onClick={e => e.stopPropagation()} style={{ maxWidth: '400px', height: 'auto', padding: '20px' }}>
            <div className="bc-header" style={{ borderBottom: 'none', padding: '0 0 12px 0' }}>
              <div className="bc-header-left">
                <div className="bc-avatar" style={{ background: 'linear-gradient(135deg, #ff6b00 0%, #dd5700 100%)' }}><i className="fas fa-magic"></i></div>
                <div>
                  <div className="bc-title">{presetModalTitle}</div>
                  <div className="bc-sub">AI Preset Customization</div>
                </div>
              </div>
              <button className="bc-close" onClick={() => setShowPresetModal(false)}>✕</button>
            </div>
            
            <div style={{ display: 'flex', flexDirection: 'column', gap: '14px', marginTop: '8px' }}>
              <label style={{ fontSize: '12px', fontWeight: 600, color: '#475569' }}>
                Masukkan Topik/Detail:
              </label>
              <textarea
                value={presetModalValue}
                onChange={e => setPresetModalValue(e.target.value)}
                placeholder={presetModalPlaceholder}
                className="bc-input"
                style={{ 
                  width: '100%', 
                  minHeight: '80px', 
                  padding: '10px', 
                  borderRadius: '8px', 
                  border: '1.2px solid #cbd5e1', 
                  fontSize: '13px',
                  boxSizing: 'border-box'
                }}
                rows={3}
                onKeyDown={e => {
                  if (e.key === 'Enter' && !e.shiftKey) {
                    e.preventDefault();
                    if (presetModalValue.trim()) {
                      presetModalCallback?.(presetModalValue);
                      setShowPresetModal(false);
                    }
                  }
                }}
              />
              <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px', marginTop: '4px' }}>
                <button 
                  onClick={() => setShowPresetModal(false)}
                  style={{
                    padding: '8px 16px',
                    background: '#f1f5f9',
                    color: '#475569',
                    border: 'none',
                    borderRadius: '8px',
                    fontSize: '12px',
                    fontWeight: 600,
                    cursor: 'pointer'
                  }}
                >
                  Batal
                </button>
                <button 
                  onClick={() => {
                    if (presetModalValue.trim()) {
                      presetModalCallback?.(presetModalValue);
                      setShowPresetModal(false);
                    }
                  }}
                  disabled={!presetModalValue.trim()}
                  style={{
                    padding: '8px 16px',
                    background: 'linear-gradient(135deg, #ff6b00 0%, #dd5700 100%)',
                    color: '#fff',
                    border: 'none',
                    borderRadius: '8px',
                    fontSize: '12px',
                    fontWeight: 600,
                    cursor: 'pointer',
                    opacity: presetModalValue.trim() ? 1 : 0.6
                  }}
                >
                  Generate
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default DocumentEditor;
