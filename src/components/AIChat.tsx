import React, { useState, useRef, useEffect, useCallback } from 'react';
import ReactMarkdown from 'react-markdown';
import remarkGfm from 'remark-gfm';
import rehypeHighlight from 'rehype-highlight';
import 'highlight.js/styles/github.css';
import { Send, Bot, User, X, Minimize2, Maximize2, Trash2, Settings, AlertCircle, FileText, Upload, File, FileSpreadsheet, FileJson, Sparkles } from 'lucide-react';
import { type ChatMessage, type CellAttachment, type FileAttachment } from '../types/ai.types';
import { GeminiService } from '../services/gemini/geminiService';
import { createSpreadsheetHandlers, type SpreadsheetOperations } from '../services/gemini/spreadsheetHandlers';
import { FileParser } from '../utils/fileParser';

interface AIChatProps {
  spreadsheetOperations: SpreadsheetOperations;
  hideFab?: boolean;
}

const AIChat: React.FC<AIChatProps> = ({ spreadsheetOperations, hideFab = false }) => {
  const [isOpen, setIsOpen] = useState(false);
  const [isMinimized, setIsMinimized] = useState(false);
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [inputValue, setInputValue] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [apiKey, setApiKey] = useState('');
  const [showSettings, setShowSettings] = useState(false);
  const [tempApiKey, setTempApiKey] = useState('');
  const [error, setError] = useState<string | null>(null);
  const [cellContexts, setCellContexts] = useState<Array<{
    id: string;
    range: string;
    values: string | number | (string | number | null)[][] | null;
  }>>([]);
  const [fileContexts, setFileContexts] = useState<FileAttachment[]>([]);
  const [isDragging, setIsDragging] = useState(false);
  const [attachmentModalData, setAttachmentModalData] = useState<{
    visible: boolean;
    cellAttachments?: CellAttachment[];
    fileAttachments?: FileAttachment[];
  }>({ visible: false });

  const messagesEndRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const geminiServiceRef = useRef<GeminiService | null>(null);

  // Paste handler to accept images/files from clipboard
  const handlePasteIntoChat = useCallback(async (e: React.ClipboardEvent<HTMLDivElement | HTMLInputElement>) => {
    try {
      const clipboardData = e.clipboardData;
      if (!clipboardData) return;
      const items = clipboardData.items;
      const files: File[] = [];
      for (let i = 0; i < items.length; i++) {
        const item = items[i];
        if (item.kind === 'file') {
          const f = item.getAsFile();
          if (f) files.push(f);
        }
      }
      const newAttachments: FileAttachment[] = [];
      if (files.length > 0) {
        e.preventDefault();
        for (const f of files) {
          const parsed = await FileParser.parseFile(f);
          newAttachments.push({
            id: `file_${Date.now()}_${Math.random()}`,
            name: parsed.name,
            type: parsed.format,
            size: parsed.size,
            content: parsed.content,
            preview: parsed.preview,
          });
        }
      } else {
        // Some apps paste images as data URLs in text/plain
        const text = clipboardData.getData('text/plain');
        const isImageDataUrl = /^data:image\/(png|jpeg|jpg|webp|gif|bmp);base64,.+/i.test(text);
        if (isImageDataUrl) {
          e.preventDefault();
          const mime = text.substring(5, text.indexOf(';')) || 'image/png';
          const ext = mime.includes('jpeg') || mime.includes('jpg') ? 'jpg' : mime.includes('webp') ? 'webp' : mime.includes('gif') ? 'gif' : mime.includes('bmp') ? 'bmp' : 'png';
          const commaIdx = text.indexOf(',');
          const b64 = commaIdx >= 0 ? text.substring(commaIdx + 1) : text;
          const approxSize = Math.floor((b64.length * 3) / 4);
          newAttachments.push({
            id: `file_${Date.now()}_${Math.random()}`,
            name: `pasted-image.${ext}`,
            type: 'image',
            size: approxSize,
            content: text,
            preview: `Pasted image (${mime})`,
          });
        }
      }
      if (newAttachments.length > 0) {
        setFileContexts(prev => [...prev, ...newAttachments]);
      }
    } catch (err) {
      console.error('Paste handling failed:', err);
    }
  }, []);

  // Initialize Gemini service with spreadsheet handlers
  const initializeGeminiService = useCallback((key: string) => {
    try {
      geminiServiceRef.current = new GeminiService(key);
      
      // Register all spreadsheet function handlers with error handling
      const handlers = createSpreadsheetHandlers(spreadsheetOperations);
      Object.entries(handlers).forEach(([name, handler]) => {
        const wrappedHandler = async (args: Record<string, unknown>) => {
          try {
            const result = await (handler as (args: Record<string, unknown>) => Promise<unknown>)(args);
            return result;
          } catch (error) {
            console.error(`Error in ${name}:`, error);
            return {
              success: false,
              error: String(error),
              message: `Failed to execute ${name}`
            };
          }
        };
        geminiServiceRef.current?.registerFunctionHandler(name, wrappedHandler);
      });

      // Add welcome message
      if (messages.length === 0) {
        setMessages([{
          id: `msg_${Date.now()}`,
          role: 'assistant',
          content: "Ask me anything about your spreadsheet",
          timestamp: new Date(),
        }]);
      }

      setError(null);
    } catch (err) {
      setError('Failed to initialize AI service');
      console.error('Error initializing Gemini service:', err);
    }
  }, [spreadsheetOperations, messages.length]);

  // Initialize Gemini service when API key is set
  useEffect(() => {
    const savedApiKey = localStorage.getItem('gemini_api_key');
    if (savedApiKey) {
      setApiKey(savedApiKey);
      initializeGeminiService(savedApiKey);
    }
  }, [initializeGeminiService]);

  // Listen for custom events from spreadsheet
  useEffect(() => {
    const handleAddCellContext = (event: CustomEvent) => {
      const context = event.detail;
      if (context) {
        // Add new context with unique ID
        const newContext = {
          id: `ctx_${Date.now()}`,
          range: context.range,
          values: context.values
        };
        setCellContexts(prev => [...prev, newContext]);
      }
    };

    const handleOpenChat = () => {
      setIsOpen(true);
      setIsMinimized(false);
    };

    window.addEventListener('addCellContext', handleAddCellContext as EventListener);
    window.addEventListener('openAIChat', handleOpenChat);

    return () => {
      window.removeEventListener('addCellContext', handleAddCellContext as EventListener);
      window.removeEventListener('openAIChat', handleOpenChat);
    };
  }, [inputValue]);

  // Scroll to bottom when new messages arrive
  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages]);

  // Handle sending message
  const handleSendMessage = async () => {
    if (!inputValue.trim() || isLoading) return;

    if (!apiKey || !geminiServiceRef.current) {
      setError('Please set your Gemini API key in settings');
      setShowSettings(true);
      return;
    }

    // Prepare message content with all contexts for API
    let messageContentForAPI = inputValue;
    
    if (cellContexts.length > 0) {
      const contextInfo = cellContexts.map(ctx => 
        `\n[Context: Selected cells ${ctx.range} with values: ${JSON.stringify(ctx.values)}]`
      ).join('');
      messageContentForAPI += contextInfo;
    }
    
    if (fileContexts.length > 0) {
      const fileInfo = fileContexts.map(file => {
        const isImage = file.type === 'image' && typeof file.content === 'string';
        const summary = isImage
          ? '(image attached)'
          : (typeof file.content === 'string' ? file.content.substring(0, 1000) : JSON.stringify(file.content));
        return `\n[File: ${file.name} (${file.type}): ${summary}]`;
      }).join('');
      messageContentForAPI += fileInfo;
    }

    // Create user message with attachments stored separately
    const userMessage: ChatMessage = {
      id: `msg_${Date.now()}`,
      role: 'user',
      content: inputValue, // Store only the user's text
      cellAttachments: cellContexts.length > 0 ? cellContexts : undefined,
      fileAttachments: fileContexts.length > 0 ? fileContexts : undefined,
      timestamp: new Date(),
    };

    const updatedMessages = [...messages, userMessage];
    const attachmentsForSend = fileContexts; // preserve before clearing
    setMessages(updatedMessages);
    setInputValue('');
    setCellContexts([]); // Clear contexts after sending
    setFileContexts([]); // Clear file contexts after sending
    setIsLoading(true);
    setError(null);

    try {
      // Send the full context to API but display nicely in UI
      const responses = await geminiServiceRef.current.sendMessage(
        messageContentForAPI,
        updatedMessages,
        attachmentsForSend
      );
      setMessages(prev => [...prev, ...responses]);
    } catch (err) {
      console.error('Error sending message:', err);
      setError('Failed to send message. Please try again.');
      setMessages(prev => [...prev, {
        id: `msg_${Date.now()}`,
        role: 'assistant',
        content: 'I encountered an error processing your request. Please try again.',
        timestamp: new Date(),
      }]);
    } finally {
      setIsLoading(false);
    }
  };

  // Handle API key save
  const handleSaveApiKey = () => {
    if (tempApiKey.trim()) {
      localStorage.setItem('gemini_api_key', tempApiKey);
      setApiKey(tempApiKey);
      initializeGeminiService(tempApiKey);
      setShowSettings(false);
      setTempApiKey('');
    }
  };

  // Clear chat history
  const handleClearChat = () => {
    setMessages([{
      id: `msg_${Date.now()}`,
      role: 'assistant',
      content: "🔄 Chat cleared. Ready for complex operations!\n🎯 Try: 'Add 2025 data with dummy values' or 'Sort and format as report'",
      timestamp: new Date(),
    }]);
    geminiServiceRef.current?.resetChat();
  };

  // Markdown rendering configuration (trusted libs)
  const markdownComponents = {
    a: (props: React.AnchorHTMLAttributes<HTMLAnchorElement>) => (
      <a target="_blank" rel="noopener noreferrer" {...props} />
    ),
    code: ({ className, children, ...props }: { className?: string; children?: React.ReactNode }) => (
      <code className={className} {...props}>{children}</code>
    ),
  } as const;

  const renderMarkdown = (content: string) => (
    <ReactMarkdown
      remarkPlugins={[remarkGfm]}
      rehypePlugins={[rehypeHighlight]}
      components={markdownComponents}
    >
      {content}
    </ReactMarkdown>
  );



  // Handle file drop
  const handleFileDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
    
    const files = Array.from(e.dataTransfer.files);
    const newFileContexts: FileAttachment[] = [];
    
    for (const file of files) {
      const parsedFile = await FileParser.parseFile(file);
      const fileAttachment: FileAttachment = {
        id: `file_${Date.now()}_${Math.random()}`,
        name: parsedFile.name,
        type: parsedFile.format,
        size: parsedFile.size,
        content: parsedFile.content,
        preview: parsedFile.preview
      };
      newFileContexts.push(fileAttachment);
    }
    
    setFileContexts(prev => [...prev, ...newFileContexts]);
  };

  // Handle drag events
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    
    // Only set isDragging to false if we're leaving the entire chat container
    const rect = e.currentTarget.getBoundingClientRect();
    const x = e.clientX;
    const y = e.clientY;
    
    if (x <= rect.left || x >= rect.right || y <= rect.top || y >= rect.bottom) {
      setIsDragging(false);
    }
  };

  if (!isOpen) {
    if (hideFab) return null;
    return (
      <button
        onClick={() => setIsOpen(true)}
        className="ai-chat-fab"
        style={{
          position: 'fixed',
          bottom: '20px',
          right: '20px',
          width: '56px',
          height: '56px',
          borderRadius: '28px',
          background: 'linear-gradient(135deg, #0d6efd 0%, #0a58ca 100%)',
          border: 'none',
          boxShadow: '0 4px 12px rgba(13,110,253,0.3)',
          cursor: 'pointer',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 1000,
          transition: 'all 0.3s ease',
        }}
        onMouseEnter={(e) => {
          e.currentTarget.style.transform = 'scale(1.1)';
          e.currentTarget.style.boxShadow = '0 6px 20px rgba(13,110,253,0.4)';
        }}
        onMouseLeave={(e) => {
          e.currentTarget.style.transform = 'scale(1)';
          e.currentTarget.style.boxShadow = '0 4px 12px rgba(13,110,253,0.3)';
        }}
      >
        <Sparkles size={24} color="white" />
      </button>
    );
  }

  return (
    <>
      <div
        className="ai-chat-container"
        style={{
          position: 'fixed',
          bottom: isMinimized ? '20px' : '20px',
          right: '20px',
          width: isMinimized ? '320px' : '380px',
          height: isMinimized ? '60px' : '600px',
          maxHeight: '80vh',
          backgroundColor: isDragging ? 'rgba(255, 255, 255, 0.95)' : 'white',
          borderRadius: '12px',
          boxShadow: isDragging 
            ? '0 10px 40px rgba(102, 126, 234, 0.3)' 
            : '0 10px 40px rgba(0,0,0,0.15)',
          display: 'flex',
          flexDirection: 'column',
          zIndex: 1000,
          transition: 'all 0.3s ease',
          overflow: 'hidden',
          border: isDragging ? '2px dashed #667eea' : '1px solid #dee2e6',
        }}
        onPaste={handlePasteIntoChat}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleFileDrop}
      >
        {/* Header */}
        <div
          style={{
            padding: '12px 16px',
            background: 'linear-gradient(135deg, #007bff 0%, #0056b3 100%)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            borderRadius: '12px 12px 0 0',
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <Bot size={20} color="white" />
            <span style={{ color: 'white', fontWeight: 600, fontSize: '14px' }}>
              Anna
            </span>
          </div>
          <div style={{ display: 'flex', gap: '8px' }}>
            <button
              onClick={() => setShowSettings(!showSettings)}
              style={{
                background: 'rgba(255,255,255,0.2)',
                border: 'none',
                borderRadius: '4px',
                padding: '4px',
                cursor: 'pointer',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                transition: 'background 0.2s',
              }}
              onMouseEnter={(e) => e.currentTarget.style.background = 'rgba(255,255,255,0.3)'}
              onMouseLeave={(e) => e.currentTarget.style.background = 'rgba(255,255,255,0.2)'}
            >
              <Settings size={16} color="white" />
            </button>
            <button
              onClick={handleClearChat}
              style={{
                background: 'rgba(255,255,255,0.2)',
                border: 'none',
                borderRadius: '4px',
                padding: '4px',
                cursor: 'pointer',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                transition: 'background 0.2s',
              }}
              onMouseEnter={(e) => e.currentTarget.style.background = 'rgba(255,255,255,0.3)'}
              onMouseLeave={(e) => e.currentTarget.style.background = 'rgba(255,255,255,0.2)'}
            >
              <Trash2 size={16} color="white" />
            </button>
            <button
              onClick={() => setIsMinimized(!isMinimized)}
              style={{
                background: 'rgba(255,255,255,0.2)',
                border: 'none',
                borderRadius: '4px',
                padding: '4px',
                cursor: 'pointer',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                transition: 'background 0.2s',
              }}
              onMouseEnter={(e) => e.currentTarget.style.background = 'rgba(255,255,255,0.3)'}
              onMouseLeave={(e) => e.currentTarget.style.background = 'rgba(255,255,255,0.2)'}
            >
              {isMinimized ? <Maximize2 size={16} color="white" /> : <Minimize2 size={16} color="white" />}
            </button>
            <button
              onClick={() => setIsOpen(false)}
              style={{
                background: 'rgba(255,255,255,0.2)',
                border: 'none',
                borderRadius: '4px',
                padding: '4px',
                cursor: 'pointer',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                transition: 'background 0.2s',
              }}
              onMouseEnter={(e) => e.currentTarget.style.background = 'rgba(255,255,255,0.3)'}
              onMouseLeave={(e) => e.currentTarget.style.background = 'rgba(255,255,255,0.2)'}
            >
              <X size={16} color="white" />
            </button>
          </div>
        </div>

        {!isMinimized && (
          <>
            {/* Drag Overlay */}
            {isDragging && (
              <div
                style={{
                  position: 'absolute',
                  top: 0,
                  left: 0,
                  right: 0,
                  bottom: 0,
                  backgroundColor: 'rgba(102, 126, 234, 0.1)',
                  backdropFilter: 'blur(4px)',
                  display: 'flex',
                  flexDirection: 'column',
                  alignItems: 'center',
                  justifyContent: 'center',
                  zIndex: 10000,
                  pointerEvents: 'none',
                }}
              >
                <div
                  style={{
                    backgroundColor: 'white',
                    borderRadius: '16px',
                    padding: '32px',
                    boxShadow: '0 8px 32px rgba(102, 126, 234, 0.2)',
                    display: 'flex',
                    flexDirection: 'column',
                    alignItems: 'center',
                    gap: '16px',
                    animation: 'bounceIn 0.3s ease-out',
                  }}
                >
                  <Upload size={48} color="#667eea" />
                  <div
                    style={{
                      fontSize: '18px',
                      fontWeight: 600,
                      color: '#2d3748',
                    }}
                  >
                    Drop files here
                  </div>
                  <div
                    style={{
                      fontSize: '13px',
                      color: '#718096',
                      textAlign: 'center',
                    }}
                  >
                    CSV • Excel • JSON • TSV • TXT
                  </div>
                </div>
              </div>
            )}
            
            {/* Settings Panel */}
            {showSettings && (
              <div
                style={{
                  padding: '16px',
                  backgroundColor: '#f8f9fa',
                  borderBottom: '1px solid #dee2e6',
                }}
              >
                <div style={{ marginBottom: '8px' }}>
                  <label style={{ fontSize: '12px', color: '#495057', fontWeight: 600 }}>
                    Gemini API Key
                  </label>
                  <input
                    type="password"
                    value={tempApiKey || apiKey}
                    onChange={(e) => setTempApiKey(e.target.value)}
                    placeholder="Enter your Gemini API key"
                    style={{
                      width: '100%',
                      padding: '8px 12px',
                      marginTop: '4px',
                      border: '1px solid #ced4da',
                      borderRadius: '4px',
                      fontSize: '13px',
                      color: '#495057',
                    }}
                  />
                </div>
                <div style={{ display: 'flex', gap: '8px', justifyContent: 'flex-end' }}>
                  <button
                    onClick={() => {
                      setShowSettings(false);
                      setTempApiKey('');
                    }}
                    style={{
                      padding: '6px 12px',
                      border: '1px solid #ced4da',
                      borderRadius: '4px',
                      background: 'white',
                      color: '#495057',
                      fontSize: '12px',
                      cursor: 'pointer',
                    }}
                  >
                    Cancel
                  </button>
                  <button
                    onClick={handleSaveApiKey}
                    style={{
                      padding: '6px 12px',
                      border: 'none',
                      borderRadius: '4px',
                      background: '#007bff',
                      color: 'white',
                      fontSize: '12px',
                      cursor: 'pointer',
                    }}
                  >
                    Save
                  </button>
                </div>
                <div style={{ marginTop: '8px', fontSize: '11px', color: '#6c757d' }}>
                  Get your API key from{' '}
                  <a
                    href="https://makersuite.google.com/app/apikey"
                    target="_blank"
                    rel="noopener noreferrer"
                    style={{ color: '#007bff', textDecoration: 'none' }}
                  >
                    Google AI Studio
                  </a>
                </div>
              </div>
            )}

            {/* Error Alert */}
            {error && (
              <div
                style={{
                  padding: '8px 12px',
                  backgroundColor: '#fff3cd',
                  borderBottom: '1px solid #ffc107',
                  display: 'flex',
                  alignItems: 'center',
                  gap: '8px',
                  fontSize: '12px',
                  color: '#856404',
                }}
              >
                <AlertCircle size={14} />
                {error}
              </div>
            )}

            {/* Messages Area */}
            <div
              style={{
                flex: 1,
                overflowY: 'auto',
                padding: '16px',
                backgroundColor: '#fafafa',
              }}
            >
              {messages.map((message) => (
                <div
                  key={message.id}
                  style={{
                    marginBottom: '12px',
                    display: 'flex',
                    gap: '8px',
                    alignItems: 'flex-start',
                    animation: 'messageSlideIn 0.3s ease',
                  }}
                >
                  {message.role !== 'function' ? (
                    <div
                      style={{
                        width: '28px',
                        height: '28px',
                        borderRadius: '14px',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        flexShrink: 0,
                        background: message.role === 'user' 
                          ? '#e9ecef' 
                          : 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
                      }}
                    >
                      {message.role === 'user' ? (
                        <User size={14} color="#495057" />
                      ) : (
                        <Bot size={14} color="white" />
                      )}
                    </div>
                  ) : (
                    <div style={{ width: '28px', flexShrink: 0 }} />
                  )}
                  <div style={{ flex: 1 }}>
                    <div
                      style={{
                        background: message.role === 'user' ? '#e9ecef' : message.role === 'function' ? 'transparent' : 'white',
                        padding: message.role === 'function' ? '4px 0' : '8px 12px',
                        borderRadius: '8px',
                        fontSize: '13px',
                        color: '#333',
                        lineHeight: '1.5',
                        border: message.role === 'user' || message.role === 'function' ? 'none' : '1px solid #dee2e6',
                        position: 'relative',
                      }}
                    >
                      {/* Cell Attachments as Compact Badges */}
                      {message.cellAttachments && message.cellAttachments.length > 0 && (
                        <div
                          style={{
                            position: 'absolute',
                            top: '-8px',
                            right: '8px',
                            display: 'flex',
                            gap: '4px',
                            alignItems: 'center',
                          }}
                        >
                          {/* Show first 2 attachments as small badges */}
                          {message.cellAttachments.slice(0, 2).map((attachment) => (
                            <div
                              key={attachment.id}
                              style={{
                                padding: '2px 6px',
                                backgroundColor: '#667eea',
                                color: 'white',
                                borderRadius: '10px',
                                fontSize: '10px',
                                fontWeight: 600,
                                display: 'flex',
                                alignItems: 'center',
                                gap: '2px',
                                cursor: 'pointer',
                                boxShadow: '0 1px 3px rgba(0,0,0,0.2)',
                                transition: 'all 0.2s ease',
                                animation: 'fadeIn 0.2s ease-out',
                              }}
                              onClick={() => {
                                const highlightEvent = new CustomEvent('highlightCells', {
                                  detail: { range: attachment.range }
                                });
                                window.dispatchEvent(highlightEvent);
                              }}
                              onMouseEnter={(e) => {
                                e.currentTarget.style.backgroundColor = '#5a67d8';
                                e.currentTarget.style.transform = 'scale(1.05)';
                              }}
                              onMouseLeave={(e) => {
                                e.currentTarget.style.backgroundColor = '#667eea';
                                e.currentTarget.style.transform = 'scale(1)';
                              }}
                              title={`${attachment.range}`}
                            >
                              <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                                <rect x="3" y="3" width="18" height="18" rx="2"/>
                                <line x1="9" y1="3" x2="9" y2="21"/>
                                <line x1="3" y1="9" x2="21" y2="9"/>
                              </svg>
                              <span>{attachment.range}</span>
                            </div>
                          ))}
                          
                          {/* Show ellipsis if more than 2 attachments */}
                          {message.cellAttachments.length > 2 && (
                            <div
                              style={{
                                padding: '2px 8px',
                                backgroundColor: '#9f7aea',
                                color: 'white',
                                borderRadius: '10px',
                                fontSize: '10px',
                                fontWeight: 700,
                                cursor: 'pointer',
                                boxShadow: '0 1px 3px rgba(0,0,0,0.2)',
                                transition: 'all 0.2s ease',
                                display: 'flex',
                                alignItems: 'center',
                                gap: '2px',
                              }}
                              onClick={() => {
                                setAttachmentModalData({
                                  visible: true,
                                  cellAttachments: message.cellAttachments
                                });
                              }}
                              onMouseEnter={(e) => {
                                e.currentTarget.style.backgroundColor = '#8b5cf6';
                                e.currentTarget.style.transform = 'scale(1.05)';
                              }}
                              onMouseLeave={(e) => {
                                e.currentTarget.style.backgroundColor = '#9f7aea';
                                e.currentTarget.style.transform = 'scale(1)';
                              }}
                              title={`View all ${message.cellAttachments.length} attachments`}
                            >
                              <span>+{message.cellAttachments.length - 2}</span>
                              <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
                                <circle cx="12" cy="12" r="10"/>
                                <path d="M8 12h8M12 8v8"/>
                              </svg>
                            </div>
                          )}
                        </div>
                      )}
                      
                      {/* File Attachments as Compact Badges */}
                      {message.fileAttachments && message.fileAttachments.length > 0 && (
                        <div
                          style={{
                            position: 'absolute',
                            top: message.cellAttachments && message.cellAttachments.length > 0 ? '12px' : '-8px',
                            right: '8px',
                            display: 'flex',
                            gap: '4px',
                            alignItems: 'center',
                          }}
                        >
                          {/* Show first 2 file attachments as small badges */}
                          {message.fileAttachments.slice(0, 2).map((attachment) => (
                            <div
                              key={attachment.id}
                              style={{
                                padding: '2px 6px',
                                backgroundColor: '#48bb78',
                                color: 'white',
                                borderRadius: '10px',
                                fontSize: '10px',
                                fontWeight: 600,
                                display: 'flex',
                                alignItems: 'center',
                                gap: '2px',
                                cursor: 'pointer',
                                boxShadow: '0 1px 3px rgba(0,0,0,0.2)',
                                transition: 'all 0.2s ease',
                                animation: 'fadeIn 0.2s ease-out',
                              }}
                              onClick={() => {
                                setAttachmentModalData({
                                  visible: true,
                                  fileAttachments: [attachment]
                                });
                              }}
                              onMouseEnter={(e) => {
                                e.currentTarget.style.backgroundColor = '#38a169';
                                e.currentTarget.style.transform = 'scale(1.05)';
                              }}
                              onMouseLeave={(e) => {
                                e.currentTarget.style.backgroundColor = '#48bb78';
                                e.currentTarget.style.transform = 'scale(1)';
                              }}
                              title={attachment.name}
                            >
                              <FileText size={10} />
                              <span style={{ maxWidth: '60px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                                {attachment.name.split('.')[0]}
                              </span>
                            </div>
                          ))}
                          
                          {/* Show ellipsis if more than 2 file attachments */}
                          {message.fileAttachments.length > 2 && (
                            <div
                              style={{
                                padding: '2px 8px',
                                backgroundColor: '#2d9c5e',
                                color: 'white',
                                borderRadius: '10px',
                                fontSize: '10px',
                                fontWeight: 700,
                                cursor: 'pointer',
                                boxShadow: '0 1px 3px rgba(0,0,0,0.2)',
                                transition: 'all 0.2s ease',
                                display: 'flex',
                                alignItems: 'center',
                                gap: '2px',
                              }}
                              onClick={() => {
                                setAttachmentModalData({
                                  visible: true,
                                  fileAttachments: message.fileAttachments
                                });
                              }}
                              onMouseEnter={(e) => {
                                e.currentTarget.style.backgroundColor = '#22543d';
                                e.currentTarget.style.transform = 'scale(1.05)';
                              }}
                              onMouseLeave={(e) => {
                                e.currentTarget.style.backgroundColor = '#2d9c5e';
                                e.currentTarget.style.transform = 'scale(1)';
                              }}
                              title={`View all ${message.fileAttachments.length} files`}
                            >
                              <span>+{message.fileAttachments.length - 2}</span>
                              <FileText size={10} />
                            </div>
                          )}
                        </div>
                      )}
                      
                      {message.role === 'function' ? (
                        <div style={{ 
                          fontSize: '11px', 
                          color: '#6b7280',
                          fontStyle: 'italic',
                          opacity: 0.8
                        }}>
                          {message.content}
                        </div>
                      ) : renderMarkdown(message.content)}



                    </div>
                    {message.role !== 'function' && (
                      <div
                        style={{
                          fontSize: '10px',
                          color: '#6c757d',
                          marginTop: '4px',
                          marginLeft: '12px',
                        }}
                      >
                        {message.timestamp.toLocaleTimeString()}
                      </div>
                    )}
                  </div>
                </div>
              ))}
              {isLoading && (
                <div
                  style={{
                    display: 'flex',
                    gap: '8px',
                    alignItems: 'flex-start',
                    padding: '12px 16px',
                    animation: 'messageSlideIn 0.3s ease',
                  }}
                >
                  <div
                    style={{
                      width: '28px',
                      height: '28px',
                      borderRadius: '50%',
                      background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      animation: 'toolExecuting 1.5s ease-in-out infinite',
                    }}
                  >
                    <Bot size={14} color="white" />
                  </div>
                  <div style={{ flex: 1 }}>
                    <div
                      style={{
                        display: 'inline-flex',
                        alignItems: 'center',
                        gap: '8px',
                        padding: '6px 10px',
                        backgroundColor: '#f9fafb',
                        borderRadius: '8px',
                        border: '1px solid #e5e7eb',
                      }}
                    >
                      <div style={{ display: 'flex', gap: '3px' }}>
                        <div
                          style={{
                            width: '3px',
                            height: '3px',
                            borderRadius: '50%',
                            backgroundColor: '#667eea',
                            animation: 'toolCallPulse 1.4s ease-in-out infinite',
                          }}
                        />
                        <div
                          style={{
                            width: '3px',
                            height: '3px',
                            borderRadius: '50%',
                            backgroundColor: '#667eea',
                            animation: 'toolCallPulse 1.4s ease-in-out 0.2s infinite',
                          }}
                        />
                        <div
                          style={{
                            width: '3px',
                            height: '3px',
                            borderRadius: '50%',
                            backgroundColor: '#667eea',
                            animation: 'toolCallPulse 1.4s ease-in-out 0.4s infinite',
                          }}
                        />
                      </div>
                    </div>
                  </div>
                </div>
              )}
              <div ref={messagesEndRef} />
            </div>

            {/* Input Area */}
            <div
              style={{
                padding: '12px',
                borderTop: '1px solid #dee2e6',
                backgroundColor: 'white',
              }}
            >
              {/* Cell Context Indicators */}
              {cellContexts.length > 0 && (
                <div
                  style={{
                    marginBottom: '12px',
                    display: 'flex',
                    flexWrap: 'wrap',
                    gap: '8px',
                  }}
                >
                  {cellContexts.map((context) => (
                    <div
                      key={context.id}
                      style={{
                        padding: '8px 12px',
                        backgroundColor: '#4a90e2',
                        color: 'white',
                        borderRadius: '16px',
                        fontSize: '12px',
                        fontWeight: 600,
                        display: 'flex',
                        alignItems: 'center',
                        gap: '6px',
                        boxShadow: '0 2px 4px rgba(0,0,0,0.15)',
                        animation: 'bounceIn 0.3s ease-out',
                        cursor: 'pointer',
                        transition: 'all 0.2s ease',
                      }}
                      onClick={() => {
                        // Highlight cells in spreadsheet
                        const highlightEvent = new CustomEvent('highlightCells', {
                          detail: { range: context.range }
                        });
                        window.dispatchEvent(highlightEvent);
                      }}
                      onMouseEnter={(e) => {
                        e.currentTarget.style.backgroundColor = '#357abd';
                        e.currentTarget.style.transform = 'scale(1.05)';
                      }}
                      onMouseLeave={(e) => {
                        e.currentTarget.style.backgroundColor = '#4a90e2';
                        e.currentTarget.style.transform = 'scale(1)';
                      }}
                      title={`Click to focus cells ${context.range}`}
                    >
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <rect x="3" y="3" width="18" height="18" rx="2"/>
                        <rect x="7" y="7" width="3" height="9"/>
                        <rect x="14" y="7" width="3" height="5"/>
                      </svg>
                      <span>{context.range}</span>
                      <button
                        onClick={(e) => {
                          e.stopPropagation();
                          setCellContexts(prev => prev.filter(c => c.id !== context.id));
                        }}
                        style={{
                          background: 'rgba(255, 255, 255, 0.2)',
                          border: 'none',
                          borderRadius: '50%',
                          width: '16px',
                          height: '16px',
                          display: 'flex',
                          alignItems: 'center',
                          justifyContent: 'center',
                          cursor: 'pointer',
                          marginLeft: '4px',
                          padding: 0,
                          transition: 'background 0.2s ease',
                        }}
                        onMouseEnter={(e) => {
                          e.currentTarget.style.background = 'rgba(255, 255, 255, 0.4)';
                        }}
                        onMouseLeave={(e) => {
                          e.currentTarget.style.background = 'rgba(255, 255, 255, 0.2)';
                        }}
                        title="Remove this attachment"
                      >
                        <X size={10} color="white" />
                      </button>
                    </div>
                  ))}
                </div>
              )}
              
              {/* File Context Indicators - Improved UI */}
              {fileContexts.length > 0 && (
                <div
                  style={{
                    marginBottom: '12px',
                    padding: '12px',
                    backgroundColor: '#f7fafc',
                    borderRadius: '8px',
                    border: '1px solid #e2e8f0',
                  }}
                >
                  <div
                    style={{
                      fontSize: '11px',
                      fontWeight: 600,
                      color: '#718096',
                      marginBottom: '8px',
                      textTransform: 'uppercase',
                      letterSpacing: '0.5px',
                    }}
                  >
                    📎 Attached Files ({fileContexts.length})
                  </div>
                  <div
                    style={{
                      display: 'flex',
                      flexDirection: 'column',
                      gap: '8px',
                    }}
                  >
                    {fileContexts.map((file) => {
                      const getFileIcon = () => {
                        if (file.type === 'csv' || file.type === 'excel') return <FileSpreadsheet size={16} />;
                        if (file.type === 'json') return <FileJson size={16} />;
                        if (file.type === 'text') return <FileText size={16} />;
                        return <File size={16} />;
                      };
                      
                      const getFileColor = () => {
                        if (file.type === 'csv' || file.type === 'excel') return '#10b981';
                        if (file.type === 'json') return '#8b5cf6';
                        if (file.type === 'text') return '#3b82f6';
                        return '#6b7280';
                      };
                      
                      return (
                        <div
                          key={file.id}
                          style={{
                            display: 'flex',
                            alignItems: 'center',
                            padding: '10px',
                            backgroundColor: 'white',
                            borderRadius: '6px',
                            border: '1px solid #e2e8f0',
                            transition: 'all 0.2s ease',
                          }}
                          onMouseEnter={(e) => {
                            e.currentTarget.style.borderColor = getFileColor();
                            e.currentTarget.style.boxShadow = `0 2px 8px ${getFileColor()}20`;
                          }}
                          onMouseLeave={(e) => {
                            e.currentTarget.style.borderColor = '#e2e8f0';
                            e.currentTarget.style.boxShadow = 'none';
                          }}
                        >
                          <div
                            style={{
                              width: '36px',
                              height: '36px',
                              backgroundColor: `${getFileColor()}15`,
                              borderRadius: '6px',
                              display: 'flex',
                              alignItems: 'center',
                              justifyContent: 'center',
                              color: getFileColor(),
                              flexShrink: 0,
                            }}
                          >
                            {getFileIcon()}
                          </div>
                          <div
                            style={{
                              flex: 1,
                              marginLeft: '12px',
                              minWidth: 0,
                            }}
                          >
                            <div
                              style={{
                                fontSize: '13px',
                                fontWeight: 600,
                                color: '#1a202c',
                                whiteSpace: 'nowrap',
                                overflow: 'hidden',
                                textOverflow: 'ellipsis',
                              }}
                            >
                              {file.name}
                            </div>
                            <div
                              style={{
                                fontSize: '11px',
                                color: '#718096',
                                marginTop: '2px',
                                whiteSpace: 'nowrap',
                                overflow: 'hidden',
                                textOverflow: 'ellipsis',
                              }}
                            >
                              {file.preview}
                            </div>
                            <div
                              style={{
                                fontSize: '10px',
                                color: '#a0aec0',
                                marginTop: '2px',
                              }}
                            >
                              {FileParser.formatFileSize(file.size)} • {file.type.toUpperCase()}
                            </div>
                          </div>
                          <button
                            onClick={() => {
                              setFileContexts(prev => prev.filter(f => f.id !== file.id));
                            }}
                            style={{
                              background: 'transparent',
                              border: '1px solid #e2e8f0',
                              borderRadius: '6px',
                              width: '28px',
                              height: '28px',
                              display: 'flex',
                              alignItems: 'center',
                              justifyContent: 'center',
                              cursor: 'pointer',
                              transition: 'all 0.2s ease',
                              flexShrink: 0,
                              marginLeft: '8px',
                            }}
                            onMouseEnter={(e) => {
                              e.currentTarget.style.backgroundColor = '#fee2e2';
                              e.currentTarget.style.borderColor = '#ef4444';
                            }}
                            onMouseLeave={(e) => {
                              e.currentTarget.style.backgroundColor = 'transparent';
                              e.currentTarget.style.borderColor = '#e2e8f0';
                            }}
                            title="Remove this file"
                          >
                            <X size={14} color="#6b7280" />
                          </button>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
              
              <div style={{ display: 'flex', gap: '8px' }}>
                <input
                  type="file"
                  id="file-upload"
                  multiple
                  accept="image/*,.csv,.xlsx,.xls,.json,.txt,.tsv,.tab"
                  style={{ display: 'none' }}
                  onChange={async (e) => {
                    const files = Array.from(e.target.files || []);
                    const newFileContexts: FileAttachment[] = [];
                    
                    for (const file of files) {
                      const parsedFile = await FileParser.parseFile(file);
                      const fileAttachment: FileAttachment = {
                        id: `file_${Date.now()}_${Math.random()}`,
                        name: parsedFile.name,
                        type: parsedFile.format,
                        size: parsedFile.size,
                        content: parsedFile.content,
                        preview: parsedFile.preview
                      };
                      newFileContexts.push(fileAttachment);
                    }
                    
                    setFileContexts(prev => [...prev, ...newFileContexts]);
                    e.target.value = ''; // Reset input
                  }}
                />
                <button
                  onClick={() => document.getElementById('file-upload')?.click()}
                  style={{
                    padding: '8px',
                    border: '1px solid #ced4da',
                    borderRadius: '4px',
                    background: 'white',
                    cursor: 'pointer',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    transition: 'all 0.2s ease',
                  }}
                  onMouseEnter={(e) => {
                    e.currentTarget.style.backgroundColor = '#f8f9fa';
                    e.currentTarget.style.borderColor = '#007bff';
                  }}
                  onMouseLeave={(e) => {
                    e.currentTarget.style.backgroundColor = 'white';
                    e.currentTarget.style.borderColor = '#ced4da';
                  }}
                  title="Attach files"
                >
                  <Upload size={16} color="#6c757d" />
                </button>
                <input
                  ref={inputRef}
                  type="text"
                  value={inputValue}
                  onChange={(e) => setInputValue(e.target.value)}
                  onPaste={handlePasteIntoChat}
                  onKeyPress={(e) => e.key === 'Enter' && handleSendMessage()}
                  placeholder={apiKey ? "Ask me anything about your sheet..." : "Set API key in settings first..."}
                  disabled={!apiKey || isLoading}
                  style={{
                    flex: 1,
                    padding: '8px 12px',
                    border: '1px solid #ced4da',
                    borderRadius: '4px',
                    fontSize: '13px',
                    color: 'white',
                    outline: 'none',
                  }}
                />
                <button
                  onClick={handleSendMessage}
                  disabled={!apiKey || isLoading || !inputValue.trim()}
                  style={{
                    padding: '8px 16px',
                    border: 'none',
                    borderRadius: '4px',
                    background: !apiKey || isLoading || !inputValue.trim() 
                      ? '#e9ecef' 
                      : 'linear-gradient(135deg, #007bff 0%, #0056b3 100%)',
                    color: !apiKey || isLoading || !inputValue.trim() ? '#6c757d' : 'white',
                    cursor: !apiKey || isLoading || !inputValue.trim() ? 'not-allowed' : 'pointer',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '4px',
                    transition: 'all 0.2s',
                  }}
                >
                  <Send size={14} />
                </button>
              </div>
            </div>
          </>
        )}
      </div>

      {/* Add animations */}
      <style>{`
        @keyframes messageSlideIn {
          from {
            opacity: 0;
            transform: translateY(10px);
          }
          to {
            opacity: 1;
            transform: translateY(0);
          }
        }

        @keyframes spin {
          from {
            transform: rotate(0deg);
          }
          to {
            transform: rotate(360deg);
          }
        }
        
        @keyframes toolCallFade {
          from {
            opacity: 0;
            transform: translateX(-10px);
          }
          to {
            opacity: 0.8;
            transform: translateX(0);
          }
        }
        
        @keyframes toolCallPulse {
          0%, 100% {
            opacity: 0.4;
          }
          50% {
            opacity: 1;
          }
        }
        
        @keyframes toolExecuting {
          0% {
            transform: scale(0.8);
            opacity: 0.3;
          }
          50% {
            transform: scale(1.1);
            opacity: 0.8;
          }
          100% {
            transform: scale(1);
            opacity: 0.6;
          }
        }

        /* New animations for enhanced UI */
        @keyframes slideIn {
          from {
            opacity: 0;
            transform: translateX(-20px);
          }
          to {
            opacity: 1;
            transform: translateX(0);
          }
        }
        
        @keyframes bounceIn {
          0% {
            transform: scale(0.3);
            opacity: 0;
          }
          50% {
            transform: scale(1.05);
          }
          70% {
            transform: scale(0.9);
          }
          100% {
            transform: scale(1);
            opacity: 1;
          }
        }
        
        @keyframes glow {
          0%, 100% {
            box-shadow: 0 0 5px rgba(102, 126, 234, 0.3);
          }
          50% {
            box-shadow: 0 0 20px rgba(102, 126, 234, 0.6);
          }
        }
        
        @keyframes pulse {
          0%, 100% {
            opacity: 1;
            transform: scale(1);
          }
          50% {
            opacity: 0.8;
            transform: scale(1.05);
          }
        }
        
        /* Scrollbar styling */
        .ai-chat-container ::-webkit-scrollbar {
          width: 6px;
        }
        
        .ai-chat-container ::-webkit-scrollbar-track {
          background: #f1f1f1;
        }
        
        .ai-chat-container ::-webkit-scrollbar-thumb {
          background: #888;
          border-radius: 3px;
        }
        
        .ai-chat-container ::-webkit-scrollbar-thumb:hover {
          background: #555;
        }
      `}</style>

      {/* Attachments Modal */}
      {attachmentModalData.visible && (
        <div
          style={{
            position: 'fixed',
            top: 0,
            left: 0,
            right: 0,
            bottom: 0,
            backgroundColor: 'rgba(0, 0, 0, 0.5)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 100000,
            animation: 'fadeIn 0.2s ease-out',
          }}
          onClick={() => setAttachmentModalData({ visible: false })}
        >
          <div
            style={{
              backgroundColor: 'white',
              borderRadius: '12px',
              padding: '20px',
              maxWidth: '500px',
              width: '90%',
              maxHeight: '70vh',
              overflow: 'auto',
              boxShadow: '0 10px 40px rgba(0, 0, 0, 0.2)',
              animation: 'slideIn 0.3s ease-out',
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div
              style={{
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center',
                marginBottom: '16px',
                borderBottom: '2px solid #e2e8f0',
                paddingBottom: '12px',
              }}
            >
              <h3
                style={{
                  margin: 0,
                  fontSize: '18px',
                  fontWeight: 700,
                  color: '#1a202c',
                  display: 'flex',
                  alignItems: 'center',
                  gap: '8px',
                }}
              >
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#667eea" strokeWidth="2">
                  <rect x="3" y="3" width="18" height="18" rx="2"/>
                  <line x1="9" y1="3" x2="9" y2="21"/>
                  <line x1="15" y1="3" x2="15" y2="21"/>
                  <line x1="3" y1="9" x2="21" y2="9"/>
                  <line x1="3" y1="15" x2="21" y2="15"/>
                </svg>
                Attachments ({(attachmentModalData.cellAttachments?.length || 0) + (attachmentModalData.fileAttachments?.length || 0)})
              </h3>
              <button
                onClick={() => setAttachmentModalData({ visible: false })}
                style={{
                  background: 'transparent',
                  border: 'none',
                  cursor: 'pointer',
                  padding: '4px',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  borderRadius: '4px',
                  transition: 'background 0.2s ease',
                }}
                onMouseEnter={(e) => {
                  e.currentTarget.style.background = '#f0f4f8';
                }}
                onMouseLeave={(e) => {
                  e.currentTarget.style.background = 'transparent';
                }}
              >
                <X size={20} color="#718096" />
              </button>
            </div>
            
            <div
              style={{
                display: 'grid',
                gap: '12px',
              }}
            >
              {/* Cell Attachments */}
              {attachmentModalData.cellAttachments?.map((attachment, index) => (
                <div
                  key={attachment.id}
                  style={{
                    padding: '12px',
                    backgroundColor: '#f7fafc',
                    borderRadius: '8px',
                    border: '1px solid #e2e8f0',
                    cursor: 'pointer',
                    transition: 'all 0.2s ease',
                  }}
                  onClick={() => {
                    const highlightEvent = new CustomEvent('highlightCells', {
                      detail: { range: attachment.range }
                    });
                    window.dispatchEvent(highlightEvent);
                    setAttachmentModalData({ visible: false });
                  }}
                  onMouseEnter={(e) => {
                    e.currentTarget.style.backgroundColor = '#edf2f7';
                    e.currentTarget.style.borderColor = '#667eea';
                    e.currentTarget.style.transform = 'translateY(-2px)';
                    e.currentTarget.style.boxShadow = '0 4px 12px rgba(102, 126, 234, 0.15)';
                  }}
                  onMouseLeave={(e) => {
                    e.currentTarget.style.backgroundColor = '#f7fafc';
                    e.currentTarget.style.borderColor = '#e2e8f0';
                    e.currentTarget.style.transform = 'translateY(0)';
                    e.currentTarget.style.boxShadow = 'none';
                  }}
                >
                  <div
                    style={{
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'space-between',
                      marginBottom: '8px',
                    }}
                  >
                    <div
                      style={{
                        display: 'flex',
                        alignItems: 'center',
                        gap: '8px',
                      }}
                    >
                      <div
                        style={{
                          width: '32px',
                          height: '32px',
                          backgroundColor: '#667eea',
                          borderRadius: '6px',
                          display: 'flex',
                          alignItems: 'center',
                          justifyContent: 'center',
                        }}
                      >
                        <span
                          style={{
                            color: 'white',
                            fontSize: '14px',
                            fontWeight: 700,
                          }}
                        >
                          {index + 1}
                        </span>
                      </div>
                      <div>
                        <div
                          style={{
                            fontSize: '14px',
                            fontWeight: 600,
                            color: '#2d3748',
                          }}
                        >
                          {attachment.range}
                        </div>
                        <div
                          style={{
                            fontSize: '11px',
                            color: '#718096',
                          }}
                        >
                          {attachment.range.includes(':') ? 'Cell Range' : 'Single Cell'}
                        </div>
                      </div>
                    </div>
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#cbd5e0" strokeWidth="2">
                      <polyline points="9 18 15 12 9 6"/>
                    </svg>
                  </div>
                  
                  {attachment.values && (
                    <div
                      style={{
                        padding: '8px',
                        backgroundColor: 'white',
                        borderRadius: '4px',
                        fontSize: '12px',
                        fontFamily: 'monospace',
                        color: '#4a5568',
                        maxHeight: '100px',
                        overflow: 'auto',
                      }}
                    >
                      {Array.isArray(attachment.values) 
                        ? attachment.values.map((row, i) => (
                            <div key={i}>
                              {Array.isArray(row) 
                                ? row.map(v => v ?? '(empty)').join(' | ')
                                : row ?? '(empty)'}
                            </div>
                          ))
                        : attachment.values || '(empty)'}
                    </div>
                  )}
                </div>
              ))}
              
              {/* File Attachments */}
              {attachmentModalData.fileAttachments?.map((file) => (
                <div
                  key={file.id}
                  style={{
                    padding: '12px',
                    backgroundColor: '#f0fff4',
                    borderRadius: '8px',
                    border: '1px solid #9ae6b4',
                    transition: 'all 0.2s ease',
                  }}
                  onMouseEnter={(e) => {
                    e.currentTarget.style.backgroundColor = '#e6fffa';
                    e.currentTarget.style.borderColor = '#48bb78';
                    e.currentTarget.style.transform = 'translateY(-2px)';
                    e.currentTarget.style.boxShadow = '0 4px 12px rgba(72, 187, 120, 0.15)';
                  }}
                  onMouseLeave={(e) => {
                    e.currentTarget.style.backgroundColor = '#f0fff4';
                    e.currentTarget.style.borderColor = '#9ae6b4';
                    e.currentTarget.style.transform = 'translateY(0)';
                    e.currentTarget.style.boxShadow = 'none';
                  }}
                >
                  <div
                    style={{
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'space-between',
                      marginBottom: '8px',
                    }}
                  >
                    <div
                      style={{
                        display: 'flex',
                        alignItems: 'center',
                        gap: '8px',
                      }}
                    >
                      <div
                        style={{
                          width: '32px',
                          height: '32px',
                          backgroundColor: '#48bb78',
                          borderRadius: '6px',
                          display: 'flex',
                          alignItems: 'center',
                          justifyContent: 'center',
                        }}
                      >
                        <FileText size={18} color="white" />
                      </div>
                      <div>
                        <div
                          style={{
                            fontSize: '14px',
                            fontWeight: 600,
                            color: '#2d3748',
                          }}
                        >
                          {file.name}
                        </div>
                        <div
                          style={{
                            fontSize: '11px',
                            color: '#718096',
                          }}
                        >
                          {file.type} • {(file.size / 1024).toFixed(1)} KB
                        </div>
                      </div>
                    </div>
                  </div>
                  
                  {file.content && (
                    file.type === 'image' && typeof file.content === 'string' && file.content.startsWith('data:image') ? (
                      <div style={{ padding: '8px', backgroundColor: 'white', borderRadius: '4px' }}>
                        <img
                          src={file.content}
                          alt={file.name}
                          style={{ maxWidth: '100%', maxHeight: '320px', borderRadius: '6px', display: 'block', margin: '0 auto' }}
                        />
                      </div>
                    ) : (
                      <div
                        style={{
                          padding: '8px',
                          backgroundColor: 'white',
                          borderRadius: '4px',
                          fontSize: '12px',
                          fontFamily: 'monospace',
                          color: '#4a5568',
                          maxHeight: '160px',
                          overflow: 'auto',
                        }}
                      >
                        {Array.isArray(file.content)
                          ? (file.content as unknown[]).slice(0, 8).map((row, i) => (
                              <div key={i}>
                                {Array.isArray(row)
                                  ? (row as unknown[]).map(v => (v ?? '') as string).join(' | ')
                                  : String(row)}
                              </div>
                            ))
                          : typeof file.content === 'string'
                            ? (file.type === 'json'
                                ? (() => { try { const obj = JSON.parse(file.content as string); return JSON.stringify(obj, null, 2).substring(0, 1000); } catch { return (file.content as string).substring(0, 1000); } })()
                                : (file.content as string).substring(0, 1000))
                            : 'Binary content'}
                      </div>
                    )
                  )}
                </div>
              ))}
            </div>
            
            <div
              style={{
                marginTop: '16px',
                paddingTop: '12px',
                borderTop: '1px solid #e2e8f0',
                fontSize: '12px',
                color: '#718096',
                textAlign: 'center',
              }}
            >
              {attachmentModalData.cellAttachments?.length ? 'Click cell attachments to focus in spreadsheet' : 'File contents shown above'}
            </div>
          </div>
        </div>
      )}
    </>
  );
};

export default AIChat;
