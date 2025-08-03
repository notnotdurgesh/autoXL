import React, { useState, useRef, useEffect, memo, useCallback } from 'react';
import {
  AlignLeft, AlignCenter, AlignRight,
  AlignStartVertical, AlignCenterVertical, AlignEndVertical,
  Bold, Italic, Underline, Strikethrough, Type, Palette, 
  Plus, Minus, ChevronDown, Undo2, Redo2,
  Copy, Clipboard, DollarSign, Percent, Link,
  Hash, Calendar, Clock
} from 'lucide-react';
import type { CellFormatting } from './ExcelSpreadsheet';

interface ExcelToolbarProps {
  onFormat: (formatting: Partial<CellFormatting>) => void;
  onUndo: () => void;
  onRedo: () => void;
  onCopy?: () => void;
  onPaste?: () => void;
  currentFormatting: CellFormatting;
  canUndo: boolean;
  canRedo: boolean;
  canPaste?: boolean;
  selectedCellValue?: string | number | null;
  onInsertLink: (url: string, text?: string) => void;
}

// Memoized constants for better performance
const FONT_FAMILIES = Object.freeze([
  'Arial', 'Calibri', 'Times New Roman', 'Georgia', 'Verdana', 
  'Tahoma', 'Trebuchet MS', 'Impact', 'Comic Sans MS', 'Courier New'
]);

const FONT_SIZES = Object.freeze([8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72]);

const COLORS = Object.freeze([
  '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
  '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#C0C0C0', '#808080',
  '#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080', '#0066CC', '#CCCCFF',
  '#FF6600', '#339966', '#3366FF', '#FF3399', '#99CC00', '#CC9900', '#6600CC', '#FF9900'
]);

const NUMBER_FORMATS = Object.freeze([
  { value: 'general', label: 'General', icon: Hash },
  { value: 'number', label: 'Number', icon: Hash },
  { value: 'currency', label: 'Currency', icon: DollarSign },
  { value: 'percentage', label: 'Percentage', icon: Percent },
  { value: 'date', label: 'Date', icon: Calendar },
  { value: 'time', label: 'Time', icon: Clock }
]);

const ExcelToolbar: React.FC<ExcelToolbarProps> = memo(({
  onFormat,
  onUndo,
  onRedo,
  onCopy,
  onPaste,
  currentFormatting,
  canUndo,
  canRedo,
  canPaste = true,
  selectedCellValue,
  onInsertLink
}) => {
  const [showFontFamilyDropdown, setShowFontFamilyDropdown] = useState(false);
  const [showFontSizeDropdown, setShowFontSizeDropdown] = useState(false);
  const [showTextColorPicker, setShowTextColorPicker] = useState(false);
  const [showBgColorPicker, setShowBgColorPicker] = useState(false);
  const [showNumberFormatDropdown, setShowNumberFormatDropdown] = useState(false);
  const [showLinkDialog, setShowLinkDialog] = useState(false);
  const [linkUrl, setLinkUrl] = useState('');
  const [linkText, setLinkText] = useState('');
  
  const fontFamilyRef = useRef<HTMLDivElement>(null);
  const fontSizeRef = useRef<HTMLDivElement>(null);
  const textColorRef = useRef<HTMLDivElement>(null);
  const bgColorRef = useRef<HTMLDivElement>(null);
  const numberFormatRef = useRef<HTMLDivElement>(null);

  const handleFontSizeChange = useCallback((delta: number) => {
    const currentSize = currentFormatting.fontSize || 12;
    const newSize = Math.max(6, Math.min(72, currentSize + delta));
    onFormat({ fontSize: newSize });
  }, [currentFormatting.fontSize, onFormat]);

  // Memoized event handlers for better performance
  const handleBoldClick = useCallback(() => {
    onFormat({ bold: !currentFormatting.bold });
  }, [onFormat, currentFormatting.bold]);

  const handleItalicClick = useCallback(() => {
    onFormat({ italic: !currentFormatting.italic });
  }, [onFormat, currentFormatting.italic]);

  const handleUnderlineClick = useCallback(() => {
    onFormat({ underline: !currentFormatting.underline });
  }, [onFormat, currentFormatting.underline]);

  const handleStrikethroughClick = useCallback(() => {
    onFormat({ strikethrough: !currentFormatting.strikethrough });
  }, [onFormat, currentFormatting.strikethrough]);

  const handleFontFamilyClick = useCallback(() => {
    setShowFontFamilyDropdown(!showFontFamilyDropdown);
  }, [showFontFamilyDropdown]);

  const handleFontSizeClick = useCallback(() => {
    setShowFontSizeDropdown(!showFontSizeDropdown);
  }, [showFontSizeDropdown]);

  const handleTextColorClick = useCallback(() => {
    setShowTextColorPicker(!showTextColorPicker);
  }, [showTextColorPicker]);

  const handleBgColorClick = useCallback(() => {
    setShowBgColorPicker(!showBgColorPicker);
  }, [showBgColorPicker]);

  const handleNumberFormatClick = useCallback(() => {
    setShowNumberFormatDropdown(!showNumberFormatDropdown);
  }, [showNumberFormatDropdown]);

  const handleDecimalIncrease = useCallback(() => {
    const currentDecimals = currentFormatting.decimalPlaces ?? 2;
    if (currentDecimals < 10) {
      onFormat({ decimalPlaces: currentDecimals + 1 });
    }
  }, [currentFormatting.decimalPlaces, onFormat]);

  const handleDecimalDecrease = useCallback(() => {
    const currentDecimals = currentFormatting.decimalPlaces ?? 2;
    if (currentDecimals > 0) {
      onFormat({ decimalPlaces: currentDecimals - 1 });
    }
  }, [currentFormatting.decimalPlaces, onFormat]);

  const handleCurrencyFormat = useCallback(() => {
    if (currentFormatting.numberFormat === 'currency') {
      // Remove currency formatting
      onFormat({ 
        numberFormat: 'general',
        decimalPlaces: undefined
      });
    } else {
      // Apply currency formatting
      onFormat({ 
        numberFormat: 'currency',
        decimalPlaces: currentFormatting.decimalPlaces ?? 2
      });
    }
  }, [onFormat, currentFormatting.numberFormat, currentFormatting.decimalPlaces]);

  const handlePercentageFormat = useCallback(() => {
    if (currentFormatting.numberFormat === 'percentage') {
      // Remove percentage formatting
      onFormat({ 
        numberFormat: 'general',
        decimalPlaces: undefined
      });
    } else {
      // Apply percentage formatting
      onFormat({ 
        numberFormat: 'percentage',
        decimalPlaces: currentFormatting.decimalPlaces ?? 2
      });
    }
  }, [onFormat, currentFormatting.numberFormat, currentFormatting.decimalPlaces]);

  const handleLinkClick = useCallback(() => {
    setLinkText(String(selectedCellValue || ''));
    setLinkUrl('');
    setShowLinkDialog(true);
  }, [selectedCellValue]);

  const normalizeUrl = useCallback((url: string): string => {
    if (!url) return '';
    
    // If URL already has protocol, return as is
    if (url.match(/^https?:\/\//i)) {
      return url;
    }
    
    // If it looks like an email, add mailto:
    if (url.includes('@') && !url.includes(' ')) {
      return `mailto:${url}`;
    }
    
    // Default to https:// for web URLs
    return `https://${url}`;
  }, []);

  const handleLinkSubmit = useCallback(() => {
    if (linkUrl.trim()) {
      const normalizedUrl = normalizeUrl(linkUrl.trim());
      onInsertLink(normalizedUrl, linkText.trim() || undefined);
      setShowLinkDialog(false);
      setLinkUrl('');
      setLinkText('');
    }
  }, [linkUrl, linkText, onInsertLink, normalizeUrl]);

  // Close dropdowns when clicking outside
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      const target = event.target as Element;
      if (!target.closest('.font-selector') && !target.closest('.dropdown-menu')) {
        setShowFontFamilyDropdown(false);
        setShowFontSizeDropdown(false);
      }
      if (!target.closest('[data-color-picker]') && !target.closest('.dropdown-menu')) {
        setShowTextColorPicker(false);
        setShowBgColorPicker(false);
      }
      if (!target.closest('[data-number-format]') && !target.closest('.dropdown-menu')) {
        setShowNumberFormatDropdown(false);
      }
      if (!target.closest('.link-dialog') && !target.closest('.link-dialog-overlay')) {
        setShowLinkDialog(false);
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const ToolbarButton: React.FC<{
    icon: React.ReactNode;
    title: string;
    onClick: () => void;
    active?: boolean;
    disabled?: boolean;
    className?: string;
  }> = memo(({ icon, title, onClick, active = false, disabled = false, className = '' }) => (
    <button
      className={`toolbar-btn ${active ? 'active' : ''} ${disabled ? 'disabled' : ''} ${className}`}
      onClick={onClick}
      disabled={disabled}
      title={title}
    >
      {icon}
    </button>
  ));

  const Dropdown: React.FC<{
    isOpen: boolean;
    onClose: () => void;
    children: React.ReactNode;
    reference: React.RefObject<HTMLDivElement | null>;
  }> = memo(({ isOpen, onClose, children, reference }) => {
    if (!isOpen) return null;
    
    return (
      <>
        <div className="dropdown-overlay" onClick={onClose} />
        <div className="dropdown-menu" style={{ 
          ...(() => {
            if (reference.current) {
              const rect = reference.current.getBoundingClientRect();
              return {
                top: rect.bottom + 5,
                left: rect.left,
                position: 'fixed' as const,
                zIndex: 1001
              };
            }
            return {
              top: 0,
              left: 0,
              position: 'fixed' as const,
              zIndex: 1001
            };
          })()
        }}>
          {children}
        </div>
      </>
    );
  });

  const ColorPalette: React.FC<{
    onColorSelect: (color: string) => void;
  }> = memo(({ onColorSelect }) => (
    <div className="color-palette">
      <div className="color-grid">
        {COLORS.map(color => (
          <div
            key={color}
            className="color-option"
            style={{ backgroundColor: color }}
            onClick={() => onColorSelect(color)}
            title={color}
          />
        ))}
      </div>
    </div>
  ));

  return (
    <>
      <style>{`
        .excel-toolbar {
          background: #f8f9fa;
          border-bottom: 1px solid #dee2e6;
          padding: 8px 12px;
          display: flex;
          flex-wrap: wrap;
          gap: 4px;
          align-items: center;
          position: relative;
          box-shadow: 0 2px 4px rgba(0,0,0,0.1);
          z-index: 100;
        }

        .toolbar-section {
          display: flex;
          align-items: center;
          gap: 3px;
          padding: 0 6px;
          border-right: 1px solid #dee2e6;
          height: 32px;
          min-width: max-content;
        }

        .toolbar-section:last-child {
          border-right: none;
        }

        .toolbar-section.compact {
          gap: 2px;
          padding: 0 4px;
        }

        .toolbar-btn {
          display: flex;
          align-items: center;
          justify-content: center;
          width: 28px;
          height: 28px;
          border: 1px solid transparent;
          background: transparent;
          border-radius: 4px;
          cursor: pointer;
          transition: all 0.2s ease;
          color: #495057;
          padding: 4px;
        }

        .toolbar-btn:hover:not(.disabled) {
          background: #e9ecef;
          border-color: #ced4da;
        }

        .toolbar-btn.active {
          background: #007bff;
          color: white;
          border-color: #0056b3;
        }

        .toolbar-btn.disabled {
          opacity: 0.5;
          cursor: not-allowed;
        }

        .font-selector {
          position: relative;
          display: flex;
          align-items: center;
          gap: 4px;
        }

        .font-family-btn, .font-size-btn {
          display: flex;
          align-items: center;
          gap: 4px;
          padding: 4px 8px;
          border: 1px solid #ced4da;
          background: white;
          border-radius: 4px;
          cursor: pointer;
          font-size: 12px;
          min-width: 80px;
          height: 28px;
          justify-content: space-between;
          color: #333;
        }

        .font-family-btn {
          min-width: 85px;
        }

        .font-size-btn {
          min-width: 50px;
        }

        .number-format-btn {
          min-width: 70px;
          display: flex;
          align-items: center;
          gap: 4px;
          padding: 4px 8px;
          border: 1px solid #ced4da;
          background: white;
          border-radius: 4px;
          cursor: pointer;
          font-size: 12px;
          height: 28px;
          justify-content: space-between;
          color: #333;
        }

        .font-family-btn:hover, .font-size-btn:hover, .number-format-btn:hover {
          border-color: #007bff;
          background: #f8f9fa;
        }

        .font-size-controls {
          display: flex;
          gap: 2px;
        }

        .size-control-btn {
          width: 20px;
          height: 28px;
          display: flex;
          align-items: center;
          justify-content: center;
          border: 1px solid #ced4da;
          background: white;
          cursor: pointer;
          font-size: 12px;
          font-weight: bold;
          color: #333;
        }

        .size-control-btn:first-child {
          border-radius: 4px 0 0 4px;
        }

        .size-control-btn:last-child {
          border-radius: 0 4px 4px 0;
          border-left: none;
        }

        .size-control-btn:hover {
          background: #e9ecef;
        }

        .dropdown-overlay {
          position: fixed;
          top: 0;
          left: 0;
          right: 0;
          bottom: 0;
          z-index: 999;
          background: transparent;
        }

        .dropdown-menu {
          position: fixed;
          background: white;
          border: 1px solid #ced4da;
          border-radius: 4px;
          box-shadow: 0 4px 12px rgba(0,0,0,0.15);
          z-index: 1001;
          max-height: 200px;
          overflow-y: auto;
          min-width: 120px;
        }

        .dropdown-item {
          padding: 8px 12px;
          cursor: pointer;
          font-size: 12px;
          border-bottom: 1px solid #f8f9fa;
          color: #333;
          background: white;
          transition: all 0.15s ease;
        }

        .dropdown-item:hover {
          background: #e9ecef;
          color: #000;
        }

        .dropdown-item.selected {
          background: #007bff;
          color: white;
        }

        .dropdown-item.selected:hover {
          background: #0056b3;
          color: white;
        }

        .dropdown-item:last-child {
          border-bottom: none;
        }

        .color-palette {
          padding: 8px;
        }

        .color-grid {
          display: grid;
          grid-template-columns: repeat(8, 1fr);
          gap: 2px;
          width: 160px;
        }

        .color-option {
          width: 18px;
          height: 18px;
          border: 1px solid #dee2e6;
          cursor: pointer;
          border-radius: 2px;
        }

        .color-option:hover {
          border-color: #007bff;
          box-shadow: 0 0 0 1px #007bff;
        }

        .color-preview {
          width: 16px;
          height: 3px;
          border: 1px solid #333;
          margin-top: 2px;
          border-radius: 1px;
          display: block;
        }

        .alignment-grid {
          display: grid;
          grid-template-columns: repeat(3, 1fr);
          gap: 2px;
        }

        .toolbar-label {
          font-size: 11px;
          color: #6c757d;
          margin-right: 4px;
          white-space: nowrap;
        }

        .decimal-controls {
          display: flex;
          gap: 0px;
        }

        .decimal-btn {
          width: 24px;
          height: 28px;
          display: flex;
          align-items: center;
          justify-content: center;
          border: 1px solid #ced4da;
          background: white;
          cursor: pointer;
          font-size: 11px;
          color: #333;
          transition: all 0.15s ease;
        }

        .decimal-btn:first-child {
          border-radius: 4px 0 0 4px;
        }

        .decimal-btn:last-child {
          border-radius: 0 4px 4px 0;
          border-left: none;
        }

        .decimal-btn:hover {
          background: #e9ecef;
        }

        .decimal-btn:disabled {
          opacity: 0.5;
          cursor: not-allowed;
          background: #f8f9fa;
        }

        .decimal-btn:disabled:hover {
          background: #f8f9fa;
        }

        .link-dialog-overlay {
          position: fixed;
          top: 0;
          left: 0;
          right: 0;
          bottom: 0;
          z-index: 1000;
          background: rgba(0,0,0,0.4);
          backdrop-filter: blur(4px);
          -webkit-backdrop-filter: blur(4px);
          display: flex;
          align-items: center;
          justify-content: center;
          animation: overlayFadeIn 0.2s ease-out;
        }

        @keyframes overlayFadeIn {
          from {
            opacity: 0;
            backdrop-filter: blur(0px);
            -webkit-backdrop-filter: blur(0px);
          }
          to {
            opacity: 1;
            backdrop-filter: blur(4px);
            -webkit-backdrop-filter: blur(4px);
          }
        }

        .link-dialog {
          background: white;
          padding: 24px;
          border-radius: 12px;
          box-shadow: 0 20px 60px rgba(0,0,0,0.12), 0 8px 25px rgba(0,0,0,0.08);
          min-width: 450px;
          max-width: 550px;
          border: 1px solid rgba(0,0,0,0.05);
          position: relative;
          animation: modalSlideIn 0.25s ease-out;
        }

        @keyframes modalSlideIn {
          from {
            opacity: 0;
            transform: scale(0.95) translateY(-20px);
          }
          to {
            opacity: 1;
            transform: scale(1) translateY(0);
          }
        }

        .link-dialog h3 {
          margin: 0 0 20px 0;
          font-size: 18px;
          font-weight: 600;
          color: #2d3748;
          display: flex;
          align-items: center;
          gap: 8px;
        }

        .link-dialog h3::before {
          content: "🔗";
          font-size: 16px;
        }

        .link-dialog-field {
          margin-bottom: 16px;
        }

        .link-dialog-field label {
          display: block;
          margin-bottom: 6px;
          font-size: 13px;
          font-weight: 600;
          color: #4a5568;
          letter-spacing: 0.025em;
        }

        .link-dialog-field input {
          width: 100%;
          padding: 12px 16px;
          border: 2px solid #e2e8f0;
          border-radius: 8px;
          font-size: 14px;
          box-sizing: border-box;
          transition: all 0.2s ease;
          background: #fafafa;
        }

        /* Make the text black for inputs in the link dialog */
        .link-dialog-field input.link-dialog-text-input {
          color: #2d3748;
          font-weight: 500;
        }

        .link-dialog-field input:focus {
          outline: none;
          border-color: #007bff;
          box-shadow: 0 0 0 3px rgba(0,123,255,0.1);
          background: white;
          transform: translateY(-1px);
        }

        .link-dialog-field input:hover:not(:focus) {
          border-color: #cbd5e0;
          background: white;
        }

        .link-dialog-help {
          font-size: 11px;
          color: #718096;
          margin-top: 6px;
          display: flex;
          align-items: center;
          gap: 4px;
        }

        .link-dialog-help::before {
          content: "💡";
          font-size: 10px;
        }

        .link-dialog-buttons {
          display: flex;
          gap: 12px;
          justify-content: flex-end;
          margin-top: 24px;
          padding-top: 16px;
          border-top: 1px solid #e2e8f0;
        }

        .link-dialog-btn {
          padding: 10px 20px;
          border: 2px solid transparent;
          border-radius: 8px;
          cursor: pointer;
          font-size: 13px;
          font-weight: 600;
          transition: all 0.2s ease;
          min-width: 90px;
          display: flex;
          align-items: center;
          justify-content: center;
          gap: 6px;
        }

        .link-dialog-btn.primary {
          background: linear-gradient(135deg, #007bff 0%, #0056b3 100%);
          color: white;
          border-color: #007bff;
          box-shadow: 0 2px 8px rgba(0,123,255,0.2);
        }

        .link-dialog-btn.primary:hover:not(:disabled) {
          background: linear-gradient(135deg, #0056b3 0%, #004085 100%);
          transform: translateY(-1px);
          box-shadow: 0 4px 12px rgba(0,123,255,0.3);
        }

        .link-dialog-btn.primary:disabled {
          background: #e2e8f0;
          color: #a0aec0;
          cursor: not-allowed;
          transform: none;
          box-shadow: none;
        }

        .link-dialog-btn.secondary {
          background: white;
          color: #4a5568;
          border-color: #e2e8f0;
        }

        .link-dialog-btn.secondary:hover {
          background: #f7fafc;
          border-color: #cbd5e0;
          transform: translateY(-1px);
        }

        @media (max-width: 1200px) {
          .excel-toolbar {
            padding: 6px 8px;
            gap: 2px;
          }
          
          .toolbar-section {
            padding: 0 4px;
            gap: 2px;
          }
          
          .font-family-btn {
            min-width: 70px;
          }
          
          .font-size-btn {
            min-width: 45px;
          }

          .number-format-btn {
            min-width: 60px;
          }
        }

        @media (max-width: 768px) {
          .excel-toolbar {
            padding: 4px 6px;
            gap: 1px;
          }
          
          .toolbar-section {
            padding: 0 3px;
            gap: 1px;
          }
          
          .font-family-btn {
            min-width: 60px;
          }
          
          .font-size-btn {
            min-width: 40px;
          }

          .number-format-btn {
            min-width: 50px;
          }

          .link-dialog {
            min-width: 320px;
            margin: 20px;
          }
        }
      `}</style>

      <div className="excel-toolbar">


        {/* Font Section */}
        <div className="toolbar-section">
          <div className="font-selector">
            <div ref={fontFamilyRef} className="font-family-btn" onClick={handleFontFamilyClick}>
              <span>{currentFormatting.fontFamily || 'Arial'}</span>
              <ChevronDown size={10} />
            </div>
            
            <div ref={fontSizeRef} className="font-size-btn" onClick={handleFontSizeClick}>
              <span>{currentFormatting.fontSize || 12}</span>
              <ChevronDown size={10} />
            </div>

            <div className="font-size-controls">
              <div className="size-control-btn" onClick={() => handleFontSizeChange(-1)} title="Decrease font size">
                <Minus size={10} />
              </div>
              <div className="size-control-btn" onClick={() => handleFontSizeChange(1)} title="Increase font size">
                <Plus size={10} />
              </div>
            </div>
          </div>
        </div>

        {/* Text Formatting */}
        <div className="toolbar-section compact">
          <ToolbarButton
            icon={<Bold size={14} />}
            title="Bold (Ctrl+B)"
            onClick={handleBoldClick}
            active={currentFormatting.bold}
          />
          <ToolbarButton
            icon={<Italic size={14} />}
            title="Italic (Ctrl+I)"
            onClick={handleItalicClick}
            active={currentFormatting.italic}
          />
          <ToolbarButton
            icon={<Underline size={14} />}
            title="Underline (Ctrl+U)"
            onClick={handleUnderlineClick}
            active={currentFormatting.underline}
          />
          <ToolbarButton
            icon={<Strikethrough size={14} />}
            title="Strikethrough"
            onClick={handleStrikethroughClick}
            active={currentFormatting.strikethrough}
          />
        </div>

        {/* Colors Section */}
        <div className="toolbar-section">
          <div ref={textColorRef} data-color-picker style={{ position: 'relative' }}>
            <ToolbarButton
              icon={
                <div>
                  <Type size={14} />
                  <div className="color-preview" style={{ backgroundColor: currentFormatting.color || '#000000' }} />
                </div>
              }
              title="Text Color"
              onClick={handleTextColorClick}
            />
          </div>
          
          <div ref={bgColorRef} data-color-picker style={{ position: 'relative' }}>
            <ToolbarButton
              icon={
                <div>
                  <Palette size={14} />
                  <div className="color-preview" style={{ backgroundColor: currentFormatting.backgroundColor || '#FFFFFF' }} />
                </div>
              }
              title="Background Color"
              onClick={handleBgColorClick}
            />
          </div>
        </div>

        {/* Number Formatting */}
        <div className="toolbar-section">
          <div ref={numberFormatRef} data-number-format className="number-format-btn" onClick={handleNumberFormatClick}>
            <span>{NUMBER_FORMATS.find(f => f.value === currentFormatting.numberFormat)?.label || 'General'}</span>
            <ChevronDown size={10} />
          </div>
          
          <ToolbarButton
            icon={<DollarSign size={14} />}
            title="Format as Currency"
            onClick={handleCurrencyFormat}
            active={currentFormatting.numberFormat === 'currency'}
          />
          
          <ToolbarButton
            icon={<Percent size={14} />}
            title="Format as Percentage"
            onClick={handlePercentageFormat}
            active={currentFormatting.numberFormat === 'percentage'}
          />

          <div className="decimal-controls">
            <div 
              className="decimal-btn" 
              onClick={handleDecimalDecrease} 
              title={`Decrease decimal places (${currentFormatting.decimalPlaces ?? 2})`}
              style={{ 
                opacity: (currentFormatting.decimalPlaces ?? 2) <= 0 ? 0.5 : 1,
                cursor: (currentFormatting.decimalPlaces ?? 2) <= 0 ? 'not-allowed' : 'pointer'
              }}
            >
              .0
            </div>
            <div 
              className="decimal-btn" 
              onClick={handleDecimalIncrease} 
              title={`Increase decimal places (${currentFormatting.decimalPlaces ?? 2})`}
              style={{ 
                opacity: (currentFormatting.decimalPlaces ?? 2) >= 10 ? 0.5 : 1,
                cursor: (currentFormatting.decimalPlaces ?? 2) >= 10 ? 'not-allowed' : 'pointer'
              }}
            >
              .00
            </div>
          </div>
        </div>

        {/* Alignment Section */}
        <div className="toolbar-section compact">
          <ToolbarButton
            icon={<AlignLeft size={14} />}
            title="Align Left"
            onClick={() => onFormat({ textAlign: 'left' })}
            active={currentFormatting.textAlign === 'left'}
          />
          <ToolbarButton
            icon={<AlignCenter size={14} />}
            title="Align Center"
            onClick={() => onFormat({ textAlign: 'center' })}
            active={currentFormatting.textAlign === 'center'}
          />
          <ToolbarButton
            icon={<AlignRight size={14} />}
            title="Align Right"
            onClick={() => onFormat({ textAlign: 'right' })}
            active={currentFormatting.textAlign === 'right'}
          />
          
          <ToolbarButton
            icon={<AlignStartVertical size={14} />}
            title="Align Top"
            onClick={() => onFormat({ verticalAlign: 'top' })}
            active={currentFormatting.verticalAlign === 'top'}
          />
          <ToolbarButton
            icon={<AlignCenterVertical size={14} />}
            title="Align Middle"
            onClick={() => onFormat({ verticalAlign: 'middle' })}
            active={currentFormatting.verticalAlign === 'middle'}
          />
          <ToolbarButton
            icon={<AlignEndVertical size={14} />}
            title="Align Bottom"
            onClick={() => onFormat({ verticalAlign: 'bottom' })}
            active={currentFormatting.verticalAlign === 'bottom'}
          />
        </div>

        {/* Links & Actions */}
        <div className="toolbar-section compact">
          <ToolbarButton
            icon={<Link size={14} />}
            title="Insert Link"
            onClick={handleLinkClick}
          />
        </div>

        {/* Clipboard Section */}
        <div className="toolbar-section compact">
          <ToolbarButton
            icon={<Copy size={14} />}
            title="Copy (Ctrl+C)"
            onClick={onCopy || (() => {})}
            disabled={!onCopy}
          />
          <ToolbarButton
            icon={<Clipboard size={14} />}
            title="Paste (Ctrl+V)"
            onClick={onPaste || (() => {})}
            disabled={!canPaste || !onPaste}
          />
        </div>

        {/* Undo/Redo Section */}
        <div className="toolbar-section compact">
          <ToolbarButton
            icon={<Undo2 size={14} />}
            title="Undo (Ctrl+Z)"
            onClick={onUndo}
            disabled={!canUndo}
          />
          <ToolbarButton
            icon={<Redo2 size={14} />}
            title="Redo (Ctrl+Y)"
            onClick={onRedo}
            disabled={!canRedo}
          />
        </div>
      </div>

      {/* Dropdowns */}
      <Dropdown
        isOpen={showFontFamilyDropdown}
        onClose={() => setShowFontFamilyDropdown(false)}
        reference={fontFamilyRef}
      >
        {FONT_FAMILIES.map(font => (
          <div
            key={font}
            className={`dropdown-item ${(currentFormatting.fontFamily || 'Arial') === font ? 'selected' : ''}`}
            style={{ fontFamily: font }}
            onClick={() => {
              if ((currentFormatting.fontFamily || 'Arial') === font) {
                // Reset to default font
                onFormat({ fontFamily: 'Arial' });
              } else {
                // Apply new font
                onFormat({ fontFamily: font });
              }
              setShowFontFamilyDropdown(false);
            }}
          >
            {font}
          </div>
        ))}
      </Dropdown>

      <Dropdown
        isOpen={showFontSizeDropdown}
        onClose={() => setShowFontSizeDropdown(false)}
        reference={fontSizeRef}
      >
        {FONT_SIZES.map(size => (
          <div
            key={size}
            className={`dropdown-item ${(currentFormatting.fontSize || 12) === size ? 'selected' : ''}`}
            onClick={() => {
              if ((currentFormatting.fontSize || 12) === size) {
                // Reset to default size
                onFormat({ fontSize: 12 });
              } else {
                // Apply new size
                onFormat({ fontSize: size });
              }
              setShowFontSizeDropdown(false);
            }}
          >
            {size}
          </div>
        ))}
      </Dropdown>

      <Dropdown
        isOpen={showTextColorPicker}
        onClose={() => setShowTextColorPicker(false)}
        reference={textColorRef}
      >
        <ColorPalette
          onColorSelect={(color) => {
            onFormat({ color });
            setShowTextColorPicker(false);
          }}
        />
      </Dropdown>

      <Dropdown
        isOpen={showBgColorPicker}
        onClose={() => setShowBgColorPicker(false)}
        reference={bgColorRef}
      >
        <ColorPalette
          onColorSelect={(color) => {
            onFormat({ backgroundColor: color });
            setShowBgColorPicker(false);
          }}
        />
      </Dropdown>

      {/* Number Format Dropdown */}
      <Dropdown
        isOpen={showNumberFormatDropdown}
        onClose={() => setShowNumberFormatDropdown(false)}
        reference={numberFormatRef}
      >
        {NUMBER_FORMATS.map(format => {
          const IconComponent = format.icon;
          const isSelected = (currentFormatting.numberFormat || 'general') === format.value;
          return (
            <div
              key={format.value}
              className={`dropdown-item ${isSelected ? 'selected' : ''}`}
              onClick={() => {
                if (isSelected && format.value !== 'general') {
                  // Remove formatting by setting to general
                  onFormat({ 
                    numberFormat: 'general',
                    decimalPlaces: undefined
                  });
                } else {
                  // Apply formatting
                  onFormat({ 
                    numberFormat: format.value as CellFormatting['numberFormat'],
                    decimalPlaces: format.value === 'general' ? undefined : (currentFormatting.decimalPlaces ?? 2)
                  });
                }
                setShowNumberFormatDropdown(false);
              }}
            >
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <IconComponent size={14} />
                <span>{format.label}</span>
                {isSelected && (
                  <div style={{ marginLeft: 'auto', fontSize: '12px', opacity: 0.8 }}>
                    ✓
                  </div>
                )}
              </div>
            </div>
          );
        })}
      </Dropdown>

      {/* Link Dialog */}
      {showLinkDialog && (
        <div className="link-dialog-overlay">
          <div className="link-dialog">
            <h3>Insert Link</h3>
            <div className="link-dialog-field">
              <label>Text to display:</label>
              <input
                type="text"
                className="link-dialog-text-input"
                value={linkText}
                onChange={(e) => setLinkText(e.target.value)}
                placeholder="Text to display"
                autoFocus
              />
            </div>
            <div className="link-dialog-field">
              <label>URL:</label>
              <input
                type="text"
                value={linkUrl}
                className="link-dialog-text-input"
                onChange={(e) => setLinkUrl(e.target.value)}
                placeholder="www.google.com or https://example.com"
                onKeyDown={(e) => {
                  if (e.key === 'Enter') {
                    handleLinkSubmit();
                  } else if (e.key === 'Escape') {
                    setShowLinkDialog(false);
                  }
                }}
              />
            </div>
            <div className="link-dialog-buttons">
              <button 
                className="link-dialog-btn secondary"
                onClick={() => setShowLinkDialog(false)}
              >
                ✕ Cancel
              </button>
              <button 
                className="link-dialog-btn primary"
                onClick={handleLinkSubmit}
                disabled={!linkUrl.trim()}
              >
                ✓ Insert Link
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
});

// Set display name for better debugging
ExcelToolbar.displayName = 'ExcelToolbar';

export default ExcelToolbar;