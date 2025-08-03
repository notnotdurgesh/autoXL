import React, { useState, useRef, useEffect, memo, useCallback } from 'react';
import {
  AlignLeft, AlignCenter, AlignRight,
  AlignStartVertical, AlignCenterVertical, AlignEndVertical,
  Bold, Italic, Underline, Type, Palette, 
  Plus, Minus, ChevronDown, Undo2, Redo2,
  Copy, Clipboard
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

const ExcelToolbar: React.FC<ExcelToolbarProps> = memo(({
  onFormat,
  onUndo,
  onRedo,
  onCopy,
  onPaste,
  currentFormatting,
  canUndo,
  canRedo,
  canPaste = true
}) => {
  const [showFontFamilyDropdown, setShowFontFamilyDropdown] = useState(false);
  const [showFontSizeDropdown, setShowFontSizeDropdown] = useState(false);
  const [showTextColorPicker, setShowTextColorPicker] = useState(false);
  const [showBgColorPicker, setShowBgColorPicker] = useState(false);
  
  const fontFamilyRef = useRef<HTMLDivElement>(null);
  const fontSizeRef = useRef<HTMLDivElement>(null);
  const textColorRef = useRef<HTMLDivElement>(null);
  const bgColorRef = useRef<HTMLDivElement>(null);

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
          gap: 4px;
          padding: 0 8px;
          border-right: 1px solid #dee2e6;
          height: 32px;
        }

        .toolbar-section:last-child {
          border-right: none;
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
          min-width: 100px;
        }

        .font-size-btn {
          min-width: 60px;
        }

        .font-family-btn:hover, .font-size-btn:hover {
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
        }

        .dropdown-item:hover {
          background: #e9ecef;
          color: #000;
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

        @media (max-width: 768px) {
          .excel-toolbar {
            padding: 6px 8px;
            gap: 2px;
          }
          
          .toolbar-section {
            padding: 0 4px;
            gap: 2px;
          }
          
          .font-family-btn {
            min-width: 80px;
          }
          
          .font-size-btn {
            min-width: 50px;
          }
        }
      `}</style>

      <div className="excel-toolbar">


        {/* Font Section */}
        <div className="toolbar-section">
          {/* <span className="toolbar-label">Font</span> */}
          <div className="font-selector">
            <div ref={fontFamilyRef} className="font-family-btn" onClick={handleFontFamilyClick}>
              <span>{currentFormatting.fontFamily || 'Arial'}</span>
              <ChevronDown size={12} />
            </div>
            
            <div ref={fontSizeRef} className="font-size-btn" onClick={handleFontSizeClick}>
              <span>{currentFormatting.fontSize || 12}</span>
              <ChevronDown size={12} />
            </div>

            <div className="font-size-controls">
              <div className="size-control-btn" onClick={() => handleFontSizeChange(-1)} title="Decrease font size">
                <Minus size={12} />
              </div>
              <div className="size-control-btn" onClick={() => handleFontSizeChange(1)} title="Increase font size">
                <Plus size={12} />
              </div>
            </div>
          </div>

          <ToolbarButton
            icon={<Bold size={16} />}
            title="Bold"
            onClick={handleBoldClick}
            active={currentFormatting.bold}
          />
          <ToolbarButton
            icon={<Italic size={16} />}
            title="Italic"
            onClick={handleItalicClick}
            active={currentFormatting.italic}
          />
          <ToolbarButton
            icon={<Underline size={16} />}
            title="Underline"
            onClick={handleUnderlineClick}
            active={currentFormatting.underline}
          />
        </div>

        {/* Colors Section */}
        <div className="toolbar-section">
          {/* <span className="toolbar-label">Colors</span> */}
          <div ref={textColorRef} data-color-picker style={{ position: 'relative' }}>
            <ToolbarButton
              icon={
                <div>
                  <Type size={16} />
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
                  <Palette size={16} />
                  <div className="color-preview" style={{ backgroundColor: currentFormatting.backgroundColor || '#FFFFFF' }} />
                </div>
              }
              title="Background Color"
              onClick={handleBgColorClick}
            />
          </div>
        </div>

        {/* Alignment Section */}
        <div className="toolbar-section">
          {/* <span className="toolbar-label">Alignment</span> */}
          <ToolbarButton
            icon={<AlignLeft size={16} />}
            title="Align Left"
            onClick={() => onFormat({ textAlign: 'left' })}
            active={currentFormatting.textAlign === 'left'}
          />
          <ToolbarButton
            icon={<AlignCenter size={16} />}
            title="Align Center"
            onClick={() => onFormat({ textAlign: 'center' })}
            active={currentFormatting.textAlign === 'center'}
          />
          <ToolbarButton
            icon={<AlignRight size={16} />}
            title="Align Right"
            onClick={() => onFormat({ textAlign: 'right' })}
            active={currentFormatting.textAlign === 'right'}
          />
          
          <ToolbarButton
            icon={<AlignStartVertical size={16} />}
            title="Align Top"
            onClick={() => onFormat({ verticalAlign: 'top' })}
            active={currentFormatting.verticalAlign === 'top'}
          />
          <ToolbarButton
            icon={<AlignCenterVertical size={16} />}
            title="Align Middle"
            onClick={() => onFormat({ verticalAlign: 'middle' })}
            active={currentFormatting.verticalAlign === 'middle'}
          />
          <ToolbarButton
            icon={<AlignEndVertical size={16} />}
            title="Align Bottom"
            onClick={() => onFormat({ verticalAlign: 'bottom' })}
            active={currentFormatting.verticalAlign === 'bottom'}
          />
        </div>

        {/* Clipboard Section */}
        <div className="toolbar-section">
          <ToolbarButton
            icon={<Copy size={16} />}
            title="Copy (Ctrl+C)"
            onClick={onCopy || (() => {})}
            disabled={!onCopy}
          />
          <ToolbarButton
            icon={<Clipboard size={16} />}
            title="Paste (Ctrl+V)"
            onClick={onPaste || (() => {})}
            disabled={!canPaste || !onPaste}
          />
        </div>

        {/* Undo/Redo Section */}
        <div className="toolbar-section">
          {/* <span className="toolbar-label">History</span> */}
          <ToolbarButton
            icon={<Undo2 size={16} />}
            title="Undo (Ctrl+Z)"
            onClick={onUndo}
            disabled={!canUndo}
          />
          <ToolbarButton
            icon={<Redo2 size={16} />}
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
            className="dropdown-item"
            style={{ fontFamily: font }}
            onClick={() => {
              onFormat({ fontFamily: font });
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
            className="dropdown-item"
            onClick={() => {
              onFormat({ fontSize: size });
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
    </>
  );
});

// Set display name for better debugging
ExcelToolbar.displayName = 'ExcelToolbar';

export default ExcelToolbar;