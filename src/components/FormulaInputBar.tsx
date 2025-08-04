import React, { useRef, useEffect, useCallback } from 'react';

interface FormulaInputBarProps {
  value: string;
  onChange: (value: string) => void;
  onKeyDown?: (e: React.KeyboardEvent<HTMLInputElement>) => void;
  onFocus?: () => void;
  onBlur?: () => void;
  isActive: boolean;
  selectedCellAddress: string;
  className?: string;
}

const FormulaInputBar: React.FC<FormulaInputBarProps> = ({
  value,
  onChange,
  onKeyDown,
  onFocus,
  onBlur,
  isActive,
  selectedCellAddress,
  className = ''
}) => {
  const inputRef = useRef<HTMLInputElement>(null);

  // Focus the input when it becomes active - but preserve cursor position
  useEffect(() => {
    if (isActive && inputRef.current && document.activeElement !== inputRef.current) {
      inputRef.current.focus();
      // Don't force cursor position - let user click where they want
    }
  }, [isActive]);

  // Handle input changes with perfect sync
  const handleChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    onChange(e.target.value);
  }, [onChange]);

  // Handle key events
  const handleKeyDown = useCallback((e: React.KeyboardEvent<HTMLInputElement>) => {
    // Allow default behavior for navigation and editing
    if (onKeyDown) {
      onKeyDown(e);
    }
  }, [onKeyDown]);

  // Handle click to allow natural cursor positioning
  const handleClick = useCallback((e: React.MouseEvent<HTMLInputElement>) => {
    e.stopPropagation(); // Prevent any parent handlers from interfering
    // Let the browser handle natural cursor positioning
  }, []);

  return (
    <div className={`formula-input-bar ${className}`}>
      <div className="formula-bar-container">
        {/* Cell address indicator */}
        <div className="cell-address-display">
          {selectedCellAddress}
        </div>
        
        {/* Formula/value input */}
        <input
          ref={inputRef}
          type="text"
          className="formula-input"
          value={value}
          onChange={handleChange}
          onKeyDown={handleKeyDown}
          onClick={handleClick}
          onFocus={onFocus}
          onBlur={onBlur}
          placeholder="Select a cell to edit"
          spellCheck={true}
          autoComplete="true"
        />
      </div>
      
      <style>{`
        .formula-input-bar {
          display: flex;
          align-items: center;
          background: white;
          border: 1px solid #d4d4d4;
          border-top: none;
          height: 24px;
          flex-shrink: 0;
          z-index: 10;
          box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05);
        }
        
        .formula-bar-container {
          display: flex;
          align-items: center;
          width: 100%;
          height: 100%;
        }
        
        .cell-address-display {
          background: linear-gradient(180deg, #f8f9fa 0%, #e9ecef 100%);
          border-right: 1px solid #c6c6c6;
          padding: 0 12px;
          height: 100%;
          display: flex;
          align-items: center;
          font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
          font-size: 11px;
          font-weight: 600;
          color: #333;
          min-width: 70px;
          justify-content: center;
          white-space: nowrap;
          cursor: default;
          user-select: none;
          box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.5);
        }
        
        .formula-input {
          flex: 1;
          height: 100%;
          border: none;
          outline: none;
          padding: 0 8px;
          font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
          font-size: 11px;
          background: white;
          color: #333;
          line-height: 1;
          resize: none;
          transition: background-color 0.15s ease;
        }
        
        .formula-input:focus {
          background: white;
          box-shadow: inset 0 0 0 2px #217346;
        }
        
        .formula-input:hover:not(:focus) {
          background: #fafbfc;
        }
        
        .formula-input::placeholder {
          color: #999;
          font-style: italic;
        }
        
        .formula-input::selection {
          background: #217346;
          color: white;
        }
        
        /* When the input bar is active (editing) */
        .formula-input-bar.active {
          border-color: #217346;
          box-shadow: 0 0 0 1px #217346, 0 1px 3px rgba(33, 115, 70, 0.15);
        }
        
        .formula-input-bar.active .cell-address-display {
          background: linear-gradient(180deg, #e8f5e8 0%, #d1f2d1 100%);
          color: #217346;
          border-right-color: #217346;
          font-weight: 700;
        }
        
        /* Excel-like enhancement: subtle animation on focus */
        .formula-input-bar {
          transition: border-color 0.15s ease, box-shadow 0.15s ease;
        }
        
        .cell-address-display {
          transition: background 0.15s ease, color 0.15s ease, border-color 0.15s ease;
        }
        
        /* Accessibility improvements */
        .formula-input:focus-visible {
          outline: 2px solid #217346;
          outline-offset: -2px;
        }
        
        /* High contrast mode support */
        @media (prefers-contrast: high) {
          .formula-input-bar {
            border: 2px solid;
          }
          
          .cell-address-display {
            border-right: 2px solid;
          }
        }
        
        /* Reduced motion support */
        @media (prefers-reduced-motion: reduce) {
          .formula-input-bar,
          .cell-address-display,
          .formula-input {
            transition: none;
          }
        }
      `}</style>
    </div>
  );
};

export default FormulaInputBar;