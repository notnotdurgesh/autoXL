import React, { memo, useRef } from 'react';
import {
  // ChevronLeft,
  // ChevronRight,
  Save,
  FileDown,
  FileText,
  Image as ImageIcon,
  Upload
} from 'lucide-react';

interface ActionSidebarProps {
  isOpen: boolean;
  onSave?: () => void;
  onSaveAsExcel?: () => void;
  onExportPDF?: () => void;
  onExportPNG?: () => void;
  onImportExcelFile?: (file: File) => void;
  versionHistory?: Array<{ id: string; timestamp: number; description: string }>;
  onPreviewVersion?: (id: string) => void;
  onRestoreVersion?: (id: string) => void;
}

const ActionSidebar: React.FC<ActionSidebarProps> = memo(({ 
  isOpen,
  onSave,
  onSaveAsExcel,
  onExportPDF,
  onExportPNG,
  onImportExcelFile,
  versionHistory = [],
  onPreviewVersion,
  onRestoreVersion
}) => {
  const fileInputRef = useRef<HTMLInputElement>(null);

  const triggerFilePick = () => {
    fileInputRef.current?.click();
  };

  return (
    <>
      <style>{`
        .action-sidebar {
          position: absolute;
          top: 0;
          right: 0;
          height: 100%;
          width: ${isOpen ? 240 : 44}px;
          transition: width 0.2s ease;
          background: #f8f9fa;
          border-left: 1px solid #dee2e6;
          box-shadow: ${isOpen ? '-4px 0 12px rgba(0,0,0,0.08)' : 'none'};
          z-index: ${isOpen ? 2000 : 200};
          display: flex;
          flex-direction: column;
        }
        .action-sidebar-header {
          height: 36px;
          display: flex;
          align-items: center;
          justify-content: ${isOpen ? 'flex-start' : 'center'};
          padding: 0 8px;
          border-bottom: 1px solid #dee2e6;
          background: #ffffff;
        }
        .action-sidebar-title {
          font-size: 12px;
          color: #495057;
          font-weight: 600;
        }
        
        .action-sidebar-content {
          padding: 8px;
          overflow: auto;
          display: ${isOpen ? 'block' : 'none'};
        }
        .action-btn {
          width: 100%;
          display: flex;
          align-items: center;
          gap: 8px;
          padding: 8px 10px;
          border: 1px solid #dee2e6;
          background: white;
          border-radius: 6px;
          font-size: 12px;
          color: #333;
          cursor: pointer;
          margin-bottom: 8px;
        }
        .action-btn:hover { background: #f1f3f5; }
        .action-icon { font-size: 12px; }
        .action-desc { color: #6c757d; font-size: 11px; }
        .action-group-label { font-size: 11px; color: #6c757d; margin: 10px 0 6px; }
        .history-list { display: flex; flex-direction: column; gap: 8px; }
        .history-item { display: flex; align-items: center; justify-content: space-between; gap: 8px; padding: 10px; border: 1px solid #e9ecef; background: #ffffff; border-radius: 8px; }
        .history-main { display: flex; flex-direction: column; gap: 2px; min-width: 0; }
        .history-desc { font-size: 12px; color: #2d3748; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; font-weight: 600; }
        .history-time { font-size: 10px; color: #6c757d; }
        .history-actions { display: flex; gap: 6px; white-space: nowrap; }
        .history-btn { border: 1px solid #ced4da; background: #ffffff; padding: 6px 10px; border-radius: 6px; font-size: 11px; cursor: pointer; color: #333; }
        .history-btn:hover { background: #f8f9fa; }
        .history-btn.primary { border-color: #0d6efd; background: #0d6efd; color: #ffffff; }
        .history-btn.primary:hover { background: #0b5ed7; }
      `}</style>

      <aside className="action-sidebar">
        <div className="action-sidebar-header">
          {isOpen && <div className="action-sidebar-title">Actions</div>}
        </div>
        <div className="action-sidebar-content">
          <div className="action-group-label">File</div>
          <button className="action-btn" onClick={onSave} disabled={!onSave} title="Save JSON">
            <Save size={14} className="action-icon" />
            <div>
              <div>Save</div>
              <div className="action-desc">Download JSON snapshot</div>
            </div>
          </button>
          <button className="action-btn" onClick={onSaveAsExcel} disabled={!onSaveAsExcel} title="Save as .xlsx">
            <FileDown size={14} className="action-icon" />
            <div>
              <div>Save as Excel (.xlsx)</div>
              <div className="action-desc">Export grid to workbook</div>
            </div>
          </button>
          <button className="action-btn" onClick={onExportPDF} disabled={!onExportPDF} title="Export PDF">
            <FileText size={14} className="action-icon" />
            <div>
              <div>Export PDF</div>
              <div className="action-desc">Paginated capture</div>
            </div>
          </button>
          <button className="action-btn" onClick={onExportPNG} disabled={!onExportPNG} title="Export PNG">
            <ImageIcon size={14} className="action-icon" />
            <div>
              <div>Export PNG</div>
              <div className="action-desc">High-resolution image</div>
            </div>
          </button>

          <div className="action-group-label">Import</div>
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            style={{ display: 'none' }}
            onChange={(e) => {
              const file = e.target.files?.[0];
              if (file && onImportExcelFile) onImportExcelFile(file);
              if (e.target) e.target.value = '';
            }}
          />
          <button className="action-btn" onClick={triggerFilePick} title="Import from Excel (.xlsx)">
            <Upload size={14} className="action-icon" />
            <div>
              <div>Import from Excel</div>
              <div className="action-desc">Load workbook into grid</div>
            </div>
          </button>

          <div className="action-group-label">History</div>
          <div className="history-list">
            {versionHistory.length === 0 && (
              <div className="action-desc">No versions yet</div>
            )}
            {versionHistory.map((v) => (
              <div key={v.id} className="history-item" title={v.description}>
                <div className="history-main">
                  <div className="history-desc">{v.description}</div>
                  <div className="history-time">{new Date(v.timestamp).toLocaleString()}</div>
                </div>
                <div className="history-actions">
                  <button
                    className="history-btn"
                    onClick={() => onPreviewVersion && onPreviewVersion(v.id)}
                    disabled={!onPreviewVersion}
                  >
                    Preview
                  </button>
                  <button
                    className="history-btn primary"
                    onClick={() => onRestoreVersion && onRestoreVersion(v.id)}
                    disabled={!onRestoreVersion}
                  >
                    Restore
                  </button>
                </div>
              </div>
            ))}
          </div>
        </div>
      </aside>
    </>
  );
});

ActionSidebar.displayName = 'ActionSidebar';

export default ActionSidebar;


