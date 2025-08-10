import React, { useState } from 'react';
import { Heart, X, ArrowUpRight } from 'lucide-react';

interface LoveBadgeProps {
  xUrl?: string;
  linkedinUrl?: string;
  handle?: string;
  githubUrl?: string;
  chessUrl?: string;
}

const LoveBadge: React.FC<LoveBadgeProps> = ({
  xUrl = 'https://x.com/durgihere',
  linkedinUrl = 'https://www.linkedin.com/in/itsdurgesh/',
  handle = '@durgihere',
  githubUrl = 'https://github.com/notnotdurgesh',
  chessUrl = 'https://www.chess.com/member/notdurgesh',
}) => {
  const [open, setOpen] = useState(false);

  return (
    <>
      <button
        aria-label="Made with love"
        onClick={() => setOpen(true)}
        style={{
          position: 'fixed',
          left: '20px',
          bottom: '20px',
          display: 'inline-flex',
          alignItems: 'center',
          justifyContent: 'center',
          gap: '8px',
          padding: '10px',
          borderRadius: '999px',
          border: '1px solid #ffccd5',
          background: 'linear-gradient(135deg, #ff6b6b 0%, #ff3b3b 100%)',
          color: 'white',
          boxShadow: '0 6px 24px rgba(255, 59, 59, 0.35)',
          cursor: 'pointer',
          zIndex: 1200,
          transition: 'transform 0.2s ease, box-shadow 0.2s ease',
        }}
        onMouseEnter={(e) => {
          (e.currentTarget as HTMLButtonElement).style.transform = 'translateY(-2px) scale(1.03)';
          (e.currentTarget as HTMLButtonElement).style.boxShadow = '0 10px 30px rgba(255, 59, 59, 0.45)';
        }}
        onMouseLeave={(e) => {
          (e.currentTarget as HTMLButtonElement).style.transform = 'translateY(0) scale(1)';
          (e.currentTarget as HTMLButtonElement).style.boxShadow = '0 6px 24px rgba(255, 59, 59, 0.35)';
        }}
      >
        <Heart size={16} style={{ transform: 'translateY(-1px)' }} />
        {/* <span style={{ fontWeight: 700, fontSize: 12, letterSpacing: 0.2 }}>Made with love</span> */}
      </button>

      {open && (
        <div
          role="dialog"
          aria-modal="true"
          onClick={() => setOpen(false)}
          style={{
            position: 'fixed',
            inset: 0,
            background: 'rgba(0,0,0,0.35)',
            backdropFilter: 'blur(2px)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 2000,
            animation: 'fadeIn 180ms ease-out',
          }}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            style={{
              width: 'min(92vw, 460px)',
              background: 'white',
              borderRadius: 16,
              border: '1px solid #ffe3e8',
              boxShadow: '0 16px 60px rgba(0,0,0,0.18)',
              overflow: 'hidden',
              transform: 'translateY(8px)',
              animation: 'slideUp 200ms ease-out forwards',
            }}
          >
            <div
              style={{
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
                padding: '14px 16px',
                background: 'linear-gradient(135deg, #fff5f7 0%, #ffe3e8 100%)',
                borderBottom: '1px solid #ffd6de',
              }}
            >
              <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                <div
                  style={{
                    display: 'inline-flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    width: 28,
                    height: 28,
                    borderRadius: 14,
                    background: '#ff6b6b',
                    color: 'white',
                    boxShadow: '0 3px 10px rgba(255,107,107,0.4)',
                  }}
                >
                  <Heart size={16} />
                </div>
                <div>
                  <div style={{ fontWeight: 800, fontSize: 14, color: '#1a1a1a' }}>Made with love</div>
                  <div style={{ fontSize: 12, color: '#6b7280' }}>by {handle}</div>
                </div>
              </div>
              <button
                aria-label="Close"
                onClick={() => setOpen(false)}
                style={{
                  background: 'transparent',
                  border: 'none',
                  width: 28,
                  height: 28,
                  borderRadius: 6,
                  cursor: 'pointer',
                }}
              >
                <X size={18} color="#6b7280" />
              </button>
            </div>

            <div style={{ padding: 16 }}>
              <div style={{ fontSize: 13, color: '#374151', lineHeight: 1.55 }}>
                Thanks for trying AutoXL. Connect with me:
              </div>

              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10, marginTop: 14 }}>
                <a
                  href={xUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                  style={{
                    display: 'inline-flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    padding: '10px 12px',
                    borderRadius: 10,
                    border: '1px solid #e2e8f0',
                    background: 'linear-gradient(135deg, #0f172a 0%, #111827 100%)',
                    color: 'white',
                    textDecoration: 'none',
                    boxShadow: '0 6px 18px rgba(15,23,42,0.35)',
                    transition: 'transform 0.2s ease, box-shadow 0.2s ease',
                  }}
                  onMouseEnter={(e) => {
                    (e.currentTarget as HTMLAnchorElement).style.transform = 'translateY(-2px)';
                    (e.currentTarget as HTMLAnchorElement).style.boxShadow = '0 10px 26px rgba(15,23,42,0.45)';
                  }}
                  onMouseLeave={(e) => {
                    (e.currentTarget as HTMLAnchorElement).style.transform = 'translateY(0)';
                    (e.currentTarget as HTMLAnchorElement).style.boxShadow = '0 6px 18px rgba(15,23,42,0.35)';
                  }}
                >
                  <span style={{ fontWeight: 700, fontSize: 12 }}>X (Twitter)</span>
                  <ArrowUpRight size={16} />
                </a>

                <a
                  href={linkedinUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                  style={{
                    display: 'inline-flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    padding: '10px 12px',
                    borderRadius: 10,
                    border: '1px solid #e2e8f0',
                    background: 'linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%)',
                    color: 'white',
                    textDecoration: 'none',
                    boxShadow: '0 6px 18px rgba(37,99,235,0.35)',
                    transition: 'transform 0.2s ease, box-shadow 0.2s ease',
                  }}
                  onMouseEnter={(e) => {
                    (e.currentTarget as HTMLAnchorElement).style.transform = 'translateY(-2px)';
                    (e.currentTarget as HTMLAnchorElement).style.boxShadow = '0 10px 26px rgba(37,99,235,0.45)';
                  }}
                  onMouseLeave={(e) => {
                    (e.currentTarget as HTMLAnchorElement).style.transform = 'translateY(0)';
                    (e.currentTarget as HTMLAnchorElement).style.boxShadow = '0 6px 18px rgba(37,99,235,0.35)';
                  }}
                >
                  <span style={{ fontWeight: 700, fontSize: 12 }}>LinkedIn</span>
                  <ArrowUpRight size={16} />
                </a>

                <a
                  href={githubUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                  style={{
                    display: 'inline-flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    padding: '10px 12px',
                    borderRadius: 10,
                    border: '1px solid #e2e8f0',
                    background: 'linear-gradient(135deg, #0f172a 0%, #1f2937 100%)',
                    color: 'white',
                    textDecoration: 'none',
                    boxShadow: '0 6px 18px rgba(17,24,39,0.35)',
                    transition: 'transform 0.2s ease, box-shadow 0.2s ease',
                  }}
                  onMouseEnter={(e) => {
                    (e.currentTarget as HTMLAnchorElement).style.transform = 'translateY(-2px)';
                    (e.currentTarget as HTMLAnchorElement).style.boxShadow = '0 10px 26px rgba(17,24,39,0.45)';
                  }}
                  onMouseLeave={(e) => {
                    (e.currentTarget as HTMLAnchorElement).style.transform = 'translateY(0)';
                    (e.currentTarget as HTMLAnchorElement).style.boxShadow = '0 6px 18px rgba(17,24,39,0.35)';
                  }}
                >
                  <span style={{ fontWeight: 700, fontSize: 12 }}>GitHub</span>
                  <ArrowUpRight size={16} />
                </a>

                <a
                  href={chessUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                  style={{
                    display: 'inline-flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    padding: '10px 12px',
                    borderRadius: 10,
                    border: '1px solid #e2e8f0',
                    background: 'linear-gradient(135deg, #10b981 0%, #059669 100%)',
                    color: 'white',
                    textDecoration: 'none',
                    boxShadow: '0 6px 18px rgba(16,185,129,0.35)',
                    transition: 'transform 0.2s ease, box-shadow 0.2s ease',
                  }}
                  onMouseEnter={(e) => {
                    (e.currentTarget as HTMLAnchorElement).style.transform = 'translateY(-2px)';
                    (e.currentTarget as HTMLAnchorElement).style.boxShadow = '0 10px 26px rgba(16,185,129,0.45)';
                  }}
                  onMouseLeave={(e) => {
                    (e.currentTarget as HTMLAnchorElement).style.transform = 'translateY(0)';
                    (e.currentTarget as HTMLAnchorElement).style.boxShadow = '0 6px 18px rgba(16,185,129,0.35)';
                  }}
                >
                  <span style={{ fontWeight: 700, fontSize: 12 }}>Chess.com</span>
                  <ArrowUpRight size={16} />
                </a>
              </div>
            </div>
          </div>
        </div>
      )}

      <style>{`
        @keyframes fadeIn {
          from { opacity: 0; }
          to { opacity: 1; }
        }
        @keyframes slideUp {
          from { opacity: 0; transform: translateY(12px); }
          to { opacity: 1; transform: translateY(0); }
        }
      `}</style>
    </>
  );
};

export default LoveBadge;


