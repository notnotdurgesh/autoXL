import React, { useState, useEffect } from 'react';
import { X, AlertCircle, Mail, Bug } from 'lucide-react';

const DemoDisclaimer: React.FC = () => {
  const [isVisible, setIsVisible] = useState(false);

  useEffect(() => {
    // Check if disclaimer has been shown in this session
    const hasShownDisclaimer = sessionStorage.getItem('demoDisclaimerShown');
    
    if (!hasShownDisclaimer) {
      // Show disclaimer after a small delay for better UX
      setTimeout(() => {
        setIsVisible(true);
      }, 500);
    }
  }, []);

  const handleClose = () => {
    setIsVisible(false);
    // Mark as shown in session storage
    sessionStorage.setItem('demoDisclaimerShown', 'true');
  };

  if (!isVisible) return null;

  return (
    <>
      {/* Backdrop */}
      <div 
        className="demo-disclaimer-backdrop"
        onClick={handleClose}
        style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          backgroundColor: 'rgba(0, 0, 0, 0.4)',
          backdropFilter: 'blur(4px)',
          WebkitBackdropFilter: 'blur(4px)',
          zIndex: 1000,
          animation: 'overlayFadeIn 0.2s ease-out',
        }}
      />
      
      {/* Modal */}
      <div 
        className="demo-disclaimer-modal"
        style={{
          position: 'fixed',
          top: '50%',
          left: '50%',
          transform: 'translate(-50%, -50%)',
          backgroundColor: 'white',
          borderRadius: '12px',
          padding: '24px',
          width: 'calc(100% - 40px)',
          maxWidth: '550px',
          maxHeight: '90vh',
          overflowY: 'auto',
          boxShadow: '0 20px 60px rgba(0,0,0,0.12), 0 8px 25px rgba(0,0,0,0.08)',
          border: '1px solid rgba(0,0,0,0.05)',
          zIndex: 1001,
          animation: 'modalSlideIn 0.25s ease-out',
        }}
      >
        {/* Close button */}
        <button
          onClick={handleClose}
          className="demo-close-btn"
          style={{
            position: 'absolute',
            top: '12px',
            right: '12px',
            background: 'transparent',
            border: '1px solid transparent',
            borderRadius: '4px',
            cursor: 'pointer',
            padding: '4px',
            width: '28px',
            height: '28px',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            transition: 'all 0.2s ease',
            color: '#495057',
          }}
          onMouseEnter={(e) => {
            e.currentTarget.style.backgroundColor = '#e9ecef';
            e.currentTarget.style.borderColor = '#ced4da';
          }}
          onMouseLeave={(e) => {
            e.currentTarget.style.backgroundColor = 'transparent';
            e.currentTarget.style.borderColor = 'transparent';
          }}
        >
          <X size={16} />
        </button>

        {/* Header */}
        <h3 style={{ 
          margin: '0 0 20px 0',
          fontSize: '18px',
          fontWeight: '600',
          color: '#2d3748',
          display: 'flex',
          alignItems: 'center',
          gap: '8px'
        }}>
          🚀 Welcome to AutoXL Beta
        </h3>

        {/* Version info */}
        <div style={{
          backgroundColor: '#f8f9fa',
          border: '1px solid #dee2e6',
          borderRadius: '4px',
          padding: '8px 12px',
          marginBottom: '16px',
          fontSize: '12px',
          color: '#495057'
        }}>
          <strong>Version:</strong> 0.6.9 - Early Access Demo
        </div>

        {/* Alert box */}
        <div style={{
          backgroundColor: '#fff3cd',
          border: '1px solid #ffc107',
          borderRadius: '4px',
          padding: '12px',
          marginBottom: '20px',
          display: 'flex',
          gap: '8px',
          alignItems: 'flex-start'
        }}>
          <AlertCircle size={16} color="#856404" style={{ flexShrink: 0, marginTop: '2px' }} />
          <div>
            <p style={{
              margin: 0,
              fontSize: '13px',
              color: '#856404',
              lineHeight: '1.5'
            }}>
              <strong>Early Demo Version:</strong> This product is in very early development. 
              You may encounter bugs, unexpected behavior, or incomplete features.
            </p>
          </div>
        </div>

        {/* Main content */}
        <div style={{ marginBottom: '20px' }}>
          <p style={{
            margin: '0 0 16px 0',
            fontSize: '14px',
            color: '#333',
            lineHeight: '1.6'
          }}>
            We're actively improving AutoXL. Your feedback is valuable!
          </p>

          {/* Bug report section */}
          <div style={{
            backgroundColor: '#f8f9fa',
            border: '1px solid #dee2e6',
            borderRadius: '4px',
            padding: '16px',
          }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '6px', marginBottom: '10px' }}>
              <Bug size={14} color="#495057" />
              <h4 style={{ 
                margin: 0,
                fontSize: '14px',
                fontWeight: '600',
                color: '#333'
              }}>
                Found a bug?
              </h4>
            </div>
            
            <p style={{
              margin: '0 0 10px 0',
              fontSize: '12px',
              color: '#495057',
              lineHeight: '1.5'
            }}>
              Please report any issues, bugs, or suggestions to:
            </p>
            
            <div className="email-section" style={{
              display: 'flex',
              alignItems: 'center',
              gap: '6px',
              backgroundColor: 'white',
              padding: '8px 12px',
              borderRadius: '4px',
              border: '1px solid #ced4da',
              overflowX: 'auto',
              whiteSpace: 'nowrap'
            }}>
              <Mail size={14} color="#007bff" style={{ flexShrink: 0 }} />
              <a 
                href="mailto:durgeshwantreddy@gmail.com?subject=AutoXL Bug Report"
                style={{
                  color: '#007bff',
                  textDecoration: 'none',
                  fontSize: '12px',
                  fontWeight: '500',
                  wordBreak: 'break-word'
                }}
              >
                durgeshwantreddy@gmail.com
              </a>
            </div>

            {/* Sweet surprise note */}
            <p style={{
              marginTop: '12px',
              marginBottom: 0,
              fontSize: '11px',
              color: '#6c757d',
              fontStyle: 'italic'
            }}>
              💝 Sweet surprise for helpful bug reports!
            </p>
          </div>
        </div>

        {/* Action buttons */}
        <div style={{
          display: 'flex',
          gap: '12px',
          justifyContent: 'flex-end',
          paddingTop: '16px',
          borderTop: '1px solid #e2e8f0'
        }}>
          <button
            onClick={handleClose}
            className="demo-btn-primary"
            style={{
              padding: '10px 20px',
              background: 'linear-gradient(135deg, #007bff 0%, #0056b3 100%)',
              color: 'white',
              border: '2px solid #007bff',
              borderRadius: '8px',
              fontSize: '13px',
              fontWeight: '600',
              cursor: 'pointer',
              transition: 'all 0.2s ease',
              minWidth: '90px',
              boxShadow: '0 2px 8px rgba(0,123,255,0.2)',
            }}
            onMouseEnter={(e) => {
              e.currentTarget.style.background = 'linear-gradient(135deg, #0056b3 0%, #004085 100%)';
              e.currentTarget.style.transform = 'translateY(-1px)';
              e.currentTarget.style.boxShadow = '0 4px 12px rgba(0,123,255,0.3)';
            }}
            onMouseLeave={(e) => {
              e.currentTarget.style.background = 'linear-gradient(135deg, #007bff 0%, #0056b3 100%)';
              e.currentTarget.style.transform = 'translateY(0)';
              e.currentTarget.style.boxShadow = '0 2px 8px rgba(0,123,255,0.2)';
            }}
          >
            Get Started
          </button>
        </div>
      </div>

      {/* Add CSS animations - matching ExcelToolbar patterns */}
      <style>{`
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

        @keyframes modalSlideIn {
          from {
            opacity: 0;
            transform: translate(-50%, -40%) scale(0.95);
          }
          to {
            opacity: 1;
            transform: translate(-50%, -50%) scale(1);
          }
        }

        /* Ensure scrollbar styling for modal */
        .demo-disclaimer-modal::-webkit-scrollbar {
          width: 8px;
        }
        
        .demo-disclaimer-modal::-webkit-scrollbar-track {
          background: #f1f1f1;
          border-radius: 4px;
        }
        
        .demo-disclaimer-modal::-webkit-scrollbar-thumb {
          background: #888;
          border-radius: 4px;
        }
        
        .demo-disclaimer-modal::-webkit-scrollbar-thumb:hover {
          background: #555;
        }

        /* Tablet styles */
        @media (max-width: 768px) {
          .demo-disclaimer-modal {
            width: calc(100% - 32px) !important;
            padding: 20px !important;
            max-height: 85vh !important;
          }
          
          .demo-disclaimer-modal h3 {
            font-size: 16px !important;
            margin-bottom: 16px !important;
          }
          
          .demo-disclaimer-modal p {
            font-size: 13px !important;
          }
          
          .demo-close-btn {
            top: 8px !important;
            right: 8px !important;
          }
        }

        /* Mobile styles */
        @media (max-width: 480px) {
          .demo-disclaimer-modal {
            width: calc(100% - 24px) !important;
            padding: 16px !important;
            max-height: 80vh !important;
            border-radius: 8px !important;
          }
          
          .demo-disclaimer-modal h3 {
            font-size: 15px !important;
            margin-bottom: 12px !important;
            padding-right: 20px !important;
          }
          
          .demo-disclaimer-modal h4 {
            font-size: 13px !important;
          }
          
          .demo-disclaimer-modal p {
            font-size: 12px !important;
            line-height: 1.4 !important;
          }
          
          .demo-disclaimer-modal > div {
            margin-bottom: 12px !important;
          }
          
          .demo-disclaimer-modal button {
            padding: 8px 16px !important;
            font-size: 12px !important;
            min-width: 80px !important;
          }
          
          .demo-close-btn {
            width: 24px !important;
            height: 24px !important;
            padding: 2px !important;
          }
          
          /* Adjust padding for inner sections */
          .demo-disclaimer-modal > div > div {
            padding: 12px !important;
          }
          
          /* Email section responsive */
          .email-section {
            flex-direction: column !important;
            align-items: flex-start !important;
            gap: 8px !important;
          }
          
          .email-section a {
            word-break: break-all !important;
            white-space: normal !important;
          }
        }

        /* Very small mobile styles */
        @media (max-width: 360px) {
          .demo-disclaimer-modal {
            width: calc(100% - 16px) !important;
            padding: 12px !important;
            max-height: 75vh !important;
          }
          
          .demo-disclaimer-modal h3 {
            font-size: 14px !important;
            margin-bottom: 10px !important;
          }
          
          .demo-disclaimer-modal p {
            font-size: 11px !important;
          }
          
          .demo-disclaimer-modal button {
            padding: 6px 12px !important;
            font-size: 11px !important;
          }
          
          /* Email section responsive */
          .demo-disclaimer-modal a {
            font-size: 11px !important;
            word-break: break-all !important;
          }
        }
        
        /* Landscape mobile */
        @media (max-height: 500px) and (max-width: 768px) {
          .demo-disclaimer-modal {
            max-height: 70vh !important;
            padding: 12px 16px !important;
          }
          
          .demo-disclaimer-modal h3 {
            margin-bottom: 8px !important;
          }
          
          .demo-disclaimer-modal > div {
            margin-bottom: 8px !important;
          }
        }
      `}</style>
    </>
  );
};

export default DemoDisclaimer;
