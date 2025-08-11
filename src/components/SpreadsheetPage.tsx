import React, { Suspense, useState, useEffect } from 'react'
import ExcelSpreadsheet from './ExcelSpreadsheet'
import DemoDisclaimer from './DemoDisclaimer'
import LoveBadge from './LoveBadge'
import Navigation from './Navigation'

// Error Boundary Component
class SpreadsheetErrorBoundary extends React.Component<
  { children: React.ReactNode },
  { hasError: boolean; error?: Error }
> {
  constructor(props: { children: React.ReactNode }) {
    super(props)
    this.state = { hasError: false }
  }

  static getDerivedStateFromError(error: Error) {
    return { hasError: true, error }
  }

  componentDidCatch(error: Error, errorInfo: React.ErrorInfo) {
    console.error('Spreadsheet Error:', error, errorInfo)
    
    // Log additional error details for debugging
    if (error.stack) {
      console.error('Error Stack:', error.stack)
    }
    
    // Log component stack if available
    if (errorInfo.componentStack) {
      console.error('Component Stack:', errorInfo.componentStack)
    }
  }

  render() {
    if (this.state.hasError) {
      return (
        <div style={{
          display: 'flex',
          justifyContent: 'center',
          alignItems: 'center',
          height: '100vh',
          flexDirection: 'column',
          fontFamily: 'Arial, sans-serif',
          backgroundColor: '#f5f5f5'
        }}>
          <div style={{
            textAlign: 'center',
            padding: '2rem',
            backgroundColor: 'white',
            borderRadius: '8px',
            boxShadow: '0 2px 10px rgba(0,0,0,0.1)',
            maxWidth: '500px'
          }}>
            <h2 style={{ color: '#e74c3c', marginBottom: '1rem' }}>Something went wrong</h2>
            <p style={{ color: '#666', marginBottom: '1rem' }}>
              The spreadsheet encountered an error. Please refresh the page to try again.
            </p>
            <button 
              onClick={() => window.location.reload()}
              style={{
                backgroundColor: '#3498db',
                color: 'white',
                border: 'none',
                padding: '10px 20px',
                borderRadius: '5px',
                cursor: 'pointer',
                fontSize: '16px'
              }}
            >
              Refresh Page
            </button>
          </div>
        </div>
      )
    }

    return this.props.children
  }
}

// Loading Component
const SpreadsheetLoader = () => (
  <div style={{
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    height: '100vh',
    flexDirection: 'column',
    fontFamily: 'Arial, sans-serif',
    backgroundColor: '#f5f5f5',
    background: 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)'
  }}>
    <div style={{
      textAlign: 'center',
      padding: '3rem',
      backgroundColor: 'white',
      borderRadius: '16px',
      boxShadow: '0 20px 40px rgba(0,0,0,0.1)',
      maxWidth: '500px',
      border: '1px solid rgba(255, 255, 255, 0.2)'
    }}>
      <div style={{
        width: '80px',
        height: '80px',
        backgroundColor: '#3498db',
        borderRadius: '50%',
        margin: '0 auto 1.5rem',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        color: 'white',
        fontSize: '2rem',
        fontWeight: 'bold'
      }}>
        XL
      </div>
      
      <h2 style={{ 
        color: '#2c3e50', 
        marginBottom: '1rem',
        fontSize: '2rem',
        fontWeight: '300'
      }}>
        Loading AutoXL
      </h2>
      
      <p style={{ 
        color: '#7f8c8d', 
        marginBottom: '2rem',
        fontSize: '1rem',
        lineHeight: '1.6'
      }}>
        Initializing spreadsheet components...
      </p>
      
      <div style={{
        width: '60px',
        height: '60px',
        border: '4px solid #f3f3f3',
        borderTop: '4px solid #3498db',
        borderRadius: '50%',
        animation: 'spin 1s linear infinite',
        margin: '0 auto'
      }}></div>
      
      <p style={{ 
        marginTop: '1.5rem', 
        fontSize: '0.9rem', 
        color: '#95a5a6' 
      }}>
        This may take a few seconds on first load
      </p>
    </div>
    
    <style>{`
      @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
      }
    `}</style>
  </div>
)

// Fallback Component for when spreadsheet fails to load
const SpreadsheetFallback = () => {
  const [retryCount, setRetryCount] = useState(0)
  
  const handleRetry = () => {
    setRetryCount(prev => prev + 1)
    window.location.reload()
  }
  
  const handleGoHome = () => {
    window.location.href = '/'
  }
  
  return (
    <div style={{
      display: 'flex',
      justifyContent: 'center',
      alignItems: 'center',
      height: '100vh',
      flexDirection: 'column',
      fontFamily: 'Arial, sans-serif',
      backgroundColor: '#f5f5f5',
      background: 'linear-gradient(135deg, #ff6b6b 0%, #ee5a24 100%)'
    }}>
      <div style={{
        textAlign: 'center',
        padding: '3rem',
        backgroundColor: 'white',
        borderRadius: '16px',
        boxShadow: '0 20px 40px rgba(0,0,0,0.1)',
        maxWidth: '500px',
        border: '1px solid rgba(255, 255, 255, 0.2)'
      }}>
        <div style={{
          width: '80px',
          height: '80px',
          backgroundColor: '#e74c3c',
          borderRadius: '50%',
          margin: '0 auto 1.5rem',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          color: 'white',
          fontSize: '2rem',
          fontWeight: 'bold'
        }}>
          ⚠️
        </div>
        
        <h2 style={{ 
          color: '#2c3e50', 
          marginBottom: '1rem',
          fontSize: '2rem',
          fontWeight: '300'
        }}>
          Spreadsheet Unavailable
        </h2>
        
        <p style={{ 
          color: '#7f8c8d', 
          marginBottom: '2rem',
          fontSize: '1rem',
          lineHeight: '1.6'
        }}>
          The spreadsheet couldn't be loaded. This might be due to a temporary issue or browser compatibility.
        </p>
        
        <div style={{
          display: 'flex',
          gap: '1rem',
          justifyContent: 'center',
          flexWrap: 'wrap'
        }}>
          <button 
            onClick={handleRetry}
            style={{
              backgroundColor: '#3498db',
              color: 'white',
              border: 'none',
              padding: '12px 24px',
              borderRadius: '8px',
              cursor: 'pointer',
              fontSize: '16px',
              fontWeight: '500',
              transition: 'all 0.2s ease'
            }}
            onMouseEnter={(e) => {
              e.currentTarget.style.backgroundColor = '#2980b9'
              e.currentTarget.style.transform = 'translateY(-2px)'
            }}
            onMouseLeave={(e) => {
              e.currentTarget.style.backgroundColor = '#3498db'
              e.currentTarget.style.transform = 'translateY(0)'
            }}
          >
            Try Again
          </button>
          
          <button 
            onClick={handleGoHome}
            style={{
              backgroundColor: '#95a5a6',
              color: 'white',
              border: 'none',
              padding: '12px 24px',
              borderRadius: '8px',
              cursor: 'pointer',
              fontSize: '16px',
              fontWeight: '500',
              transition: 'all 0.2s ease'
            }}
            onMouseEnter={(e) => {
              e.currentTarget.style.backgroundColor = '#7f8c8d'
              e.currentTarget.style.transform = 'translateY(-2px)'
            }}
            onMouseLeave={(e) => {
              e.currentTarget.style.backgroundColor = '#95a5a6'
              e.currentTarget.style.transform = 'translateY(0)'
            }}
          >
            Go Home
          </button>
        </div>
        
        {retryCount > 0 && (
          <p style={{ 
            marginTop: '1.5rem', 
            fontSize: '0.9rem', 
            color: '#e74c3c' 
          }}>
            Retry attempt: {retryCount}
          </p>
        )}
      </div>
    </div>
  )
}

const SpreadsheetPage = () => {
  const [hasError, setHasError] = useState(false)
  
  useEffect(() => {
    // Add error listener for unhandled errors
    const handleError = (event: ErrorEvent) => {
      console.error('Unhandled error:', event.error)
      setHasError(true)
    }
    
    const handleUnhandledRejection = (event: PromiseRejectionEvent) => {
      console.error('Unhandled promise rejection:', event.reason)
      setHasError(true)
    }
    
    window.addEventListener('error', handleError)
    window.addEventListener('unhandledrejection', handleUnhandledRejection)
    
    return () => {
      window.removeEventListener('error', handleError)
      window.removeEventListener('unhandledrejection', handleUnhandledRejection)
    }
  }, [])
  
  if (hasError) {
    return <SpreadsheetFallback />
  }
  
  return (
    <SpreadsheetErrorBoundary>
      <Suspense fallback={<SpreadsheetLoader />}>
        <div style={{ 
          width: '100vw', 
          height: '100vh', 
          margin: 0,
          padding: 0,
          overflow: 'hidden',
          position: 'fixed',
          top: 0,
          left: 0
        }}>
          <Navigation />
          <div style={{ width: '100%', height: '100%' }}>
            <ExcelSpreadsheet />
            <DemoDisclaimer />
            <LoveBadge />
          </div>
        </div>
      </Suspense>
    </SpreadsheetErrorBoundary>
  )
}

export default SpreadsheetPage
