import { useEffect, useState } from 'react';

const LandingPage = () => {
  const [countdown, setCountdown] = useState(3);
  const [isRedirecting, setIsRedirecting] = useState(false);

  useEffect(() => {
    const timer = setInterval(() => {
      setCountdown((prev) => {
        if (prev <= 1) {
          setIsRedirecting(true);
          // Redirect to the external AutoXL homepage
          window.location.href = 'https://autoxl-home.vercel.app/';
          return 0;
        }
        return prev - 1;
      });
    }, 1000);

    return () => clearInterval(timer);
  }, []);

  const handleManualRedirect = () => {
    setIsRedirecting(true);
    window.location.href = 'https://autoxl-home.vercel.app/';
  };

  return (
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
        
        <h1 style={{ 
          color: '#2c3e50', 
          marginBottom: '1rem',
          fontSize: '2.5rem',
          fontWeight: '300'
        }}>
          AutoXL
        </h1>
        
        <p style={{ 
          color: '#7f8c8d', 
          marginBottom: '2rem',
          fontSize: '1.1rem',
          lineHeight: '1.6'
        }}>
          Redirecting you to the AutoXL homepage in{' '}
          <span style={{ 
            color: '#3498db', 
            fontWeight: 'bold',
            fontSize: '1.2rem'
          }}>
            {countdown}
          </span>{' '}
          seconds...
        </p>
        
        <div style={{
          width: '60px',
          height: '60px',
          border: '4px solid #f3f3f3',
          borderTop: '4px solid #3498db',
          borderRadius: '50%',
          animation: 'spin 1s linear infinite',
          margin: '0 auto 1.5rem'
        }}></div>
        
        <p style={{ 
          marginBottom: '1.5rem', 
          fontSize: '0.9rem', 
          color: '#95a5a6' 
        }}>
          You'll be redirected to the full AutoXL experience
        </p>
        
        <button 
          onClick={handleManualRedirect}
          disabled={isRedirecting}
          style={{
            backgroundColor: '#3498db',
            color: 'white',
            border: 'none',
            padding: '12px 24px',
            borderRadius: '8px',
            cursor: isRedirecting ? 'not-allowed' : 'pointer',
            fontSize: '16px',
            fontWeight: '500',
            transition: 'all 0.2s ease',
            opacity: isRedirecting ? 0.6 : 1
          }}
          onMouseEnter={(e) => {
            if (!isRedirecting) {
              e.currentTarget.style.backgroundColor = '#2980b9'
              e.currentTarget.style.transform = 'translateY(-2px)'
            }
          }}
          onMouseLeave={(e) => {
            if (!isRedirecting) {
              e.currentTarget.style.backgroundColor = '#3498db'
              e.currentTarget.style.transform = 'translateY(0)'
            }
          }}
        >
          {isRedirecting ? 'Redirecting...' : 'Go Now'}
        </button>
        
        <p style={{ 
          marginTop: '1.5rem', 
          fontSize: '0.8rem', 
          color: '#bdc3c7' 
        }}>
          Or visit{' '}
          <a 
            href="https://autoxl-home.vercel.app/" 
            target="_blank"
            rel="noopener noreferrer"
            style={{ 
              color: '#3498db', 
              textDecoration: 'none',
              fontWeight: '500'
            }}
          >
            autoxl-home.vercel.app
          </a>
        </p>
      </div>
      
      <style>{`
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
      `}</style>
    </div>
  );
};

export default LandingPage;
