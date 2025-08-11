import { useEffect } from 'react';

const LandingPage = () => {
  useEffect(() => {
    // Redirect to the external AutoXL homepage
    window.location.href = 'https://autoxl-home.vercel.app/';
  }, []);

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
        maxWidth: '400px'
      }}>
        <h1 style={{ color: '#333', marginBottom: '1rem' }}>AutoXL</h1>
        <p style={{ color: '#666', marginBottom: '1rem' }}>
          Redirecting you to the AutoXL homepage...
        </p>
        <div style={{
          width: '40px',
          height: '40px',
          border: '4px solid #f3f3f3',
          borderTop: '4px solid #3498db',
          borderRadius: '50%',
          animation: 'spin 1s linear infinite',
          margin: '0 auto'
        }}></div>
        <p style={{ marginTop: '1rem', fontSize: '0.9rem', color: '#999' }}>
          If you're not redirected automatically,{' '}
          <a 
            href="https://autoxl-home.vercel.app/" 
            style={{ color: '#3498db', textDecoration: 'none' }}
          >
            click here
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
