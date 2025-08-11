import React from 'react'
import { Link } from 'react-router-dom'

const NotFound = () => {
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
          404
        </div>
        
        <h1 style={{ 
          color: '#2c3e50', 
          marginBottom: '1rem',
          fontSize: '2.5rem',
          fontWeight: '300'
        }}>
          Page Not Found
        </h1>
        
        <p style={{ 
          color: '#7f8c8d', 
          marginBottom: '2rem',
          fontSize: '1.1rem',
          lineHeight: '1.6'
        }}>
          The page you're looking for doesn't exist. 
          You can go back to the homepage or visit the AutoXL spreadsheet.
        </p>
        
        <div style={{
          display: 'flex',
          gap: '1rem',
          justifyContent: 'center',
          flexWrap: 'wrap'
        }}>
          <Link 
            to="/"
            style={{
              backgroundColor: '#3498db',
              color: 'white',
              padding: '12px 24px',
              borderRadius: '8px',
              textDecoration: 'none',
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
            Go Home
          </Link>
          
          <Link 
            to="/sheet"
            style={{
              backgroundColor: '#27ae60',
              color: 'white',
              padding: '12px 24px',
              borderRadius: '8px',
              textDecoration: 'none',
              fontSize: '16px',
              fontWeight: '500',
              transition: 'all 0.2s ease'
            }}
            onMouseEnter={(e) => {
              e.currentTarget.style.backgroundColor = '#229954'
              e.currentTarget.style.transform = 'translateY(-2px)'
            }}
            onMouseLeave={(e) => {
              e.currentTarget.style.backgroundColor = '#27ae60'
              e.currentTarget.style.transform = 'translateY(0)'
            }}
          >
            Open Spreadsheet
          </Link>
        </div>
        
        <p style={{ 
          marginTop: '2rem', 
          fontSize: '0.9rem', 
          color: '#95a5a6' 
        }}>
          Or visit the{' '}
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
            AutoXL homepage
          </a>
        </p>
      </div>
    </div>
  );
};

export default NotFound;
