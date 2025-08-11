
import { Link } from 'react-router-dom'

const Navigation = () => {
  return (
    <div style={{
      position: 'fixed',
      top: '10px',
      right: '10px',
      zIndex: 1000,
      backgroundColor: 'rgba(255, 255, 255, 0.95)',
      borderRadius: '8px',
      padding: '8px',
      boxShadow: '0 2px 10px rgba(0,0,0,0.1)',
      backdropFilter: 'blur(10px)',
      border: '1px solid rgba(255, 255, 255, 0.2)'
    }}>
      <div style={{
        display: 'flex',
        gap: '8px',
        alignItems: 'center'
      }}>
        <Link 
          to="/"
          style={{
            backgroundColor: '#3498db',
            color: 'white',
            padding: '6px 12px',
            borderRadius: '4px',
            textDecoration: 'none',
            fontSize: '12px',
            fontWeight: '500',
            transition: 'all 0.2s ease'
          }}
          onMouseEnter={(e) => {
            e.currentTarget.style.backgroundColor = '#2980b9'
            e.currentTarget.style.transform = 'translateY(-1px)'
          }}
          onMouseLeave={(e) => {
            e.currentTarget.style.backgroundColor = '#3498db'
            e.currentTarget.style.transform = 'translateY(0)'
          }}
        >
          Home
        </Link>
        <a 
          href="https://autoxl-home.vercel.app/"
          target="_blank"
          rel="noopener noreferrer"
          style={{
            backgroundColor: '#27ae60',
            color: 'white',
            padding: '6px 12px',
            borderRadius: '4px',
            textDecoration: 'none',
            fontSize: '12px',
            fontWeight: '500',
            transition: 'all 0.2s ease'
          }}
          onMouseEnter={(e) => {
            e.currentTarget.style.backgroundColor = '#229954'
            e.currentTarget.style.transform = 'translateY(-1px)'
          }}
          onMouseLeave={(e) => {
            e.currentTarget.style.backgroundColor = '#27ae60'
            e.currentTarget.style.transform = 'translateY(0)'
          }}
        >
          AutoXL Home
        </a>
        <div style={{
          backgroundColor: '#f39c12',
          color: 'white',
          padding: '6px 12px',
          borderRadius: '4px',
          fontSize: '12px',
          fontWeight: '500',
          cursor: 'default'
        }}>
          Sheet
        </div>
      </div>
    </div>
  )
}

export default Navigation
