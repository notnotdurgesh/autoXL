import { BrowserRouter as Router, Routes, Route, Navigate } from 'react-router-dom'
import LandingPage from './components/LandingPage'
import SpreadsheetPage from './components/SpreadsheetPage'
import './App.css'

function App() {
  return (
    <Router>
      <Routes>
        {/* Root route redirects to external AutoXL homepage */}
        <Route path="/" element={<LandingPage />} />
        
        {/* Sheet route shows the spreadsheet */}
        <Route path="/sheet" element={<SpreadsheetPage />} />
        
        {/* Catch all other routes and redirect to root */}
        <Route path="*" element={<Navigate to="/" replace />} />
      </Routes>
    </Router>
  )
}

export default App
