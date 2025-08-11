import { BrowserRouter as Router, Routes, Route } from 'react-router-dom'
import LandingPage from './components/LandingPage'
import SpreadsheetPage from './components/SpreadsheetPage'
import NotFound from './components/NotFound'
import './App.css'

function App() {
  return (
    <Router>
      <Routes>
        {/* Root route redirects to external AutoXL homepage */}
        <Route path="/" element={<LandingPage />} />
        
        {/* Sheet route shows the spreadsheet */}
        <Route path="/sheet" element={<SpreadsheetPage />} />
        
        {/* Catch all other routes and show 404 page */}
        <Route path="*" element={<NotFound />} />
      </Routes>
    </Router>
  )
}

export default App
