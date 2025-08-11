import ExcelSpreadsheet from './ExcelSpreadsheet'
import DemoDisclaimer from './DemoDisclaimer'
import LoveBadge from './LoveBadge'

const SpreadsheetPage = () => {
  return (
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
      <ExcelSpreadsheet />
      <DemoDisclaimer />
      <LoveBadge />
    </div>
  )
}

export default SpreadsheetPage
