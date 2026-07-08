import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.tsx'
import './index.css'

/* Ensure Office API is loaded before rendering React */
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    console.log('Outlook add-in initialized successfully');
  }

  ReactDOM.createRoot(document.getElementById('root')!).render(
    <React.StrictMode>
      <App />
    </React.StrictMode>,
  )
});
