import { createRoot } from 'react-dom/client';
import App from './App.tsx';
import './index.css';
import { initializeOfficeContext } from './office/context.ts';

// Render App after initiating Office initialization
const container = document.getElementById('root')!;
const root = createRoot(container);

// Render application
root.render(<App />);

// Kick off Office.js ready handler in background
initializeOfficeContext().then((info) => {
  console.log('[Office AI Agent] Initialized context:', info);
});
