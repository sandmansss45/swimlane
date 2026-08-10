import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import './index.css';
import App from './App.tsx';
import { msalInstance, initMsal } from './swimlane/auth/msalInstance';

initMsal().then(() => {
  createRoot(document.getElementById('root')!).render(
    <StrictMode>
      <App msalInstance={msalInstance} />
    </StrictMode>
  );
});
