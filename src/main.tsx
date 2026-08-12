import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { initializeIcons } from '@fluentui/react';
import './index.css';
import App from './App.tsx';
import { msalInstance, initMsal } from './swimlane/auth/msalInstance';

// Registers the Fabric MDL2 icon font Fluent's Icon/IconButton components
// render glyphs from - without this call, every icon (delete, close, etc.)
// exists in the DOM correctly but renders as a blank/invisible glyph.
initializeIcons();

initMsal().then(() => {
  createRoot(document.getElementById('root')!).render(
    <StrictMode>
      <App msalInstance={msalInstance} />
    </StrictMode>
  );
});
