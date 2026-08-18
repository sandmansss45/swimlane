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

initMsal()
  .then(() => {
    createRoot(document.getElementById('root')!).render(
      <StrictMode>
        <App msalInstance={msalInstance} />
      </StrictMode>
    );
  })
  .catch((err: unknown) => {
    // Startup failures here (almost always an MSAL/auth config problem)
    // previously left the page blank forever with nothing but a silent
    // console error - genuinely hard to diagnose without already knowing
    // to check DevTools. This bypasses React entirely (deliberately -
    // something in the app's own startup just failed, so don't depend on
    // any of its code paths being trustworthy) and just writes a plain,
    // visible error directly into the page.
    // eslint-disable-next-line no-console
    console.error('App failed to initialize', err);
    const root = document.getElementById('root');
    if (!root) return;
    const message = err instanceof Error ? err.message : String(err);
    const escaped = message.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    root.innerHTML = `
      <div style="font-family: system-ui, sans-serif; max-width: 640px; margin: 80px auto; padding: 24px; border: 1px solid #dde3ea; border-radius: 12px; background: #fff; box-shadow: 0 1px 2px rgba(16,24,40,0.06);">
        <h2 style="color: #142d50; margin: 0 0 8px;">Swimlane Studio failed to start</h2>
        <p style="color: #16233d; margin: 0 0 12px;">Something went wrong during startup, before the app could render anything - almost always a sign-in (MSAL) configuration issue.</p>
        <pre style="background: #eaf2fd; padding: 12px; border-radius: 8px; overflow: auto; color: #16233d; white-space: pre-wrap; margin: 0 0 12px;">${escaped}</pre>
        <p style="color: #5b6b84; font-size: 13px; margin: 0;">Open DevTools (F12) → Console for the full error and stack trace.</p>
      </div>
    `;
  });
