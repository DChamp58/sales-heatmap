import { createRoot } from 'react-dom/client';
import { SalesHeatmap } from './components/SalesHeatmap';
import { ErrorBoundary } from './components/ErrorBoundary';
import './index.css';

const rootEl = document.getElementById('root')!;

// Render a readable message instead of a blank white page for any error that
// escapes React's render cycle (module init, event handlers, async rejections).
function showFatal(title: string, detail: string) {
  const box = document.createElement('div');
  box.style.cssText =
    'max-width:820px;margin:40px auto;padding:20px;font-family:system-ui;color:#7f1d1d;background:#fef2f2;border:1px solid #fecaca;border-radius:12px';
  const h = document.createElement('div');
  h.style.cssText = 'font-weight:700;font-size:16px;margin-bottom:8px';
  h.textContent = title;
  const pre = document.createElement('pre');
  pre.style.cssText =
    'white-space:pre-wrap;word-break:break-word;font-size:12px;color:#991b1b;background:#fff;border:1px solid #fecaca;border-radius:8px;padding:12px;margin:0';
  pre.textContent = detail; // textContent avoids HTML-injection / mis-rendering
  box.append(h, pre);
  rootEl.replaceChildren(box);
}

window.addEventListener('error', (e) => {
  showFatal('The app hit an error', String(e.error?.stack || e.message || e));
});
window.addEventListener('unhandledrejection', (e) => {
  const r = e.reason;
  showFatal('Unhandled promise rejection', String((r && (r.stack || r.message)) || r));
});

try {
  createRoot(rootEl).render(
    <ErrorBoundary>
      <SalesHeatmap />
    </ErrorBoundary>
  );
} catch (err) {
  showFatal('The app failed to start', err instanceof Error ? err.stack || err.message : String(err));
}
