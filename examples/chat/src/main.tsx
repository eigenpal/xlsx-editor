/** @jsxImportSource react */
import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import '@xlsx-editor.dev/react/styles.css';
import './styles.css';
import { App } from './App.tsx';

const container = document.getElementById('root');
if (container)
  createRoot(container).render(
    <StrictMode>
      <App />
    </StrictMode>
  );
