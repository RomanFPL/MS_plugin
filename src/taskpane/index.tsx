import React from 'react';
import { createRoot } from 'react-dom/client';
import TaskPane from './TaskPane';

// Fallback if Office is not available (for testing in browser)
const initApp = () => {
  const container = document.getElementById('root');
  if (container) {
    const root = createRoot(container);
    root.render(<TaskPane />);
  }
};

// Check if Office.js is loaded
if (typeof Office !== 'undefined' && Office.onReady) {
  Office.onReady(initApp);
} else {
  // For testing outside Office environment
  document.addEventListener('DOMContentLoaded', initApp);
}