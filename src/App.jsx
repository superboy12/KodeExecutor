import { useState, useRef } from 'react';
import * as docx from 'docx';
import { saveAs } from 'file-saver';
import { Play, CheckCircle2, AlertCircle, TerminalSquare } from 'lucide-react';
import './index.css';

function App() {
  const [code, setCode] = useState('');
  const [isExecuting, setIsExecuting] = useState(false);
  const [status, setStatus] = useState({ type: 'idle', message: '' });

  const handleExecute = async () => {
    if (!code.trim()) return;

    setIsExecuting(true);
    setStatus({ type: 'idle', message: 'Executing script...' });
    
    try {
      let isFileSaved = false;

      // Create fake require function to mock Node.js environment
      const fakeRequire = (moduleName) => {
        if (moduleName === 'docx') {
          return docx;
        }
        if (moduleName === 'fs') {
          return {
            writeFileSync: (filePath, data) => {
              const fileName = filePath.split(/[\\/]/).pop() || 'document.docx';
              
              // docx in browser returns Uint8Array/ArrayBuffer, which we can Blob directly
              const blob = new Blob([data], { 
                type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' 
              });
              
              saveAs(blob, fileName);
              isFileSaved = true;
              setStatus({ type: 'success', message: `Berhasil! File diunduh sebagai: ${fileName}` });
              setIsExecuting(false);
            }
          };
        }
        if (moduleName === 'path') {
          return {
            join: (...args) => args.join('/')
          };
        }
        
        console.warn(`Module "${moduleName}" is not mocked.`);
        return {};
      };

      // Create a secure wrapper to evaluate the code
      const wrapperCode = `
        return (async function(require, __dirname) {
          try {
            ${code}
          } catch(err) {
            throw err;
          }
        })(require, __dirname);
      `;

      // Execute the script
      const runner = new Function('require', '__dirname', wrapperCode);
      await runner(fakeRequire, '/executor');

      // Safety timeout in case the script is purely async but we don't catch the promise
      setTimeout(() => {
        if (isExecuting && !isFileSaved) {
            setIsExecuting(false);
        }
      }, 8000);

    } catch (error) {
      console.error("Script Execution Error:", error);
      setStatus({ type: 'error', message: error.toString() });
      setIsExecuting(false);
    }
  };

  return (
    <div className="app-container">
      <header className="header">
        <h1 className="title">
          <TerminalSquare size={40} color="#60a5fa" />
          KodeExecutor
        </h1>
        <p className="subtitle">Paste script Node.js Claude Anda di bawah ini dan saya akan mengeksekusinya untuk men-download DOCX-nya!</p>
      </header>

      <main className="editor-panel">
        <div className="textarea-wrapper">
          <textarea
            className="code-input"
            placeholder="// Paste full source code generate.js dari Claude di sini...&#10;const { Document, Packer, Paragraph } = require('docx');&#10;const fs = require('fs');&#10;..."
            value={code}
            onChange={(e) => setCode(e.target.value)}
            spellCheck="false"
          />
        </div>
        
        {status.message && status.type !== 'idle' && (
          <div className={`status-panel ${status.type}`}>
            {status.type === 'success' ? <CheckCircle2 size={20} /> : <AlertCircle size={20} />}
            <span>{status.message}</span>
          </div>
        )}

        <div className="actions">
          <button 
            className={`btn-primary ${status.type === 'success' ? 'success' : ''}`}
            onClick={handleExecute}
            disabled={!code.trim() || isExecuting}
          >
            {status.type === 'success' ? (
              <>
                <CheckCircle2 size={20} />
                Berhasil Dieksekusi!
              </>
            ) : (
              <>
                <Play size={20} />
                {isExecuting ? 'Mengeksekusi...' : 'Jalankan & Download DOCX'}
              </>
            )}
          </button>
        </div>
      </main>
    </div>
  );
}

export default App;
