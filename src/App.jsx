import { useState, useRef } from 'react';
import * as docx from 'docx';
import PptxGenJS from 'pptxgenjs';
import { saveAs } from 'file-saver';
import JSZip from 'jszip';
import { Play, CheckCircle2, AlertCircle, TerminalSquare, FileText, FileDown, Archive } from 'lucide-react';
import './index.css';

// Helper function to format file sizes
function formatBytes(bytes, decimals = 2) {
  if (!+bytes) return '0 Bytes';
  const k = 1024;
  const dm = decimals < 0 ? 0 : decimals;
  const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return `${parseFloat((bytes / Math.pow(k, i)).toFixed(dm))} ${sizes[i]}`;
}

function App() {
  const [code, setCode] = useState('');
  const [isExecuting, setIsExecuting] = useState(false);
  const [status, setStatus] = useState({ type: 'idle', message: '' });
  const [logs, setLogs] = useState([]);
  const [downloadedFiles, setDownloadedFiles] = useState([]);

  const appendLog = (msg) => {
    setLogs(prev => [...prev, String(msg)]);
  };

  const handleExecute = async () => {
    if (!code.trim()) return;

    setIsExecuting(true);
    setStatus({ type: 'idle', message: 'Mengeksekusi script...' });
    setLogs([]);
    setDownloadedFiles([]);

    const filesDownloaded = [];

    try {
      // Intercept console.log/error/warn for display
      const fakeConsole = {
        log: (...args) => appendLog(args.join(' ')),
        error: (...args) => appendLog('[ERROR] ' + args.join(' ')),
        warn: (...args) => appendLog('[WARN] ' + args.join(' ')),
      };

      // Create fake require function to mock Node.js environment
      const fakeRequire = (moduleName) => {
        if (moduleName === 'pptxgenjs') {
          return class MockPptxGenJS extends PptxGenJS {
            constructor() {
              super();
            }
            async writeFile(options) {
              const fileName = (options && options.fileName) || 'presentation.pptx';
              // Intercept writeFile to track it in our UI
              const blob = await this.write('blob');
              saveAs(blob, fileName);
              filesDownloaded.push({ name: fileName, blob, size: blob.size });
              setDownloadedFiles([...filesDownloaded]);
              appendLog(`✅ File diunduh: ${fileName}`);
              return fileName;
            }
          };
        }
        if (moduleName === 'docx') {
          return {
            ...docx,
            Packer: {
              ...docx.Packer,
              toBuffer: async (doc) => {
                // In browser, use toBlob instead of toBuffer
                return docx.Packer.toBlob(doc);
              }
            }
          };
        }
        if (moduleName === 'fs') {
          return {
            writeFileSync: (filePath, data) => {
              const fileName = filePath.split(/[/\\]/).pop() || 'document.docx';
              
              let mimeType = 'application/octet-stream';
              if (fileName.endsWith('.docx')) {
                mimeType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
              } else if (fileName.endsWith('.pptx')) {
                mimeType = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
              } else if (fileName.endsWith('.xlsx')) {
                mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
              }

              const blob = data instanceof Blob
                ? data
                : new Blob([data], { type: mimeType });

              saveAs(blob, fileName);
              filesDownloaded.push({ name: fileName, blob, size: blob.size });
              setDownloadedFiles([...filesDownloaded]);
              appendLog(`✅ File diunduh: ${fileName}`);
            }
          };
        }
        if (moduleName === 'path') {
          return {
            join: (...args) => args.join('/')
          };
        }

        fakeConsole.warn(`Module "${moduleName}" tidak tersedia di browser.`);
        return {};
      };

      // Preprocess: add 'await' to the last top-level function call in the script.
      const lines = code.split('\n');
      let patched = false;
      for (let i = lines.length - 1; i >= 0 && !patched; i--) {
        const trimmed = lines[i].trim();
        // Match a standalone function call line (not a definition, not inside a block)
        if (
          trimmed &&
          /^[a-zA-Z_$][a-zA-Z0-9_$]*\s*\(/.test(trimmed) &&
          !trimmed.startsWith('//') &&
          !trimmed.startsWith('*') &&
          !trimmed.startsWith('function') &&
          !trimmed.startsWith('const ') &&
          !trimmed.startsWith('let ') &&
          !trimmed.startsWith('var ') &&
          !trimmed.startsWith('return ') &&
          !trimmed.startsWith('await ')
        ) {
          lines[i] = lines[i].replace(/^(\s*)/, '$1await ');
          patched = true;
          appendLog(`ℹ️ Auto-await ditambahkan pada: ${trimmed}`);
        }
      }
      let processedCode = lines.join('\n');

      // Add simple auto-fix for deprecated docx syntax
      if (/new\s+(?:docx\.)?PageNumber\s*\(\s*\)/i.test(processedCode)) {
        processedCode = processedCode.replace(/new\s+docx\.PageNumber\s*\(\s*\)/gi, 'docx.PageNumber.CURRENT');
        processedCode = processedCode.replace(/new\s+PageNumber\s*\(\s*\)/gi, 'PageNumber.CURRENT');
        appendLog(`ℹ️ Auto-fix: Mengubah 'new PageNumber()' menjadi 'PageNumber.CURRENT' untuk kompatibilitas docx v9+`);
      }

      // Wrap the code, inject fake console and process.exit
      const wrapperCode = `
        return (async function(require, __dirname, console, process) {
          try {
            ${processedCode}
          } catch(err) {
            throw err;
          }
        })(require, __dirname, fakeConsole, { exit: () => {} });
      `;

      const runner = new Function('require', '__dirname', 'fakeConsole', wrapperCode);
      await runner(fakeRequire, '/executor', fakeConsole);

      if (filesDownloaded.length > 0) {
        setStatus({
          type: 'success',
          message: `Berhasil! ${filesDownloaded.length} file DOCX telah di-generate.`
        });
      } else {
        setStatus({
          type: 'error',
          message: 'Script selesai dijalankan, tapi tidak ada file yang diunduh. Pastikan ada pemanggilan fs.writeFileSync().'
        });
      }

    } catch (error) {
      console.error('Script Execution Error:', error);
      setStatus({ type: 'error', message: error.toString() });
    } finally {
      setIsExecuting(false);
    }
  };

  const handleDownloadZip = async () => {
    try {
      setStatus(prev => ({ ...prev, message: 'Membuat file ZIP...' }));
      const JSZip = (await import('jszip')).default;
      const zip = new JSZip();
      
      downloadedFiles.forEach(f => {
        zip.file(f.name, f.blob);
      });
      
      const content = await zip.generateAsync({ type: "blob" });
      saveAs(content, "documents.zip");
      setStatus({ type: 'success', message: `Berhasil! ZIP berisi ${downloadedFiles.length} file telah diunduh.` });
    } catch (error) {
      console.error("ZIP Error:", error);
      setStatus({ type: 'error', message: `Gagal membuat ZIP: ${error.message}` });
    }
  };

  const handleDownloadSingle = (file) => {
    saveAs(file.blob, file.name);
  };

  const handleReset = () => {
    setStatus({ type: 'idle', message: '' });
    setLogs([]);
    setDownloadedFiles([]);
  };

  return (
    <div className="app-container">
      <header className="header">
        <h1 className="title">
          <TerminalSquare size={40} color="#60a5fa" />
          KodeExecutor
        </h1>
        <p className="subtitle">
          Paste script Node.js Claude Anda di bawah ini dan saya akan mengeksekusinya untuk men-download DOCX-nya!
        </p>
      </header>

      <main className="editor-panel">
        <div className="textarea-wrapper">
          <textarea
            className="code-input"
            placeholder={"// Paste full source code generate.js dari Claude di sini...\nconst { Document, Packer, Paragraph } = require('docx');\nconst fs = require('fs');\n..."}
            value={code}
            onChange={(e) => setCode(e.target.value)}
            spellCheck="false"
          />
        </div>

        {/* Status panel */}
        {status.message && status.type !== 'idle' && (
          <div className={`status-panel ${status.type}`}>
            {status.type === 'success' ? <CheckCircle2 size={20} /> : <AlertCircle size={20} />}
            <span>{status.message}</span>
          </div>
        )}

        {/* Downloaded files list */}
        {downloadedFiles.length > 0 && (
          <div className="files-list">
            <div className="files-header">
              <p className="files-title">Dokumen yang Dihasilkan ({downloadedFiles.length}):</p>
              {downloadedFiles.length > 1 && (
                <button className="btn-zip" onClick={handleDownloadZip}>
                  <Archive size={14} /> Download Semua (.ZIP)
                </button>
              )}
            </div>
            
            {downloadedFiles.map((f, i) => (
              <div key={i} className="file-item">
                <div className="file-info">
                  <FileText size={18} color="#10b981" />
                  <span className="file-name" title={f.name}>{f.name}</span>
                  <span className="file-size">{formatBytes(f.size)}</span>
                </div>
                <button 
                  className="btn-icon" 
                  onClick={() => handleDownloadSingle(f)} 
                  title="Download Ulang"
                >
                  <FileDown size={18} />
                </button>
              </div>
            ))}
          </div>
        )}

        {/* Console log output */}
        {logs.length > 0 && (
          <div className="log-panel">
            <p className="log-title">Output Console:</p>
            <div className="log-content">
              {logs.map((l, i) => (
                <div key={i} className="log-line">{l}</div>
              ))}
            </div>
          </div>
        )}

        <div className="actions">
          {status.type !== 'idle' && (
            <button className="btn-secondary" onClick={handleReset}>
              Reset
            </button>
          )}
          <button
            className={`btn-primary ${status.type === 'success' ? 'success' : ''}`}
            onClick={handleExecute}
            disabled={!code.trim() || isExecuting}
          >
            {status.type === 'success' ? (
              <>
                <CheckCircle2 size={20} />
                Eksekusi Ulang
              </>
            ) : (
              <>
                <Play size={20} />
                {isExecuting ? 'Mengeksekusi...' : 'Jalankan Script'}
              </>
            )}
          </button>
        </div>
      </main>
    </div>
  );
}

export default App;
