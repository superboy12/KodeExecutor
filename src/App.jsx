import { useState, useRef } from 'react';
import * as docx from 'docx';
import PptxGenJS from 'pptxgenjs';
import { saveAs } from 'file-saver';
import JSZip from 'jszip';
import { Play, CheckCircle2, AlertCircle, TerminalSquare, FileText, FileDown, Archive, Image, Eye, EyeOff, Code2, RotateCcw, ExternalLink } from 'lucide-react';
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
  const [svgPreviews, setSvgPreviews] = useState({});

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
              } else if (fileName.endsWith('.svg')) {
                mimeType = 'image/svg+xml';
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

      // Calculate brace depth at the start of each line, skipping braces inside strings
      const braceDepthAtLine = [];
      let depth = 0;
      for (let i = 0; i < lines.length; i++) {
        braceDepthAtLine[i] = depth;
        let inString = false;
        let stringChar = '';
        let escaped = false;
        for (let j = 0; j < lines[i].length; j++) {
          const ch = lines[i][j];
          if (escaped) { escaped = false; continue; }
          if (ch === '\\') { escaped = true; continue; }
          if (inString) {
            if (ch === stringChar) inString = false;
            continue;
          }
          if (ch === '"' || ch === "'" || ch === '`') {
            inString = true;
            stringChar = ch;
          } else if (ch === '{') {
            depth++;
          } else if (ch === '}') {
            depth--;
          }
        }
      }

      for (let i = lines.length - 1; i >= 0 && !patched; i--) {
        const trimmed = lines[i].trim();
        // Only patch top-level statements (brace depth 0)
        if (braceDepthAtLine[i] !== 0) continue;
        // Match a standalone function call line, including dotted names like Packer.toBuffer()
        if (
          trimmed &&
          /^[a-zA-Z_$][a-zA-Z0-9_$.]*\s*\(/.test(trimmed) &&
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
          message: `Berhasil! ${filesDownloaded.length} file telah di-generate.`
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
    // Revoke SVG preview URLs to prevent memory leaks
    Object.values(svgPreviews).forEach(url => URL.revokeObjectURL(url));
    setStatus({ type: 'idle', message: '' });
    setLogs([]);
    setDownloadedFiles([]);
    setSvgPreviews({});
  };

  const isSvgFile = (fileName) => fileName.toLowerCase().endsWith('.svg');

  const toggleSvgPreview = (index, file) => {
    setSvgPreviews(prev => {
      if (prev[index]) {
        // Close preview, revoke URL
        URL.revokeObjectURL(prev[index]);
        const next = { ...prev };
        delete next[index];
        return next;
      } else {
        // Open preview
        const url = URL.createObjectURL(file.blob);
        return { ...prev, [index]: url };
      }
    });
  };

  return (
    <>
      {/* ===== NAVBAR ===== */}
      <nav className="navbar">
        <div className="navbar-brand">
          <div className="navbar-logo">
            <Code2 size={18} />
          </div>
          <span className="navbar-title">KodeExecutor</span>
          <span className="navbar-badge">Beta</span>
        </div>
        <div className="navbar-links">
          <div className="feature-badges">
            <span className="badge badge-docx">DOCX</span>
            <span className="badge badge-pptx">PPTX</span>
            <span className="badge badge-svg">SVG</span>
          </div>
          <a className="navbar-link" href="https://github.com/superboy12/KodeExecutor" target="_blank" rel="noopener noreferrer">
            <ExternalLink size={14} />
            <span>GitHub</span>
          </a>
        </div>
      </nav>

      {/* ===== MAIN LAYOUT ===== */}
      <div className="app-container">
        <div className="main-content">

          {/* ===== EDITOR PANE ===== */}
          <div className="editor-pane">
            <div className="pane-header">
              <div className="pane-title">
                <TerminalSquare size={14} />
                Script Editor
              </div>
              <div className="pane-actions">
                {status.type !== 'idle' && (
                  <button className="btn-reset" onClick={handleReset}>
                    <RotateCcw size={14} />
                    Reset
                  </button>
                )}
                <button
                  className={`btn-execute ${status.type === 'success' ? 'success' : ''}`}
                  onClick={handleExecute}
                  disabled={!code.trim() || isExecuting}
                >
                  {isExecuting ? (
                    <>
                      <div className="spinner" />
                      Mengeksekusi...
                    </>
                  ) : status.type === 'success' ? (
                    <>
                      <CheckCircle2 size={15} />
                      Eksekusi Ulang
                    </>
                  ) : (
                    <>
                      <Play size={15} />
                      Jalankan
                    </>
                  )}
                </button>
              </div>
            </div>

            <textarea
              className="code-input"
              placeholder={"// Paste script Node.js dari Claude di sini...\n\nconst { Document, Packer, Paragraph } = require('docx');\nconst fs = require('fs');\n\n// Script Anda akan dieksekusi di browser"}
              value={code}
              onChange={(e) => setCode(e.target.value)}
              spellCheck="false"
            />

            <div className="editor-footer">
              <div className="editor-hint">
                Supports <kbd>docx</kbd> <kbd>pptxgenjs</kbd> <kbd>fs</kbd> <kbd>path</kbd>
              </div>
              <span className="char-count">{code.length > 0 ? `${code.split('\n').length} lines` : ''}</span>
            </div>
          </div>

          {/* ===== OUTPUT PANE ===== */}
          <div className="output-pane">
            <div className="pane-header">
              <div className="pane-title">
                <FileDown size={14} />
                Output
              </div>
            </div>

            {/* Empty state */}
            {status.type === 'idle' && downloadedFiles.length === 0 && logs.length === 0 && (
              <div className="output-empty">
                <div className="output-empty-icon">
                  <Play size={28} color="var(--text-tertiary)" />
                </div>
                <h3>Belum ada output</h3>
                <p>Paste script di editor lalu klik Jalankan untuk melihat hasilnya di sini.</p>
              </div>
            )}

            {/* Has content */}
            {(status.type !== 'idle' || downloadedFiles.length > 0 || logs.length > 0) && (
              <div className="output-content">

                {/* Status bar */}
                {status.message && status.type !== 'idle' && (
                  <div className={`status-bar ${status.type}`}>
                    {status.type === 'success' ? <CheckCircle2 size={16} /> : <AlertCircle size={16} />}
                    <span>{status.message}</span>
                  </div>
                )}

                {/* Downloaded files */}
                {downloadedFiles.length > 0 && (
                  <div className="files-section">
                    <div className="section-header">
                      <div className="section-title">
                        File yang Dihasilkan
                        <span className="section-count">{downloadedFiles.length}</span>
                      </div>
                      {downloadedFiles.length > 1 && (
                        <button className="btn-zip" onClick={handleDownloadZip}>
                          <Archive size={13} /> Download ZIP
                        </button>
                      )}
                    </div>

                    <div className="files-grid">
                      {downloadedFiles.map((f, i) => (
                        <div key={i} className="file-item-wrapper">
                          <div className="file-item">
                            <div className="file-info">
                              <div className={`file-icon ${isSvgFile(f.name) ? 'svg' : 'doc'}`}>
                                {isSvgFile(f.name) ? <Image size={16} /> : <FileText size={16} />}
                              </div>
                              <div className="file-meta">
                                <span className="file-name" title={f.name}>{f.name}</span>
                                <span className="file-size">{formatBytes(f.size)}</span>
                              </div>
                            </div>
                            <div className="file-actions">
                              {isSvgFile(f.name) && (
                                <button
                                  className="btn-icon btn-preview"
                                  onClick={() => toggleSvgPreview(i, f)}
                                  title={svgPreviews[i] ? 'Tutup Preview' : 'Preview SVG'}
                                >
                                  {svgPreviews[i] ? <EyeOff size={16} /> : <Eye size={16} />}
                                </button>
                              )}
                              <button
                                className="btn-icon"
                                onClick={() => handleDownloadSingle(f)}
                                title="Download"
                              >
                                <FileDown size={16} />
                              </button>
                            </div>
                          </div>
                          {svgPreviews[i] && (
                            <div className="svg-preview">
                              <img src={svgPreviews[i]} alt={f.name} />
                            </div>
                          )}
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {/* Console logs */}
                {logs.length > 0 && (
                  <div className="log-section">
                    <div className="section-title">
                      Console
                      <span className="section-count">{logs.length}</span>
                    </div>
                    <div className="log-content">
                      {logs.map((l, i) => (
                        <div key={i} className="log-line">{l}</div>
                      ))}
                    </div>
                  </div>
                )}

              </div>
            )}
          </div>

        </div>
      </div>
    </>
  );
}

export default App;
