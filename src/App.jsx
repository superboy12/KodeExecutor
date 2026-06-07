import { useState } from 'react';
import { Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle } from 'docx';
import { saveAs } from 'file-saver';
import { FileCode2, Download, CheckCircle2 } from 'lucide-react';
import './index.css';

function App() {
  const [code, setCode] = useState('');
  const [isGenerating, setIsGenerating] = useState(false);
  const [isSuccess, setIsSuccess] = useState(false);

  const handleGenerate = async () => {
    if (!code.trim()) return;

    setIsGenerating(true);
    
    try {
      const lines = code.split('\n');
      
      const doc = new Document({
        creator: "Code to DOCX Generator",
        title: "Source Code Document",
        sections: [
          {
            properties: {
              page: {
                margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
              }
            },
            children: [
              new Paragraph({
                text: "Source Code",
                heading: "Heading1",
                alignment: AlignmentType.CENTER,
                spacing: { after: 400 },
                border: {
                  bottom: {
                    style: BorderStyle.SINGLE,
                    size: 12,
                    color: "000000",
                    space: 4,
                  }
                }
              }),
              ...lines.map(line => {
                // Handle tabs by replacing them with 4 spaces for better rendering in Word
                const formattedLine = line.replace(/\t/g, '    ');
                return new Paragraph({
                  spacing: { before: 0, after: 0, line: 240 },
                  children: [
                    new TextRun({
                      text: formattedLine || " ", // Ensure empty lines are rendered
                      font: "Courier New",
                      size: 20, // 10pt (half-points)
                      color: "24292e"
                    })
                  ]
                });
              })
            ]
          }
        ]
      });

      const blob = await Packer.toBlob(doc);
      saveAs(blob, "SourceCode.docx");
      
      setIsSuccess(true);
      setTimeout(() => setIsSuccess(false), 3000);
    } catch (error) {
      console.error("Error generating DOCX:", error);
      alert("Failed to generate DOCX file. Check console for details.");
    } finally {
      setIsGenerating(false);
    }
  };

  return (
    <div className="app-container">
      <header className="header">
        <h1 className="title">
          <FileCode2 size={40} color="#60a5fa" />
          Code to DOCX
        </h1>
        <p className="subtitle">Paste your beautiful code below and export it to a Word document instantly.</p>
      </header>

      <main className="editor-panel">
        <div className="textarea-wrapper">
          <textarea
            className="code-input"
            placeholder="// Paste your full source code here...&#10;function helloWorld() {&#10;  console.log('Hello, DOCX!');&#10;}"
            value={code}
            onChange={(e) => setCode(e.target.value)}
            spellCheck="false"
          />
        </div>
        
        <div className="actions">
          <button 
            className={`btn-primary ${isSuccess ? 'success' : ''}`}
            onClick={handleGenerate}
            disabled={!code.trim() || isGenerating}
          >
            {isSuccess ? (
              <>
                <CheckCircle2 size={20} />
                Successfully Downloaded!
              </>
            ) : (
              <>
                <Download size={20} />
                {isGenerating ? 'Generating...' : 'Generate DOCX'}
              </>
            )}
          </button>
        </div>
      </main>
    </div>
  );
}

export default App;
