import React, { useState } from 'react';
import { 
  TextField, 
  PrimaryButton, 
  MessageBar, 
  MessageBarType,
  Spinner,
  SpinnerSize,
  IconButton,
  Icon
} from '@fluentui/react';
import { getSurroundingData, writeToActiveCell, writeArrayToRange } from '../services/ExcelService';
import { processWithGemini } from '../services/GeminiService';
import './MainPanel.css';

interface MainPanelProps {
  apiKey: string | null;
  onApiKeyLoaded: (key: string) => void;
}

const MainPanel: React.FC<MainPanelProps> = ({ apiKey, onApiKeyLoaded }) => {
  const [prompt, setPrompt] = useState<string>('');
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);
  const [history, setHistory] = useState<Array<{ prompt: string; result: string; timestamp: Date }>>([]);
  const [isDragOver, setIsDragOver] = useState(false);

  const handleProcess = async () => {
    if (!apiKey) {
      setError('Vui lòng nạp Key để bắt đầu.');
      return;
    }
    if (!prompt.trim()) {
      setError('Vui lòng nhập yêu cầu của bạn.');
      return;
    }

    setLoading(true);
    setError(null);

    try {
      const excelContext = await getSurroundingData();
      const result = await processWithGemini(apiKey, prompt, excelContext);
      
      if (result.type === 'array' && result.values) {
        await writeArrayToRange(result.values);
        setHistory(prev => [...prev, { 
          prompt, 
          result: `[${result.values?.length || 0} giá trị]`,
          timestamp: new Date()
        }]);
      } else if (result.type === 'single' && result.value) {
        await writeToActiveCell(result.value);
        setHistory(prev => [...prev, { 
          prompt, 
          result: result.value || '',
          timestamp: new Date()
        }]);
      }
    } catch (err: any) {
      setError(`Lỗi: ${err.message || 'Không thể xử lý yêu cầu.'}`);
    } finally {
      setLoading(false);
    }
  };

  const clearHistory = () => {
    setHistory([]);
  };

  const handleFileLoad = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (event) => {
        const content = event.target?.result as string;
        if (content) onApiKeyLoaded(content.trim());
      };
      reader.readAsText(file);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
    const file = e.dataTransfer.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (event) => {
        const content = event.target?.result as string;
        if (content) onApiKeyLoaded(content.trim());
      };
      reader.readAsText(file);
    }
  };

  return (
    <div className="main-panel">
      <div className="panel-header">
        <div className="header-title">
          <span className="title-icon">⚡</span>
          <h2>Auto Excel</h2>
        </div>
        <div className="header-subtitle">AI-Powered Excel Assistant</div>
      </div>

      {!apiKey ? (
        <div 
          className={`key-loader-zone ${isDragOver ? 'drag-over' : ''}`}
          onDragOver={(e) => { e.preventDefault(); setIsDragOver(true); }}
          onDragLeave={() => setIsDragOver(false)}
          onDrop={handleDrop}
        >
          <div className="loader-content">
            <Icon iconName="AzureKeyVault" className="loader-icon" />
            <p>Kéo thả file <code>key.txt</code> hoặc chọn file</p>
            <input 
              type="file" 
              id="keyFileInput" 
              style={{ display: 'none' }} 
              onChange={handleFileLoad}
              accept=".txt"
            />
            <PrimaryButton 
              text="Chọn file" 
              onClick={() => document.getElementById('keyFileInput')?.click()}
              className="select-key-btn"
            />
          </div>
        </div>
      ) : (
        <>
          <div className="api-status-bar">
            <Icon iconName="ShieldSolid" className="status-icon" />
            <span>OK!</span>
            <IconButton 
              iconProps={{ iconName: 'SignOut' }} 
              title="Gỡ bỏ Key" 
              onClick={() => onApiKeyLoaded('')} 
              className="eject-btn"
            />
          </div>

          {history.length > 0 && (
            <div className="history-section">
              <div className="history-header">
                <span className="section-label">History</span>
                <IconButton 
                  iconProps={{ iconName: 'Clear' }} 
                  title="Clear history"
                  onClick={clearHistory}
                  className="clear-btn"
                />
              </div>
              <div className="history-list">
                {history.map((item, index) => (
                  <div key={index} className="history-item">
                    <div className="history-time">
                      {item.timestamp.toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' })}
                    </div>
                    <div className="history-prompt">{item.prompt}</div>
                    <div className="history-result">
                      <span className="result-arrow">→</span>
                      <code>{item.result}</code>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}

          <div className="input-section">
            <label className="input-label">
              <span className="label-icon">▸</span>
              Request
            </label>
            <TextField 
              multiline 
              rows={4} 
              value={prompt} 
              onChange={(_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, val?: string) => setPrompt(val || '')} 
              placeholder="Ví dụ: Tính tổng A1:A10, Lọc cột A < 10000..."
              disabled={loading}
              className="prompt-input"
            />
          </div>
          
          {error && (
            <MessageBar 
              messageBarType={MessageBarType.error} 
              onDismiss={() => setError(null)}
              className="error-bar"
            >
              {error}
            </MessageBar>
          )}

          <PrimaryButton 
            text={loading ? "Processing..." : "Execute"} 
            onClick={handleProcess} 
            disabled={loading || !prompt}
            className="execute-btn"
          />

          {loading && (
            <div className="loading-section">
              <Spinner size={SpinnerSize.medium} />
              <span className="loading-text">Đang xử lý với AI...</span>
            </div>
          )}
        </>
      )}
    </div>
  );
};

export default MainPanel;
