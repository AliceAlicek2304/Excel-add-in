import React, { useState, useEffect } from 'react';
import { 
  TextField, 
  PrimaryButton, 
  MessageBar, 
  MessageBarType,
  ChoiceGroup,
  IChoiceGroupOption
} from '@fluentui/react';
import './SettingsPanel.css';

interface SettingsPanelProps {
  theme: 'light' | 'dark';
  onThemeChange: (theme: 'light' | 'dark') => void;
}

const SettingsPanel: React.FC<SettingsPanelProps> = ({ theme, onThemeChange }) => {
  const [apiKey, setApiKey] = useState<string>('');
  const [saved, setSaved] = useState<boolean>(false);

  const themeOptions: IChoiceGroupOption[] = [
    { key: 'light', text: 'Light Mode', iconProps: { iconName: 'Sunny' } },
    { key: 'dark', text: 'Dark Mode', iconProps: { iconName: 'ClearNight' } },
  ];

  useEffect(() => {
    const storedKey = localStorage.getItem('gemini_api_key');
    if (storedKey) setApiKey(storedKey);
  }, []);

  const handleSave = () => {
    localStorage.setItem('gemini_api_key', apiKey);
    setSaved(true);
    setTimeout(() => setSaved(false), 3000);
  };

  return (
    <div className="settings-panel">
      <div className="panel-header">
        <div className="header-title">
          <span className="title-icon">⚙️</span>
          <h2>Settings</h2>
        </div>
        <div className="header-subtitle">Configure your Gemini API</div>
      </div>

      <div className="settings-content">
        <div className="setting-group">
          <label className="setting-label">
            <span className="label-icon">🎨</span>
            Appearance
          </label>
          <ChoiceGroup 
            selectedKey={theme} 
            options={themeOptions} 
            onChange={(_, option) => onThemeChange(option?.key as 'light' | 'dark')} 
            className="theme-selector"
          />
        </div>

        <div className="setting-group">
          <label className="setting-label">
            <span className="label-icon">🔑</span>
            Gemini API Key
          </label>
          <TextField 
            type="password" 
            canRevealPassword 
            value={apiKey} 
            onChange={(_: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, val?: string) => setApiKey(val || '')}
            placeholder="AIza..."
            className="api-key-input"
          />
          <div className="setting-hint">
            Get your API key from <a href="https://aistudio.google.com/apikey" target="_blank" rel="noopener noreferrer">Google AI Studio</a>
          </div>
        </div>
        
        {saved && (
          <MessageBar 
            messageBarType={MessageBarType.success}
            className="success-bar"
          >
            ✓ API Key saved successfully!
          </MessageBar>
        )}

        <PrimaryButton 
          text="Save Configuration" 
          onClick={handleSave}
          className="save-btn"
        />
      </div>
    </div>
  );
};

export default SettingsPanel;
