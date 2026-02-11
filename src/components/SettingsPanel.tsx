import React from 'react';
import { 
  ChoiceGroup,
  type IChoiceGroupOption,
  IconButton,
  Icon
} from '@fluentui/react';
import './SettingsPanel.css';

interface SettingsPanelProps {
  theme: 'light' | 'dark';
  onThemeChange: (theme: 'light' | 'dark') => void;
  apiKey: string | null;
  onApiKeyLoaded: (key: string) => void;
}

const SettingsPanel: React.FC<SettingsPanelProps> = ({ theme, onThemeChange, apiKey, onApiKeyLoaded }) => {
  const themeOptions: IChoiceGroupOption[] = [
    { key: 'light', text: 'Light Mode', iconProps: { iconName: 'Sunny' } },
    { key: 'dark', text: 'Dark Mode', iconProps: { iconName: 'ClearNight' } },
  ];

  const handleFileLoad = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (event) => {
        const content = event.target?.result as string;
        if (content) {
          onApiKeyLoaded(content.trim());
        }
      };
      reader.readAsText(file);
    }
  };

  return (
    <div className="settings-panel">
      <div className="panel-header">
        <div className="header-title">
          <span className="title-icon">⚙️</span>
          <h2>Settings</h2>
        </div>
        <div className="header-subtitle">Config & Security</div>
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
          <div className="api-input-container">
            <div className="pocket-key-status">
              {apiKey ? (
                <div className="status-active">
                  <Icon iconName="ShieldSolid" style={{ color: 'var(--accent-green)', marginRight: 8 }} />
                  Đã nạp Chìa khóa bảo mật (RAM Mode)
                </div>
              ) : (
                <div className="status-inactive">
                  <Icon iconName="Lock" style={{ marginRight: 8 }} />
                  Chưa nạp chìa khóa
                </div>
              )}
            </div>
            <input 
              type="file" 
              id="settingsKeyFile" 
              style={{ display: 'none' }} 
              onChange={handleFileLoad}
              accept=".txt"
            />
            <IconButton 
              iconProps={{ iconName: 'OpenFolderHorizontal' }} 
              title="Kéo thả file key.txt hoặc chọn file" 
              onClick={() => document.getElementById('settingsKeyFile')?.click()}
              className="file-load-btn"
            />
            {apiKey && (
              <IconButton 
                iconProps={{ iconName: 'SignOut' }} 
                title="Gỡ bỏ Key" 
                onClick={() => onApiKeyLoaded('')} 
                className="eject-btn"
              />
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

export default SettingsPanel;
