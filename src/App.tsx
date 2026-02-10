import React, { useState, useEffect } from 'react';
import { Pivot, PivotItem, Stack } from '@fluentui/react';
import MainPanel from './components/MainPanel';
import SettingsPanel from './components/SettingsPanel';

const App: React.FC = () => {
  const [selectedKey, setSelectedKey] = useState<string>('process');
  const [theme, setTheme] = useState<'light' | 'dark'>(() => {
    return (localStorage.getItem('theme') as 'light' | 'dark') || 'dark';
  });

  useEffect(() => {
    document.body.className = `theme-${theme}`;
    localStorage.setItem('theme', theme);
  }, [theme]);

  return (
    <Stack style={{ height: '100vh', width: '100%', overflow: 'hidden', backgroundColor: 'var(--bg-primary)' }}>
      <Stack.Item grow style={{ overflow: 'auto' }}>
        <Pivot 
          selectedKey={selectedKey} 
          onLinkClick={(item?: PivotItem) => setSelectedKey(item?.props.itemKey || 'process')}
          styles={{ 
            root: { 
              padding: '0 10px',
              backgroundColor: 'var(--bg-secondary)',
              borderBottom: '1px solid var(--border-color)'
            },
            link: {
              backgroundColor: 'transparent',
              color: 'var(--text-secondary)',
              selectors: {
                ':hover': {
                  backgroundColor: 'var(--bg-tertiary)',
                  color: 'var(--text-primary)'
                },
                ':active': {
                  backgroundColor: 'var(--bg-tertiary)'
                }
              }
            },
            linkIsSelected: {
              backgroundColor: 'transparent',
              color: 'var(--accent-blue)',
              selectors: {
                ':hover': {
                  backgroundColor: 'var(--bg-tertiary)',
                  color: 'var(--accent-blue)'
                },
                '::before': {
                  backgroundColor: 'var(--accent-blue)',
                  height: '2px'
                }
              }
            }
          }}
        >
          <PivotItem headerText="AI Process" itemKey="process">
            <MainPanel />
          </PivotItem>
          <PivotItem headerText="Settings" itemKey="settings">
            <SettingsPanel theme={theme} onThemeChange={setTheme} />
          </PivotItem>
        </Pivot>
      </Stack.Item>
    </Stack>
  );
};

export default App;
