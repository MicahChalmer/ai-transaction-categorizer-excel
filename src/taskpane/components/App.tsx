import * as React from "react";
import { useState } from "react";
import TextInsertion from "./TextInsertion";
import { 
  makeStyles, 
  Button, 
  Text, 
  Spinner, 
  MessageBar, 
  Input, 
  Label, 
  Radio, 
  RadioGroup,
  Divider,
  Field
} from "@fluentui/react-components";
import { Tag24Regular, Settings24Regular } from "@fluentui/react-icons";
import { insertText } from "../taskpane";
import { categorizeUncategorizedTransactions, setApiConfig } from "../services/aiCategorization";

interface AppProps {
  title: string;
}

type MessageBarIntent = "success" | "error" | "warning" | "info";

interface NotificationState {
  message: string;
  type: MessageBarIntent;
  visible: boolean;
}

interface ApiSettings {
  openaiKey: string;
  googleKey: string;
  provider: 'gemini' | 'openai';
  model: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    display: "flex",
    flexDirection: "column",
    padding: "20px",
  },
  buttonContainer: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    margin: "20px 0",
  },
  autoCatButton: {
    marginTop: "20px",
    width: "200px",
  },
  notification: {
    margin: "10px 0",
  },
  settingsContainer: {
    margin: "20px 0",
  },
  apiKeyField: {
    marginBottom: "10px",
  },
  divider: {
    margin: "20px 0",
  }
});

const App: React.FC<AppProps> = (_props: AppProps) => {
  const styles = useStyles();
  const [isLoading, setIsLoading] = useState(false);
  const [notification, setNotification] = useState<NotificationState>({
    message: "",
    type: "info",
    visible: false
  });
  
  const [apiSettings, setApiSettings] = useState<ApiSettings>({
    openaiKey: "",
    googleKey: "",
    provider: "gemini",
    model: "gpt-4o-mini"
  });
  
  const [showSettings, setShowSettings] = useState<boolean>(false);
  
  const handleApiSettingChange = (field: keyof ApiSettings, value: string) => {
    setApiSettings(prev => ({
      ...prev,
      [field]: value
    }));
    
    // Apply API settings
    setApiConfig({
      openaiKey: field === 'openaiKey' ? value : apiSettings.openaiKey,
      googleKey: field === 'googleKey' ? value : apiSettings.googleKey,
      provider: field === 'provider' ? value as 'gemini' | 'openai' : apiSettings.provider,
      model: field === 'model' ? value : apiSettings.model
    });
  };

  const handleAutoCategorize = async () => {
    // Check if API keys are set based on provider
    if ((apiSettings.provider === 'gemini' && !apiSettings.googleKey) || 
        (apiSettings.provider === 'openai' && !apiSettings.openaiKey)) {
      setNotification({
        message: `Please enter your ${apiSettings.provider === 'gemini' ? 'Google' : 'OpenAI'} API key in settings`,
        type: "warning",
        visible: true
      });
      setShowSettings(true);
      return;
    }
    
    setIsLoading(true);
    setNotification({ message: "", type: "info", visible: false });
    
    try {
      await Excel.run(async (context) => {
        // Apply current API settings
        setApiConfig(apiSettings);
        
        // Run the categorization function
        const result = await categorizeUncategorizedTransactions(context);
        
        if (result.success) {
          setNotification({
            message: result.message,
            type: "success",
            visible: true
          });
        } else {
          setNotification({
            message: result.message,
            type: "error",
            visible: true
          });
        }
      });
    } catch (error) {
      console.error("Error in handleAutoCategorize:", error);
      setNotification({
        message: error instanceof Error ? error.message : "An unknown error occurred",
        type: "error",
        visible: true
      });
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className={styles.root}>
      {notification.visible && (
        <MessageBar className={styles.notification} intent={notification.type}>
          {notification.message}
        </MessageBar>
      )}
      
      <div className={styles.buttonContainer}>
        <Text weight="semibold">Auto-categorize your transactions using AI</Text>
        <Button 
          className={styles.autoCatButton}
          appearance="primary"
          icon={<Tag24Regular />}
          onClick={handleAutoCategorize}
          disabled={isLoading}
        >
          {isLoading ? <Spinner size="tiny" /> : "AI Auto-Categorize"}
        </Button>
        
        <Button 
          style={{ marginTop: '10px' }}
          appearance="subtle"
          icon={<Settings24Regular />}
          onClick={() => setShowSettings(!showSettings)}
        >
          Settings
        </Button>
      </div>
      
      {showSettings && (
        <div className={styles.settingsContainer}>
          <Divider className={styles.divider}>
            <Text>API Settings</Text>
          </Divider>
          
          <RadioGroup
            value={apiSettings.provider}
            onChange={(_e, data) => handleApiSettingChange('provider', data.value as string)}
          >
            <Label>AI Provider</Label>
            <Radio value="gemini" label="Google Gemini" />
            <Radio value="openai" label="OpenAI" />
          </RadioGroup>
          
          {apiSettings.provider === 'gemini' && (
            <Field 
              label="Google API Key" 
              className={styles.apiKeyField}
              validationMessage={!apiSettings.googleKey ? "Required" : undefined}
            >
              <Input 
                type="password"
                value={apiSettings.googleKey}
                onChange={(_e, data) => handleApiSettingChange('googleKey', data.value)}
              />
            </Field>
          )}
          
          {apiSettings.provider === 'openai' && (
            <>
              <Field 
                label="OpenAI API Key" 
                className={styles.apiKeyField}
                validationMessage={!apiSettings.openaiKey ? "Required" : undefined}
              >
                <Input 
                  type="password"
                  value={apiSettings.openaiKey}
                  onChange={(_e, data) => handleApiSettingChange('openaiKey', data.value)}
                />
              </Field>
              
              <Field label="OpenAI Model" className={styles.apiKeyField}>
                <Input 
                  value={apiSettings.model}
                  onChange={(_e, data) => handleApiSettingChange('model', data.value)}
                />
              </Field>
            </>
          )}
        </div>
      )}
      
      <TextInsertion insertText={insertText} />
    </div>
  );
};

export default App;
