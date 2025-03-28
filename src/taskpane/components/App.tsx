import * as React from "react";
import { useState, useEffect } from "react";
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
import { categorizeUncategorizedTransactions, setApiConfig } from "../services/aiCategorization";

// Export environment variables for use in components
const ENV = {
  GOOGLE_API_KEY: process.env.GOOGLE_API_KEY || '',
  OPENAI_API_KEY: process.env.OPENAI_API_KEY || ''
};

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
  maxBatchSize: number;
  maxReferenceTransactions: number;
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
    openaiKey: ENV.OPENAI_API_KEY || "",
    googleKey: ENV.GOOGLE_API_KEY || "",
    provider: "gemini",
    model: "gpt-4o-mini",
    maxBatchSize: 50,
    maxReferenceTransactions: 2000
  });
  
  const [showSettings, setShowSettings] = useState<boolean>(false);
  
  // Apply API settings on initial load
  useEffect(() => {
    if (ENV.GOOGLE_API_KEY || ENV.OPENAI_API_KEY) {
      setApiConfig({
        openaiKey: ENV.OPENAI_API_KEY || "",
        googleKey: ENV.GOOGLE_API_KEY || "",
        provider: ENV.GOOGLE_API_KEY ? "gemini" : "openai",
        model: apiSettings.model,
        maxBatchSize: apiSettings.maxBatchSize,
        maxReferenceTransactions: apiSettings.maxReferenceTransactions
      });
    }
  }, []);
  
  const handleApiSettingChange = (field: keyof ApiSettings, value: string | number) => {
    setApiSettings(prev => ({
      ...prev,
      [field]: value
    }));
    
    // Apply API settings
    setApiConfig({
      openaiKey: field === 'openaiKey' ? value as string : apiSettings.openaiKey,
      googleKey: field === 'googleKey' ? value as string : apiSettings.googleKey,
      provider: field === 'provider' ? value as 'gemini' | 'openai' : apiSettings.provider,
      model: field === 'model' ? value as string : apiSettings.model,
      maxBatchSize: field === 'maxBatchSize' ? value as number : apiSettings.maxBatchSize,
      maxReferenceTransactions: field === 'maxReferenceTransactions' ? value as number : apiSettings.maxReferenceTransactions
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
    setNotification({ message: "Processing transactions...", type: "info", visible: true });
    
    // Apply current API settings immediately
    setApiConfig(apiSettings);
    
    // Use setTimeout to allow the UI to update before starting the intensive operation
    setTimeout(async () => {
      try {
        await Excel.run(async (context) => {
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
    }, 50); // Small delay to allow UI to update
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
          
          <Divider className={styles.divider}>
            <Text>Performance Settings</Text>
          </Divider>
          
          <Field 
            label="Max Batch Size" 
            className={styles.apiKeyField}
            hint="Maximum number of transactions to categorize in one batch"
          >
            <Input 
              type="number"
              min="1"
              max="1000"
              value={apiSettings.maxBatchSize.toString()}
              onChange={(_e, data) => {
                const value = parseInt(data.value);
                if (!isNaN(value) && value > 0) {
                  handleApiSettingChange('maxBatchSize', value);
                }
              }}
            />
          </Field>
          
          <Field 
            label="Reference Transactions" 
            className={styles.apiKeyField}
            hint="Maximum number of previously categorized transactions to use as reference"
          >
            <Input 
              type="number"
              min="100"
              max="5000"
              value={apiSettings.maxReferenceTransactions.toString()}
              onChange={(_e, data) => {
                const value = parseInt(data.value);
                if (!isNaN(value) && value >= 100) {
                  handleApiSettingChange('maxReferenceTransactions', value);
                }
              }}
            />
          </Field>
        </div>
      )}
    </div>
  );
};

export default App;
