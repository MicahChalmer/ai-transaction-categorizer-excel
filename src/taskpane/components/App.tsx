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
  Field,
  Checkbox,
  Tooltip
} from "@fluentui/react-components";
import { Tag24Regular, Settings24Regular, BugRegular, CopyRegular } from "@fluentui/react-icons";
import { 
  categorizeUncategorizedTransactions, 
  setApiConfig, 
  getLastApiInteraction,
  ApiInteraction 
} from "../services/aiCategorization";

// Export environment variables for use in components
const ENV = {
  GOOGLE_API_KEY: process.env.GOOGLE_API_KEY || '',
  OPENAI_API_KEY: process.env.OPENAI_API_KEY || ''
};

// Default settings values
const DEFAULT_SETTINGS = {
  provider: "gemini" as const,
  openaiModel: "gpt-4o-mini",
  geminiModel: "gemini-2.0-flash",
  maxBatchSize: 50,
  maxReferenceTransactions: 2000,
  updateDescriptions: false
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
  openaiModel: string;
  geminiModel: string;
  maxBatchSize: number;
  maxReferenceTransactions: number;
  updateDescriptions: boolean;
}

interface ModelOption {
  id: string;
  name: string;
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
    provider: DEFAULT_SETTINGS.provider,
    openaiModel: DEFAULT_SETTINGS.openaiModel,
    geminiModel: DEFAULT_SETTINGS.geminiModel,
    maxBatchSize: DEFAULT_SETTINGS.maxBatchSize,
    maxReferenceTransactions: DEFAULT_SETTINGS.maxReferenceTransactions,
    updateDescriptions: DEFAULT_SETTINGS.updateDescriptions
  });
  
  // State for model options
  const [openaiModels, setOpenaiModels] = useState<ModelOption[]>([]);
  const [geminiModels, setGeminiModels] = useState<ModelOption[]>([]);
  const [loadingModels, setLoadingModels] = useState<boolean>(false);
  const [modelApiError, setModelApiError] = useState<string>("");
  
  const [showSettings, setShowSettings] = useState<boolean>(false);
  const [categorizationError, setCategorizationError] = useState<string>("");
  
  // Debug panel for API interactions
  const [showApiDebug, setShowApiDebug] = useState<boolean>(false);
  const [apiInteraction, setApiInteraction] = useState<ApiInteraction | null>(null);
  
  // Function to show API debug panel
  const showLastApiInteraction = () => {
    const interaction = getLastApiInteraction();
    setApiInteraction(interaction);
    setShowApiDebug(true);
  };
  
  // Copy text to clipboard
  const [copyTooltip, setCopyTooltip] = useState("Copy to clipboard");
  
  const copyToClipboard = (text: string) => {
    navigator.clipboard.writeText(text).then(() => {
      setCopyTooltip("Copied!");
      setTimeout(() => setCopyTooltip("Copy to clipboard"), 2000);
    });
  };
  
  // Format API response to remove markdown code blocks
  const formatApiResponse = (response: string): string => {
    if (!response) return "";
    
    // Remove markdown code block delimiters
    return response.replace(/```json\s*|\s*```/g, "");
  };
  
  // Apply API settings on initial load
  useEffect(() => {
    if (ENV.GOOGLE_API_KEY || ENV.OPENAI_API_KEY) {
      setApiConfig({
        openaiKey: ENV.OPENAI_API_KEY || "",
        googleKey: ENV.GOOGLE_API_KEY || "",
        provider: ENV.GOOGLE_API_KEY ? "gemini" : "openai",
        model: ENV.GOOGLE_API_KEY ? apiSettings.geminiModel : apiSettings.openaiModel,
        maxBatchSize: apiSettings.maxBatchSize,
        maxReferenceTransactions: apiSettings.maxReferenceTransactions,
        updateDescriptions: apiSettings.updateDescriptions
      });
    }
  }, []);
  
  const handleApiSettingChange = (field: keyof ApiSettings, value: string | number | boolean) => {
    // Update the local state
    setApiSettings(prev => ({
      ...prev,
      [field]: value
    }));
    
    // Prepare the updated settings
    const updatedSettings = {
      ...apiSettings,
      [field]: value
    };
    
    // Get the correct model based on provider
    const modelToUse = updatedSettings.provider === 'gemini' 
      ? updatedSettings.geminiModel 
      : updatedSettings.openaiModel;
    
    // Apply API settings
    setApiConfig({
      openaiKey: field === 'openaiKey' ? value as string : apiSettings.openaiKey,
      googleKey: field === 'googleKey' ? value as string : apiSettings.googleKey,
      provider: field === 'provider' ? value as 'gemini' | 'openai' : apiSettings.provider,
      model: modelToUse,
      maxBatchSize: field === 'maxBatchSize' ? value as number : apiSettings.maxBatchSize,
      maxReferenceTransactions: field === 'maxReferenceTransactions' ? value as number : apiSettings.maxReferenceTransactions,
      updateDescriptions: field === 'updateDescriptions' ? value as boolean : apiSettings.updateDescriptions
    });
  };

  // Function to fetch available models from OpenAI
  const fetchOpenAIModels = async () => {
    if (!apiSettings.openaiKey) {
      setNotification({
        message: "Please enter your OpenAI API key first",
        type: "warning",
        visible: true
      });
      return [];
    }
    
    setLoadingModels(true);
    setModelApiError(""); // Clear previous errors
    
    try {
      // Create a simple fetch request to OpenAI models endpoint
      const response = await fetch('https://api.openai.com/v1/models', {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${apiSettings.openaiKey}`,
          'Content-Type': 'application/json'
        }
      });
      
      const responseText = await response.text();
      
      if (!response.ok) {
        throw new Error(`API Error (${response.status}): ${responseText}`);
      }
      
      // Parse the response text as JSON
      const data = JSON.parse(responseText);
      
      // Filter for chat models only and format them
      const chatModels = data.data
        .filter((model: any) => 
          (model.id.includes('gpt') && !model.id.includes('instruct')) || 
          model.id.includes('claude')
        )
        .map((model: any) => ({
          id: model.id,
          name: model.id
        }));
      
      setOpenaiModels(chatModels);
      return chatModels;
    } catch (error) {
      console.error("Error fetching OpenAI models:", error);
      
      // Set both notification and detailed error
      setNotification({
        message: "Failed to fetch OpenAI models. See error details in settings panel.",
        type: "error",
        visible: true
      });
      
      setModelApiError(`OpenAI API Error: ${error instanceof Error ? error.message : String(error)}`);
      return [];
    } finally {
      setLoadingModels(false);
    }
  };
  
  // Function to fetch available models from Google
  const fetchGeminiModels = async () => {
    if (!apiSettings.googleKey) {
      setNotification({
        message: "Please enter your Google API key first",
        type: "warning",
        visible: true
      });
      return [];
    }
    
    setLoadingModels(true);
    setModelApiError(""); // Clear previous errors
    
    try {
      // Use the Google AI models.list API endpoint
      const response = await fetch('https://generativelanguage.googleapis.com/v1/models?key=' + apiSettings.googleKey);
      
      const responseText = await response.text();
      
      if (!response.ok) {
        throw new Error(`API Error (${response.status}): ${responseText}`);
      }
      
      // Parse the response text as JSON
      const data = JSON.parse(responseText);
      
      // Filter for Gemini models and format them
      const geminiModels = data.models
        .filter((model: any) => model.name.includes('gemini'))
        .map((model: any) => {
          const modelId = model.name.split('/').pop();
          return {
            id: modelId,
            name: modelId.replace('gemini-', 'Gemini ').replace('-', ' ')
          };
        });
      
      setGeminiModels(geminiModels);
      return geminiModels;
    } catch (error) {
      console.error("Error fetching Gemini models:", error);
      
      // Set both notification and detailed error
      setNotification({
        message: "Failed to fetch Gemini models. See error details in settings panel.",
        type: "error",
        visible: true
      });
      
      setModelApiError(`Google AI API Error: ${error instanceof Error ? error.message : String(error)}`);
      
      // Return empty array to keep the input as free-form text
      return [];
    } finally {
      setLoadingModels(false);
    }
  };
  
  // Function to fetch models based on selected provider
  const fetchModels = async () => {
    if (apiSettings.provider === 'openai') {
      return await fetchOpenAIModels();
    } else {
      return await fetchGeminiModels();
    }
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
    setCategorizationError(""); // Clear any previous errors
    
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
            // Display error message and store detailed error if available
            setNotification({
              message: "Error categorizing transactions. See details in settings panel.",
              type: "error",
              visible: true
            });
            
            if (result.errorDetails) {
              setCategorizationError(result.errorDetails);
              setShowSettings(true); // Show settings panel with error details
            } else {
              setCategorizationError(result.message || "Unknown error occurred");
            }
          }
        });
      } catch (error) {
        console.error("Error in handleAutoCategorize:", error);
        
        // Create detailed error message
        const errorMessage = error instanceof Error ? error.message : "An unknown error occurred";
        const errorStack = error instanceof Error && error.stack ? error.stack : "";
        
        setNotification({
          message: "Error categorizing transactions. See details in settings panel.",
          type: "error",
          visible: true
        });
        
        setCategorizationError(`${errorMessage}\n\n${errorStack}`);
        setShowSettings(true); // Show settings panel with error details
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
        
        <div style={{ display: 'flex', gap: '10px', marginTop: '10px' }}>
          <Button 
            appearance="subtle"
            icon={<Settings24Regular />}
            onClick={() => {
              setShowSettings(!showSettings);
              setShowApiDebug(false);
            }}
          >
            Settings
          </Button>
          
          <Button 
            appearance="subtle"
            icon={<BugRegular />}
            onClick={showLastApiInteraction}
            disabled={!getLastApiInteraction()}
            title="Show last API interaction"
          >
            Debug
          </Button>
        </div>
      </div>
      
      {showApiDebug && apiInteraction && (
        <div className={styles.settingsContainer}>
          <Divider className={styles.divider}>
            <Text>API Debug Information</Text>
            <Button
              appearance="subtle"
              size="small"
              style={{ marginLeft: '10px' }}
              onClick={() => setShowApiDebug(false)}
            >
              Close
            </Button>
          </Divider>
          
          <div style={{ marginBottom: '10px' }}>
            <Text weight="semibold">Provider:</Text> {apiInteraction.provider}
          </div>
          
          <div style={{ marginBottom: '10px' }}>
            <Text weight="semibold">Timestamp:</Text> {apiInteraction.timestamp}
          </div>
          
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', margin: '10px 0' }}>
            <Divider style={{ flexGrow: 1 }}>Request</Divider>
            <Tooltip content={copyTooltip} relationship="label">
              <Button 
                icon={<CopyRegular />} 
                appearance="subtle"
                size="small"
                onClick={() => copyToClipboard(JSON.stringify(apiInteraction.request, null, 2))}
                style={{ marginLeft: '10px' }}
              />
            </Tooltip>
          </div>
          
          <div style={{ 
            maxHeight: '200px', 
            overflow: 'auto', 
            border: '1px solid #ccc', 
            padding: '10px',
            background: '#f5f5f5',
            borderRadius: '4px',
            marginBottom: '20px'
          }}>
            <pre style={{ margin: 0, whiteSpace: 'pre-wrap' }}>
              {JSON.stringify(apiInteraction.request, null, 2)}
            </pre>
          </div>
          
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', margin: '10px 0' }}>
            <Divider style={{ flexGrow: 1 }}>Response</Divider>
            <Tooltip content={copyTooltip} relationship="label">
              <Button 
                icon={<CopyRegular />} 
                appearance="subtle"
                size="small"
                onClick={() => {
                  const content = apiInteraction.error 
                    ? JSON.stringify(apiInteraction.error, null, 2)
                    : apiInteraction.response 
                      ? formatApiResponse(apiInteraction.response) 
                      : "";
                  copyToClipboard(content);
                }}
                style={{ marginLeft: '10px' }}
              />
            </Tooltip>
          </div>
          
          <div style={{ 
            maxHeight: '200px', 
            overflow: 'auto', 
            border: '1px solid #ccc', 
            padding: '10px',
            background: '#f5f5f5',
            borderRadius: '4px'
          }}>
            <pre style={{ margin: 0, whiteSpace: 'pre-wrap' }}>
              {apiInteraction.error ? 
                JSON.stringify(apiInteraction.error, null, 2) : 
                apiInteraction.response ? formatApiResponse(apiInteraction.response) : ""
              }
            </pre>
          </div>
        </div>
      )}
      
      {showSettings && (
        <div className={styles.settingsContainer}>
          <Divider className={styles.divider}>
            <Text>API Settings</Text>
          </Divider>
          
          {(modelApiError || categorizationError) && (
            <MessageBar 
              className={styles.notification} 
              intent="error"
              style={{ 
                marginBottom: '15px', 
                whiteSpace: 'pre-wrap', 
                overflowWrap: 'break-word',
                maxHeight: '150px',
                overflowY: 'auto'
              }}
            >
              <div>
                <strong>API Error Details:</strong>
                <pre style={{ fontSize: '12px', margin: '8px 0' }}>
                  {modelApiError || categorizationError}
                </pre>
              </div>
            </MessageBar>
          )}
          
          <RadioGroup
            value={apiSettings.provider}
            onChange={(_e, data) => handleApiSettingChange('provider', data.value as string)}
          >
            <Label>AI Provider</Label>
            <Radio value="gemini" label="Google Gemini" />
            <Radio value="openai" label="OpenAI" />
          </RadioGroup>
          
          {apiSettings.provider === 'gemini' && (
            <>
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
              
              <Field label="Gemini Model" className={styles.apiKeyField}>
                {geminiModels.length > 0 ? (
                  <select
                    style={{ width: '100%', padding: '8px' }}
                    value={apiSettings.geminiModel}
                    onChange={(e) => handleApiSettingChange('geminiModel', e.target.value)}
                  >
                    {geminiModels.map(model => (
                      <option key={model.id} value={model.id}>
                        {model.name}
                      </option>
                    ))}
                  </select>
                ) : (
                  <div style={{ display: 'flex', gap: '8px' }}>
                    <Input 
                      value={apiSettings.geminiModel}
                      onChange={(_e, data) => handleApiSettingChange('geminiModel', data.value)}
                      style={{ flexGrow: 1 }}
                    />
                    <Button 
                      onClick={fetchGeminiModels}
                      disabled={loadingModels || !apiSettings.googleKey}
                    >
                      {loadingModels ? <Spinner size="tiny" /> : "Get Models"}
                    </Button>
                  </div>
                )}
              </Field>
            </>
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
                {openaiModels.length > 0 ? (
                  <select
                    style={{ width: '100%', padding: '8px' }}
                    value={apiSettings.openaiModel}
                    onChange={(e) => handleApiSettingChange('openaiModel', e.target.value)}
                  >
                    {openaiModels.map(model => (
                      <option key={model.id} value={model.id}>
                        {model.name}
                      </option>
                    ))}
                  </select>
                ) : (
                  <div style={{ display: 'flex', gap: '8px' }}>
                    <Input 
                      value={apiSettings.openaiModel}
                      onChange={(_e, data) => handleApiSettingChange('openaiModel', data.value)}
                      style={{ flexGrow: 1 }}
                    />
                    <Button 
                      onClick={fetchOpenAIModels}
                      disabled={loadingModels || !apiSettings.openaiKey}
                    >
                      {loadingModels ? <Spinner size="tiny" /> : "Get Models"}
                    </Button>
                  </div>
                )}
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
              type="text"
              value={apiSettings.maxBatchSize.toString()}
              onChange={(_e, data) => {
                const value = parseInt(data.value);
                if (!isNaN(value)) {
                  handleApiSettingChange('maxBatchSize', value || DEFAULT_SETTINGS.maxBatchSize);
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
              type="text"
              value={apiSettings.maxReferenceTransactions.toString()}
              onChange={(_e, data) => {
                const value = parseInt(data.value);
                if (!isNaN(value)) {
                  handleApiSettingChange('maxReferenceTransactions', value || DEFAULT_SETTINGS.maxReferenceTransactions);
                }
              }}
            />
          </Field>
          
          <Divider className={styles.divider}>
            <Text>Content Settings</Text>
          </Divider>
          
          <Checkbox
            label="Update transaction descriptions"
            checked={apiSettings.updateDescriptions}
            onChange={(_e, data) => handleApiSettingChange('updateDescriptions', data.checked || false)}
            style={{ marginBottom: '10px' }}
          />
          <Text size={100} style={{ color: '#666', marginLeft: '24px', marginBottom: '15px', display: 'block' }}>
            When unchecked, only categories will be updated. When checked, both categories and descriptions will be updated.
          </Text>
        </div>
      )}
    </div>
  );
};

export default App;
