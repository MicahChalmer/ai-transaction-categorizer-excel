import OpenAI from 'openai';
import { GoogleGenerativeAI } from '@google/generative-ai';

// API Keys - These should be set by the user at runtime
let OPENAI_API_KEY = '';
let GOOGLE_API_KEY = '';

// LLM To Use
let AI_PROVIDER = 'gemini'; // Can be 'gemini' or 'openai'
let GPT_MODEL = 'gpt-4o-mini'; // Can be any openai model designator

// Batch Processing Settings
let MAX_BATCH_SIZE = 50; // Max number of transactions to categorize in one batch
let MAX_REFERENCE_TRANSACTIONS = 2000; // Max number of reference transactions to include

// Content Settings
let UPDATE_DESCRIPTIONS = false; // Whether to update transaction descriptions or just categories

// Debugging - API interaction logs
export interface ApiInteraction {
  timestamp: string;
  provider: string;
  request: any;
  response?: any;
  error?: any;
}

let lastApiInteraction: ApiInteraction | null = null;

export function getLastApiInteraction(): ApiInteraction | null {
  return lastApiInteraction;
}

// API clients - initialized on-demand when keys are available
let openai: OpenAI | null = null;
let genAI: GoogleGenerativeAI | null = null;

// Function to set API keys and config at runtime
export function setApiConfig(config: {
  openaiKey?: string;
  googleKey?: string;
  provider?: 'gemini' | 'openai';
  model?: string;
  maxBatchSize?: number;
  maxReferenceTransactions?: number;
  updateDescriptions?: boolean;
}) {
  // Update keys and settings
  if (config.openaiKey) OPENAI_API_KEY = config.openaiKey;
  if (config.googleKey) GOOGLE_API_KEY = config.googleKey;
  if (config.provider) AI_PROVIDER = config.provider;
  if (config.model) GPT_MODEL = config.model;
  if (config.maxBatchSize) MAX_BATCH_SIZE = config.maxBatchSize;
  if (config.maxReferenceTransactions) MAX_REFERENCE_TRANSACTIONS = config.maxReferenceTransactions;
  if (config.updateDescriptions !== undefined) UPDATE_DESCRIPTIONS = config.updateDescriptions;
  
  // Reset clients so they'll be re-initialized with new keys when needed
  openai = null;
  genAI = null;
}

// Column Names
const TRANSACTION_ID_COL_NAME = "Transaction ID";
const ORIGINAL_DESCRIPTION_COL_NAME = "Full Description";
const DESCRIPTION_COL_NAME = "Description";
const CATEGORY_COL_NAME = "Category";
const DATE_COL_NAME = "Date";
const AMOUNT_COL_NAME = "Amount";
const AI_TOUCHED_COL_NAME = "AI Touched";

// Fallback Transaction Category
const FALLBACK_CATEGORY = "To Be Categorized";

// Prompting
// Shared prompt template for both APIs
function generateCategorizePrompt(categoryList: string[]): string {
  return `
    Act as an API that categorizes and cleans up bank transaction descriptions for for a personal finance app. Respond with only JSON.

    Reference the following list of allowed_categories:
    ${JSON.stringify(categoryList)}

    You will be given JSON input with a list of uncategorized transactions and a set of previously categorized reference transactions in the following format:
        {"transactions": [
          {
            "transaction_id": "A unique ID for this transaction"
            "original_description": "The original raw transaction description"
          }
        ],
        "reference_transactions": [
          {
            "original_description": "The original description of a previously categorized transaction",
            "updated_description": "The cleaned up description used previously",
            "category": "The category that was previously assigned",
            "amount": "The amount of the transaction"
          }
        ]}
        
        For EACH transaction in the transactions list, follow these instructions:
        (1) First check if there are any similar transactions in the reference_transactions list.
            If you find a good match in this list, use the same category and a similar updated_description.
            Match transactions based on merchant name, description patterns, and similar text.

            When looking in this list for good matches, ignore descriptions of the payment method such as "Zelle" or "PayPal", or generic terms like "Payment" or "Transfer"; if those words are present, look instead at
            who is being corresponded with.  For instance, when looking at a transaction described as "Zelle payment to Alice Bobson", you should look for other transactions involving
            "Alice Bobson" (even if they are not Zelle) but not match other "Zelle payment" transactions that don't involve Alice Bobson.  Same goes for "PayPal", "Check", and other 
            descriptions of payment methods rather than counterparties.
            
        (2) If there are no similar transactions that match well, suggest a better "updated_description" according to the following rules:
            (a) Use all of your knowledge to propose a friendly, human readable updated_description.
                The original_description often contains a merchant name - if you recognize it, use that merchant name.
            (b) Keep the suggested description as simple as possible. Remove punctuation, extraneous
                numbers, location information, abbreviations such as "Inc." or "LLC", IDs and account numbers.
                
        (3) For the transaction, suggest a "category" from the allowed_categories list provided.
        
        (4) If you are not confident in the suggested category after using your own knowledge and the similar transactions provided, 
            use the category "${FALLBACK_CATEGORY}".

        (4) Your response should be a JSON object and no other text.  The response object should be of the form:
        {"suggested_transactions": [
          {
            "transaction_id": "The unique ID previously provided for this transaction",
            "updated_description": "The cleaned up version of the description",
            "category": "A category selected from the allowed_categories list"
          }
        ]}
  `;
}

// Other Parameters

// Transaction interfaces
interface Transaction {
  transaction_id: string;
  original_description: string;
  amount?: number;
  date?: string;
}

interface CategorizedTransaction {
  original_description: string;
  updated_description: string;
  category: string;
  amount: number;
  date?: string;
}

interface SuggestedTransaction {
  transaction_id: string;
  updated_description: string;
  category: string;
}

interface AIError {
  message: string;
  details: string;
  source: 'api' | 'client';
  statusCode?: number;
}


// Function to look up categories and descriptions using Gemini
export async function lookupDescAndCategoryGemini(
  transactionList: Transaction[],
  categoryList: string[],
  categorizedTransactions: CategorizedTransaction[]
): Promise<SuggestedTransaction[] | null> {
  if (!GOOGLE_API_KEY) {
    throw new Error("Google API key not found. Please set it in the settings panel.");
  }

  try {
    // Initialize the API client if needed
    if (!genAI) {
      genAI = new GoogleGenerativeAI(GOOGLE_API_KEY);
    }

    const transactionDict = {
      transactions: transactionList,
      reference_transactions: categorizedTransactions,
    };

    const model = genAI.getGenerativeModel({ model: GPT_MODEL });
    
    // Get the shared prompt
    const prompt = generateCategorizePrompt(categoryList);
    
    // Record API request for debugging
    const apiRequest = {
      model: GPT_MODEL,
      prompt: prompt,
      data: transactionDict
    };

    try {
      const result = await model.generateContent({
        contents: [{ role: "user", parts: [{ text: JSON.stringify(transactionDict) }] }],
        systemInstruction: prompt,
      });
  
      const response = result.response;
      const text = response.text();
      
      // Extract JSON from the response 
      const jsonStart = text.indexOf("{");
      const jsonEnd = text.lastIndexOf("}") + 1; 
      const jsonText = text.substring(jsonStart, jsonEnd);
      
      // Record API interaction
      lastApiInteraction = {
        timestamp: new Date().toISOString(),
        provider: 'Gemini',
        request: {
          model: GPT_MODEL,
          prompt: prompt,
          data: transactionDict
        },
        response: text
      };
      
      // Parse the JSON response
      const parsedResponse = JSON.parse(jsonText);
      return parsedResponse.suggested_transactions;
    } catch (error) {
      // Record API error
      lastApiInteraction = {
        timestamp: new Date().toISOString(),
        provider: 'Gemini',
        request: {
          model: GPT_MODEL,
          prompt: prompt, 
          data: transactionDict
        },
        response: null,
        error: error instanceof Error ? { message: error.message, stack: error.stack } : error
      };
      throw error;
    }
  } catch (error) {
    console.error("Error using Gemini API:", error);
    
    // Throw detailed error to be captured in the main function
    const errorMessage = error instanceof Error ? error.message : String(error);
    const errorDetails = error instanceof Error && error.stack ? error.stack : "";
    
    const enhancedError = new Error(`Gemini API Error: ${errorMessage}`);
    enhancedError.stack = errorDetails;
    throw enhancedError;
  }
}

// Function to look up categories and descriptions using OpenAI
export async function lookupDescAndCategoryOpenAI(
  transactionList: Transaction[],
  categoryList: string[],
  categorizedTransactions: CategorizedTransaction[]
): Promise<SuggestedTransaction[] | null> {
  if (!OPENAI_API_KEY) {
    throw new Error("OpenAI API key not found. Please set it in the settings panel.");
  }

  try {
    // Initialize the OpenAI client if needed
    if (!openai) {
      openai = new OpenAI({
        apiKey: OPENAI_API_KEY,
        dangerouslyAllowBrowser: true // Required for browser environments
      });
    }

    const transactionDict = {
      transactions: transactionList,
      reference_transactions: categorizedTransactions,
    };

    // Record API request for debugging
    const apiRequest = {
      model: GPT_MODEL,
      temperature: 0.2,
      top_p: 0.1,
      seed: 1,
      data: transactionDict
    };
    
    try {
      // Get the shared prompt
      const prompt = generateCategorizePrompt(categoryList);
      
      const completion = await openai.chat.completions.create({
        model: GPT_MODEL,
        temperature: 0.2,
        top_p: 0.1,
        seed: 1,
        response_format: { type: "json_object" },
        messages: [
          {
            role: "system",
            content: prompt
          },
          {
            role: "user",
            content: JSON.stringify(transactionDict),
          }
        ]
      });

      const response = completion.choices[0].message.content;
      if (!response) {
        throw new Error("No response from OpenAI API");
      }
      
      // Record API interaction
      lastApiInteraction = {
        timestamp: new Date().toISOString(),
        provider: 'OpenAI',
        request: {
          model: GPT_MODEL,
          temperature: 0.2,
          top_p: 0.1,
          seed: 1,
          prompt: prompt,
          data: transactionDict
        },
        response: response
      };
      
      const parsedResponse = JSON.parse(response);
      return parsedResponse.suggested_transactions;
    } catch (error) {
      // Record API error
      lastApiInteraction = {
        timestamp: new Date().toISOString(),
        provider: 'OpenAI',
        request: {
          model: GPT_MODEL,
          temperature: 0.2,
          top_p: 0.1,
          seed: 1,
          prompt: generateCategorizePrompt(categoryList),
          data: transactionDict
        },
        response: null,
        error: error instanceof Error ? { message: error.message, stack: error.stack } : error
      };
      throw error;
    }
  } catch (error) {
    console.error("Error using OpenAI API:", error);
    
    // Throw detailed error to be captured in the main function
    const errorMessage = error instanceof Error ? error.message : String(error);
    const errorDetails = error instanceof Error && error.stack ? error.stack : "";
    
    const enhancedError = new Error(`OpenAI API Error: ${errorMessage}`);
    enhancedError.stack = errorDetails;
    throw enhancedError;
  }
}

// Main function to categorize transactions
export async function categorizeUncategorizedTransactions(context: Excel.RequestContext) {
  try {
    // Get the Transactions table
    const transactionsTable = context.workbook.tables.getItem("Transactions");
    await context.sync();
    
    if (!transactionsTable) {
      throw new Error("Transactions table not found in the workbook");
    }
    
    // Get the Categories table
    const categoriesTable = context.workbook.tables.getItem("Categories");
    await context.sync();
    
    if (!categoriesTable) {
      throw new Error("Categories table not found in the workbook");
    }
    
    // Get headers
    const headerRange = transactionsTable.getHeaderRowRange().load("values");
    await context.sync();
    const headers = headerRange.values[0];
    
    // Find column indices
    const idColIndex = headers.indexOf(TRANSACTION_ID_COL_NAME);
    const origDescColIndex = headers.indexOf(ORIGINAL_DESCRIPTION_COL_NAME);
    const descColIndex = headers.indexOf(DESCRIPTION_COL_NAME);
    const categoryColIndex = headers.indexOf(CATEGORY_COL_NAME);
    const aiTouchedColIndex = headers.indexOf(AI_TOUCHED_COL_NAME);
    
    if (idColIndex === -1 || origDescColIndex === -1 || descColIndex === -1 || categoryColIndex === -1) {
      throw new Error("Required columns not found in Transactions table");
    }
    
    // Don't require AI Touched column to be present, but log if it's missing
    if (aiTouchedColIndex === -1) {
      console.warn("AI Touched column not found in Transactions table");
    }
    
    // Get visible rows data
    const visibleRange = transactionsTable.getDataBodyRange().getVisibleView().load("values");
    await context.sync();
    const rows = visibleRange.values;
    
    // Find uncategorized transactions (with original description but no category)
    const uncategorizedTransactions: Transaction[] = [];
    const rowIndices: number[] = [];
    
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const origDesc = row[origDescColIndex];
      const category = row[categoryColIndex];
      
      if (origDesc && !category) {
        const amount = parseFloat(row[headers.indexOf(AMOUNT_COL_NAME)] || "0");
        const date = row[headers.indexOf(DATE_COL_NAME)];
        
        uncategorizedTransactions.push({
          transaction_id: row[idColIndex] || `row-${i}`,
          original_description: origDesc,
          amount: amount,
          date: date
        });
        rowIndices.push(i);
        
        // Limit batch size
        if (uncategorizedTransactions.length >= MAX_BATCH_SIZE) {
          break;
        }
      }
    }
    
    if (uncategorizedTransactions.length === 0) {
      return { success: true, message: "No uncategorized transactions found" };
    }
    
    // Load multiple ranges in parallel
    const allCategorizedRange = transactionsTable.getDataBodyRange().load("values");
    const categoryColRange = categoriesTable.getDataBodyRange().load("values");
    
    // Load all data with a single sync call
    await context.sync();
    
    // Process categorized transactions
    const allRows = allCategorizedRange.values || [];
    const categorizedTransactions: CategorizedTransaction[] = [];
    
    // Only process rows if we have data
    if (allRows && allRows.length > 0) {
      for (const row of allRows) {
        if (row && row[origDescColIndex] && row[categoryColIndex]) {
          categorizedTransactions.push({
            original_description: row[origDescColIndex],
            updated_description: row[descColIndex] || row[origDescColIndex],
            category: row[categoryColIndex],
            amount: parseFloat(row[headers.indexOf(AMOUNT_COL_NAME)] || "0"),
            date: row[headers.indexOf(DATE_COL_NAME)]
          });
        }
      }
    }
    
    // Limit the number of reference transactions to avoid too large requests
    // Sort by most recent first and take up to the configured maximum
    const limitedReferenceTransactions = categorizedTransactions
      .slice(0, MAX_REFERENCE_TRANSACTIONS);
    
    // Process categories
    const categoryValues = categoryColRange.values || [];
    
    const categoryList: string[] = categoryValues && categoryValues.length > 0 
      ? categoryValues.map(row => row && row[0]).filter(Boolean)
      : [];
    
    // Call AI service to get suggestions
    let suggestedTransactions: SuggestedTransaction[] | null;
    
    if (AI_PROVIDER === 'gemini') {
      suggestedTransactions = await lookupDescAndCategoryGemini(
        uncategorizedTransactions,
        categoryList,
        limitedReferenceTransactions
      );
    } else {
      suggestedTransactions = await lookupDescAndCategoryOpenAI(
        uncategorizedTransactions,
        categoryList,
        limitedReferenceTransactions
      );
    }
    
    if (!suggestedTransactions) {
      return { success: false, message: "Failed to get suggestions from AI provider" };
    }
    
    // Create a map for quick transaction lookup by ID
    const transactionMap = new Map();
    uncategorizedTransactions.forEach((tx, index) => {
      transactionMap.set(tx.transaction_id, rowIndices[index]);
    });
    
    // Process suggestions and update cells directly
    let updatedCount = 0;
    
    // Create a collection of updates to apply
    const updates: { row: number; values: { [key: number]: any } }[] = [];
    
    for (const suggestion of suggestedTransactions || []) {
      if (!suggestion || !suggestion.transaction_id) continue;
      
      const actualRowIndex = transactionMap.get(suggestion.transaction_id);
      
      if (actualRowIndex !== undefined) {
        // Validate category
        let category = suggestion.category;
        if (!categoryList.includes(category)) {
          category = FALLBACK_CATEGORY;
        }
        
        // Collect updates for this row
        const rowUpdate = {
          row: actualRowIndex,
          values: {} as { [key: number]: any }
        };
        
        // Set only the columns we want to update
        // Only update description if the setting is enabled
        if (UPDATE_DESCRIPTIONS) {
          rowUpdate.values[descColIndex] = suggestion.updated_description;
        }
        
        // Always update category
        rowUpdate.values[categoryColIndex] = category;
        
        // Always update AI Touched timestamp
        if (aiTouchedColIndex !== -1) {
          rowUpdate.values[aiTouchedColIndex] = new Date();
        }
        
        updates.push(rowUpdate);
        updatedCount++;
      }
    }
    
    // Apply all updates
    if (updatedCount > 0) {
      // Get the original body range for direct cell access
      const dataBodyRange = transactionsTable.getDataBodyRange();
      
      // Apply each update to individual cells
      for (const update of updates) {
        if (!update || typeof update.row !== 'number') continue;
        
        // For each column to update in this row
        for (const [colIndex, value] of Object.entries(update.values || {})) {
          if (value === undefined || value === null) continue;
          
          // Use the row index directly, since update.row already refers to the position in rowIndices
          const rowIdx = update.row;
          
          // Make sure we have a valid row index in our array
          if (rowIdx < 0 || rowIdx >= rowIndices.length) continue;
          
          try {
            const cell = dataBodyRange.getCell(rowIndices[rowIdx], parseInt(colIndex));
            cell.values = [[value]];
          } catch (cellError) {
            console.error("Error updating cell:", cellError);
            // Continue with other cells even if one fails
          }
        }
      }
      
      await context.sync();
      return { success: true, message: `Updated ${updatedCount} transactions` };
    } else {
      return { success: true, message: `No transactions needed updating` };
    }
  } catch (error) {
    console.error("Error in categorizeUncategorizedTransactions:", error);
    // Capture detailed error information
    const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
    const errorStack = error instanceof Error && error.stack ? error.stack : "";
    const errorDetails = `${errorMessage}\n\n${errorStack}`;
    
    return { 
      success: false, 
      message: errorMessage,
      errorDetails: errorDetails
    };
  }
}