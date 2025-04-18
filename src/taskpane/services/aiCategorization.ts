import OpenAI from 'openai';
import { GenerateContentParameters, GoogleGenAI } from '@google/genai';
import { ChatCompletionCreateParams } from 'openai/resources';

// API Keys - These should be set by the user at runtime
let OPENAI_API_KEY = '';
let GOOGLE_API_KEY = '';

// LLM To Use
let AI_PROVIDER = 'openai'; // Can be 'gemini' or 'openai'
let GPT_MODEL = 'gpt-4.1-mini'; // Can be any openai model designator

// Batch Processing Settings
let MAX_BATCH_SIZE = 50; // Max number of transactions to categorize in one batch
let MAX_REFERENCE_TRANSACTIONS = 5000; // Max number of reference transactions to include

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
let genAI: GoogleGenAI | null = null;

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
const INSTITUTION_COL_NAME = "Institution";

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
        {
          "transactions": [
            {
              "transaction_id": "A unique ID for this transaction",
              "original_description": "The original raw transaction description",
              "institution": "The financial institution for this transaction (bank or credit card)"
            }
          ],
          "reference_transactions": [
            {
              "transaction_id": "A unique ID for this transaction",
              "original_description": "The original description of a previously categorized transaction",
              "updated_description": "The cleaned up description used previously",
              "category": "The category that was previously assigned",
              "amount": "The amount of the transaction",
              "institution": "The financial institution for this transaction (bank or credit card)"
            }
          ]
        }
        
        For EACH transaction in the transactions list, follow these instructions:
        (1) First check if there are any similar transactions in the reference_transactions list.  This is a list of prior transactions that have already been categorized.
            The goal is to categorize the "transactions" list in a similar way to the ones in the "reference_transactions" list.
            A similar transaction is one with a similar description.  If you find a good match, use the same category and a similar updated_description for this transaction.
            
            If the transaction's description is very generic like just "Payment" or "Transfer", use the institution to help match, looking for similar transactions from the same institution.

            When looking in this list for good matches, ignore descriptions of the payment method such as "Zelle" or "PayPal", or generic terms like "Payment" or "Transfer"; if those words are present, look instead at
            who is being corresponded with.  For instance, when looking at a transaction described as "Zelle payment to Alice Bobson", you should look for other transactions involving
            "Alice Bobson" (even if they are not Zelle) but not match other "Zelle payment" transactions that don't involve Alice Bobson.  Same goes for "PayPal", "Check", and other 
            descriptions of payment methods rather than counterparties.

            On the other hand, you should still find a match if there are slight differnces in the descriptions, particularly strings with numbers at the end or differences in only whitespace.
            For instance, "Alpha Bravo Charlie x737878" could match wtih "ALPHA   BRAVO    CHARLIE XXXXXX34345".
            
            If you find a transaction from the similar_transactions that matches, include the transaction_id from the matched reference transaction in the this transaction's matched_transaction_id field.
            
        (2) If there are no similar transactions that match well, suggest a better "updated_description" according to the following rules:
            (a) Use all of your knowledge to propose a friendly, human readable updated_description.
                The original_description often contains a merchant name - if you recognize it, use that merchant name.
            (b) Keep the suggested description as simple as possible. Remove punctuation, extraneous
                numbers, location information, abbreviations such as "Inc." or "LLC", IDs and account numbers.
                
        (3) For the transaction, suggest a "category" from the allowed_categories list provided.
        
        (4) If you are not confident in the suggested category after using your own knowledge and the similar transactions provided, 
            use the category "${FALLBACK_CATEGORY}".

        (4) Your response should be a JSON object and no other text.  The response object should be of the form:
        {
          "suggested_transactions": [
            {
              "transaction_id": "The unique ID previously provided for this transaction",
              "updated_description": "The cleaned up version of the description",
              "category": "A category selected from the allowed_categories list",
              "matched_transaction_id": "The transaction_id of the matching reference transaction found for this one, or null if no match was found.  If provided, this must be a matching transaction from the reference_transactions list, not the same ID as this transaction from the transactions list."
            }
          ]
        }
  `;
}

// Other Parameters

// Transaction interfaces
interface Transaction {
  transaction_id: string;
  original_description: string;
  amount?: number;
  date?: string;
  institution?: string;
}

interface CategorizedTransaction {
  transaction_id: string;
  original_description: string;
  updated_description: string;
  category: string;
  amount: number;
  date?: string;
  institution?: string;
}

interface SuggestedTransaction {
  transaction_id: string;
  updated_description: string;
  category: string;
  matched_transaction_id?: string;
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
      genAI = new GoogleGenAI({apiKey: GOOGLE_API_KEY});
    }

    const transactionDict = {
      transactions: transactionList,
      reference_transactions: categorizedTransactions,
    };

    // Get the shared prompt
    const prompt = generateCategorizePrompt(categoryList);
    
  const geminiRequest: GenerateContentParameters = {
    model: GPT_MODEL,
    contents: [
      { role: "model", parts: [{text: prompt}]},
      { role: "user", parts: [{ text: JSON.stringify(transactionDict) }] }
    ],
  };

  const requestForDebug = {...geminiRequest, contents: [{...geminiRequest.contents[0], parts: [{ text: transactionDict}]}]};
  try {
      const response = await genAI.models.generateContent(geminiRequest);
  
      const text = response.text;
      
      // Extract JSON from the response 
      const jsonStart = text.indexOf("{");
      const jsonEnd = text.lastIndexOf("}") + 1; 
      const jsonText = text.substring(jsonStart, jsonEnd);
      
      // Record API interaction
      lastApiInteraction = {
        timestamp: new Date().toISOString(),
        provider: 'Gemini',
        request: requestForDebug,
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
        request: requestForDebug,
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

  // Get the shared prompt
  const prompt = generateCategorizePrompt(categoryList);
  const openAiRequest: ChatCompletionCreateParams = {
    model: GPT_MODEL,
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
  };
  const requestForDebug = {...openAiRequest, messages: [openAiRequest.messages[0], {...openAiRequest.messages[1], content: transactionDict}]}
  try {
      
      const completion = await openai.chat.completions.create(openAiRequest);

      const response = completion.choices[0].message.content;
      if (!response) {
        throw new Error("No response from OpenAI API");
      }
      
      // Record API interaction
      lastApiInteraction = {
        timestamp: new Date().toISOString(),
        provider: 'OpenAI',
        request: requestForDebug,
        response: response
      };
      
      const parsedResponse = JSON.parse(response);
      return parsedResponse.suggested_transactions;
    } catch (error) {
      // Record API error
      lastApiInteraction = {
        timestamp: new Date().toISOString(),
        provider: 'OpenAI',
        request: requestForDebug,
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
    if (!transactionsTable) {
      throw new Error("Transactions table not found in the workbook");
    }
    
    // Get the Categories table
    const categoriesTable = context.workbook.tables.getItem("Categories");    
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
    const institutionColIndex = headers.indexOf(INSTITUTION_COL_NAME);
    const amountColIndex = headers.indexOf(AMOUNT_COL_NAME);
    const dateColIndex = headers.indexOf(DATE_COL_NAME);
    
    if (idColIndex === -1 || origDescColIndex === -1 || descColIndex === -1 || categoryColIndex === -1) {
      throw new Error("Required columns not found in Transactions table");
    }
    
    // Don't require AI Touched column to be present, but log if it's missing
    if (aiTouchedColIndex === -1) {
      console.warn("AI Touched column not found in Transactions table");
    }
    
    // Don't require Institution column to be present, but log if it's missing
    if (institutionColIndex === -1) {
      console.warn("Institution column not found in Transactions table");
    }
    
    // Get visible rows data along with cell addresses
    const visibleRangeRows = transactionsTable.getDataBodyRange().getVisibleView().load(["rows"]);
    await context.sync();
    const rowRanges = (visibleRangeRows.rows.items.map(vr => vr.getRange().load(["values"])));
    await context.sync();
    
    // Find uncategorized transactions (with original description but no category)
    const uncategorizedTransactions: Transaction[] = [];
    const idToRowRange: {[key: string]: Excel.Range} = {};
    
    for (const [i, rowRange] of rowRanges.entries()) {
      const values = rowRange.values[0];
      const origDesc = values[origDescColIndex];
      const category = values[categoryColIndex];
      
      if (origDesc && !category) {
        const amount = parseFloat(values[amountColIndex] || "0");
        const date = values[dateColIndex];
        
        const transactionId = values[idColIndex] || `row-${i}`;
        uncategorizedTransactions.push({
          transaction_id: transactionId,
          original_description: origDesc,
          amount: amount,
          date: date,
          institution: institutionColIndex !== -1 ? values[institutionColIndex] : undefined
        });
        // Store cell addresses for this row (we'll use these to update the correct cells)
        idToRowRange[transactionId] = rowRange;
        
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
            transaction_id: row[idColIndex],
            original_description: row[origDescColIndex],
            updated_description: row[descColIndex] || row[origDescColIndex],
            category: row[categoryColIndex],
            amount: parseFloat(row[headers.indexOf(AMOUNT_COL_NAME)] || "0"),
            date: row[headers.indexOf(DATE_COL_NAME)],
            institution: institutionColIndex !== -1 ? row[institutionColIndex] : undefined
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
    
    // Process suggestions and update cells directly
    let updatedCount = 0;
        
    for (const suggestion of suggestedTransactions) {
      if (!suggestion || !suggestion.transaction_id) continue;
      
      const rowRange = idToRowRange[suggestion.transaction_id];
      
      if (rowRange) {
        // Validate category
        let category = suggestion.category;
        if (!categoryList.includes(category)) {
          category = FALLBACK_CATEGORY;
        }
        
        // Create individual cell updates for this row
        
        // Only update description if the setting is enabled
        if (UPDATE_DESCRIPTIONS) {
          rowRange.getCell(0,descColIndex).values = [[suggestion.updated_description]];
        }
        
        // Always update category
        rowRange.getCell(0,categoryColIndex).values = [[category]];
        
        // Always update AI Touched timestamp with Excel's numeric date value
        if (aiTouchedColIndex !== -1) {
          // Convert to Excel numeric date (days since 1900-01-01)
          // Excel stores dates as days since 1900-01-01 with the decimal portion representing time
          const date = new Date();
          const excelDate = 25569 + (date.getTime() / (24 * 60 * 60 * 1000));
          
          rowRange.getCell(0,aiTouchedColIndex).values = [[excelDate]];
        }
        
        updatedCount++;
      }
    }
    await context.sync();
    
    // Apply all updates
    if (updatedCount > 0) {
      // Apply each cell update directly using cell addresses
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