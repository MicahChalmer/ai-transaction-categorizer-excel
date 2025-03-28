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
}) {
  // Update keys and settings
  if (config.openaiKey) OPENAI_API_KEY = config.openaiKey;
  if (config.googleKey) GOOGLE_API_KEY = config.googleKey;
  if (config.provider) AI_PROVIDER = config.provider;
  if (config.model) GPT_MODEL = config.model;
  if (config.maxBatchSize) MAX_BATCH_SIZE = config.maxBatchSize;
  if (config.maxReferenceTransactions) MAX_REFERENCE_TRANSACTIONS = config.maxReferenceTransactions;
  
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

    const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });

    const prompt = `
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
          
          For each transaction in the transactions list, follow these instructions:
          (0) First check if there are any similar transactions in the reference_transactions list.
              If you find similar transactions, use the same category and a similar updated_description.
              Match transactions based on merchant name, description patterns, and similar text.
          (1) If there are no similar reference transactions, suggest a better "updated_description" according to the following rules:
          (a) Use all of your knowledge and information to propose a friendly, human readable updated_description for the
            transaction given the original_description. The input often contains the name of a merchant name.
            If you know of a merchant it might be referring to, use the name of that merchant for the suggested description.
          (b) Keep the suggested description as simple as possible. Remove punctuation, extraneous
            numbers, location information, abbreviations such as "Inc." or "LLC", IDs and account numbers.
          (2) For each original_description, suggest a "category" for the transaction from the allowed_categories list that was provided.
          (3) If you are not confident in the suggested category after using your own knowledge and the previous transactions provided, use the cateogry "${FALLBACK_CATEGORY}"
          (4) Your response should be a JSON object and no other text.  The response object should be of the form:
          {"suggested_transactions": [
            {
              "transaction_id": "The unique ID previously provided for this transaction",
              "updated_description": "The cleaned up version of the description",
              "category": "A category selected from the allowed_categories list"
            }
          ]}
    `;

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
    
    // Parse the JSON response
    const parsedResponse = JSON.parse(jsonText);
    return parsedResponse.suggested_transactions;
  } catch (error) {
    console.error("Error using Gemini API:", error);
    return null;
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

    const completion = await openai.chat.completions.create({
      model: GPT_MODEL,
      temperature: 0.2,
      top_p: 0.1,
      seed: 1,
      response_format: { type: "json_object" },
      messages: [
        {
          role: "system",
          content: "Act as an API that categorizes and cleans up bank transaction descriptions for for a personal finance app.",
        },
        {
          role: "system",
          content: "Reference the following list of allowed_categories:\n" + JSON.stringify(categoryList),
        },
        {
          role: "system",
          content: `You will be given JSON input with a list of uncategorized transactions and a set of previously categorized reference transactions in the following format:
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
            
            For each transaction in the transactions list, follow these instructions:
            (0) First check if there are any similar transactions in the reference_transactions list.
                If you find similar transactions, use the same category and a similar updated_description.
                Match transactions based on merchant name, description patterns, and similar text.
            (1) If there are no similar reference transactions, suggest a better "updated_description" according to the following rules:
            (a) Use all of your knowledge and information to propose a friendly, human readable updated_description for the 
              transaction given the original_description. The input often contains the name of a merchant name. 
              If you know of a merchant it might be referring to, use the name of that merchant for the suggested description.
            (b) Keep the suggested description as simple as possible. Remove punctuation, extraneous 
              numbers, location information, abbreviations such as "Inc." or "LLC", IDs and account numbers.
            (2) For each original_description, suggest a "category" for the transaction from the allowed_categories list that was provided.
            (3) If you are not confident in the suggested category after using your own knowledge and the previous transactions provided, use the cateogry "${FALLBACK_CATEGORY}"

            (4) Your response should be a JSON object and no other text.  The response object should be of the form:
            {"suggested_transactions": [
              {
                "transaction_id": "The unique ID previously provided for this transaction",
                "updated_description": "The cleaned up version of the description",
                "category": "A category selected from the allowed_categories list"
              }
            ]}`,
        },
        {
          role: "user",
          content: JSON.stringify(transactionDict),
        },
      ],
    });

    const response = completion.choices[0].message.content;
    if (!response) {
      throw new Error("No response from OpenAI API");
    }
    
    const parsedResponse = JSON.parse(response);
    return parsedResponse.suggested_transactions;
  } catch (error) {
    console.error("Error using OpenAI API:", error);
    return null;
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
    const allRows = allCategorizedRange.values;
    const categorizedTransactions: CategorizedTransaction[] = [];
    
    for (const row of allRows) {
      if (row[origDescColIndex] && row[categoryColIndex]) {
        categorizedTransactions.push({
          original_description: row[origDescColIndex],
          updated_description: row[descColIndex] || row[origDescColIndex],
          category: row[categoryColIndex],
          amount: parseFloat(row[headers.indexOf(AMOUNT_COL_NAME)] || "0"),
          date: row[headers.indexOf(DATE_COL_NAME)]
        });
      }
    }
    
    // Limit the number of reference transactions to avoid too large requests
    // Sort by most recent first and take up to the configured maximum
    const limitedReferenceTransactions = categorizedTransactions
      .slice(0, MAX_REFERENCE_TRANSACTIONS);
    
    // Process categories
    const categoryValues = categoryColRange.values;
    
    const categoryList: string[] = categoryValues.map(row => row[0]).filter(Boolean);
    
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
    
    // Use the existing rows array from earlier
    
    // Create a map for quick transaction lookup by ID
    const transactionMap = new Map();
    uncategorizedTransactions.forEach((tx, index) => {
      transactionMap.set(tx.transaction_id, rowIndices[index]);
    });
    
    // Update only the in-memory values array
    let updatedCount = 0;
    for (const suggestion of suggestedTransactions) {
      const actualRowIndex = transactionMap.get(suggestion.transaction_id);
      
      if (actualRowIndex !== undefined) {
        // Validate category
        let category = suggestion.category;
        if (!categoryList.includes(category)) {
          category = FALLBACK_CATEGORY;
        }
        
        // Update the values in the array
        rows[actualRowIndex][descColIndex] = suggestion.updated_description;
        rows[actualRowIndex][categoryColIndex] = category;
        
        // Update AI Touched column with current date/time if column exists
        if (aiTouchedColIndex !== -1) {
          rows[actualRowIndex][aiTouchedColIndex] = new Date();
        }
        
        updatedCount++;
      }
    }
    
    // Batch update: write back to Excel just once
    if (updatedCount > 0) {
      // Set the values back to the visible range
      visibleRange.values = rows;
      
      // Perform a sync to update the sheet
      await context.sync();
      
      // Return success message with count of updated transactions
      return { success: true, message: `Updated ${updatedCount} transactions` };
    } else {
      return { success: true, message: `No transactions needed updating` };
    }
  } catch (error) {
    console.error("Error in categorizeUncategorizedTransactions:", error);
    return { 
      success: false, 
      message: error instanceof Error ? error.message : "Unknown error occurred" 
    };
  }
}