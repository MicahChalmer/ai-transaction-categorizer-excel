# AI Autocategorizer for Excel

An Excel add-in that helps categorize financial transactions using AI (OpenAI or Google Gemini).

## Features

- Auto-categorize transactions in a Tiller-style spreadsheet
- Works with OpenAI or Google Gemini AI models
- Automatically suggests categories and cleans up transaction descriptions

## Setup

1. Clone the repository
2. Install dependencies:
   ```
   npm install
   ```
3. Run the dev server:
   ```
   npm run dev-server
   ```
4. Build the add-in:
   ```
   npm run build:dev
   ```
5. Start the add-in in Excel:
   ```
   npm run start
   ```

## Usage

1. Ensure your Excel workbook has two tables:
   - `Transactions` table with columns:
     - Transaction ID
     - Full Description
     - Description
     - Category
     - Date
     - Amount
   - `Categories` table with a list of valid categories

2. Open the add-in task pane

3. Click the "AI Auto-Categorize" button to categorize uncategorized transactions

4. The add-in will process visible rows where the "Category" column is empty, and update both the "Description" and "Category" columns based on AI suggestions.

## Configuration

Configure your API keys in the Settings panel of the add-in:

1. Click on the "Settings" button in the add-in
2. Choose your preferred AI provider (Gemini or OpenAI)
3. Enter your API key for the selected provider
4. If using OpenAI, you can also specify which model to use

API keys are stored only in the browser session and are not saved when the add-in is closed.

## Development

- `npm run dev-server`: Start development server
- `npm run build:dev`: Development build
- `npm run lint`: Check for linting issues
- `npm run lint:fix`: Fix linting issues