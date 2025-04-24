# AI Transaction Autocategorizer for Excel

An Excel add-in that helps categorize consumer financial transactions using AI (OpenAI or Google Gemini).

Inspired by @sjogreen's [repo that does the same thing for Google Sheets](https://github.com/sjogreen/tiller_ai_autocat).  (The code is completely different from sjogreen's, since the language & API has to be different.  The prompt was originally copied from his and then updated and changed.)

This whole project was "mostly vibe coded" with Claude Code.  I've split commits into "AI" and "Human" so you can look at the commit history and see what parts were done by which.  As a learning exercise, I tried to force myself to use the AI as much as I could rather than doing the code myself, but I didn't fully hold myself to that.

## Features

- Auto-categorize transactions in a spreadsheet (the assumed format for the spreadsheet is that of the transactions sheet maintained by [Tiller](https://tiller.com/), since that's what I'm personally using it with.)
- Uses transactions present in the table already categorized as reference - this should mean that if you correct its output it should see that and "learn" from it going forward
- Works with OpenAI or Google Gemini AI models
- Automatically suggests categories and, optionally, cleans up transaction descriptions

## How to add it to your own running Excel

I have not published the add-in anywhere other than Github; for that reason, the only way to install it is to "sideload" it from this code repo as if you were developing and testing it yourself.

Overall it works with the following instructions:

## Setup

1. Ensure [node.js and npm are installed](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm)
2. Clone this repository
3. Install dependencies:
   ```
   npm install
   ```
4. Build the add-in:
   ```
   npm run build:dev
   ```
5. (Optional) add a `.env` file in the root of this repo (copy the [.env.example](.env.example) file to see the variable names) and put in your OpenAI and/or Gemini API keys.  If you don't do this, you'll have to paste the API key into the task pane each time you start the add-in; it doesn't persist that anywhere, so if you paste it in the UI it will forget it each time the add-in is reopened.
5. Start the add-in in Excel:
   ```
   npm run start
   ```

By default this will start the Excel desktop app on your own machine with the add-on enabled, on a new blank sheet.  You can then open your own sheets and use the add-on there.  The add-on won't be permanently installed - if you close Excel and open it again the normal way, it won't be there.  You have to run the "start" script from here to use it.  Microsoft publishes instructions for [sideloading office add-ins for testing](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing).  This has some instructions for "sideloading" an add-on Excel for the web, which I personally haven't tried with this.

## Usage

1. Ensure your Excel workbook has two tables:
   - `Transactions` table with columns:
     - Transaction ID
     - Full Description
     - Description
     - Category
     - Date
     - Amount
     - (Optional) Institution
     - (Optional) AI Touched - the add-in will populate this with the current date/time when it updates a transaction
   - `Categories` table with a list of valid categories

2. Open the add-in task pane by clicking the "AI Transaction Autocategorizer" button that appears in the "Data" tab of the ribbon.

3. Click the "AI Auto-Categorize" button to categorize uncategorized transactions

4. The add-in will process visible rows where the "Category" column is empty, and update both the "Description" and "Category" columns based on AI suggestions.

## Configuration

The "Settings" button in the panel shows ways to configure the add-in's behavior.  Note that
these settings aren't currently persisted anywhere and will revert to their default when the
add-in is restarted.

#### API Settings and Keys

1. Choose your preferred AI provider (Gemini or OpenAI)
2. Enter your API key for the selected provider; will be prepopulated if you put it in `.env` file (recommended)
3. You can also specify which model to use.  You can enter a model name, or hit "get models" to have it fetch a list of available models from Google or OpenAI.  It doesn't do this by default because these lists are incomplete - for instance, as of April 23 2025, the Gemini 2.5 flash preview was available as "gemini-2.5-flash-preview-04-17" but that doesn't appear in the list of available models from Gemini's API.

#### Performance Settings

Max Batch Size

Maximum number of uncategorized transactions to categorize in a single run.  Each run makes a single AI API call, so trying to do too many at once probably won't work.  50 seems like a reasonable default.

Max Reference Transactions

Maximum number of already-categorized transactions to include as reference.  Trying to add too many will overrun the context window.

The "Debug" button will show the JSON of the last AI API request and response.  This is mainly useful for my own debugging of the add-in but could be of interest if you're curious about how it works.

#### Content Settings

Update transaction descriptions

Check this box to make it update the "description" column to clean it up.  It's off by default because I generally just want it to add the category, but some want the descriptions updated as well.