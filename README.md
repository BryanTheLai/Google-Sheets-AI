# Google Sheets MAHO AI Assistant

I made this project for fun, it adds an AI-powered chat assistant to Google Sheets using Google Apps Script (MAHO AI).
It lets you chat, generate formulas, and edit your sheet with the power of Gemini AI.

[Watch the video](https://www.youtube.com/watch?v=gY3jIOUUVkI)

## Setup

1. Open your Google Sheet.
2. Go to **Extensions ▶️ Apps Script** -> **Editor Tab**
3. In the Apps Script editor:
   - Press the add button and select Script.
   - Name the file Code.gs and copy contents from Code.js
   - Create a new HTML file named `Chat.html` and paste code from `Chat.html`
4. Set your API key: (Optional)
   - In Apps Script, open **Project Settings ▶️ Script Properties**.
   - Add a property named `GEMINI_API_KEY` and paste your Gemini API key.
5. Save and **Deploy ▶️ Test deployments** to grant permissions.
6. Reload your Google Sheet.
7. Click the **MAHO AI** menu ▶️ **Show Chat Assistant** to open the sidebar.

## Usage

- Type your prompt in the chat sidebar and send to analyze or edit your sheet.
- Or use the custom function in any cell:
  ```
  =MAHO("Summarize this data", A1:B10)
  ```

That's it! Enjoy your AI-powered spreadsheet assistant.

## Example Prompt
```
Create a "Customer Analysis" sheet with columns A-E for ID, Sign-up Date, Avg Purchase Value, Purchase Frequency, and Lifespan. 

Add 10 rows of sample data with formatted headers.

Add a "Customer Lifetime Value" (CLV) column in F with the correct formula using the other columns.
```