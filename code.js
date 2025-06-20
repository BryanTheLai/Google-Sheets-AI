/**
 * @fileoverview Gemini AI Assistant for Google Sheets
 * @version 4.2
 * This script integrates the Gemini AI to provide full-spreadsheet context awareness,
 * global editing capabilities, and a user-friendly chat interface with enhanced
 * formula generation and financial understanding. (Corrects API response parsing)
 */

// --- Configuration ---
// IMPORTANT: Add your Gemini API Key in Project Settings > Script Properties.
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
// Use a more powerful model for better reasoning and formula generation.
const DEFAULT_MODEL = 'gemini-2.5-flash';
/**
 * Creates a custom menu in the spreadsheet UI when the workbook is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('MAGE AI')
    .addItem('Show Chat Assistant', 'showChatSidebar')
    .addToUi();
}

/**
 * Displays the main chat interface as a sidebar in the spreadsheet.
 */
function showChatSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Chat')
    .setTitle('MAGE Assistant')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * This is the primary function called from the client-side chat UI.
 */
function processChatMessage(prompt) {
  if (!prompt) {
    throw new Error("Prompt cannot be empty.");
  }
  try {
    const context = getAllDataFromAllSheets();
    const aiResponseJSON = queryGemini(prompt, context);
    const reply = handleGlobalAIResponse(aiResponseJSON);
    return reply;
  } catch (e) {
    console.error(`processChatMessage Error: ${e.toString()}`);
    // Propagate a user-friendly error message to the UI
    throw new Error(`An error occurred: ${e.message}`);
  }
}

/**
 * Automatically runs when a user edits any cell in the spreadsheet.
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const editedValue = e.value;

  if (!editedValue || editedValue.trim() === '' || editedValue === e.oldValue) {
    return;
  }

  const scriptLock = LockService.getScriptLock();
  if (scriptLock.tryLock(100)) {
    try {
      range.setNote('MAGE AI is thinking...');
      const prompt = `In sheet "${sheet.getName()}", a user just changed cell ${range.getA1Notation()} to "${editedValue}". Analyze this within the workbook context and perform logical follow-up edits.`;
      const context = getAllDataFromAllSheets();
      const aiResponseJSON = queryGemini(prompt, context);
      handleGlobalAIResponse(aiResponseJSON);
      range.setNote('Task complete.');
      Utilities.sleep(2000);
      range.clearNote();
    } catch (error) {
      console.error(`onEdit Error: ${error.message}`);
      range.setNote(`MAGE AI Error: ${error.message}`);
    } finally {
      scriptLock.releaseLock();
    }
  }
}

/**
 * Gathers all data from every sheet in the active spreadsheet.
 */
function getAllDataFromAllSheets() {
  const allSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  const allData = {};
  allSheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheet.getLastRow() > 0 && sheet.getLastColumn() > 0) {
      allData[sheetName] = sheet.getDataRange().getValues();
    } else {
      allData[sheetName] = [];
    }
  });
  return allData;
}

/**
 * Parses the AI's JSON response and applies the requested edits.
 */
function handleGlobalAIResponse(aiResponseJSON) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let responseData;
  try {
    // Clean the string to remove markdown fences before parsing
    const cleanedResponse = aiResponseJSON.replace(/```json\n?|```/g, '').trim();
    responseData = JSON.parse(cleanedResponse);
  } catch (e) {
    console.error("Failed to parse AI JSON response:", aiResponseJSON);
    throw new Error("Received an invalid response from the AI. Please check the logs for details.");
  }

  const edits = responseData.edits;

  if (edits && Array.isArray(edits)) {
    edits.forEach(edit => {
      if (edit.action === 'addSheet' && edit.name) {
        try {
          spreadsheet.insertSheet(edit.name);
        } catch (e) {
          console.warn(`Could not create sheet '${edit.name}'. It might already exist.`);
        }
      } else if (edit.action === 'formatCell' && edit.sheet && edit.row && edit.column && edit.style) {
          const sheet = spreadsheet.getSheetByName(edit.sheet);
          if (sheet) {
              try {
                  const range = sheet.getRange(edit.row, edit.column);
                  const style = {};
                  if (edit.style.fontColor) style.fontColor = edit.style.fontColor;
                  if (edit.style.background) style.background = edit.style.background;
                  if (edit.style.fontWeight) style.fontWeight = edit.style.fontWeight;
                  if (edit.style.fontStyle) style.fontStyle = edit.style.fontStyle;
                  range.setValues(range.getValues()).setRichTextValues(range.getRichTextValues().map(row => row.map(cell => {
                      const newStyle = cell.getTextStyle().copy();
                      if (style.fontColor) newStyle.setForegroundColor(style.fontColor);
                      if (style.fontWeight) newStyle.setBold(style.fontWeight === 'bold');
                      if (style.fontStyle) newStyle.setItalic(style.fontStyle === 'italic');
                      return SpreadsheetApp.newRichTextValue().setText(cell.getText()).setTextStyle(newStyle.build()).build();
                  })));
                  if (style.background) range.setBackground(style.background);
              } catch(e) {
                  console.error(`Failed to format ${edit.sheet}!R${edit.row}C${edit.column}: ${e.message}`);
              }
          }
      } else if (edit.action === 'setColumnWidth' && edit.sheet && edit.column && edit.width) {
          const sheet = spreadsheet.getSheetByName(edit.sheet);
          if (sheet) {
              try {
                  sheet.setColumnWidth(edit.column, edit.width);
              } catch(e) {
                  console.error(`Failed to set width for column ${edit.column} in ${edit.sheet}: ${e.message}`);
              }
          }
      } else if (edit.sheet && edit.row && edit.column && typeof edit.value !== 'undefined') {
        const sheet = spreadsheet.getSheetByName(edit.sheet);
        if (sheet) {
          try {
            // Use setFormula for values starting with '=', otherwise setValue
            if (String(edit.value).startsWith('=')) {
              sheet.getRange(edit.row, edit.column).setFormula(edit.value);
            } else {
              sheet.getRange(edit.row, edit.column).setValue(edit.value);
            }
          } catch(e) {
            console.error(`Failed to edit ${edit.sheet}!R${edit.row}C${edit.column}: ${e.message}`);
          }
        } else {
          console.error(`Sheet with name "${edit.sheet}" not found.`);
        }
      }
    });
  }
  return responseData.reply || null;
}

/**
 * Queries the Gemini API with a prompt and the full spreadsheet context.
 * This version contains a highly detailed system prompt for accurate formula generation.
 */
function queryGemini(prompt, context) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${DEFAULT_MODEL}:generateContent?key=${GEMINI_API_KEY}`;

  const systemPrompt = `You are an expert financial analyst and Google Sheets assistant called MAGE AI.
Your goal is to help users by performing calculations, edits, and formatting accurately.

**CRITICAL INSTRUCTIONS:**
1.  **RESPONSE FORMAT:** Your response MUST be a single, valid JSON object with "edits" (an array) and "reply" (a string).
2.  **AVAILABLE ACTIONS:** You can perform the following actions in the "edits" array:
    *   **Edit Cell Value:** \`{"sheet": "SheetName", "row": 1, "column": 1, "value": "=B2+B3"}\`
    *   **Add New Sheet:** \`{"action": "addSheet", "name": "New Report"}\`
    *   **Format a Cell:** \`{"action": "formatCell", "sheet": "SheetName", "row": 1, "column": 1, "style": {"fontColor": "#ff0000", "background": "#f0f0f0", "fontWeight": "bold", "fontStyle": "italic"}}\`
        *   Supported styles: \`fontColor\` (hex), \`background\` (hex), \`fontWeight\` (\'bold\' or \'normal\'), \`fontStyle\` (\'italic\' or \'normal\').
    *   **Set Column Width:** \`{"action": "setColumnWidth", "sheet": "SheetName", "column": 3, "width": 150}\` (width is in pixels).
3.  **FORMULA SYNTAX:** When you create a Google Sheets formula, you MUST follow these rules:
    -   All formulas must start with an equals sign \`=\`.
    -   **SAME-SHEET REFERENCES:** When referencing a cell *on the same sheet* you are editing, use A1 notation directly (e.g., \`=B2/B1\`). **DO NOT** include the sheet name (e.g., do not write \`=Sheet1!B2/Sheet1!B1\`). This is the most common mistake.
    -   **CROSS-SHEET REFERENCES:** Only include the sheet name when referencing a *different* sheet (e.g., \`'Data Sheet'!A1\`).
    -   **FINANCIAL RATIOS:** Be precise. For example, a Gross Profit Ratio is typically \`(Revenue - Cost of Goods Sold) / Revenue\` or \`Gross Profit / Revenue\`. Use the correct cells based on the provided data.
4.  **CONVERSATION:**
    -   Always explain what you did in the "reply" field.
    -   If a user's request is vague (e.g., "add 500 rows"), you MUST ask for clarification in the "reply" and make NO edits.
    -   If a user asks "why", explain your own reasoning as an AI assistant.
5.  **USER CONTEXT:**
    -   The user's immediate request is: "${prompt}".
    -   The entire spreadsheet's data is provided below for your analysis.

Your task is to analyze the user's request and the data, then generate the appropriate JSON response.`;

  const requestBody = {
    "system_instruction": { "parts": [{ "text": systemPrompt }] },
    "contents": [{
      "parts": [
        { "text": "User Prompt: " + prompt },
        { "text": "Spreadsheet Data Context: " + JSON.stringify(context) }
      ]
    }],
    "generation_config": {
      "response_mime_type": "application/json",
    }
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(requestBody),
    'muteHttpExceptions': true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseText = response.getContentText();
  
  if (response.getResponseCode() !== 200) {
    console.error("API Error Response:", responseText);
    throw new Error(`The AI model returned an error (HTTP ${response.getResponseCode()}). See logs.`);
  }
  
  const jsonResponse = JSON.parse(responseText);
  
  // *** THIS IS THE CORRECTED LINE ***
  // We must access the first element of the 'candidates' array and the 'parts' array.
  return jsonResponse.candidates[0].content.parts[0].text;
}

/**
 * Custom function to use Gemini AI directly in a sheet cell.
 * Example: =MAGE("Summarize this data", A1:B10)
 *
 * @param {string} prompt The prompt to send to the AI.
 * @param {string[][]} [range] Optional. The range of cells to provide as context.
 * @return The AI's response as a string.
 * @customfunction
 */
function MAGE(prompt, range) {
  if (!prompt) {
    return "Error: Prompt cannot be empty.";
  }
  try {
    const context = range ? { "selected_range": range } : { "active_sheet": getActiveSheetData() };
    const aiResponse = queryGeminiForCellValue(prompt, context);
    return aiResponse;
  } catch (e) {
    return `Error: ${e.message}`;
  }
}

/**
 * Helper to get active sheet data, limited for performance.
 */
function getActiveSheetData() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (sheet.getLastRow() > 0 && sheet.getLastColumn() > 0) {
      const maxRows = Math.min(sheet.getLastRow(), 100); // Limit rows
      const maxCols = Math.min(sheet.getLastColumn(), 25); // Limit columns
      return sheet.getRange(1, 1, maxRows, maxCols).getValues();
    }
    return [];
}

/**
 * Queries Gemini for a direct cell value, not for edits.
 */
function queryGeminiForCellValue(prompt, context) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${DEFAULT_MODEL}:generateContent?key=${GEMINI_API_KEY}`;

  const systemPrompt = `You are a Google Sheets AI assistant inside a custom function.
- Your purpose is to answer a question or perform a calculation based on the user's prompt and optional data context.
- Your response MUST be a single, concise value that can be displayed in a single spreadsheet cell.
- Do NOT output JSON. Do NOT explain yourself. Just provide the final value.
- If the user asks for a formula, provide only the formula string, starting with '='.
- If the user asks a question, provide a direct answer.`;

  const requestBody = {
    "system_instruction": { "parts": [{ "text": systemPrompt }] },
    "contents": [{
      "parts": [
        { "text": "User Prompt: " + prompt },
        { "text": "Data Context: " + JSON.stringify(context) }
      ]
    }],
    "generation_config": {
      "response_mime_type": "text/plain",
    }
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(requestBody),
    'muteHttpExceptions': true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseText = response.getContentText();

  if (response.getResponseCode() !== 200) {
    console.error("API Error Response:", responseText);
    throw new Error(`AI model error (HTTP ${response.getResponseCode()}). Check logs.`);
  }

  const jsonResponse = JSON.parse(responseText);
  if (jsonResponse.candidates && jsonResponse.candidates.length > 0 &&
      jsonResponse.candidates[0].content && jsonResponse.candidates[0].content.parts &&
      jsonResponse.candidates[0].content.parts.length > 0) {
    return jsonResponse.candidates[0].content.parts[0].text;
  } else {
    console.error("Unexpected AI response structure:", responseText);
    throw new Error("Received an invalid or empty response from the AI.");
  }
}