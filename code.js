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
    .createMenu('Gemini AI')
    .addItem('Show Chat Assistant', 'showChatSidebar')
    .addToUi();
}

/**
 * Displays the main chat interface as a sidebar in the spreadsheet.
 */
function showChatSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Chat')
    .setTitle('Gemini Assistant')
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
      range.setNote('Gemini is thinking...');
      const prompt = `In sheet "${sheet.getName()}", a user just changed cell ${range.getA1Notation()} to "${editedValue}". Analyze this within the workbook context and perform logical follow-up edits.`;
      const context = getAllDataFromAllSheets();
      const aiResponseJSON = queryGemini(prompt, context);
      handleGlobalAIResponse(aiResponseJSON);
      range.setNote('Task complete.');
      Utilities.sleep(2000);
      range.clearNote();
    } catch (error) {
      console.error(`onEdit Error: ${error.message}`);
      range.setNote(`AI Error: ${error.message}`);
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
    responseData = JSON.parse(aiResponseJSON);
  } catch (e) {
    console.error("Failed to parse AI JSON response:", aiResponseJSON);
    throw new Error("Received an invalid response from the AI.");
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

  const systemPrompt = `You are an expert financial analyst and Google Sheets assistant.
Your goal is to help users by performing calculations and edits accurately.

**CRITICAL INSTRUCTIONS:**
1.  **RESPONSE FORMAT:** Your response MUST be a single, valid JSON object with "edits" (an array) and "reply" (a string).
2.  **FORMULA SYNTAX:** When you create a Google Sheets formula, you MUST follow these rules:
    -   All formulas must start with an equals sign \`=\`.
    -   **SAME-SHEET REFERENCES:** When referencing a cell *on the same sheet* you are editing, use A1 notation directly (e.g., \`=B2/B1\`). **DO NOT** include the sheet name (e.g., do not write \`=Sheet1!B2/Sheet1!B1\`). This is the most common mistake.
    -   **CROSS-SHEET REFERENCES:** Only include the sheet name when referencing a *different* sheet (e.g., \`'Data Sheet'!A1\`).
    -   **FINANCIAL RATIOS:** Be precise. For example, a Gross Profit Ratio is typically \`(Revenue - Cost of Goods Sold) / Revenue\` or \`Gross Profit / Revenue\`. Use the correct cells based on the provided data.
3.  **EDITING CELLS:** To edit a cell, use the format: \`{"sheet": "SheetName", "row": 1, "column": 1, "value": "=B2+B3"}\`.
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