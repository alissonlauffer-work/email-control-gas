/**
 * @fileoverview Spreadsheet navigation utilities
 *
 * This file contains utilities for managing spreadsheet navigation,
 * specifically for moving the cursor to the last available row.
 * It ensures users always have a blank row ready for data entry.
 *
 * @author Email Control Gas Team
 * @version 1.0.0
 */

/**
 * Moves the cursor to the last row of the active spreadsheet.
 *
 * This function ensures that users always have a blank row available for data entry.
 * If the spreadsheet is full (no empty rows), it automatically appends a new row
 * before moving the cursor to maintain a smooth user experience.
 *
 * The function follows this logic:
 * 1. Gets the active spreadsheet and sheet from the event
 * 2. Determines the last row with content
 * 3. Checks if the sheet is at maximum capacity
 * 4. If needed, appends a new empty row
 * 5. Moves the cursor to the first column of the new row
 *
 * @param e The Google Apps Script event object passed when the spreadsheet opens
 * @returns void
 */
function moveToLastRow(e: GoogleAppsScript.Events.SheetsOnOpen): void {
  // Extract spreadsheet and sheet information from the event
  const spreadsheet = e.source;
  const sheet = spreadsheet.getActiveSheet();

  // Get the last row that contains data
  let lastRow = spreadsheet.getLastRow();

  // Check if the sheet is completely filled (no empty rows available)
  if (sheet.getMaxRows() === lastRow) {
    // Add a new empty row to accommodate new data entry
    sheet.appendRow([""]);
  }

  // Move to the row after the last content row
  lastRow = lastRow + 1;

  // Create a range for the first column of the new row and set it as active
  const range = sheet.getRange(`A${lastRow}:A${lastRow}`);
  sheet.setActiveRange(range);
}
