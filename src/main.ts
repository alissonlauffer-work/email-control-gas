/**
 * @fileoverview Main entry point for the Google Apps Script project
 *
 * This file contains the primary functions that initialize the spreadsheet when opened,
 * including setting up custom menu controls and navigating to the last row.
 * It serves as the entry point for the Email Control Gas application.
 *
 * @author Email Control Gas Team
 * @version 1.0.0
 */

// Global menu reference that will be used across the application
const mainMenu = SpreadsheetApp.getUi().createMenu("Ferramentas adicionais");

// Import the signed proposal menu function
// Note: In Google Apps Script, functions are globally available

/**
 * Function that runs when the spreadsheet is opened.
 * This is the main entry point that initializes the application by:
 * 1. Moving the cursor to the last available row
 *
 * @param e The event object passed by Google Apps Script when the spreadsheet opens
 * @returns void
 */
function onOpen(e: GoogleAppsScript.Events.SheetsOnOpen): void {
  moveToLastRow(e);
}
