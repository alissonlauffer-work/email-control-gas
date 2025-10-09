/**
 * @fileoverview Email validation utilities
 *
 * This file contains utilities for validating that the email in the sheet header
 * matches the logged Gmail account. This ensures data integrity and prevents
 * users from accidentally working with the wrong spreadsheet.
 *
 * @author Email Control Gas Team
 * @version 1.0.0
 */

/**
 * Validates that the email in the sheet header matches the logged Gmail account.
 *
 * This function extracts the email from the sheet header (expected format: "Controle E-mail - email@example.com")
 * and compares it with the currently logged Gmail account. If they don't match, it shows an alert
 * in the spreadsheet and throws an error to prevent users from working with the wrong spreadsheet.
 *
 * The function performs the following steps:
 * 1. Gets the active spreadsheet and sheet
 * 2. Reads the header from row 1, column 1 (A1)
 * 3. Extracts the email using regex pattern matching
 * 4. Gets the currently logged Gmail account email
 * 5. Compares the two emails and shows an alert if they don't match
 *
 * @throws {Error} If the header format is invalid or emails don't match
 * @returns void
 */
function validateSheetEmail(): void {
  console.log("Validating email in sheet header against logged Gmail account");

  // Get the active spreadsheet and sheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();

  // Read the header from cell A1 (row 1, column 1)
  const headerCell = sheet.getRange(1, 1);
  const headerValue = headerCell.getValue().toString();

  console.log(`Sheet header found: "${headerValue}"`);

  // Extract email from header using regex
  // Expected format: "Controle E-mail - email@example.com"
  const emailMatch = headerValue.match(/Controle E-mail\s*-\s*([^\s]+)/);

  if (!emailMatch) {
    const alertMessage = `Formato de cabeçalho inválido. Formato esperado: "Controle E-mail - email@example.com"\nEncontrado: "${headerValue}"`;
    console.error(alertMessage);

    // Show alert in the spreadsheet instead of throwing an error
    SpreadsheetApp.getUi().alert(alertMessage);

    // Still throw an error to prevent further execution if needed
    throw new Error(alertMessage);
  }

  const sheetEmail = emailMatch[1];
  console.log(`Email extracted from header: ${sheetEmail}`);

  // Get the currently logged Gmail account email
  const currentUserEmail = Session.getActiveUser().getEmail();
  console.log(`Currently logged Gmail account: ${currentUserEmail}`);

  // Compare the emails
  if (sheetEmail !== currentUserEmail) {
    const alertMessage = `Incompatibilidade de e-mail detectada!\n\nE-mail da planilha: ${sheetEmail}\nConta Gmail logada: ${currentUserEmail}\n\nPor favor, use a planilha associada à sua conta Gmail ou faça login com a conta correta.`;
    console.error(alertMessage);

    // Show alert in the spreadsheet instead of throwing an error
    SpreadsheetApp.getUi().alert(alertMessage);

    return;
  }

  console.log("Email validation passed: Sheet email matches logged Gmail account");
}
