/**
 * @fileoverview Gmail Email Processor for HS Consórcios Signed Documents
 *
 * This Google Apps Script processes Gmail threads to identify when all parties
 * have signed a proposal and updates column F with "ok" status.
 *
 * The script performs the following main functions:
 * - Searches Gmail for HS Consórcios document completion emails
 * - Extracts proposal numbers from email subjects
 * - Finds corresponding proposals in the spreadsheet
 * - Updates column F with "ok" when all parties have signed
 *
 * @author Email Control Gas Team
 * @version 1.0.0
 */

// Gmail search query to find HS Consórcios document completion emails
// Uses wildcard pattern to match emails with any proposal number
const GMAIL_SIGNED_SEARCH_QUERY: string =
  'subject:"O documento Transferência de Cotas" subject:"foi assinado por todos."';

// Maximum number of emails to process (as requested)
const MAX_EMAILS_TO_PROCESS: number = 200;

/**
 * Processes Gmail threads to find signed completion emails and update proposal status.
 *
 * This function implements a sophisticated algorithm to:
 * 1. Search Gmail for emails indicating all parties have signed
 * 2. Extract proposal numbers using regex pattern matching
 * 3. Find corresponding proposals in the spreadsheet
 * 4. Update column F with "ok" status for signed proposals
 * 5. Provide user feedback with processing summary
 *
 * The function handles edge cases such as:
 * - No matching emails found
 * - Proposal numbers not found in spreadsheet
 * - Already processed proposals (already marked as "ok")
 * - Various email subject formats
 *
 * @returns void
 */
function processSignedProposals(): void {
  console.log("Iniciando processamento de emails para encontrar propostas assinadas");

  // Validate that the sheet email matches the logged Gmail account
  const isValidEmail = validateSheetEmail();

  // Stop processing if validation fails
  if (!isValidEmail) {
    console.log("Email validation failed. Stopping signed proposal processing.");
    return;
  }

  // Get the active spreadsheet and sheet for data operations
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();

  // Search Gmail for signed completion emails
  const threads = GmailApp.search(GMAIL_SIGNED_SEARCH_QUERY, 0, MAX_EMAILS_TO_PROCESS);

  if (threads.length === 0) {
    console.log("Nenhum email de conclusão de assinatura encontrado");
    SpreadsheetApp.getUi().alert("Nenhum email de conclusão de assinatura encontrado.");
    return;
  }

  console.log(`Encontrados ${threads.length} threads de emails de conclusão`);

  // Get all messages from the found threads
  const messages = GmailApp.getMessagesForThreads(threads);
  const flatMessages = messages.flat();

  // Extract proposal numbers from email subjects using regex
  // Pattern matches: "Transferência de Cotas - <number> foi assinado por todos"
  const signedProposals: string[] = [];

  flatMessages.forEach((message) => {
    const subject = message.getSubject();
    const match = subject.match(/Transferência de Cotas\s*-\s*(\d+)/);
    if (match) {
      signedProposals.push(match[1]);
    }
  });

  console.log(`Encontrados ${signedProposals.length} números de proposta assinados`);

  if (signedProposals.length === 0) {
    console.log("Nenhum número de proposta encontrado nos emails");
    SpreadsheetApp.getUi().alert("Nenhum número de proposta encontrado nos emails.");
    return;
  }

  // Get all proposal numbers from column A to find matching rows
  const lastRow = getTrueLastRow(sheet, 1);
  if (lastRow === 0) {
    console.log("Nenhuma proposta encontrada na planilha");
    SpreadsheetApp.getUi().alert("Nenhuma proposta encontrada na planilha.");
    return;
  }

  // Get all data from column A (proposal numbers) and column F (status)
  const dataRange = sheet.getRange(1, 1, lastRow, 6);
  const data = dataRange.getValues();

  // Track updates
  let updatedCount = 0;
  const updatedProposals: string[] = [];

  // Process each signed proposal
  signedProposals.forEach((signedProposalNumber) => {
    // Find the proposal in the spreadsheet
    for (let row = 0; row < data.length; row++) {
      const proposalNumber = data[row][0]?.toString();

      if (proposalNumber === signedProposalNumber) {
        // Check if column F is already "ok"
        const currentStatus = data[row][5]?.toString();

        if (currentStatus !== "ok") {
          // Update column F to "ok"
          sheet.getRange(row + 1, 6).setValue("ok");
          updatedCount++;
          updatedProposals.push(signedProposalNumber);
          console.log(`Atualizado proposta ${signedProposalNumber} na linha ${row + 1}`);
        } else {
          console.log(`Proposta ${signedProposalNumber} já está marcada como "ok"`);
        }
        break; // Move to next signed proposal
      }
    }
  });

  // Log processing summary
  console.log("\n=== RESUMO DO PROCESSAMENTO ===");
  console.log(`Total de emails processados: ${flatMessages.length}`);
  console.log(`Total de propostas assinadas encontradas: ${signedProposals.length}`);
  console.log(`Propostas atualizadas: ${updatedCount}`);

  // Show completion message to user
  const ui = SpreadsheetApp.getUi();
  let message = `Processamento concluído!\n\n`;

  if (updatedCount > 0) {
    message += `${updatedCount} proposta(s) atualizada(s) com status "ok":\n\n`;
    message += updatedProposals.join("\n");
  } else {
    message += "Nenhuma proposta necessitou atualização.\n";
    message += 'Todas as propostas encontradas já estavam marcadas como "ok".';
  }

  message += `\n\nTotal de emails processados: ${flatMessages.length}`;

  ui.alert(message);
}

/**
 * Adds the signed proposal processing menu item.
 *
 * This function adds a menu item to the custom menu that allows users
 * to trigger the signed proposal processing functionality.
 *
 * @returns void
 */
function addSignedProposalMenu(): void {
  mainMenu.addItem("Marcar Propostas Assinadas", "processSignedProposals").addToUi();
}
