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

// Chunk size for processing emails
const EMAIL_CHUNK_SIZE: number = 50;

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

  // Track overall updates
  let totalUpdatedCount = 0;
  const allUpdatedProposals: string[] = [];
  let totalEmailsProcessed = 0;
  let totalChunksProcessed = 0;

  // Process emails in chunks
  for (let startIndex = 0; startIndex < MAX_EMAILS_TO_PROCESS; startIndex += EMAIL_CHUNK_SIZE) {
    const chunkSize = Math.min(EMAIL_CHUNK_SIZE, MAX_EMAILS_TO_PROCESS - startIndex);

    console.log(
      `Processando chunk ${totalChunksProcessed + 1}: emails ${startIndex + 1}-${startIndex + chunkSize}`,
    );

    // Search Gmail for signed completion emails in this chunk
    const threads = GmailApp.search(GMAIL_SIGNED_SEARCH_QUERY, startIndex, chunkSize);

    if (threads.length === 0) {
      console.log(
        `Nenhum email de conclusão de assinatura encontrado no chunk ${totalChunksProcessed + 1}`,
      );
      break; // No more emails to process
    }

    console.log(
      `Encontrados ${threads.length} threads de emails de conclusão no chunk ${totalChunksProcessed + 1}`,
    );

    // Get all messages from the found threads
    const messages = GmailApp.getMessagesForThreads(threads);
    const flatMessages = messages.flat();
    totalEmailsProcessed += flatMessages.length;

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

    console.log(
      `Encontrados ${signedProposals.length} números de proposta assinados no chunk ${totalChunksProcessed + 1}`,
    );

    if (signedProposals.length === 0) {
      console.log(
        `Nenhum número de proposta encontrado nos emails do chunk ${totalChunksProcessed + 1}`,
      );
      totalChunksProcessed++;
      continue; // Move to next chunk
    }

    // Track updates for this chunk
    let chunkUpdatedCount = 0;
    const chunkUpdatedProposals: string[] = [];

    // Process each signed proposal in this chunk
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
            chunkUpdatedCount++;
            totalUpdatedCount++;
            chunkUpdatedProposals.push(signedProposalNumber);
            allUpdatedProposals.push(signedProposalNumber);
            console.log(`Atualizado proposta ${signedProposalNumber} na linha ${row + 1}`);
          } else {
            console.log(`Proposta ${signedProposalNumber} já está marcada como "ok"`);
          }
          break; // Move to next signed proposal
        }
      }
    });

    console.log(
      `Chunk ${totalChunksProcessed + 1} concluído: ${chunkUpdatedCount} propostas atualizadas`,
    );
    totalChunksProcessed++;

    // Short circuit: if no proposals needed updating in this chunk, stop processing early
    if (chunkUpdatedCount === 0 && signedProposals.length > 0) {
      console.log(
        `Todas as propostas no chunk ${totalChunksProcessed + 1} já estão marcadas como 'ok'. Interrompendo processamento precoce.`,
      );
      break;
    }
  }

  // Check if any emails were processed at all
  if (totalEmailsProcessed === 0) {
    console.log("Nenhum email de conclusão de assinatura encontrado");
    SpreadsheetApp.getUi().alert("Nenhum email de conclusão de assinatura encontrado.");
    return;
  }

  // Log processing summary
  console.log("\n=== RESUMO DO PROCESSAMENTO ===");
  console.log(`Total de chunks processados: ${totalChunksProcessed}`);
  console.log(`Total de emails processados: ${totalEmailsProcessed}`);
  console.log(`Propostas atualizadas: ${totalUpdatedCount}`);

  // Show completion message to user
  const ui = SpreadsheetApp.getUi();
  let message = `Processamento concluído!\n\n`;

  if (totalUpdatedCount > 0) {
    message += `${totalUpdatedCount} proposta(s) atualizada(s) com status "ok":\n\n`;
    message += allUpdatedProposals.join("\n");
  } else {
    message += "Nenhuma proposta necessitou atualização.\n";
    message += 'Todas as propostas encontradas já estavam marcadas como "ok".';
  }

  message += `\n\nTotal de emails processados: ${totalEmailsProcessed}`;
  message += `\nTotal de chunks processados: ${totalChunksProcessed}`;

  ui.alert(message);
}

mainMenu.addItem("Marcar Propostas Assinadas", "processSignedProposals").addToUi();
