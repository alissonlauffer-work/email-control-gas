/**
 * @fileoverview Gmail Email Processor for HS Consórcios Documents
 *
 * This Google Apps Script processes Gmail threads to extract proposal numbers from
 * HS Consórcios document signing emails. It reads the last proposal number from
 * column A of the spreadsheet and adds new proposals after it.
 *
 * The script performs the following main functions:
 * - Searches Gmail for HS Consórcios document signing emails
 * - Extracts proposal numbers from email subjects
 * - Compares with existing proposals in the spreadsheet
 * - Adds only new proposals to maintain data integrity
 *
 * @author Email Control Gas Team
 * @version 1.0.0
 */

// Gmail search query to find HS Consórcios document signing emails
// Excludes transfer reminder emails to focus only on new document requests
const GMAIL_SEARCH_QUERY: string =
  'subject:"hs consórcios enviou um documento para você assinar" -lembrete transferência';

// Maximum number of threads to fetch in a single Gmail API call
// Gmail API has a limit of 100 threads per search request
const BATCH_SIZE: number = 100;

/**
 * Processes Gmail threads to extract and add new proposal numbers to the spreadsheet.
 *
 * This function implements a sophisticated algorithm to:
 * 1. Read the last proposal number from column A of the active sheet
 * 2. Search Gmail for relevant emails in batches to avoid API limits
 * 3. Extract proposal numbers using regex pattern matching
 * 4. Identify which proposals are newer than the last recorded one
 * 5. Add only new proposals to maintain data integrity
 * 6. Provide user feedback with processing summary
 *
 * The function handles edge cases such as:
 * - Empty spreadsheets (no existing proposals)
 * - No matching emails found
 * - Large email volumes (processes in batches)
 * - Various email subject formats
 *
 * @returns void
 */
function processNewProposals(): void {
  console.log("Iniciando processamento de emails para encontrar novas propostas");

  // Validate that the sheet email matches the logged Gmail account
  validateSheetEmail();

  // Get the active spreadsheet and sheet for data operations
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();

  // Get the last proposal number from column A to establish our baseline
  const lastRow = sheet.getLastRow();
  let targetProposalNumber: number = 0;

  if (lastRow > 0) {
    // Extract the numeric value from the last row, first column
    const lastProposalValue = sheet.getRange(lastRow, 1).getValue();
    targetProposalNumber = parseInt(lastProposalValue.toString(), 10) || 0;
    console.log(`Encontrado último número de proposta na planilha: ${targetProposalNumber}`);
  } else {
    console.log("Nenhuma proposta existente encontrada na planilha");
  }

  // Initialize variables for batch processing
  const allProposals: { number: string; date: Date }[] = [];
  let start = 0;
  let totalEmails = 0;

  // Process emails in batches until we find our target number or exhaust results
  while (true) {
    // Search Gmail with our query, starting from the current offset
    const threads = GmailApp.search(GMAIL_SEARCH_QUERY, start, BATCH_SIZE);

    if (threads.length === 0) {
      console.log("Nenhum thread adicional encontrado na busca do Gmail");
      break;
    }

    // Get all messages from the found threads
    const messages = GmailApp.getMessagesForThreads(threads);

    // Extract proposal numbers and dates from all message subjects using regex
    // Pattern matches: "-<number>" at the end of the subject line
    messages.flat().forEach((message) => {
      const match = message.getSubject().match(/-\s*(\d+)$/);
      if (match) {
        // Convert Google Apps Script Date to JavaScript Date
        const gasDate = message.getDate();
        const emailDate = new Date(gasDate.getTime());
        // Store proposal number and date as an object
        allProposals.push({ number: match[1], date: emailDate });
      }
    });

    totalEmails += messages.flat().length;

    // Stop if we find our target proposal number (prevents unnecessary processing)
    if (allProposals.some((proposal) => parseInt(proposal.number, 10) === targetProposalNumber)) {
      console.log(`Encontrado número de proposta alvo: ${targetProposalNumber}`);
      break;
    }

    // Move to the next batch
    start += BATCH_SIZE;
  }

  // Find the index of our target proposal in the collected proposals
  const targetIndex = allProposals.findIndex(
    (proposal) => parseInt(proposal.number, 10) === targetProposalNumber,
  );

  // Get newer proposals (those that come before the target in reverse chronological order)
  // Gmail returns emails in reverse chronological order, so newer proposals appear first
  const newerProposals = targetIndex !== -1 ? allProposals.slice(0, targetIndex) : allProposals;

  // Add new proposals to the spreadsheet if any were found
  if (newerProposals.length > 0) {
    console.log(`Encontradas ${newerProposals.length} novas propostas para adicionar`);

    // Reverse to maintain chronological order (oldest to newest)
    // and add each proposal with email received date to the spreadsheet
    newerProposals.reverse().forEach((proposal) => {
      const formattedDate = proposal.date.toLocaleDateString("pt-BR");
      // Add row with: proposal number, empty columns, and email received date
      sheet.appendRow([proposal.number, "", "", "", formattedDate]);
    });

    console.log("Adicionadas novas propostas à planilha:");
    console.log(
      newerProposals
        .map((p) => p.number)
        .reverse()
        .join("\n"),
    );
  } else {
    console.log("Nenhuma nova proposta encontrada");
  }

  // Log processing summary for debugging and monitoring
  console.log("\n=== RESUMO DO PROCESSAMENTO ===");
  console.log(`Total de emails processados: ${totalEmails}`);
  console.log(`Total de números de proposta encontrados: ${allProposals.length}`);
  console.log(`Novas propostas adicionadas: ${newerProposals.length}`);

  // Show completion message to user with appropriate feedback
  const ui = SpreadsheetApp.getUi();
  const message =
    newerProposals.length > 0
      ? `Processamento concluído!\n\n${newerProposals.length} nova(s) proposta(s) adicionada(s) à planilha.\n\nTotal de emails processados: ${totalEmails}`
      : `Processamento concluído!\n\nNenhuma nova proposta encontrada.\n\nTotal de emails processados: ${totalEmails}`;

  ui.alert(message);
}

/**
 * Creates the custom menu in the Google Sheets interface.
 *
 * This function adds a custom menu item to the spreadsheet's menu bar
 * that allows users to trigger the proposal processing functionality.
 * The menu appears as "Ferramentas adicionais" with a submenu item
 * "Processar Novas Propostas" that calls the processNewProposals function.
 *
 * @returns void
 */
function addMenuControls(): void {
  mainMenu.addItem("Processar Novas Propostas", "processNewProposals").addToUi();
}
