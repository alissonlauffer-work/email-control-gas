/**
 * @fileoverview Processador de Email Gmail para Documentos HS Consórcios
 *
 * Este Apps Script processa threads do Gmail para extrair números de propostas de
 * emails de assinatura de documentos HS Consórcios. Lê o último número de proposta
 * da coluna A da planilha e adiciona novas propostas após ele.
/**
 * Cria um menu personalizado na interface do Google Sheets quando a planilha é aberta.
 */
function onOpen(): void {
  SpreadsheetApp.getUi()
    .createMenu("Controle de Email")
    .addItem("Processar Novas Propostas", "processNewProposals")
    .addToUi();
}

/**
 * Query de busca do Gmail para encontrar emails de assinatura de documentos HS Consórcios.
 * Exclui emails de lembrete de transferência para focar apenas em novas solicitações de documentos.
 */
const GMAIL_SEARCH_QUERY: string =
  'subject:"hs consórcios enviou um documento para você assinar" -lembrete transferência';

/**
 * Número máximo de threads para buscar em uma única chamada da API do Gmail.
 * A API do Gmail tem um limite de 100 threads por solicitação de busca.
 */
const BATCH_SIZE: number = 100;

/**
 * Processa threads do Gmail para extrair e adicionar novos números de propostas à planilha.
 *
 * Lê o último número de proposta da coluna A e adiciona quaisquer novas propostas encontradas nos emails.
 */
function processNewProposals(): void {
  console.log("Iniciando processamento de emails para encontrar novas propostas");

  // Obter a planilha e a aba
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();

  // Obter o último número de proposta da coluna A
  const lastRow = sheet.getLastRow();
  let targetProposalNumber: number = 0;

  if (lastRow > 0) {
    const lastProposalValue = sheet.getRange(lastRow, 1).getValue();
    targetProposalNumber = parseInt(lastProposalValue.toString(), 10) || 0;
    console.log(`Encontrado último número de proposta na planilha: ${targetProposalNumber}`);
  } else {
    console.log("Nenhuma proposta existente encontrada na planilha");
  }

  const allProposalNumbers: string[] = [];
  let start = 0;
  let totalEmails = 0;

  // Processar emails em lotes até encontrar o número alvo
  while (true) {
    const threads = GmailApp.search(GMAIL_SEARCH_QUERY, start, BATCH_SIZE);

    if (threads.length === 0) {
      console.log("Nenhum thread adicional encontrado na busca do Gmail");
      break;
    }

    const messages = GmailApp.getMessagesForThreads(threads);

    // Extrair números de propostas de todas as mensagens
    messages.flat().forEach((message) => {
      const match = message.getSubject().match(/-\s*(\d+)$/);
      if (match) {
        allProposalNumbers.push(match[1]);
      }
    });

    totalEmails += messages.flat().length;

    // Parar se encontrarmos nosso número alvo
    if (allProposalNumbers.some((num) => parseInt(num, 10) === targetProposalNumber)) {
      console.log(`Encontrado número de proposta alvo: ${targetProposalNumber}`);
      break;
    }

    start += BATCH_SIZE;
  }

  // Encontrar o índice alvo e filtrar números após ele (propostas mais recentes)
  const targetIndex = allProposalNumbers.findIndex(
    (num) => parseInt(num, 10) === targetProposalNumber,
  );

  // Obter propostas mais recentes (aquelas que vêm antes do alvo em ordem cronológica reversa)
  const newerProposals =
    targetIndex !== -1 ? allProposalNumbers.slice(0, targetIndex) : allProposalNumbers;

  // Adicionar novas propostas à planilha
  if (newerProposals.length > 0) {
    console.log(`Encontradas ${newerProposals.length} novas propostas para adicionar`);

    // Inverter para manter ordem cronológica (mais antiga para mais recente)
    newerProposals.reverse().forEach((proposalNumber) => {
      sheet.appendRow([proposalNumber, new Date().toISOString()]);
    });

    console.log("Adicionadas novas propostas à planilha:");
    console.log(newerProposals.reverse().join("\n"));
  } else {
    console.log("Nenhuma nova proposta encontrada");
  }

  console.log("\n=== RESUMO DO PROCESSAMENTO ===");
  console.log(`Total de emails processados: ${totalEmails}`);
  console.log(`Total de números de proposta encontrados: ${allProposalNumbers.length}`);
  console.log(`Novas propostas adicionadas: ${newerProposals.length}`);
}
