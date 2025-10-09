/**
 * @fileoverview Gmail Email Processor for HS Consórcios Documents
 *
 * This Apps Script processes Gmail threads to extract proposal numbers from
 * HS Consórcios document signing emails. It filters and displays proposal
 * numbers that come after a specified target number, showing them in
 * chronological order (oldest to newest).
 */

/**
 * The target proposal number to start filtering from.
 * All proposal numbers greater than this value will be displayed.
 * Modify this constant to change the filtering threshold.
 */
const TARGET_PROPOSAL_NUMBER: number = 310165;

/**
 * Gmail search query to find HS Consórcios document signing emails.
 * Excludes reminder transfer emails to focus only on new document requests.
 */
const GMAIL_SEARCH_QUERY: string =
  'subject:"hs consórcios enviou um documento para você assinar" -lembrete transferência';

/**
 * Maximum number of threads to fetch in a single Gmail API call.
 * Gmail API has a limit of 100 threads per search request.
 */
const BATCH_SIZE: number = 100;

/**
 * Processes Gmail threads to extract and filter proposal numbers from HS Consórcios emails.
 *
 * Logs the filtered proposal numbers to the console
 */
function logEmailSubjectsBatch(): void {
  console.log(`Starting email processing with target proposal number: ${TARGET_PROPOSAL_NUMBER}`);

  const allProposalNumbers: string[] = [];
  let start = 0;
  let totalEmails = 0;

  // Process emails in batches until we find the target number
  while (true) {
    const threads = GmailApp.search(GMAIL_SEARCH_QUERY, start, BATCH_SIZE);

    if (threads.length === 0) {
      console.log("No more threads found in Gmail search");
      break;
    }

    const messages = GmailApp.getMessagesForThreads(threads);

    // Extract proposal numbers from all messages
    messages.flat().forEach((message) => {
      const match = message.getSubject().match(/-\s*(\d+)$/);
      if (match) {
        allProposalNumbers.push(match[1]);
      }
    });

    totalEmails += messages.flat().length;

    // Stop if we found our target number
    if (allProposalNumbers.some((num) => parseInt(num, 10) === TARGET_PROPOSAL_NUMBER)) {
      console.log(`Found target proposal number: ${TARGET_PROPOSAL_NUMBER}`);
      break;
    }

    start += BATCH_SIZE;
  }

  // Find the target index and filter numbers after it
  const targetIndex = allProposalNumbers.findIndex(
    (num) => parseInt(num, 10) === TARGET_PROPOSAL_NUMBER,
  );
  const filteredNumbers =
    targetIndex !== -1 ? allProposalNumbers.slice(0, targetIndex).reverse() : [];

  // Output results
  console.log("\n=== PROPOSAL NUMBERS AFTER TARGET ===");
  if (filteredNumbers.length > 0) {
    console.log(filteredNumbers.join("\n"));
  } else {
    console.log("No proposal numbers found after the target number");
  }

  console.log("\n=== PROCESSING SUMMARY ===");
  console.log(`Total emails processed: ${totalEmails}`);
  console.log(`Total proposal numbers found: ${allProposalNumbers.length}`);
  console.log(`Proposal numbers shown: ${filteredNumbers.length}`);
}
