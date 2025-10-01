/**
 * Logs email subjects from Gmail threads matching a specific search query
 * Processes emails in batches to handle large volumes efficiently
 */
function logEmailSubjectsBatch() {
  const query =
    'subject:"hs consórcios enviou um documento para você assinar" -lembrete transferência';
  const batchSize = 100; // Gmail search limit per call
  let start = 0;
  let totalEmails = 0;

  while (true) {
    // Fetch threads in batches
    const threads = GmailApp.search(query, start, batchSize);

    // If no threads are found, stop the loop
    if (threads.length === 0) {
      break;
    }

    // Process each thread and log the subject of each message
    const messages = GmailApp.getMessagesForThreads(threads);
    messages.forEach((thread) => {
      thread.forEach((message) => {
        console.log("Subject:", message.getSubject());
      });
    });

    // Update counters and continue to the next batch
    totalEmails += messages.flat().length;
    start += batchSize;
  }

  console.log(`Total emails processed: ${totalEmails}`);
}
