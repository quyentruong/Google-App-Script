const MAX_EMAILS_PER_RUN = 50; // Limit threads processed per call to avoid hitting quotas
// --- Logging Levels ---
const LOG_LEVEL_NONE = 0;
const LOG_LEVEL_ERROR = 1;
const LOG_LEVEL_INFO = 2;
const LOG_LEVEL_DEBUG = 3;
// Set your desired log level here:
const LOG_LEVEL = LOG_LEVEL_INFO;

/**
 * Custom log function with log level support.
 * @param {string} message - The message to log.
 * @param {number} level - The log level (1=error, 2=info, 3=debug). Defaults to info.
 */
function log(message, level = LOG_LEVEL_INFO) {
  if (level <= LOG_LEVEL) {
    Logger.log(message);
  }
}
// --- Constants for Magic Strings and Error Messages ---
const NO_THREADS_RESULT = 'NO_THREADS';
const ERROR_ALL_FILTERS_EMPTY = "ERROR: All content filter parameters (sender, recipient, subject) are empty. " +
  "This script will not proceed to prevent accidental mass deletion of all old emails. " +
  "Please provide at least one content filter criterion.";
// --- Main Function to Process All Cleanup Tasks from Sheet ---
/**
 * Processes all email cleanup tasks defined in the active Google Sheet.
 * It reads configuration from the sheet, iterates through each task,
 * and calls the deleteOldEmails function, updating the lastRun timestamp.
 *
 * This function is designed to be run directly from within the Google Sheet's
 * Apps Script editor, using SpreadsheetApp.getActiveSpreadsheet() and
 * getActiveSheet() to reference the current sheet.
 */
function processAllCleanupTasks() {
  // Get the active spreadsheet and the active sheet where the script is running.
  // This removes the need for CONFIG_SHEET_ID and CONFIG_SHEET_NAME.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (!sheet) {
    Logger.log(`Error: No active sheet found. Please run this script from within a Google Sheet.`);
    return;
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  // Check if there's at least a header row
  if (values.length < 1) {
    Logger.log("Error: The sheet is empty. Please add headers and task data.");
    return;
  }

  const headers = values[0]; // First row is headers

  // Map header names to their column indices for easy access
  const headerMap = {};
  headers.forEach((header, index) => {
    headerMap[header.trim()] = index;
  });

  // Validate required columns
  const requiredHeaders = ['senderEmail', 'recipientEmail', 'subjectContains', 'daysOld', 'excludeStarred', 'excludeImportant', 'dryRun', 'lastRun'];
  for (const reqHeader of requiredHeaders) {
    if (headerMap[reqHeader] === undefined) {
      Logger.log(`Error: Required column "${reqHeader}" not found in the sheet. Please ensure all columns are present.`);
      return;
    }
  }

  Logger.log(`Starting processing of ${values.length - 1} cleanup tasks from sheet "${sheet.getName()}".`);

  // Prepare an array to batch update lastRun values
  const lastRunUpdates = new Array(values.length - 1).fill(null); // Only for data rows (not header)
  let stopProcessing = false;
  const now = new Date();

  for (let i = 1; i < values.length; i++) {
    if (stopProcessing) break;
    const row = values[i];
    log(`--- Processing Task Row ${i + 1} ---`, LOG_LEVEL_INFO);

    const options = {
      senderEmail: row[headerMap.senderEmail],
      recipientEmail: row[headerMap.recipientEmail],
      subjectContains: row[headerMap.subjectContains],
      daysOld: row[headerMap.daysOld],
      excludeStarred: row[headerMap.excludeStarred],
      excludeImportant: row[headerMap.excludeImportant],
      dryRun: row[headerMap.dryRun]
    };
    let lastRunTimestamp = row[headerMap.lastRun];
    // Handle possible string date from sheet
    if (typeof lastRunTimestamp === 'string') {
      lastRunTimestamp = new Date(lastRunTimestamp);
    }
    const twentyFourHoursAgo = new Date(now.getTime() - (24 * 60 * 60 * 1000));

    let shouldRun = true;
    if (lastRunTimestamp instanceof Date && !isNaN(lastRunTimestamp) && lastRunTimestamp > twentyFourHoursAgo) {
      shouldRun = false;
      Logger.log(`Task for row ${i + 1} last ran at ${lastRunTimestamp.toLocaleString()}. Skipping as less than 24 hours have passed.`);
    }

    if (shouldRun) {
      try {
        log(`Attempting to run cleanup for task on row ${i + 1}...`, LOG_LEVEL_INFO);
        const result = deleteOldEmails(options);
        if (result.threadsProcessed === 0) {
          lastRunUpdates[i - 1] = [now];
          log(`No threads found for row ${i + 1}; updated lastRun to ${now.toLocaleString()} (batched).`, LOG_LEVEL_INFO);
        } else {
          log(`Processed ${result.threadsProcessed} threads for row ${i + 1}; lastRun not updated so it can run again soon.`, LOG_LEVEL_INFO);
        }
      } catch (e) {
        if (e.message.includes("Service invoked too many times") || e.message.includes("Service using too much computer time")) {
          log(`Daily quota hit while processing task on row ${i + 1}. Stopping further tasks for today.`, LOG_LEVEL_ERROR);
          lastRunUpdates[i - 1] = [now];
          stopProcessing = true;
        } else if (e.message.includes("All content filter parameters are empty")) {
          log(`Skipping task on row ${i + 1} due to configuration error: ${e.message}`, LOG_LEVEL_ERROR);
        } else {
          log(`Unhandled error processing task on row ${i + 1}: ${e.toString()}`, LOG_LEVEL_ERROR);
          lastRunUpdates[i - 1] = [now];
        }
      }
      Utilities.sleep(1000);
    }
  }

  // Batch update lastRun column for all processed rows
  // Only update rows where lastRunUpdates[i] is not null
  const lastRunCol = headerMap.lastRun + 1;
  const updateRows = [];
  const updateValues = [];
  for (let i = 0; i < lastRunUpdates.length; i++) {
    if (lastRunUpdates[i] !== null) {
      updateRows.push(i + 2); // Sheet rows are 1-based, skip header
      updateValues.push(lastRunUpdates[i]);
    }
  }
  if (updateRows.length > 0) {
    // If updates are contiguous, can use setValues on a range
    // Otherwise, update individually (still better than inside the loop)
    if (updateRows.length === lastRunUpdates.length && updateRows[0] === 2 && updateRows[updateRows.length - 1] === values.length) {
      // All rows, contiguous
      sheet.getRange(2, lastRunCol, lastRunUpdates.length, 1).setValues(lastRunUpdates);
    } else {
      // Non-contiguous, update each row
      for (let j = 0; j < updateRows.length; j++) {
        sheet.getRange(updateRows[j], lastRunCol, 1, 1).setValues([updateValues[j]]);
      }
    }
  }
  log("Finished processing all cleanup tasks from sheet.", LOG_LEVEL_INFO);
}


/**
 * Deletes emails older than a specified number of days that match specified criteria.
 * Emails are moved to the Trash, not permanently deleted immediately.
 *
 * @param {object} options - An object containing the filter and behavior settings.
 * @param {string} [options.senderEmail] - (Optional) The email address of the sender to filter by.
 * @param {string} [options.recipientEmail] - (Optional) The email address of the recipient to filter by.
 * @param {string} [options.subjectContains] - (Optional) A string that the email subject must contain.
 * @param {number} [options.daysOld=30] - (Optional) Emails older than this many days will be targeted.
 * @param {boolean} [options.dryRun=true] - (Optional) If true, script only logs, without deleting.
 * @param {boolean} [options.excludeStarred=false] - (Optional) If true, starred emails will NOT be deleted. Defaults to false.
 * @param {boolean} [options.excludeImportant=false] - (Optional) If true, important emails will NOT be deleted. Defaults to false.
 * @throws {Error} Throws an error if all content filter parameters are empty or if a Gmail service quota is hit.
 */
function deleteOldEmails(options) {
  const {
    senderEmail,
    recipientEmail,
    subjectContains,
    daysOld = 30,
    dryRun = true,
    excludeStarred = false,
    excludeImportant = false
  } = options || {};

  const effectiveSender = senderEmail ? String(senderEmail).trim() : '';
  const effectiveRecipient = recipientEmail ? String(recipientEmail).trim() : '';
  const effectiveSubject = subjectContains ? String(subjectContains).trim() : '';

  // *** SAFETY CHECK ***
  // Ensure at least one content filter is provided to prevent accidental mass deletion.
  if (!effectiveSender && !effectiveRecipient && !effectiveSubject) {
    throw new Error(ERROR_ALL_FILTERS_EMPTY);
  }

  // Build the Gmail search query string
  let queryParts = [`older_than:${daysOld}d`];
  if (effectiveSender !== '') { queryParts.push(`from:${effectiveSender}`); }
  if (effectiveRecipient !== '') { queryParts.push(`to:${effectiveRecipient}`); }
  if (effectiveSubject !== '') { queryParts.push(`subject:("${effectiveSubject}")`); } // Use quotes for subject to handle spaces
  if (excludeStarred) { queryParts.push(`-is:starred`); }
  if (excludeImportant) { queryParts.push(`-is:important`); }
  const searchQuery = queryParts.join(' ');

  let threadsProcessed = 0;
  let threadsSkipped = 0;

  try {
    log(`    Sub-task: Searching for emails with query: "${searchQuery}"`, LOG_LEVEL_DEBUG);
    log(`    Sub-task: DRY RUN is: ${dryRun}`, LOG_LEVEL_DEBUG);

    let threads;
    try {
      // Fetch a limited number of threads per run to stay within execution limits
      threads = GmailApp.search(searchQuery, 0, MAX_EMAILS_PER_RUN);
    } catch (e) {
      Logger.log(`    Sub-task: Error during Gmail search: ${e.message}`);
      // Re-throw the error to be caught by the outer processAllCleanupTasks function
      throw e;
    }

    if (threads.length === 0) {
      log("    Sub-task: No matching email threads found for this query.", LOG_LEVEL_INFO);
      return { threadsProcessed: 0, threadsSkipped: 0 };
    }

    log(`    Sub-task: Found ${threads.length} matching email threads.`, LOG_LEVEL_INFO);

    for (let i = 0; i < threads.length; i++) {
      const thread = threads[i];
      let logDetail = '';

      try {
        // In dry run, get details from the first message for logging purposes
        if (dryRun) {
          const messages = thread.getMessages();
          if (messages.length === 0) {
            log(`    Sub-task: Warning: Thread at index ${i} has no messages. Skipping.`, LOG_LEVEL_ERROR);
            threadsSkipped++;
            continue;
          }
          const firstMessage = messages[0];
          const subject = firstMessage.getSubject();
          const messageDate = firstMessage.getDate();
          logDetail = ` "${subject}" (Date: ${messageDate.toLocaleString()})`;
        } else {
          // In actual run, just use the thread ID for logging
          logDetail = ` (Thread ID: ${thread.getId()})`;
        }

        if (dryRun) {
          log(`    Sub-task: DRY RUN: Would move thread${logDetail} to trash.`, LOG_LEVEL_DEBUG);
        } else {
          try {
            thread.moveToTrash(); // Move the thread to Gmail's trash
          } catch (e) {
            log(`    Sub-task: Error moving thread${logDetail} to trash: ${e.message}`, LOG_LEVEL_ERROR);
            throw e; // Re-throw to be caught by the outer function
          }
          log(`    Sub-task: Moved thread${logDetail} to trash.`, LOG_LEVEL_INFO);
        }
        threadsProcessed++;

      } catch (innerError) {
        log(`    Sub-task: Error processing thread ${i}: ${innerError.message}. Skipping.`, LOG_LEVEL_ERROR);
        // Re-throw the error to be caught by the outer processAllCleanupTasks function
        throw innerError;
      }

      Utilities.sleep(100); // Small pause between each thread operation to prevent hitting rate limits
    }

    log(`    Sub-task: Finished for query "${searchQuery}". Processed ${threadsProcessed} threads. Skipped ${threadsSkipped} threads.`, LOG_LEVEL_INFO);
    return { threadsProcessed, threadsSkipped };
  } catch (e) {
    log(`    Sub-task: FATAL ERROR (in deleteOldEmails): ${e.toString()}`, LOG_LEVEL_ERROR);
    throw e; // Re-throw the error for the main function to handle quota issues
  }
}
