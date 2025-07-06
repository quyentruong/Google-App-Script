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
 */
function _deleteOldEmails(options) {
    const {
        senderEmail,
        recipientEmail,
        subjectContains,
        daysOld = 30,
        dryRun = true,
        excludeStarred = false,
        excludeImportant = false,
    } = options || {};

    const MAX_EMAILS_PER_RUN = 100;

    const effectiveSender = senderEmail ? senderEmail.trim() : "";
    const effectiveRecipient = recipientEmail ? recipientEmail.trim() : "";
    const effectiveSubject = subjectContains ? subjectContains.trim() : "";

    if (!effectiveSender && !effectiveRecipient && !effectiveSubject) {
        throw new Error(
            "ERROR: All content filter parameters (sender, recipient, subject) are empty. " +
            "This script will not proceed to prevent accidental mass deletion of all old emails. " +
            "Please provide at least one content filter criterion."
        );
    }

    let queryParts = [`older_than:${daysOld}d`];
    if (effectiveSender !== "") {
        queryParts.push(`from:${effectiveSender}`);
    }
    if (effectiveRecipient !== "") {
        queryParts.push(`to:${effectiveRecipient}`);
    }
    if (effectiveSubject !== "") {
        queryParts.push(`subject:${effectiveSubject}`);
    }
    if (excludeStarred) {
        queryParts.push(`-is:starred`);
    }
    if (excludeImportant) {
        queryParts.push(`-is:important`);
    }
    const searchQuery = queryParts.join(" ");

    let threadsProcessed = 0;
    let threadsSkipped = 0;

    try {
        Logger.log(`Script Start Time: ${new Date().toLocaleString()}`);
        Logger.log(
            `Running with filters - From: "${effectiveSender || "Any"}", To: "${effectiveRecipient || "Any"
            }", Subject Contains: "${effectiveSubject || "Any"}"`
        );
        Logger.log(`Emails older than: ${daysOld} days.`);
        Logger.log(
            `Exclude Starred: ${excludeStarred}, Exclude Important: ${excludeImportant}`
        );
        Logger.log(`Searching for emails with query: "${searchQuery}"`);
        Logger.log(`DRY RUN is: ${dryRun}`);

        let threads;
        try {
            threads = GmailApp.search(searchQuery, 0, MAX_EMAILS_PER_RUN);
        } catch (e) {
            Logger.log(`Error during Gmail search: ${e.message}`);
            throw e;
        }

        if (threads.length === 0) {
            Logger.log("No matching email threads found to process.");
            return;
        }

        Logger.log(`Found ${threads.length} matching email threads.`);

        for (let i = 0; i < threads.length; i++) {
            const thread = threads[i];

            let logDetail = ""; // Initialize log detail string

            try {
                if (dryRun) {
                    // ONLY call getMessages() and extract details if it's a dry run
                    const messages = thread.getMessages(); // This call consumes quota
                    if (messages.length === 0) {
                        Logger.log(
                            `Warning: Thread at index ${i} has no messages. Skipping.`
                        );
                        threadsSkipped++;
                        continue;
                    }
                    const firstMessage = messages[0];
                    const subject = firstMessage.getSubject();
                    const messageDate = firstMessage.getDate();
                    logDetail = ` "${subject}" (${messageDate.toLocaleString()})`;
                } else {
                    // In live run, we skip getMessages() to save quota.
                    // We can still log the thread ID for reference if needed.
                    // Note: thread.getId() is usually not an API call that hits quotas like getMessages().
                    // It's part of the object properties.
                    logDetail = ` (Thread ID: ${thread.getId()})`;
                }

                if (dryRun) {
                    Logger.log(`DRY RUN: Would move thread${logDetail} to trash.`);
                } else {
                    try {
                        thread.moveToTrash(); // This call consumes quota
                    } catch (e) {
                        Logger.log(`Error during move to trash: ${e.message}`);
                        throw e;
                    }
                    Logger.log(`Moved thread${logDetail} to trash.`); // Logging less detail in live run
                }
                threadsProcessed++;
            } catch (innerError) {
                Logger.log(`Error processing thread ${i}: ${innerError.message}.`);
                throw innerError;
            }

            Utilities.sleep(100);
        }

        Logger.log(
            `Script finished for query "${searchQuery}". Processed ${threadsProcessed} threads. Skipped ${threadsSkipped} threads.`
        );
    } catch (e) {
        Logger.log(
            `FATAL ERROR (main block) for query "${searchQuery}" (or initial check): ${e.toString()}`
        );
        throw e;
    }
}
