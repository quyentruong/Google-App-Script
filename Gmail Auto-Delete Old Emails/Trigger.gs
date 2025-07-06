/**
 * Sets up a time-driven trigger to run the 'cleanupOldMassiveEmail' function
 * every 2 hours.
 *
 * This function should be run manually ONCE to establish the trigger.
 */
function createCleanupTrigger() {
    // Delete any existing triggers for this function to avoid duplicates
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === "cleanupOldMassiveEmail") {
            ScriptApp.deleteTrigger(triggers[i]);
            Logger.log("Deleted existing trigger for cleanupOldMassiveEmail.");
        }
    }
    // Create a new time-driven trigger to run every 2 hours
    ScriptApp.newTrigger("cleanupOldMassiveEmail")
        .timeBased()
        .everyHours(2)
        .create();
    Logger.log("Trigger for 'cleanupOldMassiveEmail' set to run every 2 hours.");
    Logger.log("The script will now clean up old massive emails automatically.");
}

function cleanupMultipleSenders() {
    // This function cleans up emails from multiple senders in a batch process.
    // It is designed to handle a predefined list of senders and delete emails older than a specified    
    const sendersToClean = [
        "newegg.com",
        "bestbuy.com",
        "modpizza.com",
        "fanatical.com",
        "gog.com",
        "quora.com",
        "store-news@amazon.com",
        "groupupdates@facebookmail.com",
        "friendupdates@facebookmail.com"
    ];

    // Define common options for this batch cleanup
    const commonOptions = {
        recipientEmail: null, // Apply to any recipient (or specify one if needed)
        subjectContains: null, // Any subject
        daysOld: 20, // Emails older than 45 days
        dryRun: false, // KEEP THIS AS TRUE FOR TESTING! Change to 'false' when ready for live deletion.
        excludeStarred: false, // Exclude starred emails from deletion
        excludeImportant: false, // Exclude important emails from deletion
    };

    Logger.log(`Starting batch cleanup for ${sendersToClean.length} senders.`);

    for (let i = 0; i < sendersToClean.length; i++) {
        const currentSender = sendersToClean[i];
        Logger.log(`--- Processing sender: ${currentSender} ---`);

        // Create a new options object for each call, overriding the senderEmail
        const currentOptions = { ...commonOptions, senderEmail: currentSender };

        try {
            _deleteOldEmails(currentOptions);
        } catch (e) {
            // Check if the error message indicates a quota limit
            if (
                e.message.includes("Service invoked too many times") ||
                e.message.includes("Service using too much computer time")
            ) {
                Logger.log(
                    `Daily quota hit while processing sender ${currentSender}. Stopping further actions for today.`
                );
                // Optionally, send an email notification about the quota hit
                // MailApp.sendEmail("your_email@example.com", "Apps Script Quota Hit", "Gmail cleanup script hit daily quota while processing " + currentSender);
                return; // Stop the loop and exit the function
            }
            Logger.log(
                `Error processing sender ${currentSender} (not a quota issue): ${e.toString()}`
            );
            // If it's another type of error, continue to the next sender
        }
        Utilities.sleep(500); // Small pause between each sender to respect API limits
    }
    Logger.log("Batch cleanup for multiple senders finished.");
}

function cleanupOldMassiveEmail() {
    cleanupMultipleSenders();
    try {
        _deleteOldEmails({
            senderEmail: "bankofamerica.com",
            subjectContains: "Your Available Balance",
            daysOld: 10,
            dryRun: false
        });
        // Add more function calls here for other senders if needed
    }
    catch (e) {
        if (e.message.includes("Service invoked too many times") || e.message.includes("Service using too much computer time")) {
            Logger.log(`Daily quota hit while processing. Stopping further actions for today.`);
            // Optionally, send an email notification about the quota hit
            // MailApp.sendEmail("your_email@example.com", "Apps Script Quota Hit", "Gmail cleanup script hit daily quota while processing " + currentSender);
            return; // Stop the loop and exit the function
        }
    }
}
