/**
 * Sets up a time-driven trigger to run the 'processAllCleanupTasks' function
 * every 2 hours.
 *
 * This function should be run manually ONCE to establish the trigger.
 */
function createCleanupTrigger() {
    // Delete any existing triggers for this function to avoid duplicates
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === "processAllCleanupTasks") {
            ScriptApp.deleteTrigger(triggers[i]);
            Logger.log("Deleted existing trigger for processAllCleanupTasks.");
        }
    }
    // Create a new time-driven trigger to run every 2 hours
    ScriptApp.newTrigger("processAllCleanupTasks")
        .timeBased()
        .everyHours(2)
        .create();
    Logger.log("Trigger for 'processAllCleanupTasks' set to run every 2 hours.");
    Logger.log("The script will now clean up old massive emails automatically.");
}
