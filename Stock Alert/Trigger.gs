/**
 * Sets up a time-driven trigger to run the 'checkStockPrices' function
 * every 10 minutes.
 *
 * This function should be run manually ONCE to establish the trigger.
 */
function createStockPriceTrigger() {
    // Delete any existing triggers for this function to avoid duplicates
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === "checkStockPrices") {
            ScriptApp.deleteTrigger(triggers[i]);
            Logger.log("Deleted existing trigger for checkStockPrices.");
        }
    }

    // Create a new time-driven trigger to run every 10 minutes
    ScriptApp.newTrigger("checkStockPrices")
        .timeBased()
        .everyMinutes(10)
        .create();

    Logger.log("Trigger for 'checkStockPrices' set to run every 10 minutes.");
    Logger.log("The script will now check for weekdays internally.");
}