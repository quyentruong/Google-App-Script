/**
 * Checks stock prices from a Google Sheet and sends email alerts when
 * a buy price (current price <= buy price) or a sell price (current price >= sell price) is reached.
 * This version includes detailed logging to the Apps Script execution log.
 * It also adds logic to prevent repeated email notifications based on a 'Last Notified' timestamp.
 * The email body has been updated for a more professional appearance using HTML,
 * including color coding for price values and a link to Yahoo Finance.
 *
 * This function is designed to be run by a time-driven trigger.
 * It includes a check to ensure it only runs on weekdays (Monday-Friday).
 */
function checkStockPrices() {
  const today = new Date();
  const dayOfWeek = today.getDay(); // 0 = Sunday, 1 = Monday, ..., 6 = Saturday

  // Check if it's a weekend (Sunday = 0, Saturday = 6)
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    Logger.log(
      `--- Skipping Stock Price Check: It's a weekend (${today.toLocaleDateString()}). ---`
    );
    return; // Exit the function if it's a weekend
  }

  Logger.log("--- Starting Stock Price Check ---");
  Logger.log(`  Running on a weekday: ${today.toLocaleDateString()}`);

  // Get the active spreadsheet and the sheet named 'Sheet1'.
  // IMPORTANT: Make sure your sheet is named 'Sheet1' or update this line.
  const sheetName = "Sheet1"; // Customize this if your sheet has a different name
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  if (!sheet) {
    Logger.log(
      `Error: Sheet named '${sheetName}' not found. Please check your sheet name.`
    );
    return;
  }

  // Get the recipient email from 'Sheet2', assuming it's in cell A1.
  // IMPORTANT: Create a sheet named 'Sheet2' and put the recipient email in cell A1.
  const recipientSheetName = "Sheet2";
  const recipientSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(recipientSheetName);
  let recipientEmail = "";

  if (recipientSheet) {
    recipientEmail = recipientSheet.getRange("A1").getValue();
    if (
      !recipientEmail ||
      typeof recipientEmail !== "string" ||
      !recipientEmail.includes("@")
    ) {
      Logger.log(
        `Error: Recipient email in '${recipientSheetName}' cell A1 is invalid or empty: ${recipientEmail}. Using default.`
      );
      recipientEmail = "default_email@example.com"; // Fallback email
    }
  } else {
    Logger.log(
      `Error: Recipient sheet named '${recipientSheetName}' not found. Using default email.`
    );
    recipientEmail = "default_email@example.com"; // Fallback email
  }
  Logger.log(`Recipient Email: ${recipientEmail}`);

  // Get all data from the sheet. Assumes the first row is a header.
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  // Define the notification interval (e.g., 24 hours)
  const NOTIFY_INTERVAL_HOURS = 24;
  const NOTIFY_INTERVAL_MS = NOTIFY_INTERVAL_HOURS * 60 * 60 * 1000; // Convert hours to milliseconds

  // Check if there's any data beyond the header row
  if (values.length <= 1) {
    Logger.log("No stock data found in the sheet (only header or empty).");
    return;
  }

  Logger.log(
    `Processing ${values.length - 1} stock entries from sheet '${sheetName}'.`
  );

  // Iterate through each row, starting from the second row (index 1) to skip headers.
  // Ensure your Google Sheet has the following columns (adjust indices if different):
  // Column A: Ticker (index 0)
  // Column B: Current Price (index 1) - This should be populated by =GOOGLEFINANCE()
  // Column C: Buy Price (index 2) - Your desired price to buy
  // Column D: Sell Price (index 3) - Your desired price to sell
  // Column E: Last Notified (index 4) - For the script to track notification timestamps
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const ticker = row[0]; // Column A
    const currentPrice = row[1]; // Column B
    const buyPrice = row[2]; // Column C
    const sellPrice = row[3]; // Column D
    const lastNotified = row[4]; // Column E

    Logger.log(`\n--- Checking Ticker: ${ticker} (Row ${i + 1}) ---`);
    Logger.log(`  Current Price: ${currentPrice}`);
    Logger.log(`  Buy Price: ${buyPrice}`);
    Logger.log(`  Sell Price: ${sellPrice}`);
    Logger.log(
      `  Last Notified: ${
        lastNotified ? lastNotified.toLocaleString() : "Never"
      }`
    );

    // Validate inputs to ensure they are numbers
    if (typeof currentPrice !== "number") {
      Logger.log(
        `  Warning: Current Price (${currentPrice}) for ${ticker} is not a number. Skipping.`
      );
      continue;
    }
    // Check if buyPrice is a valid number, if not, set to a very low value to avoid triggering
    const validBuyPrice = typeof buyPrice === "number" ? buyPrice : -Infinity;
    // Check if sellPrice is a valid number, if not, set to a very high value to avoid triggering
    const validSellPrice = typeof sellPrice === "number" ? sellPrice : Infinity;

    let alertType = null; // 'buy' or 'sell'
    let alertConditionMet = false;

    // Check for Buy Alert condition
    if (currentPrice <= validBuyPrice) {
      Logger.log(
        `  Potential Buy Condition Met: Current Price (${currentPrice}) <= Buy Price (${validBuyPrice}) for ${ticker}.`
      );
      alertType = "buy";
      alertConditionMet = true;
    }

    // Check for Sell Alert condition (only if buy condition isn't met, or if you want both alerts)
    // For simplicity, we'll prioritize buy if both are met, or you can send two emails.
    // Here, we check sell only if buy wasn't triggered OR if current price is significantly above buy.
    if (currentPrice >= validSellPrice) {
      Logger.log(
        `  Potential Sell Condition Met: Current Price (${currentPrice}) >= Sell Price (${validSellPrice}) for ${ticker}.`
      );
      // If both buy and sell conditions are met (e.g., current price is between buy and sell),
      // you might want to decide which alert to send or send both.
      // For this example, if both are met, it will default to 'sell' if it's checked second.
      // A more robust solution might send both or have a preference.
      alertType = "sell";
      alertConditionMet = true;
    }

    if (alertConditionMet) {
      const now = new Date();
      let shouldSendEmail = false;

      // Check if an email has been sent recently for this ticker
      if (!lastNotified || !(lastNotified instanceof Date)) {
        // If 'Last Notified' is empty or not a valid date, send email
        Logger.log(
          `  'Last Notified' is empty or invalid. Preparing to send email.`
        );
        shouldSendEmail = true;
      } else {
        // Calculate time difference in milliseconds
        const timeSinceLastNotification =
          now.getTime() - lastNotified.getTime();
        Logger.log(
          `  Time since last notification: ${
            timeSinceLastNotification / 1000 / 60
          } minutes.`
        );

        if (timeSinceLastNotification >= NOTIFY_INTERVAL_MS) {
          // If enough time has passed since the last notification, send email
          Logger.log(
            `  Enough time (${NOTIFY_INTERVAL_HOURS} hours) has passed since last notification. Preparing to send email.`
          );
          shouldSendEmail = true;
        } else {
          Logger.log(
            `  Notification for ${ticker} sent within the last ${NOTIFY_INTERVAL_HOURS} hours. Skipping email.`
          );
        }
      }

      if (shouldSendEmail) {
        let subject = "";
        let htmlBody = ""; // Use htmlBody for rich text emails

        // Define colors for better visual distinction
        const currentColor = "#007bff"; // Blue for current price
        const buyColor = "#28a745"; // Green for buy price
        const sellColor = "#dc3545"; // Red for sell price

        // Construct the Yahoo Finance link
        const yahooFinanceLink = `https://finance.yahoo.com/quote/${ticker}`;

        if (alertType === "buy") {
          subject = `Stock Alert: Potential Buy Opportunity for ${ticker}`;
          htmlBody = `
            <p>Dear Investor,</p>
            <p>This is an automated notification regarding <strong>${ticker}</strong>.</p>
            <p>The current price of <strong>${ticker}</strong> has reached or fallen below your specified <strong>Buy Price</strong>.</p>
            <table style="width:100%; border-collapse: collapse; margin-top: 15px;">
              <tr>
                <td style="padding: 8px; border: 1px solid #ddd; background-color: #f2f2f2;"><strong>Current Price:</strong></td>
                <td style="padding: 8px; border: 1px solid #ddd;"><strong style="color: ${currentColor};">$${currentPrice.toFixed(
            2
          )}</strong></td>
              </tr>
              <tr>
                <td style="padding: 8px; border: 1px solid #ddd; background-color: #f2f2f2;"><strong>Your Buy Price:</strong></td>
                <td style="padding: 8px; border: 1px solid #ddd;"><strong style="color: ${buyColor};">$${validBuyPrice.toFixed(
            2
          )}</strong></td>
              </tr>
            </table>
            <p style="margin-top: 20px;">For more details, please visit the <a href="${yahooFinanceLink}">Yahoo Finance page for ${ticker}</a> and consider this potential buying opportunity.</p>
            <p>Best regards,<br>Your Automated Stock Tracker</p>
          `;
        } else if (alertType === "sell") {
          subject = `Stock Alert: Potential Sell Opportunity for ${ticker}`;
          htmlBody = `
            <p>Dear Investor,</p>
            <p>This is an automated notification regarding <strong>${ticker}</strong>.</p>
            <p>The current price of <strong>${ticker}</strong> has reached or exceeded your specified <strong>Sell Price</strong>.</p>
            <table style="width:100%; border-collapse: collapse; margin-top: 15px;">
              <tr>
                <td style="padding: 8px; border: 1px solid #ddd; background-color: #f2f2f2;"><strong>Current Price:</strong></td>
                <td style="padding: 8px; border: 1px solid #ddd;"><strong style="color: ${currentColor};">$${currentPrice.toFixed(
            2
          )}</strong></td>
              </tr>
              <tr>
                <td style="padding: 8px; border: 1px solid #ddd; background-color: #f2f2f2;"><strong>Your Sell Price:</strong></td>
                <td style="padding: 8px; border: 1px solid #ddd;"><strong style="color: ${sellColor};">$${validSellPrice.toFixed(
            2
          )}</strong></td>
              </tr>
            </table>
            <p style="margin-top: 20px;">For more details, please visit the <a href="${yahooFinanceLink}">Yahoo Finance page for ${ticker}</a> and consider this potential selling opportunity.</p>
            <p>Best regards,<br>Your Automated Stock Tracker</p>
          `;
        }

        try {
          // Send email with HTML body
          MailApp.sendEmail(recipientEmail, subject, "", {
            htmlBody: htmlBody,
          });
          Logger.log(
            `  Email sent successfully to ${recipientEmail} for ${ticker} (${alertType} alert).`
          );
          // Update the 'Last Notified' column in your sheet with the current timestamp
          sheet.getRange(i + 1, 5).setValue(now); // Column E is index 5 (1-based)
          Logger.log(
            `  Updated 'Last Notified' timestamp for ${ticker} to ${now.toLocaleString()}.`
          );
        } catch (e) {
          Logger.log(`  Error sending email for ${ticker}: ${e.toString()}`);
        }
      }
    } else {
      Logger.log(`  Neither Buy nor Sell condition met for ${ticker}.`);
    }
  }

  Logger.log("\n--- Stock Price Check Completed ---");
}
