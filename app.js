import { sheets, SPREADSHEET_ID } from './auth.js';
import { getRecipientEmails } from './getRecipientEmail.js';
import { getNewBillToSheet } from './getNewBillToSheet.js';
import { getUnsentNotifications } from './getUnsentNotifications.js';
import { updateEmailTemplateFromRecord } from './updateEmailTemplateFromRecords.js';
import { sendBillToEmail, sendPaymentNotificationToEmail } from './emailSender.js';
import { updateNotificationStatus } from './updateNotificationStatus.js';
import { getPaymentStatusChange } from './getPaymentStatusChange.js';
import { updateLastPaymentStatus } from './updateLastPaymentStatus.js';

async function paymentNotificationManager(emailAddresses) {
  const paymentStatusChange = await getPaymentStatusChange();
  for (const item of paymentStatusChange) {
    if (item.newStatus === 'PAID') {
      await sendPaymentNotificationToEmail(emailAddresses, item.amount, item.monthYear);
      await updateLastPaymentStatus(item.key, item.newStatus);
    }
  }
}

async function mainFunction() {
  try {
    // 1. Resolve starting setup cell values dynamically across active context connections
    const initResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Set Up!A2',
    });
    
    const accountNumber = initResponse.data.values?.[0]?.[0];
    if (!accountNumber) throw new Error("Could not extract account configuration references.");
    
    const emailAddresses = await getRecipientEmails();
    console.log('Configured account reference target value: ' + accountNumber);

    // 2. Fetch remote scraping mutations
    await getNewBillToSheet(accountNumber);

    // 3. Process mailing actions
    const unsentNotifications = await getUnsentNotifications();
    for (const recordKey of unsentNotifications) {
      await updateEmailTemplateFromRecord(recordKey);
      await sendBillToEmail(accountNumber, emailAddresses);
      await updateNotificationStatus(recordKey);
    }

    // 4. Handle incoming closures 
    await paymentNotificationManager(emailAddresses);

    console.log("Execution cycle completely concluded.");
  } catch (error) {
    console.error("Critical Runtime Fault Detected inside Application Orchestration:", error);
  }
}

// Kick off the application execution cycle
mainFunction();