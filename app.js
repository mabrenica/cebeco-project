import { sheets, SPREADSHEET_ID } from './auth.js';
import { getRecipientEmails } from './getRecipientEmail.js';
import { getNewBillToSheet } from './getNewBillToSheet.js';
import { getUnsentNotifications } from './getUnsentNotifications.js';
import { updateEmailTemplateFromRecord } from './updateEmailTemplateFromRecords.js';
import { sendBillToEmail, sendPaymentNotificationToEmail } from './emailSender.js';
import { updateNotificationStatus } from './updateNotificationStatus.js';
import { getPaymentStatusChange } from './getPaymentStatusChange.js';
import { updateLastPaymentStatus } from './updateLastPaymentStatus.js';
import { createHeaderMap, getColInfo } from './sheetUtils.js';

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
    const initResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'set_up!A1:Z2',
    });
    
    const setupRows = initResponse.data.values || [];
    let accountNumber = null;

    if (setupRows.length > 0) {
      const headerMap = createHeaderMap(setupRows[0]);
      const accountCol = getColInfo(headerMap, 'account_number', 'accountnumber', 'account', 'accountref');
      
      if (accountCol && setupRows[1]) {
        accountNumber = setupRows[1][accountCol.index];
      } else {
        accountNumber = setupRows[1]?.[0] || setupRows[0]?.[0];
      }
    }

    if (!accountNumber) throw new Error("Could not extract account configuration references.");
    
    const emailAddresses = await getRecipientEmails();
    console.log('Configured account reference target value: ' + accountNumber);

    await getNewBillToSheet(accountNumber);

    const unsentNotifications = await getUnsentNotifications();
    for (const recordKey of unsentNotifications) {
      await updateEmailTemplateFromRecord(recordKey);
      await sendBillToEmail(accountNumber, emailAddresses);
      await updateNotificationStatus(recordKey);
    }

    await paymentNotificationManager(emailAddresses);

    console.log("Execution cycle completely concluded.");
  } catch (error) {
    console.error("Critical Runtime Fault Detected inside Application Orchestration:", error);
    process.exit(1);
  }
}

mainFunction();