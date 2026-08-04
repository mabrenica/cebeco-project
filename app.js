import { getAccountConfigurations } from './getRecipientEmail.js';
import { getNewBillToSheet } from './getNewBillToSheet.js';
import { getAccountBillData } from './getUnsentNotifications.js';
import { sendBillToEmail, sendPaymentNotificationToEmail } from './emailSender.js';
import { updateNotificationStatus } from './updateNotificationStatus.js';
import { getPaymentStatusChange } from './getPaymentStatusChange.js';
import { updateLastPaymentStatus } from './updateLastPaymentStatus.js';
import { processUnpaidBillReminders } from './billReminderManager.js';

async function paymentNotificationManager(accountNumber, emailAddresses) {
  const paymentStatusChange = await getPaymentStatusChange(accountNumber);
  for (const item of paymentStatusChange) {
    if (item.newStatus && item.newStatus.trim().toUpperCase() === 'PAID') {
      await sendPaymentNotificationToEmail(emailAddresses, item.amount, item.monthYear);
      await updateLastPaymentStatus(item.key, item.newStatus);
    }
  }
}

async function mainFunction() {
  try {
    const accountConfigs = await getAccountConfigurations();

    if (accountConfigs.length === 0) {
      throw new Error("Could not extract account configuration references from set_up sheet.");
    }

    for (const config of accountConfigs) {
      console.log(`Processing bill sync for account: ${config.accountNumber} (${config.accountName})`);
      
      // 1. Sync latest online bill records
      await getNewBillToSheet(config.accountNumber);

      // 2. Process unsent bill notifications for this account
      const { primaryBill, historicalBills, unsentKeys } = await getAccountBillData(config.accountNumber);

      if (primaryBill && config.emailAddresses.length > 0) {
        await sendBillToEmail(
          config.accountName,
          config.emailAddresses,
          primaryBill,
          historicalBills
        );

        for (const key of unsentKeys) {
          await updateNotificationStatus(key);
        }
      }

      // 3. Process payment status updates for this account
      if (config.emailAddresses.length > 0) {
        await paymentNotificationManager(config.accountNumber, config.emailAddresses);
      }

      // 4. Process scheduled payment reminders for unpaid bills (-3, -1, +1, +3 days)
      if (config.emailAddresses.length > 0) {
        await processUnpaidBillReminders(config);
      }
    }

    console.log("Execution cycle completely concluded.");
  } catch (error) {
    console.error("Critical Runtime Fault Detected inside Application Orchestration:", error);
    process.exit(1);
  }
}

mainFunction();