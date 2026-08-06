import nodemailer from 'nodemailer';
import fs from 'fs';
import path from 'path';
import { sheets, SPREADSHEET_ID, GMAIL_USER, GMAIL_APP_PASS } from './auth.js';
import { createHeaderMap, getColInfo } from './sheetUtils.js';

const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: GMAIL_USER,
    pass: GMAIL_APP_PASS
  }
});

function formatDueDate(dateStr) {
  if (!dateStr) return 'N/A';
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return dateStr;

  const day = String(d.getDate()).padStart(2, '0');
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const month = months[d.getMonth()];
  const year = String(d.getFullYear()).slice(-2);

  return `${day}-${month}-${year}`;
}

function getDaysDifferenceFromToday(dueDateStr) {
  if (!dueDateStr) return null;
  const dueDate = new Date(dueDateStr);
  if (isNaN(dueDate.getTime())) return null;

  const now = new Date();
  const todayMs = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  const dueMs = new Date(dueDate.getFullYear(), dueDate.getMonth(), dueDate.getDate()).getTime();

  return Math.round((todayMs - dueMs) / (1000 * 60 * 60 * 24));
}

function generateReminderHtml(accountName, bill, diffDays) {
  const templatePath = path.resolve('./billReminderTemplate.html');

  if (!fs.existsSync(templatePath)) {
    throw new Error('CRITICAL ERROR: billReminderTemplate.html missing from project root directory.');
  }

  let htmlTemplate = fs.readFileSync(templatePath, 'utf8');

  const formattedDueDate = formatDueDate(bill.dueDate);
  const onlineAccountLink = `<a href="https://www.cebeco1.online/general-services/bill-inquiry/${bill.accountNumber}" target="_blank" style="color: #2563EB; font-weight: 600; text-decoration: underline;">click here to view your account online</a>`;

  let reminderMessage = '';

  if (diffDays > 0) {
    // Overdue Message
    reminderMessage = `Good day, <strong>${accountName || 'Valued Customer'}</strong>. This is a reminder that your electricity bill for <strong>${bill.monthYear || 'N/A'}</strong> in the amount of <strong>${bill.billAmount || '₱0.00'}</strong>, which was due on <strong>${formattedDueDate}</strong>, is still unpaid. Please settle your payment to avoid service interruptions, or ${onlineAccountLink}.`;
  } else if (diffDays === 0) {
    // Due Date Day Message
    reminderMessage = `Good day, <strong>${accountName || 'Valued Customer'}</strong>. This is a reminder that your electricity bill for <strong>${bill.monthYear || 'N/A'}</strong> in the amount of <strong>${bill.billAmount || '₱0.00'}</strong> is due today (<strong>${formattedDueDate}</strong>). Please settle your payment to avoid penalties or service interruptions, or ${onlineAccountLink}.`;
  } else {
    // Upcoming Message
    reminderMessage = `Good day, <strong>${accountName || 'Valued Customer'}</strong>. This is a reminder regarding your unpaid electricity bill for <strong>${bill.monthYear || 'N/A'}</strong> in the amount of <strong>${bill.billAmount || '₱0.00'}</strong>, which is due on <strong>${formattedDueDate}</strong>. Please settle your payment to avoid penalties, or ${onlineAccountLink}.`;
  }

  return htmlTemplate
    .replace(/{{REMINDER_MESSAGE}}/g, reminderMessage)
    .replace(/{{ACCOUNT_NAME}}/g, accountName || 'Valued Customer')
    .replace(/{{MONTH_YEAR}}/g, bill.monthYear || 'N/A')
    .replace(/{{TOTAL_AMOUNT}}/g, bill.billAmount || '₱0.00')
    .replace(/{{ACCOUNT_NUMBER}}/g, bill.accountNumber || 'N/A')
    .replace(/{{DUE_DATE}}/g, formattedDueDate)
    .replace(/{{KWH_RATE}}/g, bill.kwhRate || 'N/A')
    .replace(/{{PREV_READING}}/g, bill.prevReading || 'N/A')
    .replace(/{{PRESENT_READING}}/g, bill.presentReading || 'N/A')
    .replace(/{{KWH_USED}}/g, bill.kwhUsed || '0');
}

async function updateReminderStatusCell(rowIndex, colLetter, newStatus) {
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `online_bill_records!${colLetter}${rowIndex}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[newStatus]] },
  });
}

export async function processUnpaidBillReminders(config) {
  try {
    const { accountNumber, accountName, emailAddresses } = config;
    if (!emailAddresses || emailAddresses.length === 0) return;

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'online_bill_records!A:Z',
    });

    const data = response.data.values || [];
    if (data.length < 2) return;

    const headerMap = createHeaderMap(data[0]);
    const keyCol = getColInfo(headerMap, 'key', 'recordkey');
    const monthCol = getColInfo(headerMap, 'month_year', 'billingmonth', 'monthyear');
    const accCol = getColInfo(headerMap, 'account_number', 'accountnumber', 'account');
    const presentCol = getColInfo(headerMap, 'present_reading', 'presentreading');
    const prevCol = getColInfo(headerMap, 'previous_reading', 'previousreading');
    const kwhCol = getColInfo(headerMap, 'kwh_used', 'kwhused');
    const dueDateCol = getColInfo(headerMap, 'due_date', 'duedate');
    const amountCol = getColInfo(headerMap, 'bill_amount', 'billamount');
    const statusCol = getColInfo(headerMap, 'status', 'billstatus');
    const lastPaymentCol = getColInfo(headerMap, 'last_payment_status', 'paymentstatus');
    const kwhRateCol = getColInfo(headerMap, 'kwh_rate', 'kwhrate');
    const reminderStatusCol = getColInfo(headerMap, 'notification_reminder_status', 'reminderstatus', 'notifreminderstatus');

    if (!reminderStatusCol) {
      console.warn("notification_reminder_status column not found in online_bill_records header.");
      return;
    }

    const cleanTargetAcc = String(accountNumber).trim().padStart(11, '0');

    let latestBill = null;
    let targetRowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      const item = data[i];
      const recordKey = keyCol ? item[keyCol.index] : item[0];
      let rowAcc = accCol ? String(item[accCol.index] || '').trim().padStart(11, '0') : '';

      if ((!rowAcc || rowAcc === '00000000000') && recordKey) {
        const parts = recordKey.split('_');
        if (parts.length > 1) {
          rowAcc = String(parts[parts.length - 1]).trim().padStart(11, '0');
        }
      }

      if (rowAcc === cleanTargetAcc) {
        targetRowIndex = i + 1; // 1-based index for Google Sheets
        latestBill = {
          accountNumber: rowAcc,
          monthYear: monthCol ? item[monthCol.index] : '',
          presentReading: presentCol ? item[presentCol.index] : '',
          prevReading: prevCol ? item[prevCol.index] : '',
          kwhUsed: kwhCol ? item[kwhCol.index] : '',
          dueDate: dueDateCol ? item[dueDateCol.index] : '',
          billAmount: amountCol ? item[amountCol.index] : '',
          status: statusCol ? item[statusCol.index] : '',
          lastPaymentStatus: lastPaymentCol ? item[lastPaymentCol.index] : '',
          kwhRate: kwhRateCol ? item[kwhRateCol.index] : '',
          reminderStatus: (item[reminderStatusCol.index] || 'unpaid').toString().trim().toLowerCase()
        };
        break;
      }
    }

    if (!latestBill || targetRowIndex === -1) return;

    const currentStatus = (latestBill.status || '').toString().trim().toUpperCase();
    const lastPaymentStatus = (latestBill.lastPaymentStatus || '').toString().trim().toUpperCase();

    // Rule: If bill is paid, update status to 'paid' and skip sending
    if (currentStatus === 'PAID' || lastPaymentStatus === 'PAID') {
      if (latestBill.reminderStatus !== 'paid') {
        await updateReminderStatusCell(targetRowIndex, reminderStatusCol.letter, 'paid');
      }
      return;
    }

    // Rule: If already reached 2nd_overdue_reminder, ignore
    if (latestBill.reminderStatus === '2nd_overdue_reminder' || latestBill.reminderStatus === 'paid') {
      return;
    }

    const diffDays = getDaysDifferenceFromToday(latestBill.dueDate);
    let nextReminderStatus = null;
    let reminderSubject = '';

    // Schedule evaluation with state progression
    if (diffDays === -3 && latestBill.reminderStatus === 'unpaid') {
      nextReminderStatus = '1st_reminder';
      reminderSubject = 'Upcoming Bill Reminder - Due in 3 Days';
    } else if (diffDays === -1 && latestBill.reminderStatus === '1st_reminder') {
      nextReminderStatus = '2nd_reminder';
      reminderSubject = 'Upcoming Bill Reminder - Due Tomorrow';
    } else if (diffDays === 0 && latestBill.reminderStatus === '2nd_reminder') {
      nextReminderStatus = '3rd_reminder';
      reminderSubject = 'Bill Due Today Reminder';
    } else if (diffDays === 1 && latestBill.reminderStatus === '3rd_reminder') {
      nextReminderStatus = '1st_overdue_reminder';
      reminderSubject = 'Overdue Bill Reminder - 1 Day Past Due';
    } else if (diffDays === 3 && latestBill.reminderStatus === '1st_overdue_reminder') {
      nextReminderStatus = '2nd_overdue_reminder';
      reminderSubject = 'Overdue Bill Reminder - 3 Days Past Due';
    }

    if (nextReminderStatus) {
      const htmlContent = generateReminderHtml(accountName, latestBill, diffDays);

      const mailOptions = {
        from: `"CEBECO 1 Billing" <${GMAIL_USER}>`,
        to: emailAddresses.join(','),
        subject: `CEBECO 1 - ${reminderSubject} (${latestBill.monthYear})`,
        html: htmlContent
      };

      await transporter.sendMail(mailOptions);
      await updateReminderStatusCell(targetRowIndex, reminderStatusCol.letter, nextReminderStatus);
      console.log(`Payment reminder [${nextReminderStatus}] sent for ${latestBill.monthYear} (${cleanTargetAcc}) to ${emailAddresses.join(', ')}.`);
    }
  } catch (error) {
    console.error(`Error processing bill reminders for account ${config.accountNumber}:`, error);
  }
}