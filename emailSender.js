import nodemailer from 'nodemailer';
import fs from 'fs';
import path from 'path';
import { GMAIL_USER, GMAIL_APP_PASS } from './auth.js';

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

function generateBillEmailHtml(accountName, primaryBill, historicalBills = []) {
  if (!primaryBill) {
    throw new Error("Cannot generate bill email HTML: primaryBill object is missing or undefined.");
  }

  const templatePath = path.resolve('./billNotificationTemplate.html');

  if (!fs.existsSync(templatePath)) {
    throw new Error('CRITICAL ERROR: billNotificationTemplate.html missing from project root directory.');
  }

  let htmlTemplate = fs.readFileSync(templatePath, 'utf8');

  let historyRowsHtml = '';
  if (historicalBills.length > 0) {
    historyRowsHtml = historicalBills.map(bill => `
      <tr>
        <td style="padding: 10px 12px; border-bottom: 1px solid #E2E8F0; text-align: left; color: #334155;">${bill.monthYear || '-'}</td>
        <td style="padding: 10px 12px; border-bottom: 1px solid #E2E8F0; text-align: center; color: #334155;">${bill.kwhRate || '-'}</td>
        <td style="padding: 10px 12px; border-bottom: 1px solid #E2E8F0; text-align: center; color: #334155;">${bill.kwhUsed || '-'}</td>
        <td style="padding: 10px 12px; border-bottom: 1px solid #E2E8F0; text-align: right; font-weight: bold; color: #0F172A;">${bill.billAmount || '-'}</td>
      </tr>
    `).join('');
  } else {
    historyRowsHtml = `
      <tr>
        <td colspan="4" style="padding: 14px; text-align: center; color: #94A3B8; font-style: italic;">No prior billing history available.</td>
      </tr>
    `;
  }

  return htmlTemplate
    .replace(/{{ACCOUNT_NAME}}/g, accountName || 'Valued Customer')
    .replace(/{{MONTH_YEAR}}/g, primaryBill.monthYear || 'N/A')
    .replace(/{{TOTAL_AMOUNT}}/g, primaryBill.billAmount || '₱0.00')
    .replace(/{{ACCOUNT_NUMBER}}/g, primaryBill.accountNumber || 'N/A')
    .replace(/{{DUE_DATE}}/g, formatDueDate(primaryBill.dueDate))
    .replace(/{{KWH_RATE}}/g, primaryBill.kwhRate || 'N/A')
    .replace(/{{PREV_READING}}/g, primaryBill.prevReading || 'N/A')
    .replace(/{{PRESENT_READING}}/g, primaryBill.presentReading || 'N/A')
    .replace(/{{KWH_USED}}/g, primaryBill.kwhUsed || '0')
    .replace(/{{HISTORY_ROWS}}/g, historyRowsHtml);
}

export async function sendBillToEmail(accountName, emailAddresses, primaryBill, historicalBills = []) {
  if (!primaryBill || !primaryBill.monthYear) {
    console.warn("sendBillToEmail skipped: primaryBill object is invalid or missing.");
    return;
  }

  const htmlContent = generateBillEmailHtml(accountName, primaryBill, historicalBills);

  const mailOptions = {
    from: `"CEBECO 1 Billing" <${GMAIL_USER}>`,
    to: emailAddresses.join(','),
    subject: `CEBECO 1 Bill - ${primaryBill.monthYear}`,
    html: htmlContent
  };

  await transporter.sendMail(mailOptions);
  console.log(`Bill notification for ${primaryBill.monthYear} (${primaryBill.accountNumber}) sent to ${emailAddresses.join(', ')}.`);
}

export async function sendPaymentNotificationToEmail(emailAddresses, amount, monthYear) {
  const mailOptions = {
    from: `"CEBECO 1 Billing" <${GMAIL_USER}>`,
    to: emailAddresses.join(','),
    subject: 'CEBECO I - Payment Notification',
    html: `
      <div style="font-family: Arial, sans-serif; padding: 20px; color: #334155;">
        <h2>Payment Confirmed</h2>
        <p>Thank you for your payment.</p>
        <p>We would like to confirm that your payment for <strong>${monthYear}</strong> bill with an amount of <strong>${amount}</strong> has been successfully posted. Cheers!</p>
        <hr style="border: 0; border-top: 1px solid #E2E8F0; margin: 20px 0;">
        <p style="font-size: 12px; color: #64748B;">Powered by Marnel Abrenica Automations</p>
      </div>
    `
  };

  await transporter.sendMail(mailOptions);
  console.log(`Payment notification for ${monthYear} successfully sent.`);
}