import { sheets, SPREADSHEET_ID } from './auth.js';
import { getOnlineBillRecord } from './getOnlineBillRecords.js';
import { sortBillRecord } from './sortBillRecords.js';
import { createHeaderMap, getColInfo, colIndexToLetter } from './sheetUtils.js';

function parseNumericValue(val) {
  if (val === null || val === undefined) return 0;
  const cleaned = val.toString().replace(/[^0-9.]/g, '');
  return parseFloat(cleaned) || 0;
}

function calculateKwhRate(billAmountStr, kwhUsedStr) {
  const billAmount = parseNumericValue(billAmountStr);
  const kwhUsed = parseNumericValue(kwhUsedStr);

  if (kwhUsed === 0) return '₱0.00';
  
  const rate = billAmount / kwhUsed;
  return `₱${rate.toFixed(2)}`;
}

export async function getNewBillToSheet(accountNumber) {
  try {
    let newRecordFound = 0;
    const cleanAccountNumber = String(accountNumber).trim().padStart(11, '0');

    const record = await getOnlineBillRecord(cleanAccountNumber);
    if (!record || record.length === 0) return false;

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'online_bill_records!A:Z',
    });

    const data = response.data.values || [];
    if (data.length === 0) return false;

    const headerRow = data[0];
    const headerMap = createHeaderMap(headerRow);

    const keyCol = getColInfo(headerMap, 'key', 'recordkey');
    const monthCol = getColInfo(headerMap, 'month_year', 'billingmonth', 'monthyear', 'month');
    const accountNumberCol = getColInfo(headerMap, 'account_number', 'accountnumber', 'account');
    const presentCol = getColInfo(headerMap, 'present_reading', 'presentreading', 'present');
    const prevCol = getColInfo(headerMap, 'previous_reading', 'previousreading', 'prevreading', 'previous');
    const kwhCol = getColInfo(headerMap, 'kwh_used', 'kwhused', 'kwh');
    const dueDateCol = getColInfo(headerMap, 'due_date', 'duedate', 'due');
    const amountCol = getColInfo(headerMap, 'bill_amount', 'billamount', 'amount');
    const statusCol = getColInfo(headerMap, 'status', 'billstatus');
    const notifStatusCol = getColInfo(headerMap, 'notification', 'notificationstatus', 'notifstatus');
    const lastPaymentCol = getColInfo(headerMap, 'last_payment_status', 'lastpaymentstatus', 'paymentstatus');
    const kwhRateCol = getColInfo(headerMap, 'kwh_rate', 'kwhrate', 'ratekwh', 'rate');
    const reminderStatusCol = getColInfo(headerMap, 'notification_reminder_status', 'reminderstatus', 'notifreminderstatus');

    const allColIndexes = [
      keyCol?.index, monthCol?.index, accountNumberCol?.index, presentCol?.index,
      prevCol?.index, kwhCol?.index, dueDateCol?.index, amountCol?.index,
      statusCol?.index, notifStatusCol?.index, lastPaymentCol?.index, kwhRateCol?.index,
      reminderStatusCol?.index
    ].filter(idx => idx !== undefined);

    const maxColIndex = Math.max(headerRow.length - 1, ...allColIndexes);

    for (const item of record) {
      const cleanMonthYear = item.monthYear.replace(/\s+/g, '');
      const recordKey = `${cleanMonthYear}_${cleanAccountNumber}`;
      
      // Case-insensitive check on scraped bill status
      const isPaid = (item.status || '').toString().trim().toUpperCase() === 'PAID';

      let rowIndex = -1;
      if (keyCol) {
        for (let i = 1; i < data.length; i++) {
          if (data[i][keyCol.index] === recordKey) {
            rowIndex = i;
            break;
          }
        }
      }

      const rowData = new Array(maxColIndex + 1).fill('');
      if (keyCol) rowData[keyCol.index] = recordKey;
      if (monthCol) rowData[monthCol.index] = item.monthYear;
      if (accountNumberCol) rowData[accountNumberCol.index] = cleanAccountNumber;
      if (presentCol) rowData[presentCol.index] = item.presentReading;
      if (prevCol) rowData[prevCol.index] = item.previousReading;
      if (kwhCol) rowData[kwhCol.index] = item.kwhUsed;
      if (dueDateCol) rowData[dueDateCol.index] = item.dueDate;
      if (amountCol) rowData[amountCol.index] = item.billAmount;
      if (statusCol) rowData[statusCol.index] = item.status;
      
      if (kwhRateCol) {
        rowData[kwhRateCol.index] = calculateKwhRate(item.billAmount, item.kwhUsed);
      }

      if (rowIndex !== -1) {
        // RECORD EXISTS: Retain existing initial notification & last_payment_status
        if (notifStatusCol) {
          rowData[notifStatusCol.index] = data[rowIndex][notifStatusCol.index] || '';
        }
        if (lastPaymentCol) {
          rowData[lastPaymentCol.index] = data[rowIndex][lastPaymentCol.index] || '';
        }

        // UPDATE REMINDER STATUS: If online status is now PAID, update reminder status to 'paid'
        if (reminderStatusCol) {
          const existingReminder = data[rowIndex][reminderStatusCol.index];
          if (isPaid) {
            rowData[reminderStatusCol.index] = 'paid';
          } else {
            // Keep current progression state (e.g., '1st_reminder', '2nd_reminder') if still unpaid
            rowData[reminderStatusCol.index] = (existingReminder && existingReminder.trim() !== '') ? existingReminder : 'unpaid';
          }
        }

        const endLetter = colIndexToLetter(maxColIndex);
        await sheets.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range: `online_bill_records!A${rowIndex + 1}:${endLetter}${rowIndex + 1}`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: [rowData] },
        });
      } else {
        // NEW RECORD
        if (notifStatusCol) rowData[notifStatusCol.index] = 'unsent';
        if (lastPaymentCol) rowData[lastPaymentCol.index] = item.status;
        if (reminderStatusCol) rowData[reminderStatusCol.index] = isPaid ? 'paid' : 'unpaid';

        await sheets.spreadsheets.values.append({
          spreadsheetId: SPREADSHEET_ID,
          range: 'online_bill_records!A:Z',
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: [rowData] },
        });
        newRecordFound++;
        console.log(`New bill record added for account ${cleanAccountNumber}: ${item.monthYear}`);
      }
    }

    await sortBillRecord();
    return true;

  } catch (e) {
    console.error('Error processing bill sync to sheet:', e);
  }
}