import { sheets, SPREADSHEET_ID } from './auth.js';
import { getOnlineBillRecord } from './getOnlineBillRecords.js';
import { sortBillRecord } from './sortBillRecords.js';
import { createHeaderMap, getColInfo, colIndexToLetter } from './sheetUtils.js';

export async function getNewBillToSheet(accountNumber) {
  try {
    let newRecordFound = 0;
    const record = await getOnlineBillRecord(accountNumber);
    if (!record || record.length === 0) return false;

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'OnlineBillRecords!A:Z',
    });

    const data = response.data.values || [];
    if (data.length === 0) return false;

    const headerRow = data[0];
    const headerMap = createHeaderMap(headerRow);

    const keyCol = getColInfo(headerMap, 'key', 'recordkey');
    const monthCol = getColInfo(headerMap, 'billingmonth', 'monthyear', 'month');
    const presentCol = getColInfo(headerMap, 'presentreading', 'present');
    const prevCol = getColInfo(headerMap, 'previousreading', 'prevreading', 'previous');
    const kwhCol = getColInfo(headerMap, 'kwhused', 'kwh');
    const dueDateCol = getColInfo(headerMap, 'duedate', 'due');
    const amountCol = getColInfo(headerMap, 'billamount', 'amount');
    const statusCol = getColInfo(headerMap, 'status', 'billstatus');
    // Added 'notification' alias here:
    const notifStatusCol = getColInfo(headerMap, 'notification', 'notificationstatus', 'notifstatus');
    const lastPaymentCol = getColInfo(headerMap, 'lastpaymentstatus', 'paymentstatus');

    const allColIndexes = [
      keyCol?.index, monthCol?.index, presentCol?.index, prevCol?.index,
      kwhCol?.index, dueDateCol?.index, amountCol?.index, statusCol?.index,
      notifStatusCol?.index, lastPaymentCol?.index
    ].filter(idx => idx !== undefined);

    const maxColIndex = Math.max(headerRow.length - 1, ...allColIndexes);

    for (const item of record) {
      const recordKey = '_' + item.monthYear.replace(/\s+/g, '');

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
      if (presentCol) rowData[presentCol.index] = item.presentReading;
      if (prevCol) rowData[prevCol.index] = item.previousReading;
      if (kwhCol) rowData[kwhCol.index] = item.kwhUsed;
      if (dueDateCol) rowData[dueDateCol.index] = item.dueDate;
      if (amountCol) rowData[amountCol.index] = item.billAmount;
      if (statusCol) rowData[statusCol.index] = item.status;

      if (rowIndex !== -1) {
        if (notifStatusCol) {
          const existingNotif = data[rowIndex][notifStatusCol.index];
          rowData[notifStatusCol.index] = (existingNotif && existingNotif.trim() !== '') ? existingNotif : 'unsent';
        }
        if (lastPaymentCol) {
          const existingPayment = data[rowIndex][lastPaymentCol.index];
          rowData[lastPaymentCol.index] = (existingPayment && existingPayment.trim() !== '') ? existingPayment : 'UNPAID';
        }

        const endLetter = colIndexToLetter(maxColIndex);
        await sheets.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range: `OnlineBillRecords!A${rowIndex + 1}:${endLetter}${rowIndex + 1}`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: [rowData] },
        });
      } else {
        if (notifStatusCol) rowData[notifStatusCol.index] = 'unsent';
        if (lastPaymentCol) rowData[lastPaymentCol.index] = 'UNPAID';

        await sheets.spreadsheets.values.append({
          spreadsheetId: SPREADSHEET_ID,
          range: 'OnlineBillRecords!A:Z',
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: [rowData] },
        });
        newRecordFound++;
        console.log("New bill record added: " + item.monthYear);
      }
    }

    if (newRecordFound > 0) {
      await sortBillRecord();
      return true;
    } else {
      console.log('No new bill found');
      return false;
    }
  } catch (e) {
    console.error('Error processing bill sync to sheet:', e);
  }
}