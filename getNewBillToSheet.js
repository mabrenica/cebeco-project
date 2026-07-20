import { sheets, SPREADSHEET_ID } from './auth.js';
import { getOnlineBillRecord } from './getOnlineBillRecord.js';
import { sortBillRecord } from './sortBillRecords.js';

export async function getNewBillToSheet(accountNumber) {
  try {
    let newRecordFound = 0;
    const record = await getOnlineBillRecord(accountNumber);
    if (!record || record.length === 0) return false;

    // Fetch all current values from the target sheet
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'OnlineBillRecords!A:J',
    });

    const data = response.data.values || [];
    
    for (const item of record) {
      const key = '_' + item.monthYear.replace(/\s+/g, '');
      const rowData = [key, item.monthYear, item.presentReading, item.previousReading, item.kwhUsed, item.dueDate, item.billAmount, item.status];
      const newData = [...rowData, 'unsent', 'UNPAID'];

      let rowIndex = -1;
      for (let i = 0; i < data.length; i++) {
        if (data[i].includes(key)) {
          rowIndex = i;
          break;
        }
      }

      if (rowIndex !== -1) {
        // Record exists, update columns A through H
        await sheets.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range: `OnlineBillRecords!A${rowIndex + 1}:H${rowIndex + 1}`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: [rowData] },
        });
      } else {
        // Record is new, append row to the end of sheet data boundary
        await sheets.spreadsheets.values.append({
          spreadsheetId: SPREADSHEET_ID,
          range: 'OnlineBillRecords!A:J',
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: [newData] },
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