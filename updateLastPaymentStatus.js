import { sheets, SPREADSHEET_ID } from './auth.js';

export async function updateLastPaymentStatus(recordKey, newStatus) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'OnlineBillRecords!A:J',
  });

  const data = response.data.values || [];

  for (let index = 0; index < data.length; index++) {
    if (data[index][0] === recordKey) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `OnlineBillRecords!J${index + 1}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [[newStatus]] },
      });
      console.log('Last Payment Status updated: ' + recordKey);
      break;
    }
  }
}