import { sheets, SPREADSHEET_ID } from './auth.js';

export async function updateNotificationStatus(recordKey) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'OnlineBillRecords!A:I',
  });

  const data = response.data.values || [];

  for (let index = 0; index < data.length; index++) {
    if (data[index][0] === recordKey) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `OnlineBillRecords!I${index + 1}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [['sent']] },
      });
      console.log('Notification Status updated: ' + recordKey);
      break;
    }
  }
}