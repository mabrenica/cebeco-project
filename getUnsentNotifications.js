import { sheets, SPREADSHEET_ID } from './auth.js';

export async function getUnsentNotifications() {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'OnlineBillRecords!A:I',
  });

  const data = response.data.values || [];
  const unsentNotificationKeys = [];

  data.forEach((item, index) => {
    if (index === 0) return;
    if (item[8] === 'unsent') {
      console.log('Unsent notification key identified: ' + item[0]);
      unsentNotificationKeys.push(item[0]);
    }
  });

  return unsentNotificationKeys;
}