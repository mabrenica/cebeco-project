import { sheets, SPREADSHEET_ID } from './auth.js';
import { createHeaderMap, getColInfo } from './sheetUtils.js';

export async function updateNotificationStatus(recordKey) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'OnlineBillRecords!A:Z',
  });

  const data = response.data.values || [];
  if (data.length < 2) return [];

  const headerMap = createHeaderMap(data[0]);
  const keyCol = getColInfo(headerMap, 'key', 'recordkey');
  // Added 'notification' alias here:
  const notifStatusCol = getColInfo(headerMap, 'notification', 'notificationstatus', 'notifstatus');

  if (!keyCol || !notifStatusCol) return;

  for (let index = 1; index < data.length; index++) {
    if (data[index][keyCol.index] === recordKey) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `OnlineBillRecords!${notifStatusCol.letter}${index + 1}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [['sent']] },
      });
      console.log('Notification Status updated to sent: ' + recordKey);
      break;
    }
  }
}