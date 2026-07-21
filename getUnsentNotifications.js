import { sheets, SPREADSHEET_ID } from './auth.js';
import { createHeaderMap, getColInfo } from './sheetUtils.js';

export async function getUnsentNotifications() {
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

  if (!notifStatusCol || !keyCol) return [];

  const unsentNotificationKeys = [];

  data.forEach((item, index) => {
    if (index === 0) return;
    if (item[notifStatusCol.index] === 'unsent') {
      console.log('Unsent notification key identified: ' + item[keyCol.index]);
      unsentNotificationKeys.push(item[keyCol.index]);
    }
  });

  return unsentNotificationKeys;
}