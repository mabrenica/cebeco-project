import { sheets, SPREADSHEET_ID } from './auth.js';
import { createHeaderMap, getColInfo } from './sheetUtils.js';

export async function updateEmailTemplateFromRecord(recordKey) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'OnlineBillRecords!A:Z',
  });

  const onlineBillData = response.data.values || [];
  if (onlineBillData.length < 2) return;

  const headerMap = createHeaderMap(onlineBillData[0]);
  const keyCol = getColInfo(headerMap, 'key', 'recordkey');
  const monthCol = getColInfo(headerMap, 'billingmonth', 'monthyear', 'month');
  const presentCol = getColInfo(headerMap, 'presentreading', 'present');
  const prevCol = getColInfo(headerMap, 'previousreading', 'prevreading', 'previous');
  const kwhCol = getColInfo(headerMap, 'kwhused', 'kwh');
  const dueDateCol = getColInfo(headerMap, 'duedate', 'due');
  const amountCol = getColInfo(headerMap, 'billamount', 'amount');

  if (!keyCol) return;

  for (let index = 1; index < onlineBillData.length; index++) {
    const item = onlineBillData[index];

    if (item[keyCol.index] === recordKey) {
      const dataPayload = [
        { range: 'Bill Template!C1', values: [[amountCol ? item[amountCol.index] : '']] },
        { range: 'Bill Template!F1', values: [[monthCol ? item[monthCol.index] : '']] },
        { range: 'Bill Template!C4', values: [[dueDateCol ? item[dueDateCol.index] : '']] },
        { range: 'Bill Template!F4', values: [[prevCol ? item[prevCol.index] : '']] },
        { range: 'Bill Template!F5', values: [[presentCol ? item[presentCol.index] : '']] },
        { range: 'Bill Template!F6', values: [[kwhCol ? item[kwhCol.index] : '']] }
      ];

      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          valueInputOption: 'USER_ENTERED',
          data: dataPayload
        }
      });
      break;
    }
  }
}