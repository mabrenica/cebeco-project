import { sheets, SPREADSHEET_ID } from './auth.js';

export async function updateEmailTemplateFromRecord(recordKey) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'OnlineBillRecords!A:G',
  });

  const onlineBillData = response.data.values || [];

  for (let index = 0; index < onlineBillData.length; index++) {
    if (index === 0) continue;
    const item = onlineBillData[index];

    if (item[0] === recordKey) {
      // Pack assignments into a clean transactional batchUpdate
      const dataPayload = [
        { range: 'Bill Template!C1', values: [[item[6]]] }, // billAmount
        { range: 'Bill Template!F1', values: [[item[1]]] }, // billingMonth
        { range: 'Bill Template!C4', values: [[item[5]]] }, // dueDate
        { range: 'Bill Template!F4', values: [[item[3]]] }, // previousReading
        { range: 'Bill Template!F5', values: [[item[2]]] }, // presentReading
        { range: 'Bill Template!F6', values: [[item[4]]] }  // kwhUsed
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