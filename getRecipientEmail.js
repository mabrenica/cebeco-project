import { sheets, SPREADSHEET_ID } from './auth.js';
import { createHeaderMap, getColInfo } from './sheetUtils.js';

export async function getRecipientEmails() {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'set_up!A1:Z50',
  });

  const values = response.data.values || [];
  if (values.length === 0) return [];

  const headerMap = createHeaderMap(values[0]);
  const emailCol = getColInfo(headerMap, 'email_recipients', 'emailrecipients', 'emails', 'email', 'recipientemails', 'recipients');
  
  const colIdx = emailCol ? emailCol.index : 1;
  const startRow = emailCol ? 1 : 0;

  const emailAddresses = [];

  for (let i = startRow; i < values.length; i++) {
    const cellValue = values[i]?.[colIdx];
    if (cellValue && typeof cellValue === 'string') {
      const splitEmails = cellValue
        .split(',')
        .map(email => email.trim())
        .filter(email => email.length > 0);

      emailAddresses.push(...splitEmails);
    }
  }

  return emailAddresses;
}