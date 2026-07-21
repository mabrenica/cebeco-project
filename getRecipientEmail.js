import { sheets, SPREADSHEET_ID } from './auth.js';
import { createHeaderMap, getColInfo } from './sheetUtils.js';

export async function getRecipientEmails() {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'Set Up!A1:Z50',
  });

  const values = response.data.values || [];
  if (values.length === 0) return [];

  const headerMap = createHeaderMap(values[0]);
  const emailCol = getColInfo(headerMap, 'emails', 'email', 'recipientemails', 'recipients');
  
  // Use matched header column, or default to Column B (index 1)
  const colIdx = emailCol ? emailCol.index : 1;
  // If a header was matched, skip row 1 (index 0); otherwise read from row 1
  const startRow = emailCol ? 1 : 0;

  const emailAddresses = [];

  values.forEach(row => {
    if (row[0] && typeof row[0] === 'string') {
      // Split by comma, trim spaces, and exclude empty entries
      const splitEmails = row[0]
        .split(',')
        .map(email => email.trim())
        .filter(email => email.length > 0);

      emailAddresses.push(...splitEmails);
    }
  });

  return emailAddresses;
}