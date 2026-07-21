import { sheets, SPREADSHEET_ID } from './auth.js';

export async function getRecipientEmails() {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'Set Up!B2:B15',
  });

  const values = response.data.values || [];
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