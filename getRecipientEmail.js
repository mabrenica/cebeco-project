import { sheets, SPREADSHEET_ID } from './auth.js';

export async function getRecipientEmails() {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'Set Up!B2:B15',
  });

  const values = response.data.values || [];
  const emailAddresses = [];

  values.forEach(row => {
    if (row[0] && row[0].trim() !== '') {
      emailAddresses.push(row[0].trim());
    }
  });
  return emailAddresses;
}