import { sheets, SPREADSHEET_ID } from './auth.js';
import { createHeaderMap, getColInfo } from './sheetUtils.js';

export async function getAccountConfigurations() {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'set_up!A1:Z50',
  });

  const values = response.data.values || [];
  if (values.length === 0) return [];

  const headerMap = createHeaderMap(values[0]);
  const accNumCol = getColInfo(headerMap, 'account_number', 'accountnumber', 'account');
  const accNameCol = getColInfo(headerMap, 'account_name', 'accountname', 'name');
  const emailCol = getColInfo(headerMap, 'email_recipients', 'emailrecipients', 'emails', 'email', 'recipients');

  const configs = [];

  for (let i = 1; i < values.length; i++) {
    const rawAccNum = accNumCol ? values[i]?.[accNumCol.index] : values[i]?.[0];
    const rawAccName = accNameCol ? values[i]?.[accNameCol.index] : 'Valued Customer';
    const rawEmails = emailCol ? values[i]?.[emailCol.index] : values[i]?.[1];

    if (rawAccNum && String(rawAccNum).trim() !== '') {
      const formattedAccNum = String(rawAccNum).trim().padStart(11, '0');
      const emailAddresses = (rawEmails && typeof rawEmails === 'string')
        ? rawEmails.split(',').map(e => e.trim()).filter(e => e.length > 0)
        : [];

      configs.push({
        accountNumber: formattedAccNum,
        accountName: String(rawAccName).trim(),
        emailAddresses: emailAddresses
      });
    }
  }

  return configs;
}