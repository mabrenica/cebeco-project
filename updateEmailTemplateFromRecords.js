import { sheets, SPREADSHEET_ID } from './auth.js';
import { createHeaderMap, getColInfo } from './sheetUtils.js';

function parseNumericValue(val) {
  if (val === null || val === undefined) return 0;
  const cleaned = val.toString().replace(/[^0-9.]/g, '');
  return parseFloat(cleaned) || 0;
}

function calculateKwhRate(billAmountStr, kwhUsedStr) {
  const billAmount = parseNumericValue(billAmountStr);
  const kwhUsed = parseNumericValue(kwhUsedStr);

  if (kwhUsed === 0) return '₱0.00';
  
  const rate = billAmount / kwhUsed;
  return `₱${rate.toFixed(2)}`;
}

export async function updateEmailTemplateFromRecord(recordKey) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'online_bill_records!A:Z',
  });

  const onlineBillData = response.data.values || [];
  if (onlineBillData.length < 2) return null;

  const headerMap = createHeaderMap(onlineBillData[0]);
  const keyCol = getColInfo(headerMap, 'key', 'recordkey');
  const monthCol = getColInfo(headerMap, 'month_year', 'billingmonth', 'monthyear', 'month');
  const accountNumberCol = getColInfo(headerMap, 'account_number', 'accountnumber', 'account');
  const presentCol = getColInfo(headerMap, 'present_reading', 'presentreading', 'present');
  const prevCol = getColInfo(headerMap, 'previous_reading', 'previousreading', 'prevreading', 'previous');
  const kwhCol = getColInfo(headerMap, 'kwh_used', 'kwhused', 'kwh');
  const dueDateCol = getColInfo(headerMap, 'due_date', 'duedate', 'due');
  const amountCol = getColInfo(headerMap, 'bill_amount', 'billamount', 'amount');
  const kwhRateCol = getColInfo(headerMap, 'kwh_rate', 'kwhrate', 'ratekwh', 'rate');

  if (!keyCol) return null;

  for (let index = 1; index < onlineBillData.length; index++) {
    const item = onlineBillData[index];

    if (item[keyCol.index] === recordKey) {
      const amountVal = amountCol ? item[amountCol.index] : '';
      const kwhVal = kwhCol ? item[kwhCol.index] : '';
      const accountNumberVal = accountNumberCol ? item[accountNumberCol.index] : '';

      let rateVal = kwhRateCol ? item[kwhRateCol.index] : '';
      if ((!rateVal || rateVal.trim() === '') && amountVal && kwhVal) {
        rateVal = calculateKwhRate(amountVal, kwhVal);
      }

      const dataPayload = [
        { range: 'bill_template!C1', values: [[amountVal]] },
        { range: 'bill_template!F1', values: [[monthCol ? item[monthCol.index] : '']] },
        { range: 'bill_template!C4', values: [[dueDateCol ? item[dueDateCol.index] : '']] },
        { range: 'bill_template!C5', values: [[rateVal]] },
        { range: 'bill_template!F4', values: [[prevCol ? item[prevCol.index] : '']] },
        { range: 'bill_template!F5', values: [[presentCol ? item[presentCol.index] : '']] },
        { range: 'bill_template!F6', values: [[kwhVal]] }
      ];

      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          valueInputOption: 'USER_ENTERED',
          data: dataPayload
        }
      });

      return accountNumberVal;
    }
  }

  return null;
}