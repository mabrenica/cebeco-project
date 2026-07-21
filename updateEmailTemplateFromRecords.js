import { sheets, SPREADSHEET_ID } from './auth.js';
import { createHeaderMap, getColInfo } from './sheetUtils.js';

// Helper function to clean currency strings (removes ₱, commas, spaces) and parse to float
function parseNumericValue(val) {
  if (val === null || val === undefined) return 0;
  const cleaned = val.toString().replace(/[^0-9.]/g, '');
  return parseFloat(cleaned) || 0;
}

// Calculates Rate per kWh (Bill Amount / kWh Used) formatted with ₱ and 2 decimal places
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
  const kwhRateCol = getColInfo(headerMap, 'kwhrate', 'ratekwh', 'rate');

  if (!keyCol) return;

  for (let index = 1; index < onlineBillData.length; index++) {
    const item = onlineBillData[index];

    if (item[keyCol.index] === recordKey) {
      const amountVal = amountCol ? item[amountCol.index] : '';
      const kwhVal = kwhCol ? item[kwhCol.index] : '';

      // Read stored rate from sheet, or calculate on the fly if missing/empty
      let rateVal = kwhRateCol ? item[kwhRateCol.index] : '';
      if ((!rateVal || rateVal.trim() === '') && amountVal && kwhVal) {
        rateVal = calculateKwhRate(amountVal, kwhVal);
      }

      const dataPayload = [
        { range: 'Bill Template!C1', values: [[amountVal]] },
        { range: 'Bill Template!F1', values: [[monthCol ? item[monthCol.index] : '']] },
        { range: 'Bill Template!C4', values: [[dueDateCol ? item[dueDateCol.index] : '']] },
        { range: 'Bill Template!C5', values: [[rateVal]] }, // Populates KWH rate in C5
        { range: 'Bill Template!F4', values: [[prevCol ? item[prevCol.index] : '']] },
        { range: 'Bill Template!F5', values: [[presentCol ? item[presentCol.index] : '']] },
        { range: 'Bill Template!F6', values: [[kwhVal]] }
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