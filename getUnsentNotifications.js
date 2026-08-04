import { sheets, SPREADSHEET_ID } from './auth.js';
import { createHeaderMap, getColInfo } from './sheetUtils.js';

export async function getAccountBillData(targetAccountNumber) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'online_bill_records!A:Z',
  });

  const data = response.data.values || [];
  if (data.length < 2) return { primaryBill: null, historicalBills: [], unsentKeys: [] };

  const headerMap = createHeaderMap(data[0]);
  const keyCol = getColInfo(headerMap, 'key', 'recordkey');
  const monthCol = getColInfo(headerMap, 'month_year', 'billingmonth', 'monthyear', 'month');
  const accCol = getColInfo(headerMap, 'account_number', 'accountnumber', 'account');
  const presentCol = getColInfo(headerMap, 'present_reading', 'presentreading', 'present');
  const prevCol = getColInfo(headerMap, 'previous_reading', 'previousreading', 'prevreading', 'previous');
  const kwhCol = getColInfo(headerMap, 'kwh_used', 'kwhused', 'kwh');
  const dueDateCol = getColInfo(headerMap, 'due_date', 'duedate', 'due');
  const amountCol = getColInfo(headerMap, 'bill_amount', 'billamount', 'amount');
  const notifCol = getColInfo(headerMap, 'notification', 'notificationstatus', 'notifstatus');
  const kwhRateCol = getColInfo(headerMap, 'kwh_rate', 'kwhrate', 'ratekwh', 'rate');

  const cleanTargetAcc = String(targetAccountNumber).trim().padStart(11, '0');
  
  const accountRecords = [];
  const unsentKeys = [];

  data.forEach((item, index) => {
    if (index === 0) return;

    const recordKey = keyCol ? item[keyCol.index] : item[0];
    let rowAcc = accCol ? String(item[accCol.index] || '').trim().padStart(11, '0') : '';
    
    // Fallback: extract account number from key string if account_number column cell is empty
    if ((!rowAcc || rowAcc === '00000000000') && recordKey) {
      const parts = recordKey.split('_');
      if (parts.length > 1) {
        rowAcc = String(parts[parts.length - 1]).trim().padStart(11, '0');
      }
    }

    if (rowAcc === cleanTargetAcc) {
      const rawNotif = notifCol ? (item[notifCol.index] || '') : '';
      const isUnsent = rawNotif.toString().trim().toLowerCase() === 'unsent';

      if (isUnsent && recordKey) {
        unsentKeys.push(recordKey);
      }

      accountRecords.push({
        key: recordKey,
        accountNumber: rowAcc,
        monthYear: monthCol ? item[monthCol.index] : '',
        presentReading: presentCol ? item[presentCol.index] : '',
        prevReading: prevCol ? item[prevCol.index] : '',
        kwhUsed: kwhCol ? item[kwhCol.index] : '',
        dueDate: dueDateCol ? item[dueDateCol.index] : '',
        billAmount: amountCol ? item[amountCol.index] : '',
        kwhRate: kwhRateCol ? item[kwhRateCol.index] : '',
        isUnsent: isUnsent
      });
    }
  });

  if (accountRecords.length === 0) {
    return { primaryBill: null, historicalBills: [], unsentKeys: [] };
  }

  const primaryBill = accountRecords.find(r => r.isUnsent) || null;

  if (!primaryBill) {
    return { primaryBill: null, historicalBills: [], unsentKeys: [] };
  }

  const historicalBills = accountRecords.filter(r => r.key !== primaryBill.key);

  return { primaryBill, historicalBills, unsentKeys };
}