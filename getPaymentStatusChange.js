import { sheets, SPREADSHEET_ID } from './auth.js';
import { createHeaderMap, getColInfo } from './sheetUtils.js';

export async function getPaymentStatusChange() {
  const responseValues = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'online_bill_records!A:Z',
  });
  
  const responseMeta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    ranges: ['online_bill_records!A:Z'],
    includeGridData: true
  });

  const data = responseValues.data.values || [];
  if (data.length < 2) return [];

  const headerMap = createHeaderMap(data[0]);
  const keyCol = getColInfo(headerMap, 'key', 'recordkey');
  const monthCol = getColInfo(headerMap, 'month_year', 'billingmonth', 'monthyear', 'month');
  const amountCol = getColInfo(headerMap, 'bill_amount', 'billamount', 'amount');
  const statusCol = getColInfo(headerMap, 'status', 'billstatus');
  const lastPaymentCol = getColInfo(headerMap, 'last_payment_status', 'lastpaymentstatus', 'paymentstatus');

  if (!statusCol || !lastPaymentCol) return [];

  const metaRows = responseMeta.data.sheets[0].data[0].rowData || [];
  const updatedPaymentKeys = [];

  data.forEach((item, index) => {
    if (index === 0) return;

    const currentStatus = item[statusCol.index] || '';
    const lastStatus = item[lastPaymentCol.index] || 'UNPAID';

    if (currentStatus && currentStatus !== lastStatus) {
      const recordKey = keyCol ? item[keyCol.index] : item[0];
      console.log(`Payment Status Change discovered for ${recordKey}: [${lastStatus}] -> [${currentStatus}]`);
      
      const formattedDate = (monthCol && metaRows[index]?.values?.[monthCol.index]?.formattedValue) 
        || (monthCol ? item[monthCol.index] : item[1]);
      
      updatedPaymentKeys.push({
        key: recordKey,
        newStatus: currentStatus,
        monthYear: formattedDate,
        amount: amountCol ? item[amountCol.index] : item[6]
      });
    }
  });
  return updatedPaymentKeys;
}