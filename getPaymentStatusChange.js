import { sheets, SPREADSHEET_ID } from './auth.js';

export async function getPaymentStatusChange() {
  // Read both values and raw formatting rules to ensure dates look normal
  const responseValues = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'OnlineBillRecords!A:J',
  });
  

    const responseMeta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    ranges: ['OnlineBillRecords!A:J'], // Corrected parameter (pluralized array)
    includeGridData: true
    });

  const data = responseValues.data.values || [];
  const metaRows = responseMeta.data.sheets[0].data[0].rowData || [];
  const updatedPaymentKeys = [];

  data.forEach((item, index) => {
    if (index === 0) return;
    if (item[7] !== item[9]) {
      console.log('Payment Status Change discovered: ' + item[0]);
      
      // Grab formatted display values for the date directly out of cell metadata strings
      const formattedDate = metaRows[index]?.values?.[1]?.formattedValue || item[1];
      
      updatedPaymentKeys.push({
        key: item[0],
        newStatus: item[7],
        monthYear: formattedDate,
        amount: item[6]
      });
    }
  });
  return updatedPaymentKeys;
}