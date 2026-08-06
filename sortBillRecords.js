import { sheets, SPREADSHEET_ID } from './auth.js';
import { createHeaderMap, getColInfo } from './sheetUtils.js';

export async function sortBillRecord() {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'online_bill_records!A:Z',
    });

    const data = response.data.values || [];
    if (data.length <= 2) return; // No sorting needed for 0 or 1 data row

    const headerRow = data[0];
    const dataRows = data.slice(1);

    const headerMap = createHeaderMap(headerRow);
    const monthCol = getColInfo(headerMap, 'month_year', 'billingmonth', 'monthyear', 'month');
    const dueDateCol = getColInfo(headerMap, 'due_date', 'duedate', 'due');

    // Parses string dates (e.g., "September 2025" or "October 25, 2025") into Epoch timestamps
    const parseRowTimestamp = (row) => {
      let d = null;
      if (monthCol && row[monthCol.index]) {
        d = new Date(row[monthCol.index]);
      }
      if ((!d || isNaN(d.getTime())) && dueDateCol && row[dueDateCol.index]) {
        d = new Date(row[dueDateCol.index]);
      }
      return (d && !isNaN(d.getTime())) ? d.getTime() : 0;
    };

    // Sort descending (latest date at the top)
    dataRows.sort((a, b) => parseRowTimestamp(b) - parseRowTimestamp(a));

    // Overwrite sheet rows starting at row 2 with sorted data
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: 'online_bill_records!A2',
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: dataRows },
    });

    //console.log('online_bill_records sheet successfully sorted (latest on top).');
  } catch (error) {
    console.error('Error sorting online_bill_records sheet:', error);
  }
}