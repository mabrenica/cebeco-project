import { sheets, SPREADSHEET_ID } from './auth.js';
import { createHeaderMap, getColInfo } from './sheetUtils.js';

export async function sortBillRecord() {
  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const sheetObj = spreadsheet.data.sheets.find(s => s.properties.title === 'online_bill_records');
  
  if (!sheetObj) return;
  const sheetId = sheetObj.properties.sheetId;
  const rowCount = sheetObj.properties.gridProperties.rowCount;

  if (rowCount < 2) return;

  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: 'online_bill_records!1:1',
  });

  const headerRow = response.data.values?.[0] || [];
  const headerMap = createHeaderMap(headerRow);
  const monthCol = getColInfo(headerMap, 'month_year', 'billingmonth', 'monthyear', 'month');

  const sortDimensionIndex = monthCol ? monthCol.index : 1;

  const requests = [{
    sortRange: {
      range: {
        sheetId: sheetId,
        startRowIndex: 1,
        endRowIndex: rowCount,
        startColumnIndex: 0,
        endColumnIndex: headerRow.length || 10
      },
      sortSpecs: [{
        dimensionIndex: sortDimensionIndex,
        sortOrder: 'DESCENDING'
      }]
    }
  }];

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { requests }
  });
}