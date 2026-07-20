import { sheets, SPREADSHEET_ID } from './auth.js';

export async function sortBillRecord() {
  // First, pull sheet metadata to discover the underlying numeric Sheet ID for 'OnlineBillRecords'
  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const sheetObj = spreadsheet.data.sheets.find(s => s.properties.title === 'OnlineBillRecords');
  
  if (!sheetObj) return;
  const sheetId = sheetObj.properties.sheetId;
  const rowCount = sheetObj.properties.gridProperties.rowCount;

  if (rowCount < 2) return;

  const requests = [{
    sortRange: {
      range: {
        sheetId: sheetId,
        startRowIndex: 1, // Row 2 onwards (0-indexed base)
        endRowIndex: rowCount,
        startColumnIndex: 0,
        endColumnIndex: 10
      },
      sortSpecs: [{
        dimensionIndex: 1, // Column B (0-indexed base)
        sortOrder: 'DESCENDING'
      }]
    }
  }];

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { requests }
  });
}