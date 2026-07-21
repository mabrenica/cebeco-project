import nodemailer from 'nodemailer';
import { sheets, SPREADSHEET_ID, GMAIL_USER, GMAIL_APP_PASS } from './auth.js'; 

const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: GMAIL_USER,
    pass: GMAIL_APP_PASS
  }
});

function convertA1ToIndexes(rangeStr) {
  const cleanRange = rangeStr.replace(/[^A-Z0-9:]/gi, '');
  const parts = cleanRange.split(':');
  
  const parsePart = (p) => {
    const colMatch = p.match(/[A-Z]+/i);
    const rowMatch = p.match(/[0-9]+/);
    let col = 0;
    if (colMatch) {
      const letters = colMatch[0].toUpperCase();
      for (let i = 0; i < letters.length; i++) {
        col = col * 26 + (letters.charCodeAt(i) - 64);
      }
    }
    const row = rowMatch ? parseInt(rowMatch[0], 10) : null;
    return { col: col - 1, row: row ? row - 1 : null };
  };

  const start = parsePart(parts[0]);
  const end = parts[1] ? parsePart(parts[1]) : start;
  return { start, end };
}

function isInAnyRange(row, col, ranges) {
  return ranges.some(rangeString => {
    if (rangeString.endsWith(':')) rangeString += 'Z';
    const { start, end } = convertA1ToIndexes(rangeString);
    const rowMatch = start.row === null || (row >= start.row && (end.row === null || row <= end.row));
    const colMatch = rowMatch && (col >= start.col && col <= end.col);
    return colMatch;
  });
}

export async function sendBillToEmail(accountNumber, emailAddresses) {
  const currencyRanges = ['C1', 'F11:F14', 'F17:F22', 'F26:F30', 'F33:F38', 'F41:F43', 'F46:F48'];
  const decimalRange = 'F4:F6';

  const sheetDataResponse = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    ranges: ['Bill Template'], 
    includeGridData: true
  });

  const rowData = sheetDataResponse.data.sheets[0].data[0].rowData || [];
  let monthFormatted = "Unknown Month";
  
  let htmlTable = '<table cellpadding="6" cellspacing="0" style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 13px; color: #333333; width: 100%; max-width: 650px;">';

  rowData.forEach((row, i) => {
    htmlTable += '<tr>';
    const values = row.values || [];
    
    values.forEach((cell, j) => {
      let formattedValue = cell.formattedValue !== undefined ? cell.formattedValue : '';
      const rawNumericValue = cell.effectiveValue?.numberValue;

      if (i === 0 && j === 5 && cell.formattedValue) {
         monthFormatted = cell.formattedValue;
      }

      if (typeof rawNumericValue === 'number') {
        if (isInAnyRange(i, j, currencyRanges)) {
          formattedValue = new Intl.NumberFormat('en-PH', { style: 'currency', currency: 'PHP' }).format(rawNumericValue);
        } else if (isInAnyRange(i, j, [decimalRange])) {
          formattedValue = rawNumericValue.toFixed(2);
        }
      }

      const format = cell.effectiveFormat || {};
      const textFormat = format.textFormat || {};
      let styles = [];

      const bg = format.backgroundColor || {};
      const r = Math.round((bg.red || 1) * 255);
      const g = Math.round((bg.green || 1) * 255);
      const b = Math.round((bg.blue || 1) * 255);
      styles.push(`background-color: rgb(${r}, ${g}, ${b})`);

      if (textFormat.foregroundColor) {
        const fg = textFormat.foregroundColor || {};
        const tcR = Math.round((fg.red || 0) * 255);
        const tcG = Math.round((fg.green || 0) * 255);
        const tcB = Math.round((fg.blue || 0) * 255);
        styles.push(`color: rgb(${tcR}, ${tcG}, ${tcB})`);
      }

      if (textFormat.bold) styles.push('font-weight: bold');
      if (textFormat.italic) styles.push('font-style: italic');
      if (textFormat.fontSize) styles.push(`font-size: ${textFormat.fontSize}px`);

      if (format.horizontalAlignment) {
        styles.push(`text-align: ${format.horizontalAlignment.toLowerCase()}`);
      } else {
        styles.push(typeof rawNumericValue === 'number' ? 'text-align: right' : 'text-align: left');
      }

      const mapBorder = (borderObj) => {
        if (!borderObj || borderObj.style === 'NONE') return '1px solid #E0E0E0';
        const bColor = borderObj.color || {};
        const bR = Math.round((bColor.red || 0) * 255);
        const bG = Math.round((bColor.green || 0) * 255);
        const bB = Math.round((bColor.blue || 0) * 255);
        const width = borderObj.width || 1;
        return `${width}px solid rgb(${bR}, ${bG}, ${bB})`;
      };

      styles.push(`border-top: ${mapBorder(format.borders?.top)}`);
      styles.push(`border-bottom: ${mapBorder(format.borders?.bottom)}`);
      styles.push(`border-left: ${mapBorder(format.borders?.left)}`);
      styles.push(`border-right: ${mapBorder(format.borders?.right)}`);

      const finalInlineStyle = styles.join('; ');

      if (i === 0) {
        htmlTable += `<th style="${finalInlineStyle}">${formattedValue}</th>`;
      } else {
        htmlTable += `<td style="${finalInlineStyle}">${formattedValue}</td>`;
      }
    });
    htmlTable += '</tr>';
  });
  htmlTable += '</table>';

  const mailOptions = {
    from: GMAIL_USER,
    to: emailAddresses.join(','),
    subject: "CEBECO1 BILL",
    html: `Good day! Your bill for ${monthFormatted} is ready. Please see below.<br><br>View your account online here: https://www.cebeco1.online/general-services/bill-inquiry/${accountNumber}<br><br>` + htmlTable
  };

  await transporter.sendMail(mailOptions);
  console.log(`${monthFormatted} bill has been successfully sent via Node transport.`);
}

export async function sendPaymentNotificationToEmail(emailAddresses, amount, monthYear) {
  const mailOptions = {
    from: GMAIL_USER,
    to: emailAddresses.join(','),
    subject: 'CEBECO I - Payment Notification',
    html: `Thank you for your payment. <br>We would like to confirm that your payment for ${monthYear} bill with an amount of ${amount} has been successfully posted. Cheers!`
  };

  await transporter.sendMail(mailOptions);
  console.log(`Payment notification for ${monthYear} successfully sent.`);
}