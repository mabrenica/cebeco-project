import { google } from 'googleapis';
import fs from 'fs';

// Load the downloaded service account key or OAuth credentials file
const CREDENTIALS_PATH = './credentials.json'; 

const auth = new google.auth.GoogleAuth({
  keyFile: CREDENTIALS_PATH,
  // We need full drive/sheets access to read, edit, and grab cell formatting
  scopes: [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.readonly'
  ],
});

export const sheets = google.sheets({ version: 'v4', auth });
// Replace this string with your real Google Sheet ID found in your browser URL
export const SPREADSHEET_ID = '1LCw13UoWlSWfHlz53p1EYm49orvAhMajIB_d83TIWTM';