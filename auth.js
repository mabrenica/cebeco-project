import { google } from 'googleapis';
import fs from 'fs';
import path from 'path';

const CREDENTIALS_PATH = path.resolve('./credentials.json');

if (!fs.existsSync(CREDENTIALS_PATH)) {
  throw new Error('CRITICAL ERROR: credentials.json file missing. Please ensure it exists locally or via GitHub Secrets.');
}

const credentialsData = JSON.parse(fs.readFileSync(CREDENTIALS_PATH, 'utf8'));

const auth = new google.auth.GoogleAuth({
  credentials: credentialsData,
  scopes: [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.readonly'
  ],
});

export const sheets = google.sheets({ version: 'v4', auth });
export const SPREADSHEET_ID = credentialsData.spreadsheet_id;
export const GMAIL_USER = credentialsData.gmail_user;
export const GMAIL_APP_PASS = credentialsData.gmail_app_pass;