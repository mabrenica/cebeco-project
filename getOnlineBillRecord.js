import { JSDOM } from 'jsdom';

// Helper function to wait before retrying
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

export async function getOnlineBillRecord(accountNumber) {
  const url = `https://www.cebeco1.online/general-services/bill-inquiry/${accountNumber}`;
  const maxAttempts = 3; // 1 initial attempt + 2 retries

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const response = await fetch(url);
      
      // Handle HTTP error statuses (like 500 or 503) as request errors
      if (!response.ok) {
        throw new Error(`HTTP network error status: ${response.status}`);
      }

      let responseText = await response.text();
      
      if (responseText.includes('defer')) {
        responseText = responseText.replace(/defer/g, '');
      }

      // Initialize JSDOM to parse HTML string locally
      const dom = new JSDOM(responseText);
      const document = dom.window.document;

      // Mimic your Apps Script DOM-navigation path
      const bodyContainer = document.querySelector('body > div');
      if (!bodyContainer) throw new Error("Could not find root body container.");

      const divs = bodyContainer.querySelectorAll(':scope > div');
      const targetDiv = divs[1]?.querySelector(':scope > div');
      const listItems = targetDiv?.querySelectorAll('ul > li');

      if (!listItems) {
        console.log('No billing elements found on the web page.');
        return [];
      }

      const billRecord = [];

      listItems.forEach((item, index) => {
        if (index === 0) return; // Skip header row element

        const spans = item.querySelectorAll('span');
        if (spans.length < 8) return;

        const listItem = {
          monthYear: spans[1].textContent.trim(),
          presentReading: spans[2].textContent.trim(),
          previousReading: spans[3].textContent.trim(),
          kwhUsed: spans[4].textContent.trim(),
          dueDate: spans[5].textContent.trim(),
          billAmount: spans[6].textContent.trim(),
          status: spans[7].textContent.trim()
        };

        billRecord.push(listItem);
      });

      return billRecord; // Success: exit loop and return data

    } catch (e) {
      console.warn(`Attempt ${attempt} failed fetching online bill record:`, e.message);
      
      if (attempt < maxAttempts) {
        console.log(`Waiting 1 second before retrying...`);
        await delay(1000);
      } else {
        console.error('All retry attempts exhausted.');
        return [];
      }
    }
  }
}