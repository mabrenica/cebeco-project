import { JSDOM } from 'jsdom';

const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

export async function getOnlineBillRecord(accountNumber) {
  const url = `https://www.cebeco1.online/general-services/bill-inquiry/${accountNumber}`;
  const maxAttempts = 3;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const response = await fetch(url);
      
      if (!response.ok) {
        throw new Error(`HTTP network error status: ${response.status}`);
      }

      let responseText = await response.text();
      
      if (responseText.includes('defer')) {
        responseText = responseText.replace(/defer/g, '');
      }

      const dom = new JSDOM(responseText);
      const document = dom.window.document;

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
        if (index === 0) return;

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

      return billRecord;

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