// AI Quest — Google Apps Script Score Collector
// ---------------------------------------------------
// 1. Open your "AI Quest — Quiz Results" Google Sheet
// 2. Extensions → Apps Script → paste this entire file
// 3. Click Deploy → New deployment → Web app
//    - Execute as: Me
//    - Who has access: Anyone
// 4. Copy the deployment URL
// 5. Paste it into dna-of-ai-quiz.html where it says PASTE_YOUR_APPS_SCRIPT_URL_HERE

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Write header row if the sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Timestamp', 'Player', 'Correct', 'Total',
        'Coins', 'Grade', 'Completed?'
      ]);
      sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#e8a634');
    }

    sheet.appendRow([
      new Date().toLocaleString('en-US', { timeZone: 'America/New_York' }),
      data.name     || '(unknown)',
      data.correct  || 0,
      data.total    || 15,
      data.coins    || 0,
      data.grade    || '',
      data.completed ? 'Yes' : 'No (Game Over)'
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Lets you test the endpoint is live by visiting the URL in a browser
function doGet() {
  return ContentService.createTextOutput('AI Quest Score Collector is running!');
}
