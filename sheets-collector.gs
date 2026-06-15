// AI Quest — Google Apps Script Score Collector
// ---------------------------------------------------
// 1. Open your Google Sheet
// 2. Extensions → Apps Script → replace everything with this file
// 3. Click Deploy → New deployment → Web app
//    - Execute as: Me
//    - Who has access: Anyone
// 4. Copy the /exec URL into dna-of-ai-quiz.html as SHEETS_URL

function doGet(e) {
  try {
    const p = e.parameter;

    // Health-check ping (no name param = just checking if live)
    if (!p.name) {
      return ContentService.createTextOutput('AI Quest Score Collector is running!');
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Player', 'Correct', 'Total', 'Coins', 'Grade', 'Completed?']);
      sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#e8a634');
    }

    sheet.appendRow([
      new Date().toLocaleString('en-US', { timeZone: 'America/New_York' }),
      p.name      || '(unknown)',
      p.correct   || 0,
      p.total     || 15,
      p.coins     || 0,
      p.grade     || '',
      p.completed || 'No'
    ]);

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
