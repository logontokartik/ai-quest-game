// AI Quest — Google Apps Script Score Collector
// ---------------------------------------------------
// 1. Open your Google Sheet
// 2. Extensions → Apps Script → replace everything with this file
// 3. Deploy → Manage deployments → edit existing → bump version → Deploy
//    (keep same URL — Execute as: Me, Who has access: Anyone)

const QUESTION_LABELS = [
  'Q1 AI stands for?',
  'Q2 What is a prompt?',
  'Q3 Who makes Claude?',
  'Q4 Who makes ChatGPT?',
  'Q5 What is LLM?',
  'Q6 What is hallucination?',
  'Q7 Not an AI assistant?',
  'Q8 What does Gen AI do?',
  'Q9 What is training data?',
  'Q10 GPT stands for?',
  'Q11 What is a skill?',
  'Q12 MCP stands for?',
  'Q13 Problem before MCP?',
  'Q14 When do hooks fire?',
  'Q15 AI primitives truth?'
];

function doGet(e) {
  try {
    const p = e.parameter;

    if (!p.name) {
      return ContentService.createTextOutput('AI Quest Score Collector is running!');
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Write header if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Timestamp', 'Player', 'Correct', 'Total', 'Coins', 'Grade', 'Completed?',
        ...QUESTION_LABELS
      ]);
      sheet.getRange(1, 1, 1, 7 + QUESTION_LABELS.length)
        .setFontWeight('bold')
        .setBackground('#e8a634');
    }

    // Parse per-question answers: "C:B:B,W:A:C,T:-:B,..."
    // Each segment: result(C/W/T) : chosen(A-D/-) : correct(A-D)
    const answerCols = (p.answers || '').split(',').map(seg => {
      if (!seg) return '';
      const [result, chosen, correct] = seg.split(':');
      if (result === 'C') return `✓ (${chosen})`;
      if (result === 'W') return `✗ chose ${chosen}, ans ${correct}`;
      if (result === 'T') return `⏰ ans ${correct}`;
      return '';
    });

    // Pad to full question count if player didn't finish
    while (answerCols.length < QUESTION_LABELS.length) answerCols.push('—');

    sheet.appendRow([
      new Date().toLocaleString('en-US', { timeZone: 'America/New_York' }),
      p.name      || '(unknown)',
      p.correct   || 0,
      p.total     || 15,
      p.coins     || 0,
      p.grade     || '',
      p.completed || 'No',
      ...answerCols
    ]);

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
