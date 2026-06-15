// AI Quest — Google Apps Script Score Collector
// ---------------------------------------------------
// 1. Open your Google Sheet
// 2. Extensions → Apps Script → replace everything with this file
// 3. Deploy → Manage deployments → edit existing → bump version → Deploy

const QUESTIONS = [
  {
    text: 'What does "AI" stand for?',
    opts: ['Automated Internet', 'Artificial Intelligence', 'Advanced Interface', 'Algorithmic Integration'],
    answer: 1
  },
  {
    text: 'What do you call the text you send to ask an AI a question?',
    opts: ['A query string', 'A prompt', 'A command line', 'An input file'],
    answer: 1
  },
  {
    text: 'Which company makes Claude?',
    opts: ['OpenAI', 'Google', 'Anthropic', 'Meta'],
    answer: 2
  },
  {
    text: 'ChatGPT is made by which company?',
    opts: ['Microsoft', 'OpenAI', 'Google', 'Apple'],
    answer: 1
  },
  {
    text: 'What does "LLM" stand for?',
    opts: ['Large Language Model', 'Local Learning Machine', 'Long Logic Module', 'Linked List Method'],
    answer: 0
  },
  {
    text: 'When an AI confidently gives you wrong or made-up info, it\'s called a...',
    opts: ['Glitch', 'Bug', 'Hallucination', 'Timeout'],
    answer: 2
  },
  {
    text: 'Which of these is NOT an AI assistant?',
    opts: ['Siri', 'Alexa', 'Firefox', 'Gemini'],
    answer: 2
  },
  {
    text: 'What is "Generative AI" best at?',
    opts: ['Running spreadsheets faster', 'Storing data in databases', 'Creating new content like text, images & code', 'Blocking spam emails'],
    answer: 2
  },
  {
    text: 'What does an AI model need to "learn" how to do things?',
    opts: ['A keyboard', 'Training data', 'An internet connection', 'A mobile app'],
    answer: 1
  },
  {
    text: 'GPT stands for...',
    opts: ['General Purpose Technology', 'Graphics Processing Thread', 'Generative Pre-trained Transformer', 'Global Processing Tool'],
    answer: 2
  },
  {
    text: 'What is a "skill" in an AI platform like Claude Code?',
    opts: ['A hardware accelerator chip', 'A bundle of instructions for one job, loaded only when needed', 'A database of past conversations', 'A fine-tuned version of the model'],
    answer: 1
  },
  {
    text: 'MCP stands for...',
    opts: ['Multi-Cloud Protocol', 'Model Context Protocol', 'Machine Communication Protocol', 'Managed Compute Platform'],
    answer: 1
  },
  {
    text: 'Before MCP, connecting AI platforms to external tools meant...',
    opts: ['One universal adapter for everything', 'Custom glue code for every platform+tool pair', 'Waiting for a vendor to build it', 'Retraining the model each time'],
    answer: 1
  },
  {
    text: '"Hooks" in an AI platform fire...',
    opts: ['Only when the AI model decides to call them', 'Only when the user types a trigger keyword', 'Automatically at lifecycle events — always, guaranteed', 'Only when the session runs out of tokens'],
    answer: 2
  },
  {
    text: 'Which statement about AI platform primitives is TRUE?',
    opts: ['Each AI brand has completely unique building blocks', 'Only Claude supports MCP', 'Skills, Tools, MCP & Hooks appear across all major platforms — just with different names', 'Hooks replace the need for Skills'],
    answer: 2
  }
];

function doGet(e) {
  try {
    const p = e.parameter;

    if (!p.name) {
      return ContentService.createTextOutput('AI Quest Score Collector is running!');
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Build header row if sheet is empty
    if (sheet.getLastRow() === 0) {
      const qHeaders = QUESTIONS.map((q, i) => `Q${i+1}: ${q.text.substring(0, 50)}${q.text.length > 50 ? '…' : ''}`);
      sheet.appendRow(['Timestamp', 'Player', 'Correct', 'Total', 'Coins', 'Grade', 'Completed?', ...qHeaders]);
      sheet.getRange(1, 1, 1, 7 + QUESTIONS.length).setFontWeight('bold').setBackground('#e8a634');
      sheet.setFrozenRows(1);
    }

    // Parse answers: each segment is "result:chosen_letter:correct_letter"
    // e.g. "C:B:B" = correct, chose B / "W:A:C" = wrong chose A correct C / "T:-:B" = timeout correct B
    const letterToIndex = { A: 0, B: 1, C: 2, D: 3 };
    const answerCols = (p.answers || '').split(',').map((seg, i) => {
      if (!seg || i >= QUESTIONS.length) return '';
      const [result, chosenLetter, correctLetter] = seg.split(':');
      const q = QUESTIONS[i];
      const correctText = q.opts[letterToIndex[correctLetter]] || correctLetter;
      if (result === 'C') {
        return `✓ ${correctText}`;
      }
      if (result === 'W') {
        const chosenText = q.opts[letterToIndex[chosenLetter]] || chosenLetter;
        return `✗ "${chosenText}" → correct: "${correctText}"`;
      }
      if (result === 'T') {
        return `⏰ timed out → correct: "${correctText}"`;
      }
      return '';
    });

    // Pad unanswered questions (game over mid-way)
    while (answerCols.length < QUESTIONS.length) answerCols.push('—');

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
