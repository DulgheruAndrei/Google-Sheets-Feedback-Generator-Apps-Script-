* Adds a custom menu when the spreadsheet opens.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Feedback Generator')
    .addItem('Open form…', 'showSidebar')
    .addToUi();
}

/**
 * Shows the sidebar UI defined in Sidebar.html
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Feedback Generator')
    .setWidth(380);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Main entry: builds the feedback text and appends a row to the sheet.
 * @param {Object} form - values posted from the sidebar
 */
function generateFeedback(form) {
  // --- Configuration ---
  const SHEET_NAME = form.targetSheet || 'Repository Sheet'; // default sheet name

  // Resolve sheet
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getActiveSheet();

  // Build the feedback text using a simple template
  const template = buildTemplate(form);

  // Compose values
  const timestamp = new Date();
  const userId = form.userId || 'Unknown';
  const comment = template;
  const status = form.status || 'Draft';

  // Append row: [Date and Time, User ID, Feedback Comment, Status]
  sheet.appendRow([timestamp, userId, comment, status]);

  // Optional: format first row as header if empty
  ensureHeader(sheet);

  return { ok: true, row: sheet.getLastRow() };
}

/**
 * Creates the feedback text from inputs. Adjust as you like.
 */
function buildTemplate({ userId, brand, bonusType, bonusValue, link, rep, extra }) {
  const parts = [];
  parts.push(`Greetings, ${userId || 'player'}!`);
  if (brand) parts.push(`According to the ${brand} representative${rep ? ' ' + rep : ''},`);
  if (bonusType || bonusValue) {
    parts.push(`the ${bonusValue ? bonusValue + ' ' : ''}${bonusType || 'bonus'} remains available.`);
  }
  parts.push('Please double‑check the bonus terms and conditions before claiming.');
  if (link) parts.push(`Use our verified link: ${link}`);
  if (extra) parts.push(extra);
  return parts.join(' ');
}

/**
 * Ensures the first row has headers matching the expected columns.
 */
function ensureHeader(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow === 1) {
    const firstRow = sheet.getRange(1, 1, 1, 4).getValues()[0];
    const empty = firstRow.every(v => v === '' || v == null);
    if (empty) {
      sheet.getRange(1, 1, 1, 4).setValues([
        ['Date and Time', 'User ID', 'Feedback Comment', 'Status']
      ]);
      sheet.setFrozenRows(1);
    }
  }
}
