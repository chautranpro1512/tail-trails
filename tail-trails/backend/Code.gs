/**
 * Tail Trails Client System — Google Apps Script
 * --------------------------------------------------------------
 * Automatically:
 *  1. Saves every website sign-up to the "Clients" sheet.
 *  2. Emails YOU instantly (owner, pet, service, tap-to-call link).
 *  3. Emails the CUSTOMER a friendly confirmation.
 *  4. Sets a "Contact by" day (same day before the cutoff hour, else next work day).
 *  5. Every work-day morning emails you a to-do list: new sign-ups to call,
 *     call-backs, meet & greets today, check-ins with new clients, review requests,
 *     plus your daily checklist.
 *  6. Sets follow-up dates automatically when you change a Status.
 *  7. Adds a "Tail Trails" menu for follow-up emails you approve first.
 *
 * Setup: see README.md. Edit SETTINGS, then run setup() once.
 */

const SETTINGS = {
  OWNER_EMAIL: 'YOUR_EMAIL@gmail.com',     // where alerts and the daily list go
  OWNER_NAME: 'Tail Trails team',          // how you sign customer emails
  BUSINESS_NAME: 'Tail Trails',
  BUSINESS_PHONE: '',                      // e.g. '306-555-0123'
  TIMEZONE: 'America/Regina',
  WORK_DAYS: [1, 2, 3, 4, 5, 6, 7],        // pet care runs every day; 1=Mon … 7=Sun
  SAME_DAY_CUTOFF_HOUR: 15,                // sign-ups before 3 pm → contact same day
  DIGEST_HOUR: 7,
  NO_ANSWER_RETRY_DAYS: 1,
  FIRST_WEEK_CHECKIN_DAYS: 7,              // "Booked" → check in after 7 days
  REVIEW_REQUEST_DAYS: 21,                 // "Regular client" → ask for a review after 3 weeks
  SEND_CUSTOMER_CONFIRMATION: true,
  DAILY_CHECKLIST: [
    'Call every new sign-up on today\'s list, write what you learn in "Your notes", then update the Status',
    'Confirm today\'s meet & greets by text',
    'Check today\'s walks and visits are assigned, keys ready',
    'Check the weather: under −20 °C, shorten walks and pack coats and booties',
    'Send photos and a report after every visit',
  ],
};

const SHEET_NAME = 'Clients';
const HEADERS = ['ID', 'Received', 'Owner name', 'Phone', 'Email', 'Pet type', 'Interested in',
  'Status', 'Contact by', 'Next follow-up', 'Pet name', 'Your notes', 'Last updated'];
const COL = {}; HEADERS.forEach((h, i) => COL[h] = i + 1);
const STATUSES = ['New', 'No answer', 'Meet & greet booked', 'Booked', 'Regular client', 'Paused', 'Lost'];
const SERVICE_LABELS = { walk30: 'Dog walk, 30 min', walk60: 'Dog walk, 60 min', dropin: 'Drop-in visit',
  overnight: 'Overnight house sitting', taxi: 'Pet taxi', companion: 'Companion time', notsure: 'Not sure yet' };

/* ============================ SETUP ============================ */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(SETTINGS.TIMEZONE);
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME, 0);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#1F4638').setFontColor('#F2F7F3');
  sh.setFrozenRows(1);
  sh.setColumnWidth(COL['Your notes'], 320);
  const n = sh.getMaxRows() - 1;
  sh.getRange(2, COL['Status'], n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build());
  const dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
  [COL['Contact by'], COL['Next follow-up']].forEach(c => sh.getRange(2, c, n, 1).setDataValidation(dateRule).setNumberFormat('ddd mmm d, yyyy'));
  sh.getRange(2, COL['Received'], n, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange(2, COL['Last updated'], n, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange(2, COL['Phone'], n, 1).setNumberFormat('@');

  const rng = sh.getRange(2, 1, n, HEADERS.length), L = columnLetter(COL['Status']);
  const colours = { 'New': '#FFF4DC', 'No answer': '#FFE3DF', 'Meet & greet booked': '#E1F4FF', 'Booked': '#EAE6FF', 'Regular client': '#DFF7EA', 'Paused': '#EEEEEE', 'Lost': '#F5F5F5' };
  sh.setConditionalFormatRules(Object.keys(colours).map(s => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + L + '2="' + s + '"').setBackground(colours[s]).setRanges([rng]).build()));

  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyDigest').timeBased().everyDays(1).atHour(SETTINGS.DIGEST_HOUR).inTimezone(SETTINGS.TIMEZONE).create();
  ScriptApp.newTrigger('handleEdit').forSpreadsheet(ss).onEdit().create();
  try { SpreadsheetApp.getUi().alert('Tail Trails setup complete.\n\nNext: Deploy → New deployment → Web app (README step 3).'); } catch (e) {}
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Tail Trails')
    .addItem('Email selected client: first-week check-in', 'menuCheckIn')
    .addItem('Email selected client: thank-you + review request', 'menuReviewRequest')
    .addSeparator()
    .addItem('Send me today\'s to-do email now', 'dailyDigestForce')
    .addItem('Add a test sign-up', 'testLead')
    .addItem('Run setup again', 'setup')
    .addToUi();
}

/* ======================== WEBSITE FORM ========================= */

function doPost(e) {
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d.website) return json({ ok: true });                          // spam bot
    const name = clean(d.name, 100), phone = clean(d.phone, 40), email = clean(d.email, 120);
    const petType = clean(d.petType, 40);
    if (!name || phone.replace(/\D/g, '').length < 10 || !validEmail(email) || !petType) return json({ ok: false, error: 'missing fields' });

    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    const sh = clientSheet();
    const id = 'TT' + Utilities.formatDate(new Date(), SETTINGS.TIMEZONE, 'yyMMdd') + '-' + Math.random().toString(36).slice(2, 6).toUpperCase();
    const contactBy = contactByDate(new Date());
    const row = new Array(HEADERS.length).fill('');
    const set = (h, v) => row[COL[h] - 1] = v;
    set('ID', id); set('Received', new Date());
    set('Owner name', safe(name)); set('Phone', safe(phone)); set('Email', safe(email));
    set('Pet type', safe(petType));
    set('Interested in', SERVICE_LABELS[d.service] || '');   // only filled if they clicked a service on the site
    set('Status', 'New'); set('Contact by', contactBy); set('Last updated', new Date());
    sh.appendRow(row);
    lock.releaseLock();

    const c = rowToClient(row);
    notifyOwner(c);
    if (SETTINGS.SEND_CUSTOMER_CONFIRMATION) confirmCustomer(c);
    return json({ ok: true, id: id, callByText: friendlyDay(contactBy) });
  } catch (err) {
    console.error(err);
    try { MailApp.sendEmail(SETTINGS.OWNER_EMAIL, 'Tail Trails website form error', String(err && err.stack || err)); } catch (_) {}
    return json({ ok: false, error: 'server error' });
  }
}

function doGet() { return ContentService.createTextOutput('Tail Trails sign-up endpoint is running.'); }

/* ========================== EMAILS ============================= */

function petLine(c) {
  return (c.petName ? c.petName + ' (' + c.petType + ')' : c.petType) + (c.service ? ' · interested in ' + c.service : '');
}

function notifyOwner(c) {
  const tel = c.phone.replace(/[^\d+]/g, '');
  const html =
    '<h2 style="margin:0">🐾 New sign-up: ' + esc(c.petType) + ' owner</h2>' +
    '<p><b>' + esc(c.name) + '</b> · <a href="tel:' + tel + '">' + esc(c.phone) + '</a> · <a href="mailto:' + esc(c.email) + '">' + esc(c.email) + '</a></p>' +
    '<p><b>Pet:</b> ' + esc(c.petType) + (c.service ? '<br><b>Clicked on:</b> ' + esc(c.service) : '') + '<br><b>Call by:</b> ' + friendlyDay(c.contactBy) + '</p>' +
    '<p style="color:#485E54">On the call, ask: pet name, breed and age, which service, how often, start date, neighbourhood, and anything special (medication, fears, feeding). Write it in "Your notes".</p>' +
    '<p><a href="' + sheetUrl() + '">Open the Clients sheet</a> · ID ' + c.id + '</p>';
  MailApp.sendEmail({ to: SETTINGS.OWNER_EMAIL, subject: '🐾 New sign-up: ' + c.name + ' – ' + c.petType + ' (call by ' + friendlyDay(c.contactBy) + ')', htmlBody: html, replyTo: c.email });
}

function confirmCustomer(c) {
  const first = c.name.split(/\s+/)[0];
  const pet = { 'Dog': 'your dog', 'Cat': 'your cat', 'Small pet': 'your pet' }[c.petType] || 'your pets';
  const body =
    'Hi ' + first + ',\n\n' +
    'Thanks for signing up with ' + SETTINGS.BUSINESS_NAME + '! We can\'t wait to meet ' + pet + '.\n\n' +
    'What happens next:\n' +
    '1. We\'ll call you by ' + friendlyDay(c.contactBy) + ' to hear about ' + pet + ' and set up a free meet & greet.\n' +
    '2. At the meet & greet we meet ' + pet + ', learn the routine and confirm your price.\n' +
    '3. You choose a schedule, and you get photos and a report after every visit.\n\n' +
    (SETTINGS.BUSINESS_PHONE ? 'Questions before then? Call or text ' + SETTINGS.BUSINESS_PHONE + '.\n\n' : 'Questions before then? Just reply to this email.\n\n') +
    SETTINGS.OWNER_NAME + '\n' + SETTINGS.BUSINESS_NAME + ' · Pet care in Moose Jaw, SK\n\n(Reference: ' + c.id + ')';
  MailApp.sendEmail({ to: c.email, subject: 'We got your sign-up – ' + SETTINGS.BUSINESS_NAME + ' 🐾', body: body, replyTo: SETTINGS.OWNER_EMAIL, name: SETTINGS.BUSINESS_NAME });
}

/* ======================= DAILY TO-DO EMAIL ===================== */

function dailyDigest() { sendDigest(false); }
function dailyDigestForce() { sendDigest(true); try { SpreadsheetApp.getUi().alert('Sent to ' + SETTINGS.OWNER_EMAIL); } catch (_) {} }

function sendDigest(force) {
  const now = new Date(), dow = Number(Utilities.formatDate(now, SETTINGS.TIMEZONE, 'u'));
  if (!force && SETTINGS.WORK_DAYS.indexOf(dow) === -1) return;
  const today = dayKey(now);
  const callNew = [], retry = [], mg = [], mgNoDate = [], checkins = [], reviews = [];
  let regulars = 0;
  allClients().forEach(c => {
    const cb = c.contactBy ? dayKey(c.contactBy) : '', fu = c.nextFollowUp ? dayKey(c.nextFollowUp) : '';
    switch (c.status) {
      case 'New': if (!cb || cb <= today) callNew.push(c); break;
      case 'No answer': if (!fu || fu <= today) retry.push(c); break;
      case 'Meet & greet booked': if (!fu) mgNoDate.push(c); else if (fu <= today) mg.push(c); break;
      case 'Booked': if (fu && fu <= today) checkins.push(c); break;
      case 'Regular client': regulars++; if (fu && fu <= today) reviews.push(c); break;
    }
  });
  const item = c => {
    const tel = c.phone.replace(/[^\d+]/g, ''), cb = c.contactBy ? dayKey(c.contactBy) : '';
    const late = c.status === 'New' && cb && cb < today ? ' <b style="color:#B3261E">OVERDUE</b>' : '';
    return '<li style="margin-bottom:8px"><b>' + esc(c.name) + '</b> – <a href="tel:' + tel + '">' + esc(c.phone) + '</a>' + late +
      '<br><span style="color:#485E54">' + esc(petLine(c)) + '</span>' +
      (c.myNotes ? '<br><i style="color:#777">Your notes: ' + esc(c.myNotes) + '</i>' : '') + '</li>';
  };
  const section = (title, list, hint) => list.length ? '<h3 style="margin:18px 0 6px">' + title + ' (' + list.length + ')</h3>' +
    (hint ? '<p style="margin:0 0 6px;color:#666">' + hint + '</p>' : '') + '<ul style="padding-left:18px;margin:0">' + list.map(item).join('') + '</ul>' : '';
  const total = callNew.length + retry.length + mg.length + checkins.length + reviews.length;
  const html = '<div style="font-family:Arial,sans-serif;max-width:620px">' +
    '<h2 style="margin:0">🐾 Tail Trails · ' + Utilities.formatDate(now, SETTINGS.TIMEZONE, 'EEEE, MMM d') + '</h2>' +
    '<p style="color:#485E54">' + (total ? total + ' client task' + (total > 1 ? 's' : '') + ' today.' : 'No client calls due today.') + (regulars ? ' Regular clients: ' + regulars + '.' : '') + '</p>' +
    section('📞 Contact new sign-ups', callNew, 'After each call, set Status to Meet & greet booked, No answer or Lost.') +
    section('🔁 Try again (no answer last time)', retry) +
    section('🤝 Meet & greets today', mg) +
    section('💬 First-week check-ins', checkins, 'Use the Tail Trails menu → first-week check-in, or call.') +
    section('⭐ Ask for a review', reviews, 'Use the Tail Trails menu → thank-you + review request.') +
    section('⚠️ Meet & greet booked but no date set', mgNoDate, 'Type the date in "Next follow-up".') +
    '<h3 style="margin:22px 0 6px">✅ Daily checklist</h3><ul style="padding-left:18px;margin:0">' + SETTINGS.DAILY_CHECKLIST.map(x => '<li>' + esc(x) + '</li>').join('') + '</ul>' +
    '<p style="margin-top:22px"><a href="' + sheetUrl() + '">Open the Clients sheet</a></p></div>';
  MailApp.sendEmail({ to: SETTINGS.OWNER_EMAIL, subject: 'Tail Trails today: ' + (callNew.length + retry.length) + ' to contact, ' + mg.length + ' meet & greets, ' + (checkins.length + reviews.length) + ' follow-ups', htmlBody: html });
}

/* =================== STATUS CHANGES (auto dates) ================ */

function handleEdit(e) {
  const r = e && e.range; if (!r) return;
  const sh = r.getSheet();
  if (sh.getName() !== SHEET_NAME || r.getRow() < 2 || r.getColumn() !== COL['Status'] || r.getNumColumns() > 1) return;
  const row = r.getRow(), status = String(r.getValue()), fu = sh.getRange(row, COL['Next follow-up']), now = new Date();
  switch (status) {
    case 'No answer': fu.setValue(addWorkDays(now, SETTINGS.NO_ANSWER_RETRY_DAYS)); break;
    case 'Meet & greet booked': fu.clearContent(); fu.setNote('Type the meet & greet date here.'); break;
    case 'Booked': fu.setNote(''); fu.setValue(addWorkDays(now, SETTINGS.FIRST_WEEK_CHECKIN_DAYS)); break;
    case 'Regular client': fu.setValue(addWorkDays(now, SETTINGS.REVIEW_REQUEST_DAYS)); break;
    case 'Paused': case 'Lost': fu.clearContent(); break;
  }
  sh.getRange(row, COL['Last updated']).setValue(now);
}

/* ============== MENU: follow-up emails you approve ============== */

function menuCheckIn() {
  sendTemplateToSelected('How is ' + '{pet}' + ' settling in? – Tail Trails', c =>
    'Hi ' + c.name.split(/\s+/)[0] + ',\n\nIt\'s been about a week since we started caring for ' + (c.petName || 'your pet') +
    '. How is everything going? Is there anything you\'d like us to change, like timing, routes or feeding?\n\n' +
    'Just reply to this email.\n\n' + SETTINGS.OWNER_NAME + '\n' + SETTINGS.BUSINESS_NAME + (SETTINGS.BUSINESS_PHONE ? '\n' + SETTINGS.BUSINESS_PHONE : ''),
    (sh, row) => sh.getRange(row, COL['Next follow-up']).clearContent());
}

function menuReviewRequest() {
  sendTemplateToSelected('Thank you from Tail Trails 🐾', c =>
    'Hi ' + c.name.split(/\s+/)[0] + ',\n\nThank you for trusting us with ' + (c.petName || 'your pet') + '. It\'s been a pleasure.\n\n' +
    'If you\'re happy with our care, a short review or a referral to a friend with pets would mean a lot to a small local business.\n\n' +
    SETTINGS.OWNER_NAME + '\n' + SETTINGS.BUSINESS_NAME,
    (sh, row) => sh.getRange(row, COL['Next follow-up']).clearContent());
}

function sendTemplateToSelected(subject, bodyFn, after) {
  const ui = SpreadsheetApp.getUi(), sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== SHEET_NAME) return ui.alert('Click a row in the Clients sheet first.');
  const row = sh.getActiveRange().getRow();
  if (row < 2) return ui.alert('Click a client row first.');
  const c = rowToClient(sh.getRange(row, 1, 1, HEADERS.length).getValues()[0]);
  if (!validEmail(c.email)) return ui.alert(c.name + ' has no email. Call instead: ' + c.phone);
  const subj = subject.replace('{pet}', c.petName || 'your pet'), body = bodyFn(c);
  if (ui.alert('Send this email to ' + c.name + ' <' + c.email + '>?', body, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  MailApp.sendEmail({ to: c.email, subject: subj, body: body, replyTo: SETTINGS.OWNER_EMAIL, name: SETTINGS.BUSINESS_NAME });
  const n = sh.getRange(row, COL['Your notes']);
  n.setValue((n.getValue() ? n.getValue() + '\n' : '') + Utilities.formatDate(new Date(), SETTINGS.TIMEZONE, 'MMM d') + ': emailed "' + subj + '"');
  if (after) after(sh, row);
  sh.getRange(row, COL['Last updated']).setValue(new Date());
  ui.alert('Sent.');
}

/* ============================ TEST ============================= */

function testLead() {
  const r = doPost({ postData: { contents: JSON.stringify({
    name: 'Test Customer', phone: '306-555-0100', email: SETTINGS.OWNER_EMAIL, petType: 'Dog', service: 'walk30' }) } });
  console.log(r.getContent());
}

/* =========================== HELPERS =========================== */

function clientSheet() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  return (id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet()).getSheetByName(SHEET_NAME);
}
function sheetUrl() { const sh = clientSheet(); return sh.getParent().getUrl() + '#gid=' + sh.getSheetId(); }
function allClients() {
  const sh = clientSheet(), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, HEADERS.length).getValues().filter(r => r[COL['Owner name'] - 1]).map(rowToClient);
}
function rowToClient(r) {
  const g = h => r[COL[h] - 1];
  return { id: String(g('ID')), name: String(g('Owner name')), phone: String(g('Phone')), email: String(g('Email')),
    petName: String(g('Pet name')), petType: String(g('Pet type') || 'Pet'), service: String(g('Interested in') || ''),
    status: String(g('Status') || 'New'),
    contactBy: g('Contact by'), nextFollowUp: g('Next follow-up'), myNotes: String(g('Your notes')) };
}
function isWorkDay(d) { return SETTINGS.WORK_DAYS.indexOf(Number(Utilities.formatDate(d, SETTINGS.TIMEZONE, 'u'))) !== -1; }
function keyToDate(k) { const s = String(k).split('-'); return new Date(+s[0], +s[1] - 1, +s[2], 12); }
function startOfDay(d) { return keyToDate(Utilities.formatDate(d, SETTINGS.TIMEZONE, 'yyyy-MM-dd')); }
function addWorkDays(d, n) { let x = startOfDay(d), k = 0; while (k < n) { x = new Date(x.getTime() + 864e5); if (isWorkDay(x)) k++; } return x; }
function contactByDate(now) {
  const hour = Number(Utilities.formatDate(now, SETTINGS.TIMEZONE, 'H'));
  return (isWorkDay(now) && hour < SETTINGS.SAME_DAY_CUTOFF_HOUR) ? startOfDay(now) : addWorkDays(now, 1);
}
function dayKey(d) { return d instanceof Date ? Utilities.formatDate(d, SETTINGS.TIMEZONE, 'yyyy-MM-dd') : String(d); }
function friendlyDay(d) {
  if (!(d instanceof Date)) return String(d || '');
  const k = dayKey(d), today = dayKey(new Date()), tmr = dayKey(new Date(Date.now() + 864e5)), f = Utilities.formatDate(d, SETTINGS.TIMEZONE, 'EEE MMM d');
  if (k === today) return 'today (' + f + ')';
  if (k === tmr) return 'tomorrow (' + f + ')';
  return Utilities.formatDate(d, SETTINGS.TIMEZONE, 'EEEE, MMM d');
}
function clean(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim().slice(0, max || 500); }
function safe(v) { return /^[=+\-@]/.test(v) ? "'" + v : v; }
function validEmail(v) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(v || '')); }
function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function columnLetter(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
