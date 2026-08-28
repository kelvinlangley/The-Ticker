/**
 * CRDB Brokerage — CDS account application receiver (Google Apps Script)
 * ─────────────────────────────────────────────────────────────────────
 * Receives the JSON POSTed by cds/index.html and turns each application into:
 *   1. a row in a Google Sheet (the desk's register / work queue),
 *   2. a Drive folder per applicant holding every ID photo & signature PNG,
 *   3. an e-mail to the brokerage desk with a readable summary + links.
 *
 * SETUP (≈5 minutes, one-time):
 *   1. Create a Google Sheet named e.g. "CDS Applications". Copy its ID from
 *      the URL and paste into SHEET_ID below.
 *   2. Create a Drive folder e.g. "CDS Applications - Documents". Copy its ID
 *      into FOLDER_ID below.
 *   3. Go to script.google.com → New project → paste this whole file.
 *   4. Deploy → New deployment → type "Web app" →
 *        Execute as: Me
 *        Who has access: Anyone
 *      → Deploy, authorize, and copy the Web app URL.
 *   5. Paste that URL into CONFIG.SUBMIT_ENDPOINT in cds/index.html.
 *
 * NOTES
 *   • The form sends Content-Type text/plain to avoid a CORS preflight —
 *     read the body from e.postData.contents, never e.parameter (and never
 *     name payload fields "c" or "sid"; those parameter names are reserved).
 *   • After ANY code change you must Deploy → Manage deployments → Edit →
 *     New version, or the /exec URL keeps serving the old code.
 *   • Keep the Sheet/Folder in the bank's Google Workspace account so the
 *     data stays under the bank's control (and under the Workspace DPA).
 *     Workspace accounts can send ~1,500 mails/day (consumer Gmail ~100/day)
 *     — far above any realistic application volume.
 *   • Apps Script has no documented incoming-POST cap (~50 MB by community
 *     consensus); the form compresses images client-side so a typical
 *     application is 1–4 MB. Execution limit is 6 minutes.
 *   • The endpoint is public. Set SHARED_SECRET to the same value as
 *     CONFIG.SHARED_SECRET in index.html to drop junk submissions.
 */

const SHEET_ID  = 'PASTE_SHEET_ID_HERE';
const FOLDER_ID = 'PASTE_DRIVE_FOLDER_ID_HERE';
const NOTIFY    = 'brokerage@crdbbank.co.tz';   // desk inbox (comma-separate for several)
const SHARED_SECRET = '';                        // '' disables the check

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    if (SHARED_SECRET && p.secret !== SHARED_SECRET) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'unauthorized' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    delete p.secret;
    const ref = String(p.reference || 'NO-REF').replace(/[^\w-]/g, '');

    // 1 ── save attachments + signatures to a per-application Drive folder
    const root = DriveApp.getFolderById(FOLDER_ID);
    const folder = root.createFolder(ref + ' — ' + (p.accountName || 'unnamed'));
    const links = [];
    const saveDataUrl = (dataURL, name) => {
      const m = /^data:(.+?);base64,(.*)$/.exec(dataURL || '');
      if (!m) return;
      const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name);
      links.push(name + ': ' + folder.createFile(blob).getUrl());
    };
    Object.entries(p.attachments || {}).forEach(([key, f]) => {
      const ext = /pdf/.test((f.dataURL || '').slice(0, 30)) ? '.pdf' : '.jpg';
      saveDataUrl(f.dataURL, key + ext);
    });
    (p.mandate && p.mandate.signatories || []).forEach((s, i) => {
      if (s.signature) saveDataUrl(s.signature, 'signature_' + String.fromCharCode(65 + i) + '.png');
    });
    // full JSON for the record / any future system import
    folder.createFile(ref + '.json', JSON.stringify(p, null, 2), 'application/json');

    // 2 ── append the register row
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    if (sh.getLastRow() === 0) {
      sh.appendRow(['Received', 'Reference', 'Account name', 'Type', 'Category', 'Class',
        'Primary holder', 'NIN', 'TIN', 'Mobile', 'Email', 'Bank', 'Bank acct', 'Mandate',
        'Source of funds', 'Existing CDS ID', 'Documents folder', 'Status']);
    }
    const h0 = (p.holders && p.holders[0]) || {};
    sh.appendRow([
      new Date(), ref, p.accountName || '', (p.accountType || '') + ' → ' + (p.formVariant || ''), p.category || '', p['class'] || '',
      [h0.surname, h0.firstName, h0.middleName].filter(Boolean).join(' ') || (p.company && p.company.name) || '',
      h0.nin || '', h0.tin || (p.company && p.company.tin) || '',
      h0.mobile ? '+255' + h0.mobile : (p.company && p.company.contact && p.company.contact.mobile) || '',
      h0.email || (p.company && p.company.email) || '',
      (p.bank && p.bank.bank) + ' · ' + (p.bank && p.bank.branch),
      (p.bank && p.bank.accountNumber) || '', (p.mandate && p.mandate.rule) || '',
      (p.declarations && p.declarations.sourceOfFunds) || '', p.existingCdsId || '',
      folder.getUrl(), 'NEW',
    ]);

    // 3 ── notify the desk
    MailApp.sendEmail({
      to: NOTIFY,
      subject: '[CDS] New application ' + ref + ' — ' + (p.accountName || ''),
      body: 'A new CDS account application has arrived.\n\n' +
        'Reference:   ' + ref + '\n' +
        'Account:     ' + (p.accountName || '') + ' (' + p.accountType + ' / ' + p.category + ' / ' + p['class'] + ')\n' +
        'BOT form:    ' + (p.formVariant || '') + '\n' +
        'Contact:     ' + (h0.mobile ? '+255' + h0.mobile : '') + '  ' + (h0.email || '') + '\n' +
        'Bank:        ' + (p.bank && (p.bank.bank + ' ' + p.bank.accountNumber)) + '\n' +
        (p.flags && p.flags.length ? '\nFLAGS TO CHECK:\n- ' + p.flags.join('\n- ') + '\n' : '') +
        '\nStill needed at branch: ' + ((p.branchChecklist || []).join('; ')) + '\n\n' +
        'Documents:   ' + folder.getUrl() + '\n' +
        'Register:    https://docs.google.com/spreadsheets/d/' + SHEET_ID + '\n\n' +
        (links.length ? 'Files:\n' + links.join('\n') : ''),
    });

    return ContentService.createTextOutput(JSON.stringify({ ok: true, reference: ref }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Health check: open the web app URL in a browser to confirm deployment.
function doGet() {
  return ContentService.createTextOutput('CDS application receiver is running.');
}
