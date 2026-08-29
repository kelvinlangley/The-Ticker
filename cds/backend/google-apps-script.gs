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

    // filled, printable application form as PDF (ready to print, stamp and lodge)
    let pdfBlob = null;
    try {
      pdfBlob = Utilities.newBlob(pdfHtml(p, ref), 'text/html', ref + '.html')
        .getAs('application/pdf').setName(ref + ' - Application form.pdf');
      links.unshift('Application form PDF: ' + folder.createFile(pdfBlob).getUrl());
    } catch (e2) { links.push('PDF generation failed: ' + e2); }

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
      h0.mobile ? (h0.mobile.indexOf('+') === 0 ? h0.mobile : '+255' + h0.mobile) :(p.company && p.company.contact && p.company.contact.mobile) || '',
      h0.email || (p.company && p.company.email) || '',
      (p.bank && p.bank.bank) + ' · ' + (p.bank && p.bank.branch),
      (p.bank && p.bank.accountNumber) || '', (p.mandate && p.mandate.rule) || '',
      (p.declarations && p.declarations.sourceOfFunds) || '', p.existingCdsId || '',
      folder.getUrl(), 'NEW',
    ]);

    // 3 ── notify the desk (filled form PDF attached)
    MailApp.sendEmail({
      to: NOTIFY,
      attachments: pdfBlob ? [pdfBlob] : [],
      subject: '[CDS] New application ' + ref + ' — ' + (p.accountName || ''),
      body: 'A new CDS account application has arrived.\n\n' +
        'Reference:   ' + ref + '\n' +
        'Account:     ' + (p.accountName || '') + ' (' + p.accountType + ' / ' + p.category + ' / ' + p['class'] + ')\n' +
        'BOT form:    ' + (p.formVariant || '') + '\n' +
        'Contact:     ' + (h0.mobile ? (h0.mobile.indexOf('+') === 0 ? h0.mobile : '+255' + h0.mobile) :'') + '  ' + (h0.email || '') + '\n' +
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

// Renders the application as a BOT-form-style HTML document for PDF conversion.
function pdfHtml(p, ref) {
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const dmy = iso => (!iso || !/^\d{4}-\d{2}-\d{2}$/.test(iso)) ? esc(iso) : iso.split('-').reverse().join('-');
  const row = (k, v) => '<tr><th>' + esc(k) + '</th><td>' + (v == null ? '' : v) + '</td></tr>';
  const join = a => a.filter(Boolean).join(' / ');
  const RULES = { 'any-one': 'Any one to sign', 'any-two': 'Any two to sign jointly', 'all': 'All to sign jointly' };

  const holderTbl = h => '<table>' +
    row('Name', esc(join([h.surname, h.firstName, h.middleName]).replace(/ \/ /g, ' '))) +
    row('Role', esc(h.role)) + row('Date of birth', dmy(h.dob)) +
    row('Nationality / residence', esc(join([h.nationality, h.residence])) + (h.residency ? ' (' + esc(h.residency) + ')' : '')) +
    row('NIDA NIN', esc(h.nin)) + row('TIN & place of issue', esc(join([h.tin, h.tinIssue]))) +
    row('Own CDS ID (if any)', esc(h.cdsId)) +
    row('Passport No. & place of issue', esc(join([h.passport, h.passportIssue])) + (h.passportExpiry ? ' - expires ' + dmy(h.passportExpiry) : '')) +
    row('Voter ID / Driving licence', esc(join([h.voterId, h.license]))) +
    (h.relationship ? row('Relationship to the minor', esc(h.relationship)) : '') +
    row('Occupation / employer / employment ID', esc(join([h.occupation, h.employer, h.employmentId]))) +
    row('Postal address', esc(h.postal)) + row('Physical address', esc(h.physical)) +
    row('Mobile / other phone', esc(join([h.mobile ? (h.mobile.indexOf('+') === 0 ? h.mobile : '+255' + h.mobile) : '', h.phone2]))) +
    row('E-mail', esc(h.email)) + row('Tax status', esc(h.taxStatus)) + '</table>';

  const b = p.bank || {}, m = p.mandate || {}, d = p.declarations || {}, co = p.company;
  let html = '<style>' +
    'body{font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#000;margin:24px}' +
    'h4{font-size:12px;background:#eee;border:1px solid #000;padding:4px 8px;margin:14px 0 6px}' +
    'table{width:100%;border-collapse:collapse;margin-bottom:6px}' +
    'th,td{border:1px solid #444;padding:4px 7px;text-align:left;vertical-align:top;font-weight:normal}' +
    'th{width:32%;background:#f6f6f6;font-weight:bold}' +
    '.head{text-align:center;border-bottom:2px solid #000;padding-bottom:8px;margin-bottom:12px}' +
    '.head b{font-size:14px;display:block}.fno{font-size:9px;font-family:monospace}' +
    '.mandate{border:1px solid #444;padding:8px;font-size:10px}' +
    '.office{border:2px solid #000;padding:8px;margin-top:14px}' +
    '.sig{height:46px}.doc{max-height:140px;max-width:220px;margin:4px}' +
    '</style>' +
    '<div class="head"><b>CDS SECURITIES ACCOUNT APPLICATION - GOVERNMENT SECURITIES</b>' +
    'Submitted through CRDB (CDS depository participant) to the Director Financial Markets, Bank of Tanzania' +
    '<div class="fno">Transcribe onto ' + esc(p.formVariant) + ' - Ref ' + esc(ref) + ' - ' + dmy(d.date) + ' - submitted ' + esc((p.submittedAt || '').slice(0, 16)) + '</div></div>' +
    '<h4>1. APPLICANT DETAILS</h4>' +
    '<table>' + row('Account name', '<b>' + esc(p.accountName) + '</b>') + row('Account type', esc(p.accountType)) +
    row('Existing CDS ID', esc(p.existingCdsId || 'None - new account')) + '</table>' +
    (p.holders || []).map(holderTbl).join('') +
    (co ? '<table>' + row('Company', esc(co.name)) + row('Registration No.', esc(co.regNumber)) +
      row('Date of incorporation', dmy(co.incDate)) + row('Nature of business', esc(co.business)) +
      row('Company TIN & place of issue', esc(join([co.tin, co.tinIssue]))) +
      row('Country of incorporation / residence', esc(join([co.nationality, co.residence]))) +
      row('Tax status', esc(co.taxStatus)) +
      row('Postal / physical address', esc(join([co.postal, co.physical]))) +
      row('Telephone / e-mail', esc(join([co.tel, co.email]))) +
      row('Contact person', esc(co.contact ? co.contact.name + ' (' + co.contact.role + ') ' + co.contact.mobile : '')) + '</table>' : '') +
    '<h4>2. SETTLEMENT BANK DETAILS</h4>' +
    '<table>' + row('Bank', esc(b.bank)) + row('Branch', esc(b.branch)) +
    row('Account number', esc(b.accountNumber)) + row('Account name', esc(b.accountName)) +
    row('Currency', esc(b.currency)) + row('Branch postal address', esc(b.postal)) +
    row('Branch telephone / e-mail', esc(join([b.tel, b.email]))) + '</table>' +
    '<h4>3. PERSONS AUTHORISED TO OPERATE THE CDS SECURITIES ACCOUNT</h4>' +
    '<table><tr><th style="width:6%">No.</th><th>Surname</th><th>First name</th><th>Middle name</th><th>Capacity</th><th>Specimen signature (provisional)</th></tr>' +
    (m.signatories || []).map((s, i) => '<tr><td>' + String.fromCharCode(65 + i) + '</td><td>' + esc(s.surname) + '</td><td>' + esc(s.firstName) + '</td><td>' + esc(s.middleName) + '</td><td>' + esc(s.capacity) + '</td><td>' +
      (s.signature ? '<img class="sig" src="' + s.signature + '">' : '') + '</td></tr>').join('') + '</table>' +
    '<table>' + row('Operating mandate', esc(RULES[m.rule] || m.rule)) + '</table>' +
    '<h4>4. CATEGORY OF THE CDS SECURITIES ACCOUNT HOLDER</h4>' +
    '<table>' + row('Category', esc(p.category)) + row('Class', esc(p['class'])) + '</table>' +
    '<h4>5. MANDATE FOR OPERATING CDS SECURITY ACCOUNT</h4>' +
    '<div class="mandate">I/We agree to operate the CDS securities account in accordance with the Central Depository System Dealing Agreement and the Central Depository System Rules and Operational Guidelines issued by the Bank of Tanzania. I/We declare the information in this application to be complete and true, consent to the processing of the personal data herein for the purposes of opening and operating the account (Personal Data Protection Act, 2022), and authorise CRDB to submit this application to the Bank of Tanzania on my/our behalf. Source of funds: <b>' + esc(d.sourceOfFunds) + '</b>. Date: <b>' + dmy(d.date) + '</b>.<br><br>' +
    '<b>Still to be completed at a CRDB branch:</b> ' + (p.branchChecklist || []).map(esc).join('; ') + '.' +
    ((p.flags || []).length ? '<br><br><b>FLAGS FOR THE DESK:</b> ' + p.flags.map(esc).join('; ') : '') + '</div>' +
    '<div class="office"><b>FOR OFFICIAL USE - CRDB (CDP) / BANK OF TANZANIA</b>' +
    '<table><tr><th>Originated by / date</th><td></td><th>Verified by / date</th><td></td></tr>' +
    '<tr><th>Approved by / date</th><td></td><th>Signatures witnessed by</th><td></td></tr>' +
    '<tr><th>CDP CDS ID</th><td></td><th>CDP CDS SEC. A/C</th><td></td></tr>' +
    '<tr><th>CDS account No. allocated</th><td></td><th>Remarks</th><td></td></tr></table></div>' +
    '<h4>ATTACHED DOCUMENTS</h4>';
  Object.keys(p.attachments || {}).forEach(k => {
    const f = p.attachments[k];
    html += '<p><b>' + esc(k) + '</b> - ' + esc(f.name) + '<br>' +
      ((f.dataURL || '').indexOf('data:image') === 0 ? '<img class="doc" src="' + f.dataURL + '">' : '(PDF saved as a separate file in this folder)') + '</p>';
  });
  return html;
}
