// CRDB Brokerage - staff bid-entry portal backend (Google Apps Script).
//
// One Google Sheet drives everything (create a blank sheet, put its ID below,
// set TOKEN_SECRET to any long random text, then run setup() once from the
// editor - it creates three tabs):
//   Auctions - one row per auction. The Active column is the admin's switch:
//                YES    = shown in the portal and OPEN for bidding
//                         (closes automatically at the deadline)
//                CLOSED = still shown, but staff see a thank-you note that the
//                         auction is closed (consider the next auction)
//                NO / blank = row ignored
//              Leave SettlementDate BLANK = auto: the day after AuctionDate.
//              Leave BidCutoff BLANK = auto deadline: 5:00 PM EAT the day
//              BEFORE AuctionDate. Grace period / extension = type an explicit
//              date-time in BidCutoff (e.g. 2026-09-02 11:00) - it overrides
//              the automatic deadline and is checked live on every submission.
//              AuctionDate must be a real date (yyyy-MM-dd or a date cell);
//              if it is not and BidCutoff is blank, bidding is HELD CLOSED
//              (an auction must never stay open with no deadline).
//              Next auction = add a row and type YES.
//              Title shown to staff = "AUCTION <AuctionNo> - <Security>".
//              Tenors: leave blank for a bond; for a T-bill auction list the
//              offered tenors e.g. "35,91,182,364". MaxWapTZS caps one
//              client's total WAP (price-taking) amount in this auction
//              (0/blank = no cap).
//   Staff    - Email / Name / Branch / StaffNumber / Active / Admin /
//              AdminPin. Staff log in with their CRDB e-mail (must end
//              @crdbbank.co.tz) and their staff number (digits, max 5) as the
//              PIN. Only Active=YES rows can log in. Admin=YES additionally
//              unlocks the admin portal (bids/admin.html): bid reports with
//              Excel/PDF download and an auction editor that writes to the
//              Auctions tab. STRONGLY RECOMMENDED for admins: type a long
//              AdminPin - the admin portal then requires it on sign-in, so a
//              guessable staff number alone can never open the reports.
//              Re-pasting a newer script version is safe: missing columns are
//              added to existing tabs automatically on the next request.
//   Bids     - one row per submitted bid, in the report column order
//              (File > Download > Excel).
//
// Endpoints (deploy as Web app: Execute as Me, Who has access: Anyone):
//   GET  ?action=config   -> active auction as JSON
//   POST {action:"login"} -> verifies Email+StaffNumber (rate-limited),
//                            returns a same-day token
//   POST {action:"bid"}   -> verifies token, re-validates everything, appends
//                            under a lock, idempotent per submitId
//
// After ANY code change: Deploy > Manage deployments > Edit > New version.

const SHEET_ID = 'PASTE_SHEET_ID_HERE';
const TOKEN_SECRET = 'CHANGE_ME_TO_ANY_LONG_RANDOM_TEXT';
const TZ = 'Africa/Dar_es_Salaam';
const STAFF_DOMAIN = 'crdbbank.co.tz';
const MAX_FACE_TZS = 500e9; // sanity ceiling per bid

const AUCTION_COLS = ['Active', 'AuctionNo', 'Security', 'Coupon', 'ISIN',
  'AuctionDate', 'SettlementDate', 'MaturityPeriod', 'Tenors', 'MinBidTZS',
  'MultipleTZS', 'PriceMin', 'PriceMax', 'MaxWapTZS', 'BidCutoff', 'Notes'];
const STAFF_COLS = ['Email', 'Name', 'Branch', 'StaffNumber', 'Active', 'Admin', 'AdminPin'];
const BID_COLS = ['Received', 'BidRef', 'AuctionNo', 'Security', 'StaffEmail',
  'StaffNumber', 'InvestorFullNames', 'NatureOfInvestor', 'SecuritiesAccountNumber',
  'FaceValueTZS', 'PriceType', 'CleanPricePer100', 'ConsiderationTZS',
  'AccountToDebit', 'Branch', 'ResponsiblePerson', 'ClientEmail', 'Tenor', 'SubmitID'];
const NATURES = ['Client', 'Client with Approval', 'Staff with Approval'];

function configured() {
  return SHEET_ID !== 'PASTE_SHEET_ID_HERE' && TOKEN_SECRET !== 'CHANGE_ME_TO_ANY_LONG_RANDOM_TEXT';
}

// Columns that must stay TEXT: Sheets otherwise re-types what is written or
// typed into them (Coupon "11.25%" -> 0.1125, Tenors "91,182" -> 91182,
// account numbers losing leading zeros).
const TEXT_COLS = {
  Auctions: ['AuctionNo', 'Security', 'Coupon', 'ISIN', 'MaturityPeriod', 'Tenors', 'BidCutoff', 'Notes'],
  Staff: ['StaffNumber', 'AdminPin'],
  Bids: ['StaffNumber', 'SecuritiesAccountNumber', 'AccountToDebit'],
};

// Creates missing tabs WITHOUT sample data (safe to run implicitly), appends
// headers a newer script version added, and pins text formats.
function ensureTabs() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  try { ss.setSpreadsheetTimeZone(TZ); } catch (e) {}
  const tab = (name, cols) => {
    let sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); sh.appendRow(cols); sh.setFrozenRows(1); }
    const head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
    cols.forEach(c => {
      if (head.indexOf(c) < 0) { sh.getRange(1, head.length + 1).setValue(c); head.push(c); }
    });
    (TEXT_COLS[name] || []).forEach(c => {
      const j = head.indexOf(c);
      if (j >= 0) sh.getRange(1, j + 1, sh.getMaxRows(), 1).setNumberFormat('@');
    });
    return sh;
  };
  tab('Auctions', AUCTION_COLS); tab('Staff', STAFF_COLS); tab('Bids', BID_COLS);
}

// Run this ONCE, manually, from the editor (Run > setup). Only a manual run
// seeds sample rows - web requests never do.
function setup() {
  ensureTabs();
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const auctions = ss.getSheetByName('Auctions');
  if (auctions.getLastRow() === 1) {
    // sample auction two weeks out; SettlementDate and BidCutoff left blank on
    // purpose so the automatic rules apply (settlement = day after the
    // auction, deadline = 5:00 PM EAT the day before)
    const sampleDate = Utilities.formatDate(new Date(Date.now() + 14 * 864e5), TZ, 'yyyy-MM-dd');
    auctions.appendRow(['YES', '711 (Re-open)', '11.25% 10yrs T-Bond', '11.25%',
      'TZ1996106112', sampleDate, '', '10 years', '', 1000000, 100000,
      80, 130, 0, '', 'Sample auction - edit or replace this row']);
  }
  const staff = ss.getSheetByName('Staff');
  if (staff.getLastRow() === 1) {
    // Deliberately Active=NO: these sample credentials are public (they sit
    // in the repo), so the row is a TEMPLATE only. Add your real staff with
    // Active=YES; give admins a long AdminPin.
    staff.appendRow(['sample.staff@' + STAFF_DOMAIN, 'Sample Staff (template - replace)', 'Head Office', '10001', 'NO', 'YES', '']);
  }
}

function sheetRows(name, cols) {
  let sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
  if (!sh) { ensureTabs(); sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(name); }
  const values = sh.getDataRange().getValues();
  const head = values[0].map(String);
  return values.slice(1).map(r => {
    const o = {};
    cols.forEach(c => { const i = head.indexOf(c); o[c] = i >= 0 ? r[i] : ''; });
    return o;
  });
}

// A value written to a cell must never execute as a formula.
function safe(v) {
  const s = String(v == null ? '' : v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function staffEmailOk(email) {
  return new RegExp('^[a-z0-9._%+-]+@' + STAFF_DOMAIN.replace(/\./g, '\\.') + '$').test(email);
}

function parseCutoff(v) {
  if (v instanceof Date) return v;                       // Sheet date-time cell
  const s = String(v || '').trim();
  if (!s) return null;                                    // no cutoff set
  const m = /^(\d{4}-\d{2}-\d{2})[ T](\d{2}:\d{2})(:\d{2})?$/.exec(s);
  if (!m) return new Date(0);                             // unparseable -> fail CLOSED
  return new Date(m[1] + 'T' + m[2] + (m[3] || ':00') + '+03:00'); // EAT
}

// 'yyyy-MM-dd' shifted by whole days (anchored at UTC noon, so no rollover).
function dayShift(ymd, days) {
  const t = new Date(ymd + 'T12:00:00Z');
  t.setUTCDate(t.getUTCDate() + days);
  return t.toISOString().slice(0, 10);
}

function activeAuction() {
  const rows = sheetRows('Auctions', AUCTION_COLS)
    .filter(r => ['YES', 'CLOSED'].indexOf(String(r.Active).trim().toUpperCase()) >= 0);
  if (!rows.length) return null;
  const a = rows[rows.length - 1]; // if several are marked, the newest row wins
  const statusOpen = String(a.Active).trim().toUpperCase() === 'YES';
  const fmt = v => (v instanceof Date) ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v || '');
  const auctionYmd = fmt(a.AuctionDate);
  const derivable = /^\d{4}-\d{2}-\d{2}$/.test(auctionYmd);
  // blank SettlementDate -> the day after the auction;
  // blank BidCutoff -> 5:00 PM EAT the day before the auction (an explicit
  // BidCutoff value always wins, e.g. for a grace period)
  let settlement = fmt(a.SettlementDate);
  if (!settlement && derivable) settlement = dayShift(auctionYmd, 1);
  let cutoff = parseCutoff(a.BidCutoff);
  if (!cutoff && derivable) cutoff = new Date(dayShift(auctionYmd, -1) + 'T17:00:00+03:00');
  const tenors = String(a.Tenors || '').split(',').map(t => t.trim()).filter(Boolean);
  const cfg = {
    auctionNo: String(a.AuctionNo), security: String(a.Security), coupon: String(a.Coupon),
    isin: String(a.ISIN), auctionDate: auctionYmd, settlementDate: settlement,
    maturityPeriod: String(a.MaturityPeriod), tenors: tenors,
    minBid: Number(a.MinBidTZS) || 0, multiple: Number(a.MultipleTZS) || 0,
    priceMin: Number(a.PriceMin) || 0, priceMax: Number(a.PriceMax) || 0,
    maxWap: Number(a.MaxWapTZS) || 0,
    deadline: cutoff && cutoff.getTime() > 0
      ? Utilities.formatDate(cutoff, TZ, 'd/MM/yyyy hh:mm a') : '',
    cutoff: cutoff && cutoff.getTime() > 0
      ? Utilities.formatDate(cutoff, TZ, 'yyyy-MM-dd HH:mm') : String(a.BidCutoff || ''),
    notes: String(a.Notes || ''),
  };
  // an incomplete auction row fails CLOSED, never permissive
  const complete = cfg.minBid > 0 && cfg.multiple > 0 && cfg.priceMax > cfg.priceMin && cfg.priceMin > 0;
  if (!complete) cfg.notes = (cfg.notes ? cfg.notes + ' · ' : '') +
    'Auction row incomplete (min/multiple/price band) - bidding held closed';
  // no deadline at all (blank BidCutoff and an AuctionDate the automatic rule
  // cannot read) also fails CLOSED - an auction must never stay open forever
  if (!cutoff) cfg.notes = (cfg.notes ? cfg.notes + ' · ' : '') +
    'No deadline set (AuctionDate must be yyyy-MM-dd, or fill BidCutoff) - bidding held closed';
  cfg.open = statusOpen && complete && !!cutoff && new Date() < cutoff;
  return cfg;
}

const b64u = s => Utilities.base64EncodeWebSafe(Utilities.newBlob(s).getBytes());
// Scoped same-day tokens: a staff token can never call admin endpoints.
function makeToken(email, scope) {
  const day = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const sig = Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(b64u(email) + '|' + day + '|' + scope, TOKEN_SECRET));
  return b64u(email) + '.' + day + '.' + scope + '.' + sig;
}
function checkToken(token, scope) {
  const p = String(token || '').split('.');
  if (p.length !== 4 || p[2] !== scope) return null;
  if (p[1] !== Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd')) return null; // expires midnight EAT
  let email;
  try { email = Utilities.newBlob(Utilities.base64DecodeWebSafe(p[0])).getDataAsString(); }
  catch (e) { return null; }
  return makeToken(email, scope) === token ? email : null;
}

// Brute-force brake: 5 failures locks an e-mail for 15 minutes.
function loginThrottle(email, failed) {
  const cache = CacheService.getScriptCache();
  const key = 'fail:' + email;
  const n = Number(cache.get(key) || 0);
  if (failed === undefined) return n >= 5;
  if (failed) { cache.put(key, String(n + 1), 900); return n + 1 >= 5; }
  cache.remove(key); return false;
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  if ((e && e.parameter && e.parameter.action) === 'config') {
    if (!configured()) return json({ ok: false, error: 'Backend not configured (SHEET_ID / TOKEN_SECRET)' });
    const a = activeAuction();
    return json(a ? { ok: true, auction: a }
      : { ok: false, noAuction: true, error: 'No auction is open right now' });
  }
  return ContentService.createTextOutput('CRDB bid portal backend is running.');
}

function doPost(e) {
  try {
    if (!configured()) return json({ ok: false, error: 'Backend not configured (SHEET_ID / TOKEN_SECRET)' });
    const p = JSON.parse(e.postData.contents);
    if (p.action === 'login') return handleLogin(p);
    if (p.action === 'bid') return handleBid(p);
    if (p.action === 'adminData') return handleAdminData(p);
    if (p.action === 'adminSaveAuction') return handleAdminSaveAuction(p);
    if (p.action === 'adminSetActive') return handleAdminSetActive(p);
    return json({ ok: false, error: 'Unknown action' });
  } catch (err) {
    console.error(err);
    return json({ ok: false, error: 'Server error - contact the admin' });
  }
}

function handleLogin(p) {
  const email = String(p.email || '').trim().toLowerCase();
  const staffNo = String(p.staffNo || '').trim();
  if (!staffEmailOk(email)) return json({ ok: false, error: 'Use your CRDB e-mail (…@' + STAFF_DOMAIN + ')' });
  if (!/^\d{1,5}$/.test(staffNo)) return json({ ok: false, error: 'Staff number is digits only, up to 5' });
  if (loginThrottle(email)) return json({ ok: false, error: 'Too many attempts - locked for 15 minutes' });
  const staff = sheetRows('Staff', STAFF_COLS).find(s =>
    String(s.Email).trim().toLowerCase() === email &&
    String(s.StaffNumber).trim() === staffNo &&
    String(s.Active).trim().toUpperCase() === 'YES');
  if (!staff) {
    const locked = loginThrottle(email, true);
    return json({ ok: false, error: locked ? 'Too many attempts - locked for 15 minutes' : 'E-mail and staff number do not match our staff list' });
  }
  const isAdmin = String(staff.Admin).trim().toUpperCase() === 'YES';
  // Admin portal sign-in (wantAdmin): if the row has an AdminPin, the staff
  // number alone is NOT enough - the pin must match too. Wrong pins count
  // toward the same 5-failure lockout.
  if (p.wantAdmin) {
    if (!isAdmin) {
      loginThrottle(email, false);
      return json({ ok: false, error: 'This account is not an admin. Ask the desk to mark Admin = YES for you on the staff list.' });
    }
    const pin = String(staff.AdminPin == null ? '' : staff.AdminPin).trim();
    if (pin && String(p.adminPin || '').trim() !== pin) {
      const locked = loginThrottle(email, true);
      return json({ ok: false, pinRequired: true, error: locked ? 'Too many attempts - locked for 15 minutes'
        : (String(p.adminPin || '').trim() ? 'Wrong Admin PIN' : 'This admin account also needs its Admin PIN') });
    }
    loginThrottle(email, false);
    return json({ ok: true, token: makeToken(email, 'admin'), name: String(staff.Name),
      branch: String(staff.Branch), staffNo: staffNo, admin: true });
  }
  loginThrottle(email, false);
  return json({ ok: true, token: makeToken(email, 'staff'), name: String(staff.Name),
    branch: String(staff.Branch), staffNo: staffNo, admin: isAdmin });
}

function handleBid(p) {
  const email = checkToken(p.token, 'staff');
  if (!email) return json({ ok: false, error: 'Session expired - please log in again' });
  const staff = sheetRows('Staff', STAFF_COLS).find(s =>
    String(s.Email).trim().toLowerCase() === email && String(s.Active).trim().toUpperCase() === 'YES');
  if (!staff) return json({ ok: false, error: 'Staff no longer active' });

  const a = activeAuction();
  if (!a) return json({ ok: false, error: 'No active auction' });
  if (!a.open) return json({ ok: false, error: 'Submission deadline has passed for auction ' + a.auctionNo });
  if (p.auctionNo !== a.auctionNo) return json({ ok: false, auctionChanged: true, error: 'The auction changed - the page will reload it' });

  // ── server-side re-validation of every field (all required) ──
  if (String(p.investorNames || '').trim().length < 2) return json({ ok: false, error: 'Investors full names missing' });
  if (NATURES.indexOf(String(p.nature)) < 0) return json({ ok: false, error: 'Nature of investor missing' });
  const acct = String(p.securitiesAccount || '').trim().toUpperCase();
  if (!/^(BOTCDSB026|BOTCDSCORU)\d{4,8}$/.test(acct))
    return json({ ok: false, error: 'Securities account must be BOTCDSB026 or BOTCDSCORU followed by numbers only' });
  if (a.tenors.length && a.tenors.indexOf(String(p.tenor || '').trim()) < 0)
    return json({ ok: false, error: 'Pick a tenor (' + a.tenors.join(', ') + ' days)' });
  if (!/^\d+$/.test(String(p.faceValue))) return json({ ok: false, error: 'Face value must be a whole number of shillings' });
  const amt = Number(p.faceValue);
  if (amt < a.minBid) return json({ ok: false, error: 'Face value minimum is TZS ' + a.minBid.toLocaleString() });
  if (amt > MAX_FACE_TZS) return json({ ok: false, error: 'Face value is implausibly large - check it' });
  if (a.multiple && amt % a.multiple !== 0) return json({ ok: false, error: 'Face value must be a multiple of TZS ' + a.multiple.toLocaleString() });
  const isClean = p.priceType === 'Clean Price';
  if (!isClean && p.priceType !== 'Weighted Average Price (WAP)')
    return json({ ok: false, error: 'Price choice missing' });
  if (isClean) {
    if (!/^\d+(\.\d{1,4})?$/.test(String(p.cleanPrice))) return json({ ok: false, error: 'Clean price takes at most 4 decimal places' });
    const pr = Number(p.cleanPrice);
    if (pr < a.priceMin || pr > a.priceMax)
      return json({ ok: false, error: 'Clean price must be between ' + a.priceMin + ' and ' + a.priceMax });
  }
  if (!/^\d{10,13}$/.test(String(p.accountToDebit || ''))) return json({ ok: false, error: 'Account to debit must be 10-13 digits' });
  if (!String(p.branch || '').trim()) return json({ ok: false, error: 'Branch name missing' });
  if (!String(p.responsible || '').trim()) return json({ ok: false, error: 'Responsible person missing' });
  const clientEmail = String(p.clientEmail || '').trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(clientEmail))
    return json({ ok: false, error: 'Client e-mail is required (a valid address)' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    // idempotency: a retry of the same submission returns the original result
    const cache = CacheService.getScriptCache();
    const subKey = /^[\w-]{8,64}$/.test(String(p.submitId || '')) ? 'sub:' + p.submitId : null;
    if (subKey) { const prev = cache.get(subKey); if (prev) return json(JSON.parse(prev)); }

    // WAP cap across the client's earlier WAP bids in this auction
    if (a.maxWap > 0 && !isClean) {
      const prior = sheetRows('Bids', BID_COLS)
        .filter(r => String(r.AuctionNo) === a.auctionNo &&
          String(r.SecuritiesAccountNumber).toUpperCase() === acct &&
          String(r.PriceType) !== 'Clean Price')
        .reduce((s, r) => s + (Number(r.FaceValueTZS) || 0), 0);
      if (prior + amt > a.maxWap)
        return json({ ok: false, error: 'WAP cap exceeded: this client already has TZS ' +
          prior.toLocaleString() + ' and the cap is TZS ' + a.maxWap.toLocaleString() });
    }

    const now = new Date();
    const ref = 'BID-' + Utilities.formatDate(now, TZ, 'yyyyMMdd-HHmmss') + '-' + String(staff.StaffNumber);
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Bids');
    sh.appendRow([
      now, ref, a.auctionNo, safe(a.security), email, safe(String(staff.StaffNumber)),
      safe(String(p.investorNames).trim()), String(p.nature), safe(acct),
      amt, String(p.priceType),
      isClean ? Number(p.cleanPrice) : '',
      isClean ? Math.round(amt * Number(p.cleanPrice)) / 100 : '',
      safe(String(p.accountToDebit)), safe(String(p.branch).trim()),
      safe(String(p.responsible).trim()), safe(clientEmail),
      a.tenors.length ? safe(String(p.tenor)) + ' days' : safe(a.maturityPeriod),
      String(p.submitId || ''),
    ]);

    const result = { ok: true, ref: ref,
      at: Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm:ss') };
    if (subKey) cache.put(subKey, JSON.stringify(result), 21600); // 6h
    return json(result);
  } finally {
    lock.releaseLock();
  }
}

/* ── admin portal (bids/admin.html) ──────────────────────────────────────
   Reports + auction editor. Every action re-verifies the token AND that the
   staff row is Active=YES with Admin=YES - the page itself is not trusted. */

function adminAuth(token) {
  const email = checkToken(token, 'admin');
  if (!email) return null;
  const s = sheetRows('Staff', STAFF_COLS).find(x =>
    String(x.Email).trim().toLowerCase() === email &&
    String(x.Active).trim().toUpperCase() === 'YES' &&
    String(x.Admin).trim().toUpperCase() === 'YES');
  return s ? { staff: s, email: email } : null;
}

function fmtCell(v, withTime) {
  if (v instanceof Date)
    return Utilities.formatDate(v, TZ, withTime ? 'yyyy-MM-dd HH:mm' : 'yyyy-MM-dd');
  return String(v == null ? '' : v);
}

// Every auction row WITH its sheet row number, so the editor can write back.
function auctionTable() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Auctions');
  const values = sh.getDataRange().getValues();
  const head = values[0].map(String);
  return values.slice(1).map((r, i) => {
    const o = { row: i + 2 };
    AUCTION_COLS.forEach(c => {
      const j = head.indexOf(c);
      const v = j >= 0 ? r[j] : '';
      o[c] = (c === 'BidCutoff') ? fmtCell(v, true)
        : (c === 'AuctionDate' || c === 'SettlementDate') ? fmtCell(v, false)
        : (v instanceof Date ? fmtCell(v, true) : v);
    });
    return o;
  }).filter(o => String(o.AuctionNo).trim() || String(o.Security).trim()); // skip blank ghost rows
}

// Guards a row-targeted write: the row must still hold the auction the admin
// was looking at (rows shift when someone edits the Sheet directly).
function rowStillHolds(sh, row, expectNo) {
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const j = head.indexOf('AuctionNo');
  if (j < 0) return false;
  const cur = String(sh.getRange(row, j + 1).getValue()).trim();
  return cur === String(expectNo == null ? '' : expectNo).trim();
}

function handleAdminData(p) {
  if (!adminAuth(p.token)) return json({ ok: false, error: 'Admin session expired - please log in again' });
  let bids = sheetRows('Bids', BID_COLS).map(r => {
    const o = {};
    BID_COLS.forEach(c => { o[c] = (r[c] instanceof Date)
      ? Utilities.formatDate(r[c], TZ, 'yyyy-MM-dd HH:mm:ss') : r[c]; });
    return o;
  });
  if (p.auctionNo) bids = bids.filter(b => String(b.AuctionNo) === String(p.auctionNo));
  return json({ ok: true, auctions: auctionTable(), bids: bids, live: activeAuction() });
}

// Validates one auction row from the editor; {err} or {vals} in AUCTION_COLS order.
function checkAuctionInput(a) {
  const t = k => String(a[k] == null ? '' : a[k]).trim();
  const active = t('Active').toUpperCase();
  if (['YES', 'CLOSED', 'NO'].indexOf(active) < 0) return { err: 'Active must be YES, CLOSED or NO' };
  if (!t('AuctionNo')) return { err: 'Auction number is required' };
  if (!t('Security')) return { err: 'Security is required' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(t('AuctionDate'))) return { err: 'Auction date is required (yyyy-MM-dd)' };
  if (t('SettlementDate') && !/^\d{4}-\d{2}-\d{2}$/.test(t('SettlementDate')))
    return { err: 'Settlement date must be yyyy-MM-dd, or blank = day after the auction' };
  const cut = t('BidCutoff').replace('T', ' ');
  if (cut && !/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}(:\d{2})?$/.test(cut))
    return { err: 'Deadline must be yyyy-MM-dd HH:mm, or blank = 5:00 PM the day before the auction' };
  const num = k => Number(String(a[k] == null ? '' : a[k]).replace(/[, ]/g, ''));
  const minBid = num('MinBidTZS'), mult = num('MultipleTZS');
  const pMin = num('PriceMin'), pMax = num('PriceMax');
  const maxWap = t('MaxWapTZS') === '' ? 0 : num('MaxWapTZS');
  if (!(minBid > 0)) return { err: 'Minimum bid must be a positive number' };
  if (!(mult > 0)) return { err: 'Multiples must be a positive number' };
  if (!(pMin > 0 && pMax > pMin)) return { err: 'Price band needs 0 < min < max' };
  if (!(maxWap >= 0)) return { err: 'WAP cap must be a number (0 = no cap)' };
  const tenors = t('Tenors');
  if (tenors && !/^\d+(\s*,\s*\d+)*$/.test(tenors))
    return { err: 'Tenors must be blank (bond) or a comma list of days e.g. 91,182,364' };
  return { vals: [active, safe(t('AuctionNo')), safe(t('Security')), safe(t('Coupon')), safe(t('ISIN')),
    t('AuctionDate'), t('SettlementDate'), safe(t('MaturityPeriod')), tenors,
    minBid, mult, pMin, pMax, maxWap, cut, safe(t('Notes'))] };
}

function handleAdminSaveAuction(p) {
  if (!adminAuth(p.token)) return json({ ok: false, error: 'Admin session expired - please log in again' });
  const chk = checkAuctionInput(p.auction || {});
  if (chk.err) return json({ ok: false, error: chk.err });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ensureTabs();
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Auctions');
    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    const row = Number(p.row) || 0;
    if (row && (row < 2 || row > sh.getLastRow() || !rowStillHolds(sh, row, p.expectAuctionNo)))
      return json({ ok: false, error: 'The auction list changed since your page loaded - refresh and try again' });
    // one auction number = one row; a duplicate would silently merge reports
    const newNo = String(chk.vals[1]).trim().toUpperCase();
    const dup = auctionTable().find(a => a.row !== row && String(a.AuctionNo).trim().toUpperCase() === newNo);
    if (dup) return json({ ok: false, error: 'Auction number ' + chk.vals[1] + ' is already used by another row - use a distinct number (e.g. add "(Re-open)")' });
    const target = row || sh.getLastRow() + 1;
    AUCTION_COLS.forEach((c, i) => {
      const j = head.indexOf(c);
      if (j >= 0) sh.getRange(target, j + 1).setValue(chk.vals[i]);
    });
    return json({ ok: true, row: target, live: activeAuction() });
  } finally { lock.releaseLock(); }
}

// Quick open/close/hide without editing the whole row.
function handleAdminSetActive(p) {
  if (!adminAuth(p.token)) return json({ ok: false, error: 'Admin session expired - please log in again' });
  const active = String(p.active || '').trim().toUpperCase();
  if (['YES', 'CLOSED', 'NO'].indexOf(active) < 0) return json({ ok: false, error: 'Active must be YES, CLOSED or NO' });
  const row = Number(p.row) || 0;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Auctions');
    if (row < 2 || row > sh.getLastRow() || !rowStillHolds(sh, row, p.expectAuctionNo))
      return json({ ok: false, error: 'The auction list changed since your page loaded - refresh and try again' });
    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    const j = head.indexOf('Active');
    if (j < 0) return json({ ok: false, error: 'Active column missing in the Auctions tab' });
    sh.getRange(row, j + 1).setValue(active);
    return json({ ok: true, row: row, live: activeAuction() });
  } finally { lock.releaseLock(); }
}
