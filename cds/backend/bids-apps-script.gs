// CRDB Brokerage - staff bid-entry portal backend (Google Apps Script).
//
// One Google Sheet drives everything (create a blank sheet, put its ID below,
// set TOKEN_SECRET to any long random text, then run setup() once from the
// editor - it creates three tabs and sample rows):
//   Auctions - one row per auction. The Active column is the admin's switch:
//                YES    = shown in the portal and OPEN for bidding
//                         (closes automatically at BidCutoff)
//                CLOSED = still shown, but bidding is stopped
//                NO / blank = row ignored
//              Extend or grace period = edit BidCutoff (checked live on every
//              submission). Next auction = add a row and type YES.
//              Tenors: leave blank for a bond; for a T-bill auction list the
//              offered tenors, e.g. "35,91,182,364" - staff then pick a tenor
//              per bid line. MaxNonCompTZS caps one client's total
//              non-competitive amount in this auction (0/blank = no cap).
//   Staff    - StaffID / Name / Branch / PIN / Active. Only Active=YES staff
//              can log in. StaffID must not contain '.' or '|'.
//   Bids     - every submitted bid line (File > Download > Excel).
//
// Endpoints (deploy as Web app: Execute as Me, Who has access: Anyone):
//   GET  ?action=config   -> active auction as JSON
//   POST {action:"login"} -> verifies StaffID+PIN (rate-limited), returns a
//                            same-day token
//   POST {action:"bid"}   -> verifies token, re-validates everything, appends
//                            under a lock, idempotent per submitId
//
// After ANY code change: Deploy > Manage deployments > Edit > New version.

const SHEET_ID = 'PASTE_SHEET_ID_HERE';
const TOKEN_SECRET = 'CHANGE_ME_TO_ANY_LONG_RANDOM_TEXT';
const TZ = 'Africa/Dar_es_Salaam';
const MAX_FACE_TZS = 500e9; // sanity ceiling per line

const AUCTION_COLS = ['Active', 'AuctionNo', 'Security', 'Coupon', 'ISIN',
  'AuctionDate', 'SettlementDate', 'MaturityPeriod', 'Tenors', 'MinBidTZS',
  'MultipleTZS', 'PriceMin', 'PriceMax', 'MaxNonCompTZS', 'BidCutoff', 'Notes'];
const STAFF_COLS = ['StaffID', 'Name', 'Branch', 'PIN', 'Active'];
const BID_COLS = ['Received', 'BidRef', 'AuctionNo', 'Security', 'Tenor', 'StaffID',
  'StaffName', 'Branch', 'ClientName', 'ClientCDS', 'FundingAccount', 'ClientPhone',
  'InstructionRef', 'Line', 'BidType', 'FaceValueTZS', 'PricePer100',
  'ConsiderationTZS', 'Remarks', 'SubmitID'];

function configured() {
  return SHEET_ID !== 'PASTE_SHEET_ID_HERE' && TOKEN_SECRET !== 'CHANGE_ME_TO_ANY_LONG_RANDOM_TEXT';
}

// Creates missing tabs WITHOUT sample data (safe to run implicitly).
function ensureTabs() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const tab = (name, cols) => {
    let sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); sh.appendRow(cols); sh.setFrozenRows(1); }
    return sh;
  };
  tab('Auctions', AUCTION_COLS); tab('Staff', STAFF_COLS); tab('Bids', BID_COLS);
}

// Run this ONCE, manually, from the editor (Run > setup). Only a manual run
// seeds the sample auction and sample staff - web requests never do.
function setup() {
  ensureTabs();
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const auctions = ss.getSheetByName('Auctions');
  if (auctions.getLastRow() === 1) {
    auctions.appendRow(['YES', '711 (re-open)', '10-Year Treasury Bond (re-opening)', '11.25%',
      'TZ1996106112', '2026-09-02', '2026-09-04', '10 years', '', 1000000, 100000,
      80, 130, 0, '2026-09-02 10:00', 'Sample auction - edit or replace this row']);
  }
  const staff = ss.getSheetByName('Staff');
  if (staff.getLastRow() === 1) staff.appendRow(['1001', 'Sample Staff', 'Head Office', '1234', 'YES']);
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

function parseCutoff(v) {
  if (v instanceof Date) return v;                       // Sheet date-time cell (sheet TZ)
  const s = String(v || '').trim();
  if (!s) return null;                                    // no cutoff set
  const m = /^(\d{4}-\d{2}-\d{2})[ T](\d{2}:\d{2})(:\d{2})?$/.exec(s);
  if (!m) return new Date(0);                             // unparseable -> fail CLOSED
  return new Date(m[1] + 'T' + m[2] + (m[3] || ':00') + '+03:00'); // EAT
}

function activeAuction() {
  const rows = sheetRows('Auctions', AUCTION_COLS)
    .filter(r => ['YES', 'CLOSED'].indexOf(String(r.Active).trim().toUpperCase()) >= 0);
  if (!rows.length) return null;
  const a = rows[rows.length - 1]; // if several are marked, the newest row wins
  const statusOpen = String(a.Active).trim().toUpperCase() === 'YES';
  const fmt = v => (v instanceof Date) ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v || '');
  const fmtDT = v => (v instanceof Date) ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm') : String(v || '');
  const cutoff = parseCutoff(a.BidCutoff);
  const tenors = String(a.Tenors || '').split(',').map(t => t.trim()).filter(Boolean);
  const cfg = {
    auctionNo: String(a.AuctionNo), security: String(a.Security), coupon: String(a.Coupon),
    isin: String(a.ISIN), auctionDate: fmt(a.AuctionDate), settlementDate: fmt(a.SettlementDate),
    maturityPeriod: String(a.MaturityPeriod), tenors: tenors,
    minBid: Number(a.MinBidTZS) || 0, multiple: Number(a.MultipleTZS) || 0,
    priceMin: Number(a.PriceMin) || 0, priceMax: Number(a.PriceMax) || 0,
    maxNonComp: Number(a.MaxNonCompTZS) || 0,
    cutoff: fmtDT(a.BidCutoff), notes: String(a.Notes || ''),
  };
  // an incomplete auction row fails CLOSED, never permissive
  const complete = cfg.minBid > 0 && cfg.multiple > 0 && cfg.priceMax > cfg.priceMin && cfg.priceMin > 0;
  if (!complete) cfg.notes = (cfg.notes ? cfg.notes + ' · ' : '') +
    'Auction row incomplete (min/multiple/price band) - bidding held closed';
  cfg.open = statusOpen && complete && (!cutoff || new Date() < cutoff);
  return cfg;
}

const b64u = s => Utilities.base64EncodeWebSafe(Utilities.newBlob(s).getBytes());
function makeToken(staffId) {
  const day = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const sig = Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(b64u(staffId) + '|' + day, TOKEN_SECRET));
  return b64u(staffId) + '.' + day + '.' + sig;
}
function checkToken(token) {
  const p = String(token || '').split('.');
  if (p.length !== 3) return null;
  if (p[1] !== Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd')) return null; // expires midnight EAT
  let staffId;
  try { staffId = Utilities.newBlob(Utilities.base64DecodeWebSafe(p[0])).getDataAsString(); }
  catch (e) { return null; }
  return makeToken(staffId) === token ? staffId : null;
}

// Simple brute-force brake: 5 failures locks a StaffID for 15 minutes.
function loginThrottle(staffId, failed) {
  const cache = CacheService.getScriptCache();
  const key = 'fail:' + staffId;
  const n = Number(cache.get(key) || 0);
  if (failed === undefined) return n >= 5;                 // query
  if (failed) { cache.put(key, String(n + 1), 900); return n + 1 >= 5; }
  cache.remove(key); return false;                          // success clears
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
    return json({ ok: false, error: 'Unknown action' });
  } catch (err) {
    console.error(err);
    return json({ ok: false, error: 'Server error - contact the admin' });
  }
}

function handleLogin(p) {
  const staffId = String(p.staffId || '').trim();
  if (!staffId || /[.|]/.test(staffId)) return json({ ok: false, error: 'Unknown staff ID or wrong PIN' });
  if (loginThrottle(staffId)) return json({ ok: false, error: 'Too many attempts - locked for 15 minutes' });
  const staff = sheetRows('Staff', STAFF_COLS).find(s =>
    String(s.StaffID).trim() === staffId &&
    String(s.PIN).trim() === String(p.pin).trim() &&
    String(s.Active).trim().toUpperCase() === 'YES');
  if (!staff) {
    const locked = loginThrottle(staffId, true);
    return json({ ok: false, error: locked ? 'Too many attempts - locked for 15 minutes' : 'Unknown staff ID or wrong PIN' });
  }
  loginThrottle(staffId, false);
  return json({ ok: true, token: makeToken(staffId), name: String(staff.Name), branch: String(staff.Branch) });
}

function handleBid(p) {
  const staffId = checkToken(p.token);
  if (!staffId) return json({ ok: false, error: 'Session expired - please log in again' });
  const staff = sheetRows('Staff', STAFF_COLS).find(s =>
    String(s.StaffID).trim() === staffId && String(s.Active).trim().toUpperCase() === 'YES');
  if (!staff) return json({ ok: false, error: 'Staff no longer active' });

  const a = activeAuction();
  if (!a) return json({ ok: false, error: 'No active auction' });
  if (!a.open) return json({ ok: false, error: 'Bidding for auction ' + a.auctionNo + ' is closed' });
  if (p.auctionNo !== a.auctionNo) return json({ ok: false, auctionChanged: true, error: 'The auction changed - the page will reload it' });

  // ── server-side re-validation: never trust the browser alone ──
  if (!String(p.clientName || '').trim()) return json({ ok: false, error: 'Client name missing' });
  const cds = String(p.clientCDS || '').trim();
  if (!/^(\d{1,12}|PENDING:[\w-]{4,40})$/.test(cds))
    return json({ ok: false, error: 'Client CDS must be digits, or PENDING:<application ref>' });
  if (!/^\d{10,16}$/.test(String(p.fundingAccount || ''))) return json({ ok: false, error: 'Funding account must be 10-16 digits' });
  const instr = String(p.instructionRef || '').trim();
  if (instr.length < 2 || instr.length > 60) return json({ ok: false, error: 'Client instruction reference required' });
  const lines = Array.isArray(p.lines) ? p.lines : [];
  if (!lines.length || lines.length > 4) return json({ ok: false, error: '1 to 4 bid lines required' });

  let newNonComp = 0;
  for (let i = 0; i < lines.length; i++) {
    const L = lines[i], n = 'Line ' + (i + 1) + ': ';
    if (a.tenors.length) {
      if (a.tenors.indexOf(String(L.tenor || '').trim()) < 0)
        return json({ ok: false, error: n + 'pick a tenor (' + a.tenors.join(', ') + ' days)' });
    }
    if (!/^\d+$/.test(String(L.amount))) return json({ ok: false, error: n + 'amount must be a whole number of shillings' });
    const amt = Number(L.amount);
    if (amt < a.minBid) return json({ ok: false, error: n + 'minimum is TZS ' + a.minBid.toLocaleString() });
    if (amt > MAX_FACE_TZS) return json({ ok: false, error: n + 'amount is implausibly large - check it' });
    if (a.multiple && amt % a.multiple !== 0) return json({ ok: false, error: n + 'must be a multiple of TZS ' + a.multiple.toLocaleString() });
    if (L.type === 'Competitive') {
      if (!/^\d+(\.\d{1,4})?$/.test(String(L.price))) return json({ ok: false, error: n + 'price takes at most 4 decimal places' });
      const pr = Number(L.price);
      if (pr < a.priceMin || pr > a.priceMax)
        return json({ ok: false, error: n + 'price must be between ' + a.priceMin + ' and ' + a.priceMax });
    } else if (L.type === 'Non-competitive') {
      newNonComp += amt;
    } else return json({ ok: false, error: n + 'bid type missing' });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    // idempotency: a retry of the same submission returns the original result
    const cache = CacheService.getScriptCache();
    const subKey = /^[\w-]{8,64}$/.test(String(p.submitId || '')) ? 'sub:' + p.submitId : null;
    if (subKey) { const prev = cache.get(subKey); if (prev) return json(JSON.parse(prev)); }

    // non-competitive cap across the client's earlier bids in this auction
    if (a.maxNonComp > 0 && newNonComp > 0) {
      const prior = sheetRows('Bids', BID_COLS)
        .filter(r => String(r.AuctionNo) === a.auctionNo && String(r.ClientCDS) === cds
          && String(r.BidType) === 'Non-competitive')
        .reduce((s, r) => s + (Number(r.FaceValueTZS) || 0), 0);
      if (prior + newNonComp > a.maxNonComp)
        return json({ ok: false, error: 'Non-competitive cap exceeded: this client already has TZS ' +
          prior.toLocaleString() + ' and the cap is TZS ' + a.maxNonComp.toLocaleString() });
    }

    const now = new Date();
    const ref = 'BID-' + Utilities.formatDate(now, TZ, 'yyyyMMdd-HHmmss') + '-' + staffId;
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Bids');
    const rows = lines.map((L, i) => [
      now, ref, a.auctionNo, a.security,
      a.tenors.length ? safe(L.tenor) + ' days' : safe(a.maturityPeriod),
      staffId, safe(staff.Name), safe(staff.Branch),
      safe(String(p.clientName).trim()), safe(cds), safe(String(p.fundingAccount)),
      safe(String(p.clientPhone || '')), safe(instr), i + 1, String(L.type), Number(L.amount),
      L.type === 'Competitive' ? Number(L.price) : '',
      L.type === 'Competitive' ? Math.round(Number(L.amount) * Number(L.price)) / 100 : '',
      safe(String(p.remarks || '')), String(p.submitId || ''),
    ]);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, BID_COLS.length).setValues(rows);

    const result = { ok: true, ref: ref, lines: lines.length,
      at: Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm:ss') };
    if (subKey) cache.put(subKey, JSON.stringify(result), 21600); // 6h
    return json(result);
  } finally {
    lock.releaseLock();
  }
}
