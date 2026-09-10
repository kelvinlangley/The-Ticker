// CRDB Brokerage - staff bid-entry portal backend (Google Apps Script).
//
// One Google Sheet drives everything (create a blank sheet, put its ID below,
// then run setup() once from the editor - it creates three tabs):
//   Auctions - one row per auction. The Active column is the admin's switch:
//                YES    = shown in the portal and OPEN for bidding
//                         (closes automatically at BidCutoff)
//                CLOSED = still shown, but bidding is stopped (use this the
//                         moment submission time is up)
//                NO / blank = row ignored
//              New BOT call for tender = add a row with the new auction's
//              details and type YES. No code edits, ever.
//   Staff    - StaffID / Name / Branch / PIN / Active. Only people listed here
//              (Active = YES) can log in. Add or remove staff by editing rows.
//   Bids     - every submitted bid line lands here (File > Download > Excel).
//
// Endpoints (deploy as Web app: Execute as Me, Who has access: Anyone):
//   GET  ?action=config      -> the active auction as JSON (what the portal shows)
//   POST {action:"login"}    -> verifies StaffID+PIN against the Staff tab,
//                               returns a token valid for the rest of the day
//   POST {action:"bid"}      -> verifies the token AND re-validates every rule
//                               server-side, then appends rows to Bids
//
// Security model: the login and every bid are checked HERE, on the server -
// a visitor without a valid StaffID+PIN cannot submit anything, even if they
// read the portal's source code. Set TOKEN_SECRET to any long random text.
// After ANY code change: Deploy > Manage deployments > Edit > New version.

const SHEET_ID = 'PASTE_SHEET_ID_HERE';
const TOKEN_SECRET = 'CHANGE_ME_TO_ANY_LONG_RANDOM_TEXT';
const TZ = 'Africa/Dar_es_Salaam';

const AUCTION_COLS = ['Active', 'AuctionNo', 'Security', 'Coupon', 'ISIN',
  'AuctionDate', 'SettlementDate', 'MaturityPeriod', 'MinBidTZS', 'MultipleTZS',
  'PriceMin', 'PriceMax', 'BidCutoff', 'Notes'];
const STAFF_COLS = ['StaffID', 'Name', 'Branch', 'PIN', 'Active'];
const BID_COLS = ['Received', 'BidRef', 'AuctionNo', 'Security', 'StaffID', 'StaffName',
  'Branch', 'ClientName', 'ClientCDS', 'FundingAccount', 'ClientPhone', 'Line',
  'BidType', 'FaceValueTZS', 'PricePer100', 'Remarks'];

// Run this once from the editor (Run > setup) after pasting your SHEET_ID.
function setup() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const tab = (name, cols) => {
    let sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); sh.appendRow(cols); sh.setFrozenRows(1); }
    return sh;
  };
  const auctions = tab('Auctions', AUCTION_COLS);
  if (auctions.getLastRow() === 1) {
    auctions.appendRow(['YES', '711 (re-open)', '10-Year Treasury Bond (re-opening)', '11.25%',
      'TZ1996106112', '2026-09-02', '2026-09-04', '10 years', 1000000, 100000,
      80, 130, '2026-09-02 10:00', 'Sample auction - edit or replace this row']);
  }
  const staff = tab('Staff', STAFF_COLS);
  if (staff.getLastRow() === 1) {
    staff.appendRow(['1001', 'Sample Staff', 'Head Office', '1234', 'YES']);
  }
  tab('Bids', BID_COLS);
}

function sheetRows(name, cols) {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
  if (!sh) { setup(); return sheetRows(name, cols); }
  const values = sh.getDataRange().getValues();
  const head = values[0].map(String);
  return values.slice(1).map(r => {
    const o = {};
    cols.forEach(c => { const i = head.indexOf(c); o[c] = i >= 0 ? r[i] : ''; });
    return o;
  });
}

function activeAuction() {
  const rows = sheetRows('Auctions', AUCTION_COLS)
    .filter(r => ['YES', 'CLOSED'].indexOf(String(r.Active).trim().toUpperCase()) >= 0);
  if (!rows.length) return null;
  const a = rows[rows.length - 1]; // if several are marked, the newest row wins
  const statusOpen = String(a.Active).trim().toUpperCase() === 'YES';
  const fmt = v => (v instanceof Date) ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v || '');
  const fmtDT = v => (v instanceof Date) ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm') : String(v || '');
  const cutoff = a.BidCutoff instanceof Date ? a.BidCutoff : (a.BidCutoff ? new Date(String(a.BidCutoff).replace(' ', 'T') + ':00+03:00') : null);
  return {
    auctionNo: String(a.AuctionNo), security: String(a.Security), coupon: String(a.Coupon),
    isin: String(a.ISIN), auctionDate: fmt(a.AuctionDate), settlementDate: fmt(a.SettlementDate),
    maturityPeriod: String(a.MaturityPeriod),
    minBid: Number(a.MinBidTZS) || 0, multiple: Number(a.MultipleTZS) || 0,
    priceMin: Number(a.PriceMin) || 0, priceMax: Number(a.PriceMax) || 999,
    cutoff: fmtDT(a.BidCutoff), notes: String(a.Notes || ''),
    open: statusOpen && (!cutoff || new Date() < cutoff),
  };
}

function makeToken(staffId) {
  const day = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const sig = Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(staffId + '|' + day, TOKEN_SECRET));
  return staffId + '.' + day + '.' + sig;
}
function checkToken(token) {
  const p = String(token || '').split('.');
  if (p.length !== 3) return null;
  const day = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  if (p[1] !== day) return null; // tokens expire at midnight EAT
  return makeToken(p[0]) === token ? p[0] : null;
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action === 'config') {
    const a = activeAuction();
    return json(a ? { ok: true, auction: a }
      : { ok: false, noAuction: true, error: 'No auction is open right now' });
  }
  return ContentService.createTextOutput('CRDB bid portal backend is running.');
}

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);

    if (p.action === 'login') {
      const staff = sheetRows('Staff', STAFF_COLS).find(s =>
        String(s.StaffID).trim() === String(p.staffId).trim() &&
        String(s.PIN).trim() === String(p.pin).trim() &&
        String(s.Active).trim().toUpperCase() === 'YES');
      if (!staff) return json({ ok: false, error: 'Unknown staff ID or wrong PIN' });
      return json({ ok: true, token: makeToken(String(staff.StaffID).trim()),
        name: String(staff.Name), branch: String(staff.Branch) });
    }

    if (p.action === 'bid') {
      const staffId = checkToken(p.token);
      if (!staffId) return json({ ok: false, error: 'Session expired - please log in again' });
      const staff = sheetRows('Staff', STAFF_COLS).find(s =>
        String(s.StaffID).trim() === staffId && String(s.Active).trim().toUpperCase() === 'YES');
      if (!staff) return json({ ok: false, error: 'Staff no longer active' });

      const a = activeAuction();
      if (!a) return json({ ok: false, error: 'No active auction' });
      if (!a.open) return json({ ok: false, error: 'Bidding for auction ' + a.auctionNo + ' closed at ' + a.cutoff });
      if (p.auctionNo !== a.auctionNo) return json({ ok: false, error: 'The auction changed - reload the page' });

      // server-side re-validation: never trust the browser alone
      if (!String(p.clientName || '').trim()) return json({ ok: false, error: 'Client name missing' });
      if (!/^\d{1,12}$/.test(String(p.clientCDS || ''))) return json({ ok: false, error: 'Client CDS account must be digits' });
      if (!/^\d{10,16}$/.test(String(p.fundingAccount || ''))) return json({ ok: false, error: 'Funding account must be 10-16 digits' });
      const lines = Array.isArray(p.lines) ? p.lines : [];
      if (!lines.length || lines.length > 4) return json({ ok: false, error: '1 to 4 bid lines required' });
      for (let i = 0; i < lines.length; i++) {
        const L = lines[i], n = 'Line ' + (i + 1) + ': ';
        const amt = Number(L.amount);
        if (!isFinite(amt) || amt < a.minBid) return json({ ok: false, error: n + 'minimum is TZS ' + a.minBid.toLocaleString() });
        if (a.multiple && Math.round(amt) % a.multiple !== 0) return json({ ok: false, error: n + 'must be a multiple of TZS ' + a.multiple.toLocaleString() });
        if (L.type === 'Competitive') {
          const pr = Number(L.price);
          if (!isFinite(pr) || pr < a.priceMin || pr > a.priceMax)
            return json({ ok: false, error: n + 'price must be between ' + a.priceMin + ' and ' + a.priceMax });
          if (!/^\d+(\.\d{1,4})?$/.test(String(L.price)))
            return json({ ok: false, error: n + 'price takes at most 4 decimal places' });
        } else if (L.type !== 'Non-competitive') {
          return json({ ok: false, error: n + 'bid type missing' });
        }
      }

      const ref = 'BID-' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss') + '-' + staffId;
      const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Bids');
      lines.forEach((L, i) => sh.appendRow([
        new Date(), ref, a.auctionNo, a.security, staffId, String(staff.Name), String(staff.Branch),
        String(p.clientName).trim(), String(p.clientCDS), String(p.fundingAccount),
        String(p.clientPhone || ''), i + 1, String(L.type),
        Number(L.amount), L.type === 'Competitive' ? Number(L.price) : '', String(p.remarks || ''),
      ]));
      return json({ ok: true, ref: ref, lines: lines.length,
        at: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss') });
    }

    return json({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}
