# CDS Account Opening Form — CRDB Brokerage

An online **CDS securities account application** for brokerage customers who want to buy
Government securities (Treasury bills & bonds) held in the **Bank of Tanzania Central
Depository System (CDS)**. The paper originals are BOT forms **CDS/FORM/02A** (sole
holder — individuals and entities), **02B** (joint — three holders per sheet; more
holders need an additional sheet) and **02C** (minor via guardian); this form captures
the union of their fields and tells the desk which variant to print. The Zansec online form (`zansec.co.tz/Onlineforms/cds/cds_form`) is built on
the same base form — this version restructures it so the brokerage desk receives a
complete, verified package it can act on without calling the customer back.

Official form PDFs (2022 revisions where available):

- [CDS/FORM/02A](https://www.bot.go.tz/webdocs/VariousForms/Financial%20Market%20-%20DFM/en/2022092113441553.pdf) · [02B](https://www.bot.go.tz/webdocs/VariousForms/Financial%20Market%20-%20DFM/en/2022092113450511.pdf) · [02C](https://www.bot.go.tz/webdocs/VariousForms/Financial%20Market%20-%20DFM/en/2020022109335762.pdf)
- [CDS Operational Guidelines](https://www.bot.go.tz/Publications/Acts,%20Regulations,%20Circulars,%20Guidelines/Guidelines/en/2020042813032255.pdf) (CRDB Bank Plc is registered CDS depository participant #15)

> **Template status:** this is a working template. Before public deployment it needs the
> bank's official branding (the colors here are placeholders), hosting on a
> bank-controlled HTTPS domain, PDPC registration checks, and compliance sign-off. It
> ships with no live submission endpoint configured. CRDB contact details in the page are
> placeholders — verify before shipping.

## What the customer experiences

Seven short steps with only the questions that apply to them:

1. **Account type** — Individual · Joint (2–3 in the form; the desk handles more) · Minor · Company/Institution.
   BOT category & class are derived automatically (individuals) or picked from the
   official category→class list (institutions). Maps to the right form variant
   (02A/02B/02C) automatically.
2. **Applicant details** — per holder: names, DOB, nationality/residence
   (EAC/SADC/diaspora classification derived, not asked), addresses, mobile
   (+255, `6/7` + 8 digits), e-mail, TIN (9–10 digits, auto-formatted), NIDA NIN
   (20 digits auto-formatted, soft-checked against DOB), passport + place of issue for
   non-residents, occupation/employer/employment ID, tax status. Guardian relationship
   for minors.
3. **Documents** — ID card/passport, passport photo, TIN certificate, and (per type)
   birth certificate + guardianship proof, certificate of incorporation, board
   resolution, signatories' IDs, tax-exemption certificate. Photos are **compressed in
   the browser** (≤1400 px JPEG) so a full application stays around 1–4 MB.
4. **Settlement bank** — defaults to CRDB Bank Plc; the BOT rule that the bank account
   name must equal the CDS account name is enforced with auto-copy + warning; for minors
   the account must be the child's own (flagged in-page).
5. **Signatories & specimen signatures** — up to 4 signatories (A–D), drawn signature
   pads (finger/mouse), operating mandate (any one / any two / all). The page is honest
   that BOT requires wet, witnessed signatures — the drawn ones are provisional.
6. **Mandate & declarations** — CDS Dealing Agreement mandate, truth declaration,
   **Personal Data Protection Act 2022 consent**, source of funds (AML).
7. **Review & submit** — everything on one screen with edit links and soft-check flags
   (NIN↔DOB mismatch, name mismatches, non-resident eligibility), then submit.

Extras: drafts auto-save to `localStorage` (text only — never ID images or signatures),
every application gets a reference (`CRDB-CDS-YYYYMMDD-XXXX`), and the customer can
print/save a **BOT-style PDF** of the completed application at any time.

## How the desk gets results it can act on

Every submission is one JSON document — field names mirror the BOT form sections, plus
processing metadata:

```json
{
  "reference": "CRDB-CDS-20260828-7K2Q",
  "formVariant": "CDS/FORM/02A",
  "accountType": "individual", "category": "Individuals", "class": "Individual",
  "accountName": "LANGLEY, Kelvin",
  "holders": [ { "surname": "...", "nin": "...", "tin": "...", "residency": "Resident — Tanzania", ... } ],
  "bank": { "bank": "CRDB Bank Plc", "branch": "...", "accountNumber": "...", "accountName": "..." },
  "mandate": { "rule": "any-one", "signatories": [ { "surname": "...", "signature": "data:image/png;base64,..." } ] },
  "declarations": { "sourceOfFunds": "...", "pdpaConsent": true, "date": "2026-08-28" },
  "flags": [ "…anything the desk should double-check…" ],
  "branchChecklist": [ "Sign the specimen card in ink before an officer", "Sign the Client CDS Agreement", "..." ],
  "attachments": { "id_0": { "name": "nida.jpg", "dataURL": "data:image/jpeg;base64,..." }, ... }
}
```

The printable PDF mirrors the BOT layout (sections 1–5, DD-MM-YYYY dates, officer trailer:
Originated/Verified/Approved by, CDP CDS ID, remarks) so an officer can transcribe onto
the official form — or attach it — with zero re-asking.

### Delivery options (pick one, set `CONFIG.SUBMIT_ENDPOINT` in `index.html`)

| Option | Setup | Where results land | Cost | Verdict |
|---|---|---|---|---|
| **1. Google Apps Script** (recommended start) | ~5 min — paste `backend/google-apps-script.gs`, deploy as Web App ("Execute as me" / "Anyone") | Google Sheet register + Drive folder per applicant (photos, signatures, JSON) + e-mail to the desk with links | Free, no monthly cap | The only zero-budget option that handles base64 images natively |
| **2. Supabase** | ~1 h — table + private storage bucket, anon insert with RLS | Postgres + object storage; build an admin view later | Free tier (500 MB DB / 1 GB storage; projects pause after 1 wk idle) · $25/mo Pro | Good growth path to a real system |
| **3. Microsoft Power Automate / Azure Function** (if the bank runs M365) | HTTP-trigger flow → SharePoint list + Teams/Outlook alert | SharePoint (bank tenant) | HTTP trigger is Premium-licensed | Best data-residency fit for an M365 bank |
| **4. Own API on bank infra** | Backend team | Bank database, CRM/core-banking integration, audit logs | Internal | The production end-state |
| **0. Offline fallback** (no endpoint) | Nothing | Customer downloads PDF + JSON and e-mails them to `DESK_EMAIL` | Free | Works today; also the automatic fallback if a configured endpoint fails |

**Ruled out for KYC images** (checked Aug 2026): Web3Forms (file uploads are Pro-only),
Formspree (uploads paid-only), EmailJS (≈50 KB request cap — one photo exceeds it),
Netlify Forms (requires Netlify hosting, no JSON AJAX), Basin/Getform (small free tiers).
More fundamentally, **no third-party form relay is appropriate for KYC documents** —
ID scans would transit and persist on an external processor under a consumer-grade DPA.

**Apps Script gotchas** (already handled in the shipped code): POST as
`text/plain` (Apps Script can't answer CORS preflights, so `application/json` fails);
read `e.postData.contents`, not `e.parameter`; the response is a 302 the browser follows
automatically; **redeploy a new version after every code change**; the endpoint is
public, so a shared-secret token (`CONFIG.SHARED_SECRET`) is supported end-to-end.

**PDPA compliance notes** (Act No. 11 of 2022, regulator PDPC): CRDB as data controller
must be PDPC-registered; the form carries a specific-purpose statement, an affirmative
unticked consent, and a rights notice. Storing KYC data outside Tanzania (Google,
Supabase) is a **cross-border transfer** needing explicit consent and a compliance
check — a strong argument for options 3/4, or option 1 run inside the bank's own
Workspace tenant, for anything beyond a pilot.

## Processing workflow for the desk

1. Notification arrives (e-mail / Sheet row) with the reference; the payload's `flags`
   list anything to double-check and `branchChecklist` what remains for the branch visit.
2. Verify photos & data, then invite the customer to a branch to **sign the specimen
   card and Client CDS Agreement in ink before an officer** (the guardian signs across
   both photos for a minor) — the only steps BOT does not accept online.
3. Print the application PDF or transcribe onto the official 02A/02B/02C, attach
   certified copies, and lodge with BOT Financial Markets (2 Mirambo Street, Dar es
   Salaam). BOT opens complete accounts within ~2 working days, or returns incomplete
   forms with reasons.
4. BOT confirms the CDS ID + client account number to CRDB, which must relay it to the
   customer within 3 working days — the form collects mobile + e-mail for exactly this.

## Known limitations

- Native date inputs display in the device's locale; stored values are ISO and the
  printable form renders DD-MM-YYYY as BOT expects, but very old browsers without date
  pickers fall back to free-text entry.
- Drafts keep the typed text (including ID numbers) in the browser's `localStorage` so a
  customer can resume; they are **offered** on return, never silently applied, and images
  and signatures are never stored. On genuinely shared computers customers should tap
  "Discard" when prompted — a real deployment on bank infrastructure may prefer
  server-side drafts.
- The `SHARED_SECRET` token ships in the public page source, so it filters junk POSTs;
  it is **not** authentication.
- Drawn/typed signatures are provisional by design — BOT requires wet signatures
  witnessed by the depository participant, which happens at the branch step.

## Files

- `index.html` — the whole app (no build step, no dependencies; fonts are the only
  external request). Configuration is the `CONFIG` object at the top of the script.
- `backend/google-apps-script.gs` — ready-to-paste receiver for option 1.
