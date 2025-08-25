// server.js
// Prudential WhatsApp Scheduler Bot
// Tech: Node.js, Express, Twilio WhatsApp, ExcelJS, Multer
// Single-file implementation for simplicity with a minimal HTML UI

require('dotenv').config();
const express = require('express');
const bodyParser = require('body-parser');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const { Twilio } = require('twilio');

// ---------------------- Config ----------------------
const app = express();
app.use(bodyParser.urlencoded({ extended: false }));
app.use(bodyParser.json());
app.use(express.static(path.join(__dirname, 'public')));

const DATA_DIR = process.env.DATA_DIR || __dirname;
const uploadDir = path.join(DATA_DIR, 'uploads');
const exportDir = path.join(DATA_DIR, 'exports');

if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir);
if (!fs.existsSync(exportDir)) fs.mkdirSync(exportDir);

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, uploadDir),
  filename: (req, file, cb) => cb(null, `clients_${Date.now()}${path.extname(file.originalname)}`)
});
const upload = multer({ storage });

// Twilio
const {
  TWILIO_ACCOUNT_SID,
  TWILIO_AUTH_TOKEN,
  TWILIO_WHATSAPP_NUMBER, // e.g., 'whatsapp:+14155238886' (Twilio sandbox) or your WABA number
  AGENT_DISPLAY_NAME // optional, e.g., 'Chiam from Prudential'
} = process.env;

let twilioClient = null;
if (TWILIO_ACCOUNT_SID && TWILIO_AUTH_TOKEN) {
  twilioClient = new Twilio(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN);
}

// ---------------------- In-memory State ----------------------
/**
 * availabilitySlots: Array of {
 *   id: string,              // unique
 *   start: Date,
 *   end: Date,
 *   label: string,           // human-friendly label
 *   booked: boolean,
 *   bookedBy?: string        // WhatsApp number 'whatsapp:+65...'
 * }
 */
let availabilitySlots = [];

/**
 * excelState: {
 *   filePath: string,
 *   workbook: ExcelJS.Workbook,
 *   sheet: ExcelJS.Worksheet,
 *   headerMap: { [headerName: string]: number } // column index by header name
 * }
 */
let excelState = null;

/**
 * clientsByWa: Map from whatsapp:+number -> { name, phone }
 */
let clientsByWa = new Map();

// ---------------------- Helpers ----------------------
function pad2(n) { return String(n).padStart(2, '0'); }

function parseAvailabilityLine(line) {
  // Accept formats like:
  // 25 Aug 1-5pm
  // 26 Aug 9am-12pm
  // 26 Aug 2-7pm
  const months = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
  const raw = line.trim().replace(/\s+/g, ' ');
  if (!raw) return null;

  // Extract parts: <day> <Mon> <range>
  const m = raw.match(/^(\d{1,2})\s+([A-Za-z]{3,})\s+(.*)$/);
  if (!m) return null;
  const day = parseInt(m[1], 10);
  const monStr = m[2].toLowerCase().slice(0,3);
  const monIdx = months.indexOf(monStr);
  if (monIdx < 0) return null;
  const rangeStr = m[3].toLowerCase();

  // startHour[-endHour][am/pm]
  let rangeMatch = rangeStr.match(/^(\d{1,2})(?:\:(\d{2}))?\s*(am|pm)?\s*[-–]\s*(\d{1,2})(?:\:(\d{2}))?\s*(am|pm)?$/);
  if (!rangeMatch) return null;

  let [_, sh, sm, sMer, eh, em, eMer] = rangeMatch;
  sm = sm || '00';
  em = em || '00';

  // If only one meridian present, use it for both; otherwise respect each
  if (!sMer && eMer) sMer = eMer;
  if (!eMer && sMer) eMer = sMer;

  let startH = parseInt(sh, 10);
  let endH = parseInt(eh, 10);
  const startM = parseInt(sm, 10);
  const endM = parseInt(em, 10);

  function to24(h, mer) {
    if (!mer) return h; // assume 24h
    if (mer === 'am') return h === 12 ? 0 : h;
    return h === 12 ? 12 : h + 12;
  }

  startH = to24(startH, sMer);
  endH = to24(endH, eMer);

  const now = new Date();
  const year = now.getFullYear();

  const start = new Date(year, monIdx, day, startH, startM);
  const end   = new Date(year, monIdx, day, endH, endM);
  if (end <= start) end.setHours(end.getHours() + 1);

  return { day, monIdx, start, end };
}

function humanSlotLabel(d1, d2) {
  const monthsShort = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const day = d1.getDate();
  const mon = monthsShort[d1.getMonth()];

  function hm(date) {
    let h = date.getHours();
    let mer = 'am';
    if (h === 0) { h = 12; mer = 'am'; }
    else if (h === 12) { mer = 'pm'; }
    else if (h > 12) { h = h - 12; mer = 'pm'; }
    const mm = date.getMinutes();
    return mm ? `${h}:${pad2(mm)}${mer}` : `${h}${mer}`;
  }

  const s = hm(d1);
  const e = hm(d2);
  const sMer = s.endsWith('am') ? 'am' : 'pm';
  const eMer = e.endsWith('am') ? 'am' : 'pm';

  let sCore = s.replace(/(am|pm)$/,'');
  let eCore = e;
  if (sMer === eMer) {
    eCore = e.replace(/(am|pm)$/,'');
    return `${day} ${mon} ${sCore}–${eCore}${eMer}`;
  }
  return `${day} ${mon} ${s}–${e}`;
}

function expandToHourlySlots(start, end) {
  const slots = [];
  let cursor = new Date(start);
  while (cursor < end) {
    const next = new Date(cursor);
    next.setHours(next.getHours() + 1);
    if (next > end) break; // no partial hour
    slots.push({ start: new Date(cursor), end: new Date(next) });
    cursor = next;
  }
  return slots;
}

function regenerateSlotIds() {
  availabilitySlots.forEach((s, idx) => s.id = `S${idx+1}`);
}

function listSlotsForMessage() {
  const lines = availabilitySlots
    .filter(s => !s.booked)
    .map((s, i) => `${i+1}) ${s.label}`);
  return lines.length ? lines.join('\n') : '(All slots have been booked)';
}



// JS fixed version of waFormat (override previous if typo)
function waFormat(numberRaw) {
  const digits = (numberRaw || '').toString().replace(/[^\d]/g, '');
  if (!digits) return null;
  const withCC = digits.startsWith('65') ? `+${digits}` : `+65${digits}`;
  return `whatsapp:${withCC}`;
}


async function loadExcel(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const sheet = workbook.worksheets[0];
  if (!sheet) throw new Error('Excel has no sheets');

  // 1) Ensure required headers exist (append any missing)
  const requiredHeaders = [
    'Client Name', 'Contact Number', 'Booked Date', 'Booked Time', 'Status', 'Last Notified'
  ];

  const headerRow = sheet.getRow(1);
  const existing = (headerRow.values || []).map(v => (typeof v === 'string' ? v.trim() : v));

  if (!existing || existing.length <= 1) {
    // Empty/invalid header row: write all required headers
    headerRow.values = [
      , // 1-based indexing
      'Client Name', 'Contact Number', 'Booked Date', 'Booked Time', 'Status', 'Last Notified'
    ];
    headerRow.commit();
  } else {
    // Normalize and append missing headers
    const headers = [];
    for (let i = 1; i <= headerRow.cellCount; i++) {
      const val = headerRow.getCell(i).value;
      headers.push(typeof val === 'string' ? val.trim() : String(val || ''));
    }
    requiredHeaders.forEach(h => { if (!headers.includes(h)) headers.push(h); });
    headerRow.values = [, ...headers];
    headerRow.commit();
  }

  // 2) Build headerMap after header row is finalized
  const headerMap = {};
  for (let i = 1; i <= headerRow.cellCount; i++) {
    const v = headerRow.getCell(i).value;
    if (v) headerMap[String(v).trim()] = i;
  }

  // 3) Initialize missing Status to "Pending"
  const statusIdx = headerMap['Status'];
  for (let r = 2; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const statusCell = row.getCell(statusIdx);
    if (!statusCell.value) {
      statusCell.value = 'Pending';
      row.commit();
    }
  }

  // 4) Ensure "Last Notified" column exists for all rows (blank by default)
  const lnIdx = headerMap['Last Notified'];
  for (let r = 2; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const c = row.getCell(lnIdx);
    if (c.value === undefined || c.value === null) {
      c.value = ''; // leave blank until a message is sent
      row.commit();
    }
  }

  // Persist any fixes
  await workbook.xlsx.writeFile(filePath);
  return { workbook, sheet, headerMap };
}


async function saveExcel() {
  if (!excelState) return;
  await excelState.workbook.xlsx.writeFile(excelState.filePath);
}

function findRowByPhone(phoneDigitsRaw) {
  const sheet = excelState.sheet;
  const cIdx = excelState.headerMap['Contact Number'];
  const phoneDigits = (phoneDigitsRaw || '').replace(/[^\d]/g, '');
  for (let r = 2; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const val = (row.getCell(cIdx).value || '').toString();
    const digits = val.replace(/[^\d]/g, '');
    if (digits && phoneDigits && digits.endsWith(phoneDigits)) return row;
  }
  return null;
}

async function sendWa(to, body) {
  if (!twilioClient) throw new Error('Twilio client not configured.');
  return twilioClient.messages.create({
    from: TWILIO_WHATSAPP_NUMBER,
    to,
    body
  });
}

// ---------------------- Routes ----------------------
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// 1) Upload Excel of clients
app.post('/upload-clients', upload.single('file'), async (req, res) => {
  try {
    const filePath = req.file.path;
    const { workbook, sheet, headerMap } = await loadExcel(filePath);

    // keep a handle to the workbook/sheet/headers for later writes
    excelState = { filePath, workbook, sheet, headerMap };

    // reset in-memory map
    clientsByWa.clear();

    // column indexes
    const nameIdx  = headerMap['Client Name'];
    const phoneIdx = headerMap['Contact Number'];

    // >>> paste your loop here <<<
    for (let r = 2; r <= sheet.rowCount; r++) {
      const row = sheet.getRow(r);
      const name = (row.getCell(nameIdx).value || '').toString().trim();
      const phoneRaw = (row.getCell(phoneIdx).value || '').toString().trim();
      const wa = waFormat(phoneRaw);
      const status = (row.getCell(headerMap['Status']).value || '').toString().trim();
      const lastNotified = row.getCell(headerMap['Last Notified']).value;

      if (wa) {
        clientsByWa.set(wa, { name, phone: phoneRaw, rowIndex: r, status, lastNotified });
      }
    }

    return res.json({ ok: true, message: 'Excel loaded', filePath, totalClients: clientsByWa.size });
  } catch (err) {
    console.error(err);
    return res.status(400).json({ ok: false, error: err.message });
  }
});


// 2) Set availability (text lines). Example body: { availabilityText: "25 Aug 1-5pm\n26 Aug 2-7pm" }
app.post('/set-availability', (req, res) => {
  const { availabilityText } = req.body;
  if (!availabilityText) return res.status(400).json({ ok:false, error:'availabilityText is required' });

  const lines = availabilityText.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  availabilitySlots = [];

  for (const line of lines) {
    const parsed = parseAvailabilityLine(line);
    if (!parsed) continue;
    const hourly = expandToHourlySlots(parsed.start, parsed.end);
    hourly.forEach(h => {
      availabilitySlots.push({ id: '', start: h.start, end: h.end, label: humanSlotLabel(h.start, h.end), booked: false });
    });
  }

  availabilitySlots.sort((a,b) => a.start - b.start);
  regenerateSlotIds();

  res.json({ ok:true, totalSlots: availabilitySlots.length, slots: availabilitySlots.map(s => s.label) });
});

// 3) Broadcast availability to all clients (WhatsApp)
app.post('/broadcast', async (req, res) => {
  try {
    if (!excelState) return res.status(400).json({ ok:false, error:'Upload an Excel first' });
    if (!availabilitySlots.length) return res.status(400).json({ ok:false, error:'Set availability first' });

    const agentName = AGENT_DISPLAY_NAME || 'Your Prudential Agent';
    const slotsText = listSlotsForMessage();

    const promises = [];
    for (const [wa, client] of clientsByWa.entries()) {
      const body = `Hi ${client.name}, this is ${agentName}.\nHere are the available 1-hour meeting slots:\n\n${slotsText}\n\nReply with the number of your preferred slot (e.g., 2).`;
      promises.push(sendWa(wa, body));
    }

    await Promise.all(promises);
    res.json({ ok:true, sentTo: clientsByWa.size });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok:false, error: err.message });
  }
});

// 4) Twilio inbound webhook for WhatsApp replies
// Configure Twilio to POST to: https://<your-domain>/whatsapp/inbound
app.post('/whatsapp/inbound', async (req, res) => {
  try {
    const from = req.body.From; // 'whatsapp:+65...'
    const body = (req.body.Body || '').toString().trim();

    // Acknowledge to Twilio quickly
    res.status(200).send('OK');

    if (!from || !clientsByWa.has(from)) {
      return; // Unknown sender
    }

    const choiceMatch = body.match(/(\d{1,3})/);
    if (!choiceMatch) {
      await sendWa(from, 'Please reply with the number of your preferred slot (e.g., 2).');
      return;
    }
    const idx = parseInt(choiceMatch[1], 10) - 1;
    if (isNaN(idx) || idx < 0) {
      await sendWa(from, 'Invalid slot number. Please try again.');
      return;
    }

    const openSlots = availabilitySlots.filter(s => !s.booked);
    if (idx >= openSlots.length) {
      await sendWa(from, 'That slot number is no longer available. Please pick another.');
      return;
    }
    const slot = openSlots[idx];

    if (slot.booked) {
      await sendWa(from, 'Sorry, that slot was just taken. Please choose another number.');
      return;
    }
    slot.booked = true;
    slot.bookedBy = from;

    const client = clientsByWa.get(from);
    const dateOnly = `${pad2(slot.start.getDate())} ${['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][slot.start.getMonth()]}`;
    const timeLabel = humanSlotLabel(slot.start, slot.end).split(' ').slice(2).join(' ');

    const row = findRowByPhone((client.phone || '').replace(/[^\d]/g, ''));
    if (row) {
      const { headerMap } = excelState;
      row.getCell(headerMap['Booked Date']).value = dateOnly;
      row.getCell(headerMap['Booked Time']).value = timeLabel;
      row.getCell(headerMap['Status']).value = 'Confirmed';
      row.commit();
      await saveExcel();
    }

    await sendWa(from, `Booked ✅: ${humanSlotLabel(slot.start, slot.end)}.\nThank you! We will contact you as soon as possible!.`);
  } catch (err) {
    console.error('Inbound handler error:', err);
  }
});

// 5) Download updated Excel
app.get('/download-latest', (req, res) => {
  try {
    if (!excelState) return res.status(400).json({ ok:false, error:'No Excel loaded yet.' });
    const outPath = path.join(exportDir, `bookings_${Date.now()}.xlsx`);
    fs.copyFileSync(excelState.filePath, outPath);
    res.download(outPath, 'bookings.xlsx');
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok:false, error: err.message });
  }
});

// 6) Health for Twilio
app.get('/health', (req, res) => res.json({ ok:true }));

// ---------------------- Start Server ----------------------
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server listening on port ${PORT}`);
});
