// server.js
// Prudential WhatsApp Scheduler Bot
// Tech: Node.js, Express, Twilio WhatsApp, ExcelJS, Multer

require('dotenv').config();
const express = require('express');
const bodyParser = require('body-parser');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const { Twilio } = require('twilio');

// phone-digits -> { confirmed:boolean, pending:boolean, notified:boolean, rowIndices:number[] }
let statusByDigits = new Map();

// ---------------------- App ----------------------
const app = express();
app.use(bodyParser.urlencoded({ extended: true })); // Twilio posts x-www-form-urlencoded
app.use(bodyParser.json());
app.use(express.static(path.join(__dirname, 'public')));

const DATA_DIR = process.env.DATA_DIR || __dirname;
const uploadDir = path.join(DATA_DIR, 'uploads');
const exportDir = path.join(DATA_DIR, 'exports');
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir, { recursive: true });
if (!fs.existsSync(exportDir)) fs.mkdirSync(exportDir, { recursive: true });

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, uploadDir),
  filename: (req, file, cb) => cb(null, `clients_${Date.now()}${path.extname(file.originalname)}`)
});
const upload = multer({ storage });

// ---------------------- Twilio ----------------------
// ---- Message templates (stored on disk) ----
const templatesPath = path.join(exportDir, 'templates.json');
const defaultTemplates = {
  broadcast:
    "Hi {{client.name}}, here are my available 1-hour meeting slots:\n\n{{slotsText}}\n\nReply with the number of your preferred slot (e.g., 2).",
  confirm:
    "📌 Hi {{client.name}}, your appointment is confirmed.\n\n🗓 {{slotLabel}}\n\n– Your Agent"
};

function loadTemplates() {
  try {
    if (fs.existsSync(templatesPath)) {
      const raw = JSON.parse(fs.readFileSync(templatesPath, 'utf8'));
      return {
        broadcast: typeof raw.broadcast === 'string' && raw.broadcast.trim() ? raw.broadcast : defaultTemplates.broadcast,
        confirm:   typeof raw.confirm   === 'string' && raw.confirm.trim()   ? raw.confirm   : defaultTemplates.confirm
      };
    }
  } catch (e) {
    console.error('Failed to load templates:', e.message);
  }
  return { ...defaultTemplates };
}

function saveTemplates(t) {
  const clean = {
    broadcast: typeof t.broadcast === 'string' && t.broadcast.trim() ? t.broadcast : defaultTemplates.broadcast,
    confirm:   typeof t.confirm   === 'string' && t.confirm.trim()   ? t.confirm   : defaultTemplates.confirm
  };
  fs.writeFileSync(templatesPath, JSON.stringify(clean, null, 2));
  return clean;
}

// simple {{path.to.value}} renderer (supports dot notation)
function renderTemplate(tpl, data) {
  return (tpl || '').replace(/\{\{\s*([a-zA-Z0-9_.]+)\s*\}\}/g, (_, path) => {
    const parts = path.split('.');
    let cur = data;
    for (const p of parts) {
      if (cur && Object.prototype.hasOwnProperty.call(cur, p)) cur = cur[p];
      else return `{{${path}}}`; // leave unknown placeholder as-is
    }
    return (cur ?? `{{${path}}}`).toString();
  });
}

// load on boot
let templates = loadTemplates();


const {
  TWILIO_ACCOUNT_SID,
  TWILIO_AUTH_TOKEN,
  TWILIO_WHATSAPP_NUMBER, // e.g., 'whatsapp:+14155238886'
  AGENT_DISPLAY_NAME
} = process.env;

let twilioClient = null;
if (TWILIO_ACCOUNT_SID && TWILIO_AUTH_TOKEN) {
  twilioClient = new Twilio(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN);
}

// ---------------------- State ----------------------
let availabilitySlots = [];   // [{id,start,end,label,booked,bookedBy}]
let excelState = null;        // { filePath, workbook, sheet, headerMap }
let clientsByWa = new Map();  // 'whatsapp:+65...' -> { name, phone, rowIndex, status, lastNotified }

// ---------------------- Helpers ----------------------
function expandToBufferedSlots(start, end, slotMins = 60, bufferMins = 0) {
  const slots = [];
  const step = slotMins + bufferMins; // e.g., 60 + 30 = 90 minutes between starts
  let cursor = new Date(start);

  while (cursor < end) {
    const slotEnd = new Date(cursor);
    slotEnd.setMinutes(slotEnd.getMinutes() + slotMins);
    if (slotEnd > end) break; // don’t create partial slot past the end
    slots.push({ start: new Date(cursor), end: slotEnd });
    cursor.setMinutes(cursor.getMinutes() + step); // jump by slot + buffer
  }
  return slots;
}


const pad2 = n => String(n).padStart(2, '0');

function phoneDigitsOnly(v) {
  return (v || '').toString().replace(/[^\d]/g, '');
}

function waFormat(numberRaw) {
  const digits = phoneDigitsOnly(numberRaw);
  if (!digits) return null;
  const withCC = digits.startsWith('65') ? `+${digits}` : `+65${digits}`;
  return `whatsapp:${withCC}`;
}

function parseAvailabilityLine(line) {
  const months = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
  const raw = (line || '').trim().replace(/\s+/g, ' ');
  if (!raw) return null;

  const m = raw.match(/^(\d{1,2})\s+([A-Za-z]{3,})\s+(.*)$/);
  if (!m) return null;
  const day = parseInt(m[1], 10);
  const monIdx = months.indexOf(m[2].toLowerCase().slice(0,3));
  if (monIdx < 0) return null;

  const rangeStr = m[3].toLowerCase();
  const rm = rangeStr.match(/^(\d{1,2})(?::(\d{2}))?\s*(am|pm)?\s*[-–]\s*(\d{1,2})(?::(\d{2}))?\s*(am|pm)?$/);
  if (!rm) return null;

  let [_, sh, sm, sMer, eh, em, eMer] = rm;
  sm = sm || '00';
  em = em || '00';
  if (!sMer && eMer) sMer = eMer;
  if (!eMer && sMer) eMer = sMer;

  let startH = parseInt(sh, 10);
  let endH   = parseInt(eh, 10);
  const startM = parseInt(sm, 10);
  const endM   = parseInt(em, 10);

  const to24 = (h, mer) => {
    if (!mer) return h;
    if (mer === 'am') return h === 12 ? 0 : h;
    return h === 12 ? 12 : h + 12;
  };
  startH = to24(startH, sMer);
  endH   = to24(endH, eMer);

  const year = new Date().getFullYear();
  const start = new Date(year, monIdx, day, startH, startM);
  const end   = new Date(year, monIdx, day, endH, endM);
  if (end <= start) end.setHours(end.getHours() + 1);

  return { start, end };
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
    else if (h > 12) { h -= 12; mer = 'pm'; }
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
    if (next > end) break;
    slots.push({ start: new Date(cursor), end: new Date(next) });
    cursor = next;
  }
  return slots;
}

function regenerateSlotIds() {
  availabilitySlots.forEach((s, idx) => { s.id = `S${idx + 1}`; });
}

function listSlotsForMessage() {
  const lines = availabilitySlots
    .filter(s => !s.booked)
    .map((s, i) => `${i + 1}) ${s.label}`);
  return lines.length ? lines.join('\n') : '(All slots have been booked)';
}

async function sendWa(to, body) {
  if (!twilioClient) throw new Error('Twilio client not configured.');
  return twilioClient.messages.create({
    from: TWILIO_WHATSAPP_NUMBER,
    to,
    body
  });
}

// ---------------------- Excel ----------------------
async function loadExcel(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const sheet = workbook.worksheets[0];
  if (!sheet) throw new Error('Excel has no sheets');

  const requiredHeaders = [
    'Client Name', 'Contact Number', 'Booked Date', 'Booked Time', 'Status', 'Last Notified'
  ];

  const headerRow = sheet.getRow(1);
  const existing = (headerRow.values || []).map(v => (typeof v === 'string' ? v.trim() : v));

  if (!existing || existing.length <= 1) {
    headerRow.values = [, ...requiredHeaders]; // 1-based indexing
    headerRow.commit();
  } else {
    const headers = [];
    for (let i = 1; i <= headerRow.cellCount; i++) {
      const val = headerRow.getCell(i).value;
      headers.push(typeof val === 'string' ? val.trim() : String(val || ''));
    }
    requiredHeaders.forEach(h => { if (!headers.includes(h)) headers.push(h); });
    headerRow.values = [, ...headers];
    headerRow.commit();
  }

  // Build headerMap AFTER committing headers
  const headerMap = {};
  for (let i = 1; i <= headerRow.cellCount; i++) {
    const v = headerRow.getCell(i).value;
    if (v) headerMap[String(v).trim()] = i;
  }

  // IMPORTANT: Do NOT auto-set Status to "Pending" here.
  // You want Pending to mean "already contacted", so mark it yourself or during broadcast.

  // Ensure "Last Notified" exists on all rows (leave blank)
  const lnIdx = headerMap['Last Notified'];
  for (let r = 2; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const c = row.getCell(lnIdx);
    if (c.value === undefined || c.value === null) { c.value = ''; row.commit(); }
  }

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
  const phoneDigits = phoneDigitsOnly(phoneDigitsRaw);
  for (let r = 2; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const val = (row.getCell(cIdx).value || '').toString();
    const digits = phoneDigitsOnly(val);
    if (digits && phoneDigits && digits.endsWith(phoneDigits)) return row;
  }
  return null;
}

// ---------------------- Routes ----------------------
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Upload Excel and build in-memory maps
app.post('/upload-clients', upload.single('file'), async (req, res) => {
  try {
    const filePath = req.file.path;
    const { workbook, sheet, headerMap } = await loadExcel(filePath);

    excelState = { filePath, workbook, sheet, headerMap };
    clientsByWa.clear();
    statusByDigits.clear();

    const nameIdx  = headerMap['Client Name'];
    const phoneIdx = headerMap['Contact Number'];
    const statusIdx = headerMap['Status'];
    const lnIdx = headerMap['Last Notified'];

    for (let r = 2; r <= sheet.rowCount; r++) {
      const row = sheet.getRow(r);
      const name = (row.getCell(nameIdx).value || '').toString().trim();
      const phoneRaw = (row.getCell(phoneIdx).value || '').toString().trim();
      const wa = waFormat(phoneRaw);
      const status = (row.getCell(statusIdx).value || '').toString().trim().toLowerCase();
      const lastNotified = row.getCell(lnIdx).value;
      const digits = phoneDigitsOnly(phoneRaw);

      if (wa) {
        clientsByWa.set(wa, { name, phone: phoneRaw, rowIndex: r, status, lastNotified });
      }

      if (digits) {
        let agg = statusByDigits.get(digits);
        if (!agg) agg = { confirmed: false, pending: false, notified: false, rowIndices: [] };
        if (status === 'confirmed') agg.confirmed = true;
        if (status === 'pending')   agg.pending = true;
        if (lastNotified)           agg.notified = true;
        agg.rowIndices.push(r);
        statusByDigits.set(digits, agg);
      }
    }

    res.json({ ok: true, message: 'Excel loaded', filePath, totalClients: clientsByWa.size });
  } catch (err) {
    console.error(err);
    res.status(400).json({ ok: false, error: err.message });
  }
});

// Set availability
app.post('/set-availability', (req, res) => {
  const { availabilityText, bufferMinutes } = req.body || {};
  if (!availabilityText) return res.status(400).json({ ok:false, error:'availabilityText is required' });

  // validate buffer: allow 0 (default), 30, 60
  const allowed = new Set([0, 30, 60, '0', '30', '60']);
  const bufRaw = bufferMinutes ?? 0;
  const buf = parseInt(bufRaw, 10);
  const buffer = allowed.has(bufRaw) && [0,30,60].includes(buf) ? buf : 0;

  const lines = availabilityText.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  availabilitySlots = [];

  for (const line of lines) {
    const parsed = parseAvailabilityLine(line);
    if (!parsed) continue;

    // slot duration fixed at 60 mins; buffer variable (0/30/60)
    const slots = expandToBufferedSlots(parsed.start, parsed.end, 60, buffer);
    for (const h of slots) {
      availabilitySlots.push({
        id: '',
        start: h.start,
        end: h.end,
        label: humanSlotLabel(h.start, h.end),
        booked: false
      });
    }
  }

  availabilitySlots.sort((a,b) => a.start - b.start);
  regenerateSlotIds();

  res.json({
    ok:true,
    bufferMinutes: buffer,
    totalSlots: availabilitySlots.length,
    slots: availabilitySlots.map(s => s.label)
  });
});


// Broadcast (skip phones that are Pending or Confirmed on ANY row)
app.post('/broadcast', async (req, res) => {
  try {
    if (!excelState) return res.status(400).json({ ok:false, error:'Upload an Excel first' });
    if (!availabilitySlots.length) return res.status(400).json({ ok:false, error:'Set availability first' });

    const agentName = AGENT_DISPLAY_NAME || 'Your Agent';
    const slotsText = listSlotsForMessage();
    const force = (req.query.force || '').toString().toLowerCase() === 'true';

    const toSend = [];
    const usedDigits = new Set();

    for (const [wa, client] of clientsByWa.entries()) {
      const digits = phoneDigitsOnly(client.phone);
      if (!digits || usedDigits.has(digits)) continue;

      const agg = statusByDigits.get(digits) || { confirmed: false, pending: false, notified: false, rowIndices: [] };

      // Skip if any row has Pending or Confirmed
      if (!force && (agg.pending || agg.confirmed)) {
        usedDigits.add(digits);
        continue;
      }

      usedDigits.add(digits);
      toSend.push({ wa, digits, client, rowIndices: agg.rowIndices });
    }

    if (!toSend.length) {
      return res.json({
        ok: true,
        sentTo: 0,
        skipped: clientsByWa.size,
        reason: force ? 'No eligible numbers' : 'All numbers have Pending or Confirmed status (or were duplicates).'
      });
    }

    const whenISO = new Date().toISOString();

    const tasks = toSend.map(async ({ wa, digits, client, rowIndices }) => {
      const body = renderTemplate(templates.broadcast, {
        client: { name: client.name },
        slotsText // allowed token
      });
      await sendWa(wa, body);


      // Mark Last Notified on all rows for this phone
      for (const r of rowIndices) {
        const row = excelState.sheet.getRow(r);
        row.getCell(excelState.headerMap['Last Notified']).value = whenISO;
        const sIdx = excelState.headerMap['Status'];
        for (const r of rowIndices) {
          const row = excelState.sheet.getRow(r);
          row.getCell(excelState.headerMap['Last Notified']).value = whenISO;

          const cur = (row.getCell(sIdx).value || '').toString().trim().toLowerCase();
          if (cur !== 'confirmed') {
            row.getCell(sIdx).value = 'Pending';
          }
          row.commit();
        }

        // keep in-memory aggregator in sync so future broadcasts skip immediately
        const agg = statusByDigits.get(digits) || { confirmed: false, pending: false, notified: false, rowIndices: [] };
        agg.pending = true;
        agg.notified = true;
        statusByDigits.set(digits, agg);
        client.status = 'pending';

        row.commit();
      }
    });

    await Promise.all(tasks);
    await saveExcel();

    res.json({ ok:true, sentTo: toSend.length, skipped: clientsByWa.size - toSend.length });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok:false, error: err.message });
  }
});

// Inbound webhook
app.post('/whatsapp/inbound', async (req, res) => {
  try {
    const from = req.body.From;
    const text = (req.body.Body || '').toString().trim();

    // Acknowledge Twilio immediately
    res.status(200).send('OK');

    // Only handle replies from numbers in the uploaded Excel
    if (!from || !clientsByWa.has(from)) return;

    // Expect a numeric choice (1-based)
    const m = text.match(/(\d{1,3})/);
    if (!m) {
      await sendWa(from, 'Please reply with the number of your preferred slot (e.g., 2).');
      return;
    }
    const idx = parseInt(m[1], 10) - 1;
    if (isNaN(idx) || idx < 0) {
      await sendWa(from, 'Invalid slot number. Please try again.');
      return;
    }

    // Map to the current list of open (unbooked) slots
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

    // Book it
    slot.booked = true;
    slot.bookedBy = from;

    // Update Excel row
    const client = clientsByWa.get(from);
    const dateOnly = `${pad2(slot.start.getDate())} ${['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][slot.start.getMonth()]}`;
    const timeLabel = humanSlotLabel(slot.start, slot.end).split(' ').slice(2).join(' ');

    const row = findRowByPhone(client.phone);
    if (row) {
      const { headerMap } = excelState;
      row.getCell(headerMap['Booked Date']).value = dateOnly;
      row.getCell(headerMap['Booked Time']).value = timeLabel;
      row.getCell(headerMap['Status']).value = 'Confirmed';
      row.commit();
      await saveExcel();
    }

    // ✅ Use your editable confirm template
    const slotLabel = humanSlotLabel(slot.start, slot.end);
    const body = renderTemplate(templates.confirm, {
      client: { name: client.name },
      slotLabel
    });
    await sendWa(from, body);
  } catch (err) {
    console.error('Inbound handler error:', err);
  }
});


// Download latest Excel
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

// Health
app.get('/health', (req, res) => res.json({ ok:true }));

// Start
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server listening on port ${PORT}`);
});

// Get current templates
app.get('/templates', (req, res) => {
  res.json({ ok: true, templates });
});

// Update templates
app.post('/templates', (req, res) => {
  try {
    const { broadcast, confirm } = req.body || {};
    templates = saveTemplates({ broadcast, confirm });
    res.json({ ok: true, templates });
  } catch (e) {
    console.error(e);
    res.status(400).json({ ok: false, error: e.message });
  }
});

// Serve the Format page
app.get('/format', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'format.html'));
});

