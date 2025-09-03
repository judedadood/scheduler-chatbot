// server.js
//  WhatsApp Scheduler Bot (Workspace edition)
// Tech: Node.js, Express, Twilio WhatsApp, ExcelJS, Multer

require('dotenv').config();
const express = require('express');
const bodyParser = require('body-parser');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const { Twilio } = require('twilio');

// ---------------------- App + Static ----------------------
const app = express();
app.use(bodyParser.urlencoded({ extended: true })); // Twilio posts x-www-form-urlencoded
app.use(bodyParser.json());
app.use(express.static(path.join(__dirname, 'public')));

const DATA_DIR = process.env.DATA_DIR || __dirname;
const APP_DATA_DIR = path.join(DATA_DIR, 'appdata'); // where workspaces live
if (!fs.existsSync(APP_DATA_DIR)) fs.mkdirSync(APP_DATA_DIR, { recursive: true });

const globalUploadDir = path.join(APP_DATA_DIR, 'uploads_tmp');
if (!fs.existsSync(globalUploadDir)) fs.mkdirSync(globalUploadDir, { recursive: true });

const makeStorage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, globalUploadDir),
  filename: (req, file, cb) => cb(null, `upload_${Date.now()}${path.extname(file.originalname)}`)
});
const upload = multer({ storage: makeStorage });

// ---------------------- Twilio ----------------------
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

async function sendWa(to, body, mediaUrl) {
  if (!twilioClient) throw new Error('Twilio client not configured.');
  const msg = { from: TWILIO_WHATSAPP_NUMBER, to, body };
  if (mediaUrl) msg.mediaUrl = [mediaUrl];
  return twilioClient.messages.create(msg);
}

// ---------------------- Time & Phone helpers ----------------------
const pad2 = n => String(n).padStart(2, '0');
const phoneDigitsOnly = v => (v || '').toString().replace(/[^\d]/g, '');
function waFormat(numberRaw) {
  const digits = phoneDigitsOnly(numberRaw);
  if (!digits) return null;
  const withCC = digits.startsWith('65') ? `+${digits}` : `+65${digits}`;
  return `whatsapp:${withCC}`;
}
function sgtStamp() {
  // e.g., "2025-09-03 14:23:11 SGT"
  const s = new Date().toLocaleString('en-GB', {
    timeZone: 'Asia/Singapore',
    year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit',
    hour12: false
  }).replace(',', '');
  return `${s} SGT`;
}

// ---------------------- Availability & formatting ----------------------
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

function expandToBufferedSlots(start, end, slotMins = 60, bufferMins = 0) {
  const slots = [];
  const step = slotMins + bufferMins;
  let cursor = new Date(start);
  while (cursor < end) {
    const slotEnd = new Date(cursor);
    slotEnd.setMinutes(slotEnd.getMinutes() + slotMins);
    if (slotEnd > end) break;
    slots.push({ start: new Date(cursor), end: slotEnd });
    cursor.setMinutes(cursor.getMinutes() + step);
  }
  return slots;
}

// ---------------------- Template helpers ----------------------
const defaultTemplates = {
  broadcast:
    "Hi {{client.name}}, here are my available 1-hour meeting slots:\n\n{{slotsText}}\n\nReply with the number of your preferred slot (e.g., 2).",
  confirm:
    "📌 Hi {{client.name}}, your appointment is confirmed.\n\n🗓 {{slotLabel}}\n\n– Your Agent"
};
function renderTemplate(tpl, data) {
  return (tpl || '').replace(/\{\{\s*([a-zA-Z0-9_.]+)\s*\}\}/g, (_, path) => {
    const parts = path.split('.');
    let cur = data;
    for (const p of parts) {
      if (cur && Object.prototype.hasOwnProperty.call(cur, p)) cur = cur[p];
      else return `{{${path}}}`;
    }
    return (cur ?? `{{${path}}}`).toString();
  });
}

// ---------------------- Workspaces ----------------------
/**
 * Each workspace:
 * {
 *   id, name, baseDir, uploadDir, exportDir,
 *   excelState: { filePath, workbook, sheet, headerMap },
 *   clientsByWa: Map('whatsapp:+...' -> { name, phone, rowIndex, status, lastNotified }),
 *   statusByDigits: Map(digits -> { confirmed, pending, notified, rowIndices }),
 *   availabilitySlots: [{id,start,end,label,booked,bookedBy}],
 *   lastBroadcastOrder: [slotId, ...],
 *   templatesPath, templates: { broadcast, confirm }
 * }
 */
const workspaces = new Map();
// global map from WA number to workspace id for inbound matching
const waToWs = new Map();

function makeWorkspaceDirs(id) {
  const baseDir = path.join(APP_DATA_DIR, 'workspaces', id);
  const uploadDir = path.join(baseDir, 'uploads');
  const exportDir = path.join(baseDir, 'exports');
  if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir, { recursive: true });
  if (!fs.existsSync(exportDir)) fs.mkdirSync(exportDir, { recursive: true });
  return { baseDir, uploadDir, exportDir };
}

function loadTemplates(ws) {
  const fp = path.join(ws.exportDir, 'templates.json');
  ws.templatesPath = fp;
  try {
    if (fs.existsSync(fp)) {
      const raw = JSON.parse(fs.readFileSync(fp, 'utf8'));
      ws.templates = {
        broadcast: typeof raw.broadcast === 'string' && raw.broadcast.trim() ? raw.broadcast : defaultTemplates.broadcast,
        confirm:   typeof raw.confirm   === 'string' && raw.confirm.trim()   ? raw.confirm   : defaultTemplates.confirm
      };
      return;
    }
  } catch (e) {
    console.error(`Failed to load templates for ${ws.id}:`, e.message);
  }
  ws.templates = { ...defaultTemplates };
}

function saveTemplates(ws, t) {
  const clean = {
    broadcast: typeof t.broadcast === 'string' && t.broadcast.trim() ? t.broadcast : defaultTemplates.broadcast,
    confirm:   typeof t.confirm   === 'string' && t.confirm.trim()   ? t.confirm   : defaultTemplates.confirm
  };
  fs.writeFileSync(ws.templatesPath, JSON.stringify(clean, null, 2));
  ws.templates = clean;
  return clean;
}

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
    headerRow.values = [, ...requiredHeaders];
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

async function saveExcel(ws) {
  if (!ws.excelState) return;
  await ws.excelState.workbook.xlsx.writeFile(ws.excelState.filePath);
}

function buildClientMaps(ws) {
  ws.clientsByWa = new Map();
  ws.statusByDigits = new Map();

  const { sheet, headerMap } = ws.excelState;
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
      ws.clientsByWa.set(wa, { name, phone: phoneRaw, rowIndex: r, status, lastNotified });
      waToWs.set(wa, ws.id); // global mapping for inbound handling
    }

    if (digits) {
      let agg = ws.statusByDigits.get(digits);
      if (!agg) agg = { confirmed: false, pending: false, notified: false, rowIndices: [] };
      if (status === 'confirmed') agg.confirmed = true;
      if (status === 'pending')   agg.pending = true;
      if (lastNotified)           agg.notified = true;
      agg.rowIndices.push(r);
      ws.statusByDigits.set(digits, agg);
    }
  }
}

function findRowByPhone(ws, phoneDigitsRaw) {
  const sheet = ws.excelState.sheet;
  const cIdx = ws.excelState.headerMap['Contact Number'];
  const phoneDigits = phoneDigitsOnly(phoneDigitsRaw);
  for (let r = 2; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const val = (row.getCell(cIdx).value || '').toString();
    const digits = phoneDigitsOnly(val);
    if (digits && phoneDigits && digits.endsWith(phoneDigits)) return row;
  }
  return null;
}

function refreshStatusForDigits(ws, digits) {
  if (!ws.excelState) return;
  const sheet = ws.excelState.sheet;
  const h = ws.excelState.headerMap;

  let confirmed = false, pending = false, notified = false;
  const rowIndices = [];

  for (let r = 2; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const val = (row.getCell(h['Contact Number']).value || '').toString();
    const d = phoneDigitsOnly(val);
    if (!d || !digits) continue;
    if (!d.endsWith(digits)) continue;

    rowIndices.push(r);
    const st = String(row.getCell(h['Status']).value || '').trim().toLowerCase();
    if (st === 'confirmed') confirmed = true;
    if (st === 'pending')   pending = true;
    if (row.getCell(h['Last Notified']).value) notified = true;
  }
  ws.statusByDigits.set(digits, { confirmed, pending, notified, rowIndices });
}


function listSlotsForMessage(ws) {
  // Default list from current open slots (1..n)
  const lines = ws.availabilitySlots
    .filter(s => !s.booked)
    .map((s, i) => `${i + 1}) ${s.label}`);
  return lines.length ? lines.join('\n') : '(All slots have been booked)';
}

function listSlotsStable(ws) {
  // Use lastBroadcastOrder to preserve numbering; show only available ones
  if (!ws.lastBroadcastOrder || !ws.lastBroadcastOrder.length) {
    return listSlotsForMessage(ws);
  }
  const idToSlot = new Map(ws.availabilitySlots.map(s => [s.id, s]));
  const lines = [];
  ws.lastBroadcastOrder.forEach((slotId, index) => {
    const slot = idToSlot.get(slotId);
    if (slot && !slot.booked) {
      lines.push(`${index + 1}) ${slot.label}`);
    }
  });
  return lines.length ? lines.join('\n') : '(All slots have been booked)';
}

// ---------------------- Workspace helpers ----------------------
function requireWS(req, res) {
  const wsId = req.params.ws;
  const ws = workspaces.get(wsId);
  if (!ws) {
    res.status(404).json({ ok: false, error: 'Workspace not found' });
    return null;
  }
  return ws;
}

// ---------------------- Pages (workspaces) ----------------------
app.get('/workspaces', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'workspaces.html'));
});
app.get('/w/:ws/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});
app.get('/w/:ws/format', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'format.html'));
});
app.get('/w/:ws/excel', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'excel.html'));
});
app.get('/w/:ws/followup', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'followup.html'));
});

// ---------------------- API: workspaces ----------------------
app.get('/api/workspaces', (req, res) => {
  const items = [];
  for (const ws of workspaces.values()) {
    items.push({
      id: ws.id,
      name: ws.name || `Workspace ${ws.id}`,
      workbookName: ws.excelState ? path.basename(ws.excelState.filePath) : null,
      totalClients: ws.clientsByWa ? ws.clientsByWa.size : 0
    });
  }
  res.json({ ok: true, workspaces: items });
});

// Create workspace by uploading an Excel
app.post('/api/workspaces', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) throw new Error('No Excel uploaded');

    const id = Math.random().toString(36).slice(2, 8);
    const { baseDir, uploadDir, exportDir } = makeWorkspaceDirs(id);
    const ws = {
      id, name: req.file.originalname, baseDir, uploadDir, exportDir,
      excelState: null,
      clientsByWa: new Map(),
      statusByDigits: new Map(),
      availabilitySlots: [],
      lastBroadcastOrder: [],
      templatesPath: '',
      templates: { ...defaultTemplates }
    };

    // Move uploaded file into workspace uploadDir
    const dest = path.join(uploadDir, `clients_${Date.now()}${path.extname(req.file.originalname)}`);
    fs.renameSync(req.file.path, dest);

    // Load Excel
    const loaded = await loadExcel(dest);
    ws.excelState = { filePath: dest, ...loaded };
    buildClientMaps(ws);

    // Templates
    loadTemplates(ws);

    workspaces.set(id, ws);

    res.json({ ok: true, id, workbookName: path.basename(dest), totalClients: ws.clientsByWa.size });
  } catch (err) {
    console.error(err);
    res.status(400).json({ ok: false, error: err.message });
  }
});

// Delete a workspace
app.delete('/api/workspaces/:id', (req, res) => {
  try {
    const id = req.params.id;
    const ws = workspaces.get(id);
    if (!ws) {
      return res.status(404).json({ ok: false, error: 'Workspace not found' });
    }

    // Remove WA number → workspace mapping
    if (ws.clientsByWa) {
      for (const wa of ws.clientsByWa.keys()) {
        waToWs.delete(wa);
      }
    }

    // Remove from memory
    workspaces.delete(id);

    // Delete files on disk
    try {
      fs.rmSync(ws.baseDir, { recursive: true, force: true });
    } catch (e) {
      // Non-fatal if directory already gone
      console.warn(`Failed to rm ${ws.baseDir}:`, e.message);
    }

    res.json({ ok: true });
  } catch (err) {
    console.error('Delete workspace error:', err);
    res.status(500).json({ ok: false, error: err.message });
  }
});


// ---------------------- API: per-workspace info ----------------------
app.get('/api/w/:ws/info', (req, res) => {
  const ws = requireWS(req, res); if (!ws) return;
  const hasExcel = !!ws.excelState;
  const workbookName = hasExcel ? path.basename(ws.excelState.filePath) : null;
  res.json({
    ok: true,
    hasExcel,
    workbookName,
    totalClients: hasExcel ? ws.clientsByWa.size : 0
  });
});

// Replace/Upload Excel for existing workspace (optional)
app.post('/api/w/:ws/upload-clients', upload.single('file'), async (req, res) => {
  const ws = requireWS(req, res); if (!ws) return;
  try {
    if (!req.file) throw new Error('No Excel uploaded');
    const dest = path.join(ws.uploadDir, `clients_${Date.now()}${path.extname(req.file.originalname)}`);
    fs.renameSync(req.file.path, dest);

    const loaded = await loadExcel(dest);
    ws.excelState = { filePath: dest, ...loaded };
    buildClientMaps(ws);

    res.json({ ok: true, message: 'Excel loaded', filePath: dest, totalClients: ws.clientsByWa.size });
  } catch (err) {
    console.error(err);
    res.status(400).json({ ok: false, error: err.message });
  }
});

// ---------------------- API: availability ----------------------
app.post('/api/w/:ws/set-availability', (req, res) => {
  const ws = requireWS(req, res); if (!ws) return;
  try {
    const { availabilityText, bufferMinutes } = req.body || {};
    if (!availabilityText) return res.status(400).json({ ok:false, error:'availabilityText is required' });

    const allowed = new Set([0, 30, 60, '0', '30', '60']);
    const bufRaw = bufferMinutes ?? 0;
    const buf = parseInt(bufRaw, 10);
    const buffer = allowed.has(bufRaw) && [0,30,60].includes(buf) ? buf : 0;

    ws.availabilitySlots = [];
    const lines = availabilityText.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
    for (const line of lines) {
      const parsed = parseAvailabilityLine(line);
      if (!parsed) continue;
      const slots = expandToBufferedSlots(parsed.start, parsed.end, 60, buffer);
      for (const h of slots) {
        ws.availabilitySlots.push({
          id: '',
          start: h.start,
          end: h.end,
          label: humanSlotLabel(h.start, h.end),
          booked: false
        });
      }
    }
    ws.availabilitySlots.sort((a,b) => a.start - b.start);
    ws.availabilitySlots.forEach((s, idx) => s.id = `S${idx + 1}`);
    // Reset broadcast order until next broadcast
    ws.lastBroadcastOrder = [];

    res.json({
      ok: true,
      bufferMinutes: buffer,
      totalSlots: ws.availabilitySlots.length,
      slots: ws.availabilitySlots.map(s => s.label)
    });
  } catch (err) {
    console.error(err);
    res.status(400).json({ ok:false, error: err.message });
  }
});

// ---------------------- API: broadcast ----------------------
app.post('/api/w/:ws/broadcast', async (req, res) => {
  const ws = requireWS(req, res); if (!ws) return;
  try {
    if (!ws.excelState) return res.status(400).json({ ok:false, error:'Upload an Excel first' });
    if (!ws.availabilitySlots.length) return res.status(400).json({ ok:false, error:'Set availability first' });

    const slotsText = listSlotsForMessage(ws);
    // Freeze numbering (snapshot of open slots by id)
    ws.lastBroadcastOrder = ws.availabilitySlots.filter(s => !s.booked).map(s => s.id);

    const force = (req.query.force || '').toString().toLowerCase() === 'true';
    const toSend = [];
    const usedDigits = new Set();

    for (const [wa, client] of ws.clientsByWa.entries()) {
      const digits = phoneDigitsOnly(client.phone);
      if (!digits || usedDigits.has(digits)) continue;
      const agg = ws.statusByDigits.get(digits) || { confirmed:false, pending:false, notified:false, rowIndices:[] };
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
        skipped: ws.clientsByWa.size,
        reason: force ? 'No eligible numbers' : 'All numbers have Pending or Confirmed status (or were duplicates).'
      });
    }

    const whenSGT = sgtStamp();

    const tasks = toSend.map(async ({ wa, digits, client, rowIndices }) => {
      const body = renderTemplate(ws.templates.broadcast, {
        client: { name: client.name },
        slotsText
      });
      await sendWa(wa, body);

      // Mark Last Notified (SGT) and set Status='Pending' if not Confirmed
      const h = ws.excelState.headerMap;
      for (const r of rowIndices) {
        const row = ws.excelState.sheet.getRow(r);
        row.getCell(h['Last Notified']).value = whenSGT;
        const sIdx = h['Status'];
        const cur = (row.getCell(sIdx).value || '').toString().trim().toLowerCase();
        if (cur !== 'confirmed') row.getCell(sIdx).value = 'Pending';
        row.commit();
      }

      // Update in-memory aggregator
      const agg = ws.statusByDigits.get(digits) || { confirmed:false, pending:false, notified:false, rowIndices: rowIndices.slice() };
      agg.pending = true; agg.notified = true;
      ws.statusByDigits.set(digits, agg);
      client.status = 'pending';
    });

    await Promise.all(tasks);
    await saveExcel(ws);

    res.json({ ok:true, sentTo: toSend.length, skipped: ws.clientsByWa.size - toSend.length });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok:false, error: err.message });
  }
});

// ---------------------- API: follow-up broadcast (Pending only) ----------------------
app.post('/api/w/:ws/followup-broadcast', async (req, res) => {
  const ws = requireWS(req, res); if (!ws) return;
  try {
    const { template } = req.body || {};
    if (!template || !template.trim()) return res.status(400).json({ ok:false, error:'template is required' });

    const slotsText = listSlotsStable(ws);
    const h = ws.excelState.headerMap;

    // Build recipients strictly from Excel where Status == 'pending'
    const toSend = [];
    for (const [wa, client] of ws.clientsByWa.entries()) {
      const row = findRowByPhone(ws, client.phone);
      if (!row) continue;
      const st = String(row.getCell(h['Status']).value || '').trim().toLowerCase();
      if (st === 'pending') {
        toSend.push({ wa, client });
      }
    }

    if (!toSend.length) {
      return res.json({ ok:true, sentTo: 0, skipped: ws.clientsByWa.size });
    }

    const whenSGT = sgtStamp();

    const tasks = toSend.map(async ({ wa, client }) => {
      const body = renderTemplate(template, {
        client: { name: client.name },
        slotsText
      });
      await sendWa(wa, body);

      // Update Last Notified (SGT) and keep Status as Pending
      const row = findRowByPhone(ws, client.phone);
      if (row) {
        row.getCell(h['Last Notified']).value = whenSGT;
        const cur = (row.getCell(h['Status']).value || '').toString().trim().toLowerCase();
        if (cur !== 'confirmed') row.getCell(h['Status']).value = 'Pending';
        row.commit();
      }

      // keep aggregator consistent (optional)
      const digits = phoneDigitsOnly(client.phone);
      if (digits) refreshStatusForDigits(ws, digits);
    });

    await Promise.all(tasks);
    await saveExcel(ws);

    res.json({ ok:true, sentTo: toSend.length, skipped: ws.clientsByWa.size - toSend.length });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok:false, error: err.message });
  }
});


// ---------------------- API: templates per workspace ----------------------
app.get('/api/w/:ws/templates', (req, res) => {
  const ws = requireWS(req, res); if (!ws) return;
  if (!ws.templates) loadTemplates(ws);
  res.json({ ok: true, templates: ws.templates });
});

app.post('/api/w/:ws/templates', (req, res) => {
  const ws = requireWS(req, res); if (!ws) return;
  try {
    const { broadcast, confirm } = req.body || {};
    const out = saveTemplates(ws, { broadcast, confirm });
    res.json({ ok: true, templates: out });
  } catch (e) {
    console.error(e);
    res.status(400).json({ ok: false, error: e.message });
  }
});

// ---------------------- API: Excel view/download ----------------------
app.get('/api/w/:ws/excel', (req, res) => {
  const ws = requireWS(req, res); if (!ws) return;
  try {
    if (!ws.excelState) return res.json({ ok: true, headers: [], rows: [] });
    const sheet = ws.excelState.sheet;

    const headers = [];
    const headerRow = sheet.getRow(1);
    for (let i = 1; i <= headerRow.cellCount; i++) {
      const v = headerRow.getCell(i).value;
      if (v) headers.push(String(v));
    }

    const rows = [];
    for (let r = 2; r <= sheet.rowCount; r++) {
      const row = sheet.getRow(r);
      const item = {};
      headers.forEach((h, i) => {
        const v = row.getCell(i + 1).value;
        item[h] = (v == null) ? '' : v.toString();
      });
      // Exclude completely empty rows
      const hasData = Object.values(item).some(v => String(v).trim() !== '');
      if (hasData) rows.push(item);
    }

    res.json({ ok: true, headers, rows });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok:false, error: err.message });
  }
});

app.get('/api/w/:ws/download-latest', (req, res) => {
  const ws = requireWS(req, res); if (!ws) return;
  try {
    if (!ws.excelState) return res.status(400).json({ ok:false, error:'No Excel loaded yet.' });
    const outPath = path.join(ws.exportDir, `bookings_${Date.now()}.xlsx`);
    fs.copyFileSync(ws.excelState.filePath, outPath);
    res.download(outPath, 'bookings.xlsx');
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok:false, error: err.message });
  }
});

// ---------------------- Inbound WhatsApp webhook ----------------------
const INVALID_INPUT_MSG =
  'Please reply with the number of your preferred slot (e.g., 2).';

app.post('/whatsapp/inbound', async (req, res) => {
  try {
    const from = req.body.From;          // 'whatsapp:+65...'
    const text = (req.body.Body || '').toString().trim();

    // Acknowledge Twilio immediately
    res.status(200).send('OK');

    if (!from) return;

    // Identify workspace by WA number
    let wsId = waToWs.get(from);
    let ws = wsId ? workspaces.get(wsId) : null;

    // Fallback: scan workspaces (in case mapping was lost)
    if (!ws) {
      for (const w of workspaces.values()) {
        if (w.clientsByWa && w.clientsByWa.has(from)) { ws = w; waToWs.set(from, w.id); break; }
      }
    }
    if (!ws) return; // unknown sender across all workspaces

    // Quick commands to re-show menu
    if (/\b(menu|slots|options|list)\b/i.test(text)) {
      const slotsText = listSlotsStable(ws);
      await sendWa(from, `Here are the available slots:\n\n${slotsText}\n\n${INVALID_INPUT_MSG}`);
      return;
    }

    // Require numeric choice 1-3 digits
    const m = text.match(/\b(\d{1,3})\b/);
    if (!m) {
      await sendWa(from, INVALID_INPUT_MSG);
      return;
    }

    const idx = parseInt(m[1], 10) - 1;
    if (!Number.isInteger(idx) || idx < 0) {
      await sendWa(from, INVALID_INPUT_MSG);
      return;
    }

    // Prevent double-booking if already confirmed
    const client = ws.clientsByWa.get(from);
    const rowExisting = client ? findRowByPhone(ws, client.phone) : null;
    if (rowExisting) {
      const h = ws.excelState.headerMap;
      const status = String(rowExisting.getCell(h['Status']).value || '').toLowerCase();
      const bookedDate = rowExisting.getCell(h['Booked Date']).value || '';
      const bookedTime = rowExisting.getCell(h['Booked Time']).value || '';
      if (status === 'confirmed' && (bookedDate || bookedTime)) {
        await sendWa(from, `You already have a confirmed appointment: ${bookedDate} ${bookedTime}.`);
        return;
      }
    }

    // Resolve slot using stable broadcast order if available
    let slot = null;
    if (ws.lastBroadcastOrder && ws.lastBroadcastOrder.length >= (idx + 1)) {
      const slotId = ws.lastBroadcastOrder[idx];
      slot = ws.availabilitySlots.find(s => s.id === slotId);
    }
    // Fallback to current open slots
    if (!slot) {
      const openSlots = ws.availabilitySlots.filter(s => !s.booked);
      if (idx >= openSlots.length) {
        await sendWa(from, `That slot number is no longer available.\n\n${INVALID_INPUT_MSG}`);
        return;
      }
      slot = openSlots[idx];
    }

    if (!slot || slot.booked) {
      await sendWa(from, `Sorry, that slot was just taken.\n\n${INVALID_INPUT_MSG}`);
      return;
    }

    // Book it
    slot.booked = true;
    slot.bookedBy = from;

    // Update Excel
    const dateOnly = `${pad2(slot.start.getDate())} ${['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][slot.start.getMonth()]}`;
    const timeLabel = humanSlotLabel(slot.start, slot.end).split(' ').slice(2).join(' ');
    const row = client ? findRowByPhone(ws, client.phone) : null;

    if (row) {
      const { headerMap } = ws.excelState;
      row.getCell(headerMap['Booked Date']).value = dateOnly;
      row.getCell(headerMap['Booked Time']).value = timeLabel;
      row.getCell(headerMap['Status']).value = 'Confirmed';
      row.commit();
      await saveExcel(ws);
    }
    
    // ✅ Sync in-memory aggregator & map after saving Excel
    const digits = client ? phoneDigitsOnly(client.phone) : '';
    if (digits) {
      // Prefer to mirror Excel exactly:
      refreshStatusForDigits(ws, digits);
      // (or, minimally)
      // const agg = ws.statusByDigits.get(digits) || { confirmed:false, pending:false, notified:false, rowIndices: [] };
      // agg.confirmed = true; agg.pending = false;
      // ws.statusByDigits.set(digits, agg);
    }
    if (client) client.status = 'confirmed';

    // Send confirmation using workspace template
    const slotLabel = humanSlotLabel(slot.start, slot.end);
    const body = renderTemplate(ws.templates.confirm, {
      client: { name: client ? client.name : '' },
      slotLabel
    });
    await sendWa(from, body);
  } catch (err) {
    console.error('Inbound handler error:', err);
  }
});

// ---------------------- Health ----------------------
app.get('/health', (req, res) => res.json({ ok:true }));

// ---------------------- Start ----------------------
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server listening on port ${PORT}`);
});
