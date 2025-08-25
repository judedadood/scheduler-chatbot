# Prudential WhatsApp Scheduler Bot

Schedule client meetings over WhatsApp by broadcasting available 1-hour slots and updating an Excel file with bookings.

## Features
- Upload Excel with columns: **Client Name**, **Contact Number**
- Paste availability lines (e.g., `25 Aug 1-5pm`); app expands to 1-hour slots
- Broadcast numbered slots via WhatsApp (Twilio)
- Clients reply with a number to book (first-come-first-served)
- Excel is updated: Booked Date, Booked Time, Status=Confirmed
- Download the updated Excel anytime
- Minimal web UI provided (index.html)

## Quick Start

```bash
npm i
cp .env.example .env
# fill .env with Twilio credentials
node server.js
```

Open http://localhost:${PORT:-3000}

### Endpoints
- `GET /` UI
- `POST /upload-clients` multipart form with `file` (.xlsx)
- `POST /set-availability` JSON `{ availabilityText: "25 Aug 1-5pm\n26 Aug 2-7pm" }`
- `POST /broadcast`
- `POST /whatsapp/inbound` (Twilio webhook)
- `GET /download-latest`

## Twilio / WhatsApp
- Use Twilio Sandbox or a WhatsApp-approved sender.
- Set the **inbound webhook** in Twilio to `https://<your-domain>/whatsapp/inbound`.
- For local tests, expose with ngrok: `ngrok http 3000`.
- Recipients must be opted-in per WhatsApp policy.

## Excel format
Input sheet (first sheet):
- Client Name
- Contact Number

App adds (if missing):
- Booked Date
- Booked Time
- Status

Row matching is by phone (rightmost digits).

## Deployment
See `Dockerfile` or deploy to Render/Railway with environment variables set.

## Notes
- Timezone assumed Singapore; current year inferred.
- Slot granularity is 60 minutes (change in `expandToHourlySlots`).

