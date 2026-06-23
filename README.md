# Purchasing Department Tool
IT Intern Project — PO Scanner & Monitoring Dashboard

## Quick Start

```bash
# 1. Open this folder in VS Code terminal
cd purchasing-tool

# 2. Install dependencies (one time only)
npm install (on work pc if this doesnt work try cmd /c npm install)

# 3. Start the server
node server.js
```

The terminal will print two URLs:
```
✅ Purchasing Tool running!
   Local:   http://localhost:3000       ← your computer
   Network: http://192.168.x.x:3000    ← phones on same WiFi
```

Open the **Network** URL on your phone to use the scanner.

---

## File Overview

| File | Purpose |
|------|---------|
| `server.js` | Express backend — serves files, reads/writes Excel |
| `public/index.html` | Home screen (two buttons) |
| `public/scanner.html` | PO Scanner page |
| `public/scanner.js` | Camera feed + Tesseract OCR logic |
| `public/dashboard.html` | Data visualization page |
| `public/charts.js` | Chart.js chart rendering |
| `public/style.css` | All shared styles |
| `data/monitoring.xlsx` | The Excel monitoring file |

---

## Using an Existing Excel File

If you already have a `monitoring.xlsx`:

1. Copy it into the `data/` folder
2. Make sure your sheet names match: `JAN`, `FEB`, ... `DEC`, `SUMMARY`, `COST SAVINGS`
3. Make sure your PO column header is exactly: `PO NUMBER` (all caps)

The server auto-creates a blank workbook if none exists.

---

5. The blue dashed rectangles on the scanner preview show your current zones

**Pro tip**: Use the Upload tab to test with a photo first before trying live camera.

---

## Tips for Better OCR Accuracy

- **Lighting**: Scan under even, bright light. Avoid glare.
- **Angle**: Hold the phone directly above the document (not at an angle)
- **Distance**: Fill ~80% of the frame with the PO
- **Stability**: Keep still while the scan line animates
- **Font**: Tesseract works best on clean printed text (your template should be ideal)

---

## Excel Column Reference

Each monthly sheet uses these columns:

```
PR DATE | PR DATE RECEIVED | PR NO. | REQUESTING DEPT. | PO DATE |
PO NUMBER | END USER/S | SUPPLIER'S NAME | ITEM CODE | ITEM DESCRIPTION |
SPECIFICATIONS | QTY | UoM | UNIT PRICE | AMOUNT | TOTAL AMOUNT |
PR REQUIRED DATE | DATE DELIVERED | REMARKS | PURCHASE ORDER STATUS |
ITEMS/SERVICES | ORIGINAL PRICE | TOTAL COST SAVINGS
```

---

## Accessing from Phones (Office WiFi)

1. Connect your phone to the **same WiFi** as the office computer
2. Find the computer's local IP: `ipconfig` (Windows) or `ifconfig` (Mac/Linux)
3. Go to `http://[IP]:3000` on the phone browser
4. For best camera experience, use **Chrome on Android** or **Safari on iPhone**

> Note: Camera features require HTTPS on some iOS versions. If camera doesn't work on iPhone,
> you may need to set up a self-signed certificate or use the **Upload** tab instead.

---

## Development

```bash
# Auto-restart server on file changes
npm run dev   # uses nodemon
```

---
## Sharing Your Tool Outside the Office WiFi (ngrok)

If you need to access the dashboard/scanner from outside the office network (e.g. testing from home, sharing with someone off-site), you can use ngrok to create a temporary public URL.

**Before doing anything below, make sure your local server is already running** (`node server.js`, with the ✅ message showing). ngrok just tunnels to your local server — it doesn't start it for you.

1. Make sure your server is running locally first (see Quick Start above)
2. In a **separate terminal**, run:
```bash
   ngrok http 3000
```
3. ngrok will print a public URL like `https://abcd1234.ngrok-free.app` — this is what you share
4. Keep both terminals open (server + ngrok) for as long as you need the link to work
5. Closing either terminal kills the connection

---
## Troubleshooting

### ngrok: "connection actively refused" error
Traffic successfully made it to the ngrok agent, but the agent failed to establish

a connection to the upstream web service at http://localhost:3000.

dial tcp [::1]:3000: connectex: No connection could be made because the target

machine actively refused it.
This means ngrok itself is working fine — it just can't find anything running on port 3000 to connect to. Almost always this means **the server isn't actually running yet**. Fix:
1. Check that `node server.js` is running and you see the ✅ message *before* you start ngrok
2. If it's running and you still get this error, try forcing IPv4 instead of `localhost`:
```bash
   ngrok http 127.0.0.1:3000
```
3. Double check nothing else is occupying port 3000:
```bash
   netstat -ano | findstr :3000
```
