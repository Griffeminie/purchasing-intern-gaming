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

## OCR Zone Calibration ⚠️ IMPORTANT

The scanner uses **fixed zones** (percentage positions) to know where to look on the PO image.
You **must calibrate these** to match your actual PO template.

### How to calibrate:

1. Take a photo of a real PO and open it in any image editor
2. Note the pixel coordinates of:
   - **PO Number** (top right corner)
   - **Supplier Name** (top left area)
   - **Items table** (middle section)
3. Convert to 0–1 fractions: `x / imageWidth`, `y / imageHeight`
4. Edit `ZONES` in `public/scanner.js`:

```javascript
const ZONES = {
  poNumber: { top: 0.03, left: 0.76, width: 0.22, height: 0.07 },
  supplier: { top: 0.03, left: 0.02, width: 0.30, height: 0.07 },
  items:    { top: 0.22, left: 0.02, width: 0.96, height: 0.45 },
};
```

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
