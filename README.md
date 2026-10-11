# ⚡ Meter Barcode Printer — Zebra ZT411

A web-based barcode sticker generator and direct print manager designed for **Zebra ZT411** (and compatible thermal printers) for water & electricity meter labels.

🌐 **Live Web Application (GitHub Pages):**  
👉 **[https://ectronom-create.github.io/Meter-Barcode-Printer/](https://ectronom-create.github.io/Meter-Barcode-Printer/)**

---

## 🚀 Features

- **Single Label & Bulk Batch Printing:** Print individual meter labels or batches from imported spreadsheets.
- **Excel Import (`.xlsx`):** Automatically maps:
  - `Device_pallet_number`
  - `Device_box_number`
  - `Serial_number`
  - `Module_ICCID`
- **Dual Printing Modes:**
  1. **Direct RAW ZPL Printing (Fast):** Sends ZPL commands directly to the Zebra ZT411 via Windows Print Spooler with no dialog boxes.
  2. **Browser Print Mode (Fallback / Standalone):** Formats labels precisely (`50mm × 25mm`) and triggers the standard browser print dialog. Works on any browser and device.
- **Pallet Header Stickers:** Generates pallet summary cards with total label counts.
- **Live Label Preview & Fine-tuning:** Customize offsets, label height/width, barcode narrow bar width, and font sizes directly from the UI.

---

## 🛠️ How to Use

### Option 1: Directly in the Browser (GitHub Pages)
1. Open the [Live Web App](https://ectronom-create.github.io/Meter-Barcode-Printer/).
2. Import your Excel file or enter a Serial Number & ICCID manually.
3. Click **Print Label** or **Print All** — your browser's print dialog will open with the exact label dimensions.

---

### Option 2: Direct Zebra Hardware Print (Windows)

For ultra-fast printing without browser dialog popups:

1. Clone or download this repository:
   ```bash
   git clone https://github.com/ectronom-create/Meter-Barcode-Printer.git
   cd Meter-Barcode-Printer
   ```
2. Double-click **`run.bat`** (or run `python server.py` in your terminal).
3. The app will launch at `http://localhost:5000` with the status indicator showing:  
   `🟢 Zebra ZT411 Ready (Direct RAW)`.
4. Printing will immediately send RAW ZPL instructions directly to your Zebra ZT411 printer.
