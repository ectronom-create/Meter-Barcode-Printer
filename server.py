import os
import sys
import json
import ctypes
from ctypes import wintypes
from http.server import HTTPServer, SimpleHTTPRequestHandler

# Windows Print Spooler Setup
class DOC_INFO_1(ctypes.Structure):
    _fields_ = [
        ("pDocName", wintypes.LPCWSTR),
        ("pOutputFile", wintypes.LPCWSTR),
        ("pDatatype", wintypes.LPCWSTR)
    ]

# ──────────────────────────────────────────────────────────────
# Unit conversion  (203 DPI printer)
# ──────────────────────────────────────────────────────────────
MM_TO_DOTS = 203 / 25.4   # ≈ 7.992 dots per mm

def mm_to_dots(mm, minimum=8):
    return max(minimum, round(float(mm) * MM_TO_DOTS))

def get_zebra_printer_name():
    """Detect Zebra printer installed on Windows"""
    winspool = ctypes.WinDLL('winspool.drv')
    needed = wintypes.DWORD()
    returned = wintypes.DWORD()
    winspool.EnumPrintersW(2, None, 1, None, 0, ctypes.byref(needed), ctypes.byref(returned))
    
    buf = ctypes.create_string_buffer(needed.value)
    if winspool.EnumPrintersW(2, None, 1, buf, needed.value, ctypes.byref(needed), ctypes.byref(returned)):
        # Inspect printers
        pass
    
    # Preferred default
    default_name = "ZDesigner ZT411-203dpi ZPL"
    return default_name

def print_zpl_raw(printer_name, zpl_string):
    """Sends RAW ZPL code directly to Zebra printer via Windows spooler"""
    winspool = ctypes.WinDLL('winspool.drv')
    hPrinter = wintypes.HANDLE()
    
    if not winspool.OpenPrinterW(printer_name, ctypes.byref(hPrinter), None):
        raise Exception(f"تعذر فتح الطابعة {printer_name}. تأكد من توصيلها وتشغيلها.")
        
    try:
        doc_info = DOC_INFO_1("Zebra_Label_Job", None, "RAW")
        job_id = winspool.StartDocPrinterW(hPrinter, 1, ctypes.byref(doc_info))
        if not job_id:
            raise Exception("تعذر بدء أمر الطباعة في نظام ويندوز.")
            
        try:
            if not winspool.StartPagePrinter(hPrinter):
                raise Exception("تعذر بدء صفحة الطباعة.")
            
            raw_bytes = zpl_string.encode('utf-8')
            written = wintypes.DWORD()
            if not winspool.WritePrinter(hPrinter, raw_bytes, len(raw_bytes), ctypes.byref(written)):
                raise Exception("فشل إرسال بيانات الباركود إلى الطابعة.")
                
            winspool.EndPagePrinter(hPrinter)
        finally:
            winspool.EndDocPrinter(hPrinter)
    finally:
        winspool.ClosePrinter(hPrinter)

def estimate_code128_modules(text):
    """Calculates symbol modules for Code 128 Auto (^BC ... ,,,A)"""
    clean = str(text).strip()
    modules = 11 + 11 + 13  # Start + Check + Stop
    i = 0
    in_code_c = False
    while i < len(clean):
        digits_count = 0
        while i + digits_count < len(clean) and clean[i + digits_count].isdigit():
            digits_count += 1
        if digits_count >= 4 or (digits_count >= 2 and in_code_c):
            if not in_code_c:
                modules += 11  # Switch to Code C
                in_code_c = True
            pairs = digits_count // 2
            modules += pairs * 11
            i += pairs * 2
        else:
            if in_code_c:
                modules += 11  # Switch to Code B
                in_code_c = False
            modules += 11
            i += 1
    return modules

def generate_zpl(iccid, meter_sn, box_suffix='', s=None):
    """Generates ZPL for a meter label with ICCID, Barcode, and Serial all centered."""
    s = s or {}

    pw          = mm_to_dots(s.get('widthMm',    50))
    ll          = mm_to_dots(s.get('heightMm',   25))
    by_mod      = float(s.get('byMod',       1.5))
    bc_h        = mm_to_dots(s.get('bcHeightMm', 12), minimum=20)
    iccid_font  = int(s.get('iccidFont',  22))
    sn_font     = int(s.get('snFont',     26))
    lt          = int(s.get('topOffset',   0))   # ^LT: vertical offset
    ls          = int(s.get('leftOffset',  0))   # ^LS: horizontal offset

    clean_iccid = ''.join(c for c in str(iccid) if c.isalnum())
    clean_sn    = str(meter_sn).strip()
    clean_box   = str(box_suffix).strip()
    sn_line     = f"SN: {clean_sn}  ({clean_box})" if clean_box else f"SN: {clean_sn}"

    # Calculate Barcode Width and Horizontal Centering
    mods = estimate_code128_modules(clean_iccid)
    bc_w = mods * by_mod
    # Scale module down if barcode would exceed label printable width
    if bc_w > (pw - 12) and by_mod > 1.0:
        by_mod = max(1.0, round((pw - 16) / mods, 2))
        bc_w = mods * by_mod
    bc_x = max(2, int((pw - bc_w) / 2))

    # Vertical Centering of ICCID + Barcode + Serial in the label
    min_gap = 4
    total_elements_h = iccid_font + bc_h + sn_font
    if (total_elements_h + 2 * min_gap) > ll:
        bc_h = max(20, ll - iccid_font - sn_font - 2 * min_gap)
        total_elements_h = iccid_font + bc_h + sn_font

    rem_v = max(0, ll - total_elements_h)
    gap_v = max(min_gap, rem_v // 3)
    top_margin = max(2, (ll - (iccid_font + gap_v + bc_h + gap_v + sn_font)) // 2)

    iccid_y = top_margin
    bc_y    = iccid_y + iccid_font + gap_v
    sn_y    = bc_y + bc_h + gap_v

    return f"""^XA
^PW{pw}
^LL{ll}
^LH0,0
^LT{lt}
^LS{ls}
^PR3,3
^MD5
^FO0,{iccid_y}^FB{pw},1,0,C,0^A0N,{iccid_font},{iccid_font}^FDICCID: {clean_iccid}^FS
^BY{by_mod},2.0,{bc_h}^FO{bc_x},{bc_y}^BCN,{bc_h},N,N,N,A^FD{clean_iccid}^FS
^FO0,{sn_y}^FB{pw},1,0,C,0^A0N,{sn_font},{sn_font}^FD{sn_line}^FS
^XZ
"""

def generate_pallet_header_zpl(pallet_num, total_labels, s=None):
    """Generates a PALLET HEADER sticker centered vertically and horizontally."""
    s = s or {}

    pw           = mm_to_dots(s.get('widthMm',       30))
    ll           = mm_to_dots(s.get('heightMm',       15))
    pallet_font  = int(s.get('palletNumFont', 20))
    header_font  = int(s.get('headerFont',    12))
    total_font   = int(s.get('totalFont',     14))
    lt           = int(s.get('topOffset',      0))
    ls           = int(s.get('leftOffset',     0))

    clean_pallet = str(pallet_num).strip()

    divider_h = 2
    div_w     = int(pw * 0.85)
    div_x     = (pw - div_w) // 2

    total_content_h = header_font + pallet_font + divider_h + total_font
    rem_v = max(0, ll - total_content_h)
    gap = max(2, rem_v // 4)
    top_margin = max(2, (ll - (total_content_h + 3 * gap)) // 2)

    header_y  = top_margin
    num_y     = header_y  + header_font  + gap
    divider_y = num_y     + pallet_font  + gap
    total_y   = divider_y + divider_h    + gap

    return f"""^XA
^PW{pw}
^LL{ll}
^LH0,0
^LT{lt}
^LS{ls}
^PR3,3
^MD5
^FO2,2^GB{pw - 4},{ll - 4},2^FS
^FO0,{header_y}^FB{pw},1,0,C,0^A0N,{header_font},{header_font}^FD*** PALLET ***^FS
^FO0,{num_y}^FB{pw},1,0,C,0^A0N,{pallet_font},{pallet_font}^FD{clean_pallet}^FS
^FO{div_x},{divider_y}^GB{div_w},{divider_h},{divider_h}^FS
^FO0,{total_y}^FB{pw},1,0,C,0^A0N,{total_font},{total_font}^FDTotal: {total_labels} labels^FS
^XZ
"""


class ZebraPrintHandler(SimpleHTTPRequestHandler):
    def end_headers(self):
        # Disable browser caching so JS/CSS changes take effect immediately
        self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
        self.send_header('Pragma', 'no-cache')
        self.send_header('Expires', '0')
        # Enable CORS for local fetch
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_POST(self):
        if self.path == '/api/print-zebra':
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length).decode('utf-8')
            try:
                data = json.loads(body)
                records = data.get('records', [])
                if not records and 'meterSN' in data:
                    records = [data]
                
                printer_name = data.get('printerName', get_zebra_printer_name())

                # ── Print settings from UI ──────────────────────────
                ps             = data.get('printSettings', {})
                meter_settings  = ps.get('meter',  {})
                pallet_settings = ps.get('pallet', {})
                # ────────────────────────────────────────────────────
                
                full_zpl = ""
                for rec in records:
                    rec_type = str(rec.get('type', 'label')).strip()

                    if rec_type == 'pallet_header':
                        pallet_num   = str(rec.get('palletNum',    '')).strip()
                        total_labels = int(rec.get('totalLabels',   0))
                        if pallet_num:
                            full_zpl += generate_pallet_header_zpl(pallet_num, total_labels, pallet_settings)
                    else:
                        iccid      = str(rec.get('iccid',     '')).strip()
                        sn         = str(rec.get('meterSN',   '')).strip()
                        box_suffix = str(rec.get('boxSuffix', '')).strip()
                        if iccid and sn:
                            full_zpl += generate_zpl(iccid, sn, box_suffix, meter_settings)

                # ── DEBUG: show received settings + generated ZPL ──
                print(f"\n[DEBUG] pallet_settings = {pallet_settings}")
                print(f"[DEBUG] meter_settings  = {meter_settings}")
                print(f"[DEBUG] Generated ZPL:\n{full_zpl}")
                print("─" * 50)

                if not full_zpl:
                    self.send_response(400)
                    self.send_header('Content-Type', 'application/json; charset=utf-8')
                    self.end_headers()
                    self.wfile.write(json.dumps({'error': 'لا توجد بيانات صالحة للطباعة'}).encode('utf-8'))
                    return
                
                print_zpl_raw(printer_name, full_zpl)
                
                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                self.wfile.write(json.dumps({
                    'success': True,
                    'count': len(records),
                    'printer': printer_name,
                    'message': f'تمت طباعة {len(records)} ملصق بنجاح على طابعة {printer_name}'
                }).encode('utf-8'))
                
            except Exception as e:
                self.send_response(500)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                self.wfile.write(json.dumps({'error': str(e)}).encode('utf-8'))
        else:
            self.send_error(404, "Not Found")

if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    port = 5000
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    server = HTTPServer(('127.0.0.1', port), ZebraPrintHandler)
    print("======================================================")
    print("Zebra ZT411 Direct Print Server running!")
    print(f"URL: http://localhost:{port}")
    print("======================================================")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("Stopping server...")
        server.server_close()
