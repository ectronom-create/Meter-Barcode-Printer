import os
import sys
import json
import ctypes
import time
import math
import struct
from ctypes import wintypes
from http.server import HTTPServer, SimpleHTTPRequestHandler

# Niimbot & Image dependencies
try:
    from PIL import Image, ImageDraw, ImageFont
    import barcode
    import serial.tools.list_ports
    from niimprint.packet import NiimbotPacket
    from niimprint.printer import RequestCodeEnum, InfoEnum, SerialTransport, PrinterClient
    NIIMBOT_LIB_AVAILABLE = True
except Exception as e:
    NIIMBOT_LIB_AVAILABLE = False
    print(f"[WARN] Niimbot libraries partially unavailable: {e}")

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

# ──────────────────────────────────────────────────────────────
# Zebra Printer Management
# ──────────────────────────────────────────────────────────────
def get_zebra_printer_name():
    """Detect Zebra printer installed on Windows"""
    default_name = "ZDesigner ZT411-203dpi ZPL"
    return default_name

def is_zebra_available():
    """Checks if Zebra printer is installed and can be opened"""
    try:
        winspool = ctypes.WinDLL('winspool.drv')
        hPrinter = wintypes.HANDLE()
        name = get_zebra_printer_name()
        if winspool.OpenPrinterW(name, ctypes.byref(hPrinter), None):
            winspool.ClosePrinter(hPrinter)
            return True
    except:
        pass
    return False

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
    lt          = int(s.get('topOffset',   0))
    ls          = int(s.get('leftOffset',  0))

    clean_iccid = ''.join(c for c in str(iccid) if c.isalnum())
    clean_sn    = str(meter_sn).strip()
    clean_box   = str(box_suffix).strip()
    sn_line     = f"SN: {clean_sn}  ({clean_box})" if clean_box else f"SN: {clean_sn}"

    mods = estimate_code128_modules(clean_iccid)
    bc_w = mods * by_mod
    if bc_w > (pw - 12) and by_mod > 1.0:
        by_mod = max(1.0, round((pw - 16) / mods, 2))
        bc_w = mods * by_mod
    bc_x = max(2, int((pw - bc_w) / 2))

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

    total_content_h = header_font + 3 + pallet_font + 3 + total_font
    top_margin      = max(2, (ll - total_content_h) // 2)

    hdr_y    = top_margin
    pallet_y = hdr_y + header_font + 3
    total_y  = pallet_y + pallet_font + 3

    return f"""^XA
^PW{pw}
^LL{ll}
^LH0,0
^LT{lt}
^LS{ls}
^PR3,3
^MD5
^FO0,{hdr_y}^FB{pw},1,0,C,0^A0N,{header_font},{header_font}^FDPALLET^FS
^FO0,{pallet_y}^FB{pw},1,0,C,0^A0N,{pallet_font},{pallet_font}^FD{clean_pallet}^FS
^FO0,{total_y}^FB{pw},1,0,C,0^A0N,{total_font},{total_font}^FDTotal: {total_labels} labels^FS
^XZ
"""


# ──────────────────────────────────────────────────────────────
# Niimbot B21 Printer Management
# ──────────────────────────────────────────────────────────────
def find_niimbot_port():
    """Detect Niimbot B21 COM port"""
    if not NIIMBOT_LIB_AVAILABLE:
        return None
    try:
        for p in serial.tools.list_ports.comports():
            # Check VID 3513 or B21 in description / hwid
            if '3513' in p.hwid or 'B21' in p.description or 'B21' in p.hwid or 'NIIMBOT' in p.description.upper():
                return p.device
    except:
        pass
    return None

def get_niimbot_status():
    """Returns status details of connected Niimbot printer"""
    port = find_niimbot_port()
    if not port:
        return {'available': False, 'message': 'Niimbot printer not detected via USB/COM'}
    
    try:
        transport = SerialTransport(port=port)
        client = PrinterClient(transport)
        hb = client.heartbeat()
        power = hb.get('powerlevel', 0)
        # Power level: 0 to 4 (approx 0-100%)
        battery_pct = min(100, int((power / 4.0) * 100)) if power is not None else 100
        paper_state = 'OK' if hb.get('paperstate', 0) == 0 else 'Out of Paper'
        transport._serial.close()
        return {
            'available': True,
            'port': port,
            'model': 'Niimbot B21',
            'battery': battery_pct,
            'paper': paper_state,
            'message': f'Niimbot B21 Ready on {port} (Battery: {battery_pct}%)'
        }
    except Exception as e:
        err_msg = str(e)
        if 'Access is denied' in err_msg or 'PermissionError' in err_msg:
            return {
                'available': False,
                'port': port,
                'model': 'Niimbot B21',
                'message': 'منفذ الطابعة مشغول (يرجى إغلاق تطبيق NIIMBOT الرسمي على الكمبيوتر إن كان مفتوحاً)'
            }
        return {'available': False, 'port': port, 'message': err_msg}

def generate_niimbot_label_image(iccid, meter_sn, box_suffix='', s=None):
    """Renders a pixel-perfect monochrome bitmap for Niimbot B21 (384 dots width)"""
    s = s or {}
    canvas_w = 384  # Niimbot B21 printhead width
    h_mm = float(s.get('heightMm', 25))
    canvas_h = max(100, round(h_mm * 8))

    img = Image.new('1', (canvas_w, canvas_h), 1)  # 1-bit white
    draw = ImageDraw.Draw(img)

    clean_iccid = ''.join(c for c in str(iccid) if c.isalnum())
    clean_sn = str(meter_sn).strip()
    clean_box = str(box_suffix).strip()
    sn_line = f"SN: {clean_sn}  ({clean_box})" if clean_box else f"SN: {clean_sn}"

    # Select fonts
    f_iccid_size = int(s.get('iccidFont', 20))
    f_sn_size = int(s.get('snFont', 24))
    try:
        font_iccid = ImageFont.truetype('consola.ttf', f_iccid_size)
        font_sn = ImageFont.truetype('consolab.ttf', f_sn_size)
    except:
        try:
            font_iccid = ImageFont.truetype('arial.ttf', f_iccid_size)
            font_sn = ImageFont.truetype('arialbd.ttf', f_sn_size)
        except:
            font_iccid = ImageFont.load_default()
            font_sn = font_iccid

    # 1. ICCID Header
    text_iccid = f"ICCID: {clean_iccid}"
    bbox = draw.textbbox((0, 0), text_iccid, font=font_iccid)
    w_iccid = bbox[2] - bbox[0]
    draw.text(((canvas_w - w_iccid) // 2, 8), text_iccid, fill=0, font=font_iccid)

    # 2. Barcode (Code 128)
    code128 = barcode.get_barcode_class('code128')
    bc = code128(clean_iccid)
    bc.build()
    pattern = bc.to_ascii()

    # Module width (2 dots if fits, else 1)
    mod_w = 2 if len(pattern) * 2 <= (canvas_w - 24) else 1
    bc_w = len(pattern) * mod_w
    bc_x = (canvas_w - bc_w) // 2
    bc_y = 38
    bc_h = max(30, canvas_h - 105)

    for i, ch in enumerate(pattern):
        if ch == 'X':
            draw.rectangle([bc_x + i * mod_w, bc_y, bc_x + (i + 1) * mod_w - 1, bc_y + bc_h], fill=0)

    # 3. Serial Number
    bbox_sn = draw.textbbox((0, 0), sn_line, font=font_sn)
    w_sn = bbox_sn[2] - bbox_sn[0]
    draw.text(((canvas_w - w_sn) // 2, bc_y + bc_h + 10), sn_line, fill=0, font=font_sn)

    return img

def generate_niimbot_pallet_image(pallet_num, total_labels, s=None):
    """Renders a pallet header card for Niimbot B21"""
    canvas_w = 384
    canvas_h = 200
    img = Image.new('1', (canvas_w, canvas_h), 1)
    draw = ImageDraw.Draw(img)

    try:
        font_hdr = ImageFont.truetype('arialbd.ttf', 26)
        font_num = ImageFont.truetype('arialbd.ttf', 38)
        font_tot = ImageFont.truetype('arial.ttf', 24)
    except:
        font_hdr = ImageFont.load_default()
        font_num = font_hdr
        font_tot = font_hdr

    t1 = "PALLET"
    t2 = str(pallet_num)
    t3 = f"Total: {total_labels} labels"

    for y_pos, txt, fnt in [(15, t1, font_hdr), (60, t2, font_num), (135, t3, font_tot)]:
        bbox = draw.textbbox((0, 0), txt, font=fnt)
        w = bbox[2] - bbox[0]
        draw.text(((canvas_w - w) // 2, y_pos), txt, fill=0, font=fnt)

    return img

def print_niimbot_records(records, ps=None):
    """Prints a batch of labels to Niimbot B21 via direct USB/COM"""
    if not NIIMBOT_LIB_AVAILABLE:
        raise Exception("مكتبات Niimbot غير مثبتة على السيرفر.")
        
    port = find_niimbot_port()
    if not port:
        raise Exception("لم يتم العثور على طابعة Niimbot B21 موصولة بالكمبيوتر.")

    ps = ps or {}
    meter_settings = ps.get('meter', {})
    pallet_settings = ps.get('pallet', {})
    density = int(meter_settings.get('density', 3))

    transport = SerialTransport(port=port)
    client = PrinterClient(transport)

    try:
        printed_count = 0
        for rec in records:
            rec_type = str(rec.get('type', 'label')).strip()
            if rec_type == 'pallet_header':
                pallet_num = str(rec.get('palletNum', '')).strip()
                total_labels = int(rec.get('totalLabels', 0))
                if pallet_num:
                    img = generate_niimbot_pallet_image(pallet_num, total_labels, pallet_settings)
                    client.print_image(img, density=density)
                    printed_count += 1
                    time.sleep(0.3)
            else:
                iccid = str(rec.get('iccid', '')).strip()
                sn = str(rec.get('meterSN', '')).strip()
                box_suffix = str(rec.get('boxSuffix', '')).strip()
                if iccid and sn:
                    img = generate_niimbot_label_image(iccid, sn, box_suffix, meter_settings)
                    client.print_image(img, density=density)
                    printed_count += 1
                    time.sleep(0.3)
        return printed_count
    finally:
        try:
            transport._serial.close()
        except:
            pass


# ──────────────────────────────────────────────────────────────
# HTTP Server Handler
# ──────────────────────────────────────────────────────────────
class ZebraPrintHandler(SimpleHTTPRequestHandler):
    def end_headers(self):
        # Disable caching
        self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
        self.send_header('Pragma', 'no-cache')
        self.send_header('Expires', '0')
        # Enable CORS
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_GET(self):
        # API: Printer Status
        if self.path in ('/api/status', '/api/printers'):
            zebra_ok = is_zebra_available()
            niimbot_status = get_niimbot_status()
            
            res_data = {
                'zebra': {
                    'available': zebra_ok,
                    'name': get_zebra_printer_name() if zebra_ok else None
                },
                'niimbot': niimbot_status
            }
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.end_headers()
            self.wfile.write(json.dumps(res_data).encode('utf-8'))
            return

        # Default: Static files (index.html, css, js)
        super().do_GET()

    def do_POST(self):
        if self.path in ('/api/print', '/api/print-zebra', '/api/print-niimbot'):
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length).decode('utf-8')
            try:
                data = json.loads(body)
                records = data.get('records', [])
                if not records and 'meterSN' in data:
                    records = [data]

                # Determine printer type
                target = data.get('printerType', '')
                if self.path == '/api/print-zebra':
                    target = 'zebra'
                elif self.path == '/api/print-niimbot':
                    target = 'niimbot'
                elif not target:
                    # Auto-detect: prefer Niimbot if connected, else Zebra
                    niim_info = get_niimbot_status()
                    target = 'niimbot' if niim_info.get('available') else 'zebra'

                ps = data.get('printSettings', {})

                # ── 1. NIIMBOT PRINTER ──────────────────────────────
                if target == 'niimbot':
                    count = print_niimbot_records(records, ps)
                    self.send_response(200)
                    self.send_header('Content-Type', 'application/json; charset=utf-8')
                    self.end_headers()
                    self.wfile.write(json.dumps({
                        'success': True,
                        'count': count,
                        'printer': 'Niimbot B21',
                        'message': f'تمت طباعة {count} ملصق بنجاح على طابعة Niimbot B21'
                    }).encode('utf-8'))
                    return

                # ── 2. ZEBRA PRINTER ────────────────────────────────
                printer_name = data.get('printerName', get_zebra_printer_name())
                meter_settings = ps.get('meter', {})
                pallet_settings = ps.get('pallet', {})

                full_zpl = ""
                for rec in records:
                    rec_type = str(rec.get('type', 'label')).strip()
                    if rec_type == 'pallet_header':
                        pallet_num = str(rec.get('palletNum', '')).strip()
                        total_labels = int(rec.get('totalLabels', 0))
                        if pallet_num:
                            full_zpl += generate_pallet_header_zpl(pallet_num, total_labels, pallet_settings)
                    else:
                        iccid = str(rec.get('iccid', '')).strip()
                        sn = str(rec.get('meterSN', '')).strip()
                        box_suffix = str(rec.get('boxSuffix', '')).strip()
                        if iccid and sn:
                            full_zpl += generate_zpl(iccid, sn, box_suffix, meter_settings)

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
    print("  Direct Thermal Print Server Running!")
    print(f"  URL: http://localhost:{port}")
    print("  Supported: Zebra ZT411 (ZPL) + Niimbot B21 (Direct)")
    print("======================================================")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("Stopping server...")
        server.server_close()
