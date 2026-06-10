import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import sqlite3
import datetime
import os
import webbrowser
import json
import qrcode
from tkcalendar import DateEntry
from openpyxl import Workbook, load_workbook
import sys
import threading
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import report_generator

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    return os.path.join(base_path, relative_path)

def persistent_path(relative_path):
    """ Get absolute path to persistent data, relative to executable """
    if hasattr(sys, '_MEIPASS'):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    return os.path.join(base_path, relative_path)

def center_window(win, width, height):
    win.update_idletasks()
    x = (win.winfo_screenwidth() // 2) - (width // 2)
    y = (win.winfo_screenheight() // 2) - (height // 2)
    win.geometry(f"{width}x{height}+{x}+{y}")

try:
    import win32print
    WIN32_PRINT_AVAILABLE = True
except ImportError:
    WIN32_PRINT_AVAILABLE = False

try:
    import matplotlib # type: ignore
    matplotlib.use("Agg")
    from matplotlib.backends.backend_agg import FigureCanvasAgg # type: ignore
    from matplotlib.figure import Figure # type: ignore
    import matplotlib.patches as mpatches
    from PIL import Image, ImageTk
    import numpy as np
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False

DATA_DIR = persistent_path("data")
os.makedirs(DATA_DIR, exist_ok=True)

# Migration logic for old data files
_base = os.path.abspath(os.path.join(DATA_DIR, ".."))
for _f in ["traceability_config.json", "sf_data.json", "traceability.db", "production_data.xlsx"]:
    _old = os.path.join(_base, _f)
    _new = os.path.join(DATA_DIR, _f)
    if os.path.exists(_old) and not os.path.exists(_new):
        try:
            import shutil
            shutil.move(_old, _new)
        except Exception:
            pass

CONFIG_FILE = os.path.join(DATA_DIR, "traceability_config.json")
APP_VERSION = "1.1.0"
# Premium Industrial HMI Theme

BG_COLOR = "#121212"       # Main Background
SURFACE_COLOR = "#1E1E1E"  # Panels & Cards

ACCENT_COLOR = "#00B4D8"   # Primary Action
SUCCESS_COLOR = "#2DC653"  # Running / Completed
WARNING_COLOR = "#F59E0B"  # Warning
ERROR_COLOR = "#DC2626"    # Error / Stop

BORDER_COLOR = "#334155"   # Borders

TEXT_COLOR = "#F8FAFC"     # Main Text
TEXT_MUTED = "#94A3B8"     # Secondary Text

# Status Colors
STATUS_RUNNING = "#22C55E"
STATUS_IDLE = "#64748B"
STATUS_WARNING = "#F59E0B"
STATUS_STOPPED = "#EF4444"

# Modern HMI Fonts
HMI_FONT_XL = ("Segoe UI", 20, "bold")
HMI_FONT_L = ("Segoe UI", 16, "bold")
HMI_FONT_M = ("Segoe UI", 12, "bold")
HMI_FONT_S = ("Segoe UI", 10)
HMI_FONT_XS = ("Segoe UI", 9)

# Default Hardcoded Data (Fallback)
SF_DATA_DEFAULT = {
    "MOCK-A01-10001-X-SUB": ("Alpha Subsystem Right", [("MOCK-A01-10001-Y-PRT", "Alpha Core Right"), ("MOCK-A01-10002-Z-PRT", "Alpha Buffer Right")]),
    "MOCK-A01-10003-X-SUB": ("Alpha Subsystem Aux Right", [("MOCK-A01-10003-Y-PRT", "Alpha Core (Auxiliary) Right"), ("MOCK-A01-10002-Z-PRT", "Alpha Buffer Right")]),
    "MOCK-A01-10004-X-SUB": ("Alpha Subsystem Left", [("MOCK-A01-10004-Y-PRT", "Alpha Core Left"), ("MOCK-A01-10005-Z-PRT", "Alpha Buffer Left")]),
    "MOCK-A01-10006-X-SUB": ("Alpha Subsystem Aux Left", [("MOCK-A01-10006-Y-PRT", "Alpha Core (Auxiliary) Left"), ("MOCK-A01-10005-Z-PRT", "Alpha Buffer Left")]),
    "MOCK-B02-20001-X-SUB": ("Beta Panel Right Sub", [("MOCK-B02-20001-Y-PRT", "Beta Panel Right"), ("MOCK-B02-20002-Z-PRT", "Beta Sealant")]),
    "MOCK-B02-20003-X-SUB": ("Beta Panel Left Sub", [("MOCK-B02-20003-Y-PRT", "Beta Panel Left"), ("MOCK-B02-20002-Z-PRT", "Beta Sealant")]),
    "MOCK-B02-20004-X-SUB": ("Beta Panel Right Sub V2", [("MOCK-B02-20004-Y-PRT", "Beta Panel (Upgraded) Right"), ("MOCK-B02-20002-Z-PRT", "Beta Sealant")]),
    "MOCK-B02-20005-X-SUB": ("Beta Panel Left Sub V2", [("MOCK-B02-20005-Y-PRT", "Beta Panel (Upgraded) Left"), ("MOCK-B02-20002-Z-PRT", "Beta Sealant")]),
    "MOCK-C03-30001-X-SUB": ("Gamma Switch Right Sub (Short)", [("MOCK-C03-30001-Y-PRT", "Gamma Switch Right"), ("MOCK-C03-30002-Z-PRT", "Gamma Coil Right"), ("MOCK-C03-30003-W-PRT", "Gamma Pin")]),
    "MOCK-C03-30004-X-SUB": ("Gamma Switch Left Sub (Short)", [("MOCK-C03-30004-Y-PRT", "Gamma Switch Left"), ("MOCK-C03-30005-Z-PRT", "Gamma Coil Left"), ("MOCK-C03-30003-W-PRT", "Gamma Pin")]),
    "MOCK-C03-30006-X-SUB": ("Gamma Switch Right Sub (Long)", [("MOCK-C03-30006-Y-PRT", "Gamma Switch Long Right"), ("MOCK-C03-30002-Z-PRT", "Gamma Coil Right"), ("MOCK-C03-30003-W-PRT", "Gamma Pin")]),
    "MOCK-C03-30007-X-SUB": ("Gamma Switch Left Sub (Long)", [("MOCK-C03-30007-Y-PRT", "Gamma Switch Long Left"), ("MOCK-C03-30005-Z-PRT", "Gamma Coil Left"), ("MOCK-C03-30003-W-PRT", "Gamma Pin")]),
    "MOCK-D04-40001-X-SUB": ("Delta Processor Node Right", [("MOCK-D04-40001-Y-PRT", "Delta Primary Node"), ("MOCK-D04-40002-Z-PRT", "Delta Interface Right"), ("MOCK-D04-40003-W-PRT", "Delta Rotor")]),
    "MOCK-D04-40004-X-SUB": ("Delta Processor Node Left", [("MOCK-D04-40001-Y-PRT", "Delta Primary Node"), ("MOCK-D04-40005-Z-PRT", "Delta Interface Left"), ("MOCK-D04-40003-W-PRT", "Delta Rotor")]),
    "MOCK-D04-40006-X-SUB": ("Delta Processor Array Right", [("MOCK-D04-40001-Y-PRT", "Delta Primary Node"), ("MOCK-D04-40007-Y-PRT", "Delta Secondary Node"), ("MOCK-D04-40008-Z-PRT", "Delta Interface Array Right"), ("MOCK-D04-40003-W-PRT", "Delta Rotor")]),
    "MOCK-E05-50001-X-SUB": ("Epsilon Housing Sub Right", [("MOCK-E05-50001-Y-PRT", "Epsilon Housing Right"), ("MOCK-E05-50002-Z-PRT", "Epsilon Axis")]),
    "MOCK-E05-50003-X-SUB": ("Epsilon Housing Sub Left", [("MOCK-E05-50003-Y-PRT", "Epsilon Housing Left"), ("MOCK-E05-50002-Z-PRT", "Epsilon Axis")]),
    "MOCK-F06-60001-X-SUB": ("Zeta Mount Assembly RH", [("MOCK-F06-60001-Y-PRT", "Zeta Mount RH"), ("MOCK-F06-60002-Z-PRT", "Zeta Fastener")]),
    "MOCK-F06-60003-X-SUB": ("Zeta Mount Assembly LH", [("MOCK-F06-60003-Y-PRT", "Zeta Mount LH"), ("MOCK-F06-60002-Z-PRT", "Zeta Fastener")]),
    "MOCK-G07-70001-X-SUB": ("Omega Framework Sub Left", [("MOCK-G07-70001-Y-PRT", "Omega Anchor Left"), ("MOCK-G07-70002-Y-PRT", "Omega Frame Left"), ("MOCK-A01-10005-Z-PRT", "Alpha Buffer Left"), ("MOCK-B02-20002-Z-PRT", "Beta Sealant")])
}

SF_DATA_FILE = os.path.join(DATA_DIR, "sf_data.json")
SF_DATA = {}

def load_sf_data():
    global SF_DATA
    SF_DATA = {}
    try:
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("SELECT pn_sf, name_sf, rms_json FROM products")
        rows = c.fetchall()
        for r in rows:
            pn = r[0]
            name = r[1]
            try:
                rms = json.loads(r[2])
            except:
                rms = []
            SF_DATA[pn] = (name, rms)
        conn.close()
        
        if not SF_DATA:
            SF_DATA = dict(SF_DATA_DEFAULT)
            conn = get_db_connection()
            c = conn.cursor()
            for pn, val in SF_DATA.items():
                c.execute("INSERT INTO products (pn_sf, name_sf, rms_json) VALUES (?, ?, ?)", (pn, val[0], json.dumps(val[1])))
            conn.commit()
            conn.close()
            
    except Exception as e:
        print("Error loading SF_DATA from DB:", e)
        SF_DATA = dict(SF_DATA_DEFAULT)


DB_FILE = os.path.join(DATA_DIR, "traceability.db")

def get_db_connection():
    conn = sqlite3.connect(DB_FILE, check_same_thread=False, timeout=10)
    conn.execute('PRAGMA journal_mode=WAL')
    return conn

import hashlib
import binascii

def hash_password(password):
    salt = b"subproc_trace_salt_2026"
    hash_obj = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 100000)
    return binascii.hexlify(hash_obj).decode("utf-8")

def migrate_passwords():
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("SELECT id, password FROM auth")
    rows = c.fetchall()
    for row in rows:
        uid, pwd = row
        if pwd and (len(pwd) != 64 or not all(ch in "0123456789abcdef" for ch in pwd)):
            hashed = hash_password(pwd)
            c.execute("UPDATE auth SET password = ? WHERE id = ?", (hashed, uid))
    conn.commit()
    conn.close()
EXCEL_FILE = os.path.join(DATA_DIR, "production_data.xlsx")

def init_db():
    conn = get_db_connection()
    c = conn.cursor()
    
    c.execute('''
        CREATE TABLE IF NOT EXISTS products (
            pn_sf TEXT PRIMARY KEY,
            name_sf TEXT,
            rms_json TEXT
        )
    ''')
    
    c.execute('''
        CREATE TABLE IF NOT EXISTS product_audit_trail (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            action TEXT,
            pn_sf TEXT,
            details TEXT,
            user_id TEXT,
            shift TEXT,
            timestamp TEXT
        )
    ''')
    
    c.execute('''
        CREATE TABLE IF NOT EXISTS system_access_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            event_type TEXT,
            user_id TEXT,
            shift TEXT,
            timestamp TEXT
        )
    ''')
    
    c.execute('''
        CREATE TABLE IF NOT EXISTS auth (
            id TEXT PRIMARY KEY,
            password TEXT,
            role TEXT
        )
    ''')
    
    try:
        c.executemany("INSERT OR IGNORE INTO auth (id, password, role) VALUES (?, ?, ?)", [
            ('1', '1', 'Supervisor'),
            ('999', 'hilex999', 'Supervisor'),
            ('998', 'hilex998', 'Shift Leader'),
            ('111', 'hilex111', 'Operator'),
            ('mg90', 'oi90', 'Manager')
        ])
        conn.commit()
    except Exception as e:
        print("Auth seed error:", e)
    
    migrate_passwords()
        
    c.execute("SELECT COUNT(*) FROM products")
    count = c.fetchone()[0]
    if count == 0 and os.path.exists(SF_DATA_FILE):
        try:
            with open(SF_DATA_FILE, "r", encoding="utf-8") as f:
                json_data = json.load(f)
            for pn, val in json_data.items():
                c.execute("INSERT INTO products (pn_sf, name_sf, rms_json) VALUES (?, ?, ?)", (pn, val[0], json.dumps(val[1])))
            conn.commit()
        except Exception:
            pass

    c.execute('''
        CREATE TABLE IF NOT EXISTS records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sub_batch_id TEXT NOT NULL,
            pn_sf TEXT,
            part_sf TEXT,
            rm1_pn TEXT,
            rm1_name TEXT,
            rm2_pn TEXT,
            rm2_name TEXT,
            rm3_pn TEXT,
            rm3_name TEXT,
            rm4_pn TEXT,
            rm4_name TEXT,
            batch1 TEXT,
            batch2 TEXT,
            batch3 TEXT,
            quantity INTEGER,
            shift_sp TEXT,
            op_id TEXT,
            station TEXT,
            dt_sp TEXT,
            dt_line TEXT,
            shift_line TEXT,
            remarks TEXT,
            status TEXT DEFAULT 'In Rack',
            created_at TEXT NOT NULL
        )
    ''')
    
    try:
        c.execute("ALTER TABLE records ADD COLUMN registered_by TEXT DEFAULT ''")
    except sqlite3.OperationalError:
        pass
        
    try:
        c.execute("ALTER TABLE records ADD COLUMN reprint_count INTEGER DEFAULT 0")
    except sqlite3.OperationalError:
        pass
        
    try:
        c.execute("ALTER TABLE records ADD COLUMN last_reprinted_at TEXT")
    except sqlite3.OperationalError:
        pass
        
    try:
        c.execute("ALTER TABLE records ADD COLUMN last_reprinted_by TEXT")
    except sqlite3.OperationalError:
        pass
        
    c.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_records_sub_batch_id ON records(sub_batch_id)")
    
    c.execute('''CREATE TABLE IF NOT EXISTS shift_targets (
                     id INTEGER PRIMARY KEY AUTOINCREMENT,
                     product_pn TEXT NOT NULL,
                     shift TEXT NOT NULL,
                     station TEXT,
                     target_qty INTEGER NOT NULL,
                     effective_date TEXT NOT NULL
                 )''')

    c.execute('''CREATE TABLE IF NOT EXISTS part_thresholds (
                     pn_sf TEXT PRIMARY KEY,
                     min_qty INTEGER NOT NULL DEFAULT 0
                 )''')

    c.execute('''CREATE TABLE IF NOT EXISTS inventory_snapshots (
                     snapshot_date TEXT NOT NULL,
                     pn_sf TEXT NOT NULL,
                     boxes_in_rack INTEGER,
                     total_qty_in_rack INTEGER,
                     oldest_box_age_hours REAL,
                     PRIMARY KEY (snapshot_date, pn_sf)
                 )''')
        
    conn.commit()
    conn.close()
    
    load_sf_data()


def save_to_excel(data):
    headers = [
        "SB ID", "FULL PN Semi fini", "PART NAME (SF)", 
        "RM PN", "RM Name",
        "Batch No. 1", "Batch No. 2", "Batch No. 3", "Quantity",
        "Work in Sub-Process by Shift", "Op ID", "Station",
        "Sub-Process Work Date/Time", "Production Line Entry Date/Time",
        "Work in PROD line by Shift", "Remarks", "Registered by ID"
    ]
    
    header_font = Font(name='Calibri', size=10, bold=True, color="000000")
    header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "Sub-process fill by TL"
        ws.append(headers)
        
        for col in range(1, 18):
            cell = ws.cell(row=1, column=col)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border
        
        ws.freeze_panes = "A2"
    else:
        wb = load_workbook(EXCEL_FILE)
        if "Sub-process fill by TL" in wb.sheetnames:
            ws = wb["Sub-process fill by TL"]
        else:
            ws = wb.create_sheet("Sub-process fill by TL")
            ws.append(headers)
            
        cell = ws.cell(row=1, column=17)
        if cell.value != "Registered by ID":
            cell.value = "Registered by ID"
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border

    ws.row_dimensions[1].height = 34
    widths = [20, 22, 25, 22, 25, 12, 12, 12, 10, 15, 10, 10, 20, 20, 15, 25, 15]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w
    
    sb_id = data[0] if data[0] else "-"
    sf_pn = data[1] if data[1] else "-"
    part_sf = data[2] if data[2] else "-"
    rms = []
    for i in range(4):
        rm_pn = data[3 + i*2]
        rm_name = data[4 + i*2]
        if rm_pn:
            rms.append((rm_pn if rm_pn else "-", rm_name if rm_name else "-"))
            
    if not rms:
        rms = [("-", "-")]
        
    num_rms = len(rms)
    rest_data = ["-" if (x == "" or x is None) else x for x in data[11:]]
    
    start_row = ws.max_row + 1
    end_row = start_row + num_rms - 1
    
    for i in range(num_rms):
        row_data = []
        if i == 0:
            row_data.extend([sb_id, sf_pn, part_sf])
            row_data.extend([rms[i][0], rms[i][1]])
            row_data.extend(rest_data)
        else:
            row_data.extend([None, None, None])
            row_data.extend([rms[i][0], rms[i][1]])
            row_data.extend([None] * len(rest_data))
        ws.append(row_data)
        
    if num_rms > 1:
        ws.merge_cells(start_row=start_row, start_column=1, end_row=end_row, end_column=1)
        ws.merge_cells(start_row=start_row, start_column=2, end_row=end_row, end_column=2)
        ws.merge_cells(start_row=start_row, start_column=3, end_row=end_row, end_column=3)
        for col in range(6, 18):
            ws.merge_cells(start_row=start_row, start_column=col, end_row=end_row, end_column=col)
    
    sf_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    rm_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    rest_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    row_font = Font(name='Calibri', size=10, color="000000")
    align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    for r in range(start_row, end_row + 1):
        for c in range(1, 18):
            cell = ws.cell(row=r, column=c)
            cell.font = row_font
            cell.alignment = align
            cell.border = thin_border
            if c in (1, 2, 3):
                cell.fill = sf_fill
            elif c in (4, 5):
                cell.fill = rm_fill
            else:
                cell.fill = rest_fill
            
    wb.save(EXCEL_FILE)

def generate_zpl(record_data):
    qr_payload = {
        "sub_batch_id": record_data.get('sub_batch_id', ''),
        "full_pn_sf": record_data.get('pn_sf', ''),
        "part_name_sf": record_data.get('part_sf', ''),
        "rm1_pn": record_data.get('rm1_pn', ''),
        "rm1_name": record_data.get('rm1_name', ''),
        "rm2_pn": record_data.get('rm2_pn', ''),
        "rm2_name": record_data.get('rm2_name', ''),
        "rm3_pn": record_data.get('rm3_pn', ''),
        "rm3_name": record_data.get('rm3_name', ''),
        "rm4_pn": record_data.get('rm4_pn', ''),
        "rm4_name": record_data.get('rm4_name', ''),
        "batch_no_1": record_data.get('batch1', ''),
        "batch_no_2": record_data.get('batch2', ''),
        "batch_no_3": record_data.get('batch3', ''),
        "quantity": record_data.get('quantity', 0),
        "sub_process_shift": record_data.get('shift_sp', ''),
        "op_id": record_data.get('op_id', ''),
        "station": record_data.get('station', ''),
        "sub_process_datetime": record_data.get('dt_sp', ''),
        "production_line_entry_datetime": record_data.get('dt_line', ''),
        "production_line_shift": record_data.get('shift_line', ''),
        "remarks": record_data.get('remarks', '')
    }
    json_str = json.dumps(qr_payload)
    
    def format_val(val, default="-"):
        return str(val) if val else default

    zpl = []
    zpl.append("^XA")
    zpl.append("^PW520")
    zpl.append("^LH25,0")
    zpl.append("^CI28")
    
    def draw_layout(offset_y):
        zpl.append(f"^FO80,{20+offset_y}^A0N,24,24^FDSUB-PROCESS RECORD^FS")
        zpl.append(f"^FO10,{55+offset_y}^GB460,3,3^FS")
        
        zpl.append(f"^FO10,{70+offset_y}^GB460,40,40^FS")
        zpl.append(f"^FO10,{80+offset_y}^A0N,24,24^FR^FB460,1,0,C^FD{format_val(record_data.get('sub_batch_id'))}\\&^FS")
        
        sb_id_val = format_val(record_data.get('sub_batch_id'))
        zpl.append(f"^BY2^FO15,{120+offset_y}^BCN,50,N,N,N,A^FD{sb_id_val}^FS")
        
        y = 185 + offset_y
        def add_row(label, value, font_size=18, y_inc=20):
            nonlocal y
            zpl.append(f"^FO10,{y}^A0N,{font_size},{font_size}^FD{label}^FS")
            zpl.append(f"^FO10,{y}^A0N,{font_size},{font_size}^FB460,1,0,R^FD{value}\\&^FS")
            y += y_inc

        add_row("PN Semi fini", format_val(record_data.get('pn_sf')))
        add_row("Part Name (SF)", format_val(record_data.get('part_sf')))
                
        y += 3
        zpl.append(f"^FO10,{y}^GB460,1,1^FS")
        y += 8
        
        add_row("Batch No. 1", format_val(record_data.get('batch1')))
        add_row("Batch No. 2", format_val(record_data.get('batch2')))
        add_row("Batch No. 3", format_val(record_data.get('batch3')))
        add_row("Quantity", f"{format_val(record_data.get('quantity'))} pcs", font_size=20, y_inc=22)
        
        y += 3
        zpl.append(f"^FO10,{y}^GB460,1,1^FS")
        y += 8
        
        add_row("Shift SP", format_val(record_data.get('shift_sp')))
        add_row("Op ID", format_val(record_data.get('op_id')))
        add_row("Station", format_val(record_data.get('station')))
        dt_sp = format_val(record_data.get('dt_sp'))[:16] if record_data.get('dt_sp') else '-'
        add_row("SP Date/Time", dt_sp)
        dt_line = format_val(record_data.get('dt_line'))[:16] if record_data.get('dt_line') else '-'
        add_row("Line Entry", dt_line)
        
        printed = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        add_row("Printed At", printed)
        
        y += 3
        zpl.append(f"^FO10,{y}^GB460,2,2^FS")
        y += 8
        
        reprints = int(record_data.get('reprint_count') or 0)
        if reprints > 0:
            last_by = record_data.get('last_reprinted_by', 'Unknown')
            last_at = record_data.get('last_reprinted_at', '')[:16] if record_data.get('last_reprinted_at') else ''
            audit_text = f"** REPRINT #{reprints} by {last_by} at {last_at} **"
            zpl.append(f"^FO10,{y}^A0N,14,14^FB460,1,0,C^FD{audit_text}\\&^FS")
        else:
            zpl.append(f"^FO10,{y}^A0N,14,14^FB460,1,0,C^FD-- Original Print --\\&^FS")
        y += 18
        
        zpl.append(f"^FO130,{y}^BQN,2,2^FDQA,{json_str}^FS")
        return y + 110
        
    end_y = draw_layout(0)
    
    zpl.append(f"^FO10,{end_y}^A0N,20,20^FB460,1,0,C^FD- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -^FS")
    
    draw_layout(end_y + 30)

    zpl.append("^PQ1")
    zpl.append("^XZ")
    return "\n".join(zpl).encode("utf-8")

def execute_zpl_print(zpl_string, record_data, silent=False):
    config = {}
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                config = json.load(f)
        except: pass
            
    printer_name = config.get("zebra_printer", "")
    
    try:
        qr_dir = persistent_path("qr")
        os.makedirs(qr_dir, exist_ok=True)
        sb_id_safe = str(record_data.get('sub_batch_id', 'unknown')).replace(":", "").replace("-", "")
        file_path = os.path.join(qr_dir, f"{sb_id_safe}.zpl")
        with open(file_path, "wb") as f:
            f.write(zpl_string)
    except Exception as e:
        print("Could not save ZPL file:", e)
        
    if not printer_name:
        if not silent:
            messagebox.showwarning("Setup Required", "ZPL saved to 'qr' folder. Please configure Zebra Printer in Settings to physically print.")
        return
        
    if not WIN32_PRINT_AVAILABLE:
        if not silent:
            messagebox.showerror("Error", "pywin32 is not installed. ZPL saved to 'qr' folder, but cannot print. Please run 'pip install pywin32'.")
        return
    
    try:
        hPrinter = win32print.OpenPrinter(printer_name)
        try:
            job = win32print.StartDocPrinter(hPrinter, 1, ("Sub-Process Label", None, "RAW"))
            win32print.StartPagePrinter(hPrinter)
            win32print.WritePrinter(hPrinter, zpl_string)
            win32print.EndPagePrinter(hPrinter)
            win32print.EndDocPrinter(hPrinter)
            if not silent:
                messagebox.showinfo("Success", f"Label sent to printer '{printer_name}' successfully!")
        finally:
            win32print.ClosePrinter(hPrinter)
    except Exception as e:
        if not silent:
            messagebox.showerror("Printer Error", f"Failed to send to printer '{printer_name}':\n{e}")

def print_html_slip(record_data, silent=False):
    zpl_string = generate_zpl(record_data)
    execute_zpl_print(zpl_string, record_data, silent)

def generate_pick_ticket_zpl(pn, target_qty, accumulated, picked):
    zpl = []
    zpl.append("^XA")
    zpl.append("^PW480")
    zpl.append("^LH20,0")
    zpl.append("^CI28")
    
    zpl.append("^FO0,30^A0N,30,30^FB440,1,0,C^FDFIFO PICK TICKET^FS")
    zpl.append("^FO0,70^GB440,3,3^FS")
    
    zpl.append(f"^FO5,85^A0N,22,22^FDPart: {pn}^FS")
    zpl.append(f"^FO5,115^A0N,22,22^FDTarget: {target_qty} | Actual: {accumulated}^FS")
    
    y = 150
    table_start_y = y
    
    zpl.append(f"^FO0,{y}^GB440,2,2^FS")
    y += 5
    
    zpl.append(f"^FO5,{y}^A0N,20,20^FDSub-Batch ID^FS")
    zpl.append(f"^FO215,{y}^A0N,20,20^FDQty^FS")
    zpl.append(f"^FO280,{y}^A0N,20,20^FDStored Date^FS")
    y += 25
    
    zpl.append(f"^FO0,{y}^GB440,2,2^FS")
    
    for sb, q, dt in picked:
        y += 5
        zpl.append(f"^FO5,{y}^A0N,18,18^FD{sb}^FS")
        zpl.append(f"^FO215,{y}^A0N,18,18^FD{q}^FS")
        dt_short = dt[:16] if dt else ""
        zpl.append(f"^FO280,{y}^A0N,18,18^FD{dt_short}^FS")
        y += 25
        zpl.append(f"^FO0,{y}^GB440,1,1^FS")
        
    zpl.append(f"^FO0,{table_start_y}^GB2,{y - table_start_y},2^FS")
    zpl.append(f"^FO210,{table_start_y}^GB2,{y - table_start_y},2^FS")
    zpl.append(f"^FO270,{table_start_y}^GB2,{y - table_start_y},2^FS")
    zpl.append(f"^FO440,{table_start_y}^GB2,{y - table_start_y},2^FS")
    
    y += 15
    import datetime
    printed = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    zpl.append(f"^FO5,{y}^A0N,18,18^FDPrinted: {printed}^FS")
    
    zpl.append("^PQ1")
    zpl.append("^XZ")
    return "\n".join(zpl).encode("utf-8")

def execute_ticket_print(zpl_string, pn, silent=False):
    config = {}
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                config = json.load(f)
        except: pass

    printer_name = config.get("zebra_printer", "")

    try:
        tickets_dir = persistent_path("tickets")
        os.makedirs(tickets_dir, exist_ok=True)
        safe_pn = str(pn).replace(":", "").replace("-", "")
        import datetime
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        file_path = os.path.join(tickets_dir, f"TICKET_{safe_pn}_{timestamp}.zpl")
        with open(file_path, "wb") as f:
            f.write(zpl_string)
    except Exception as e:
        print("Could not save Pick Ticket ZPL file:", e)

    if not printer_name:
        if not silent:
            messagebox.showwarning("Setup Required", "Ticket saved to 'tickets' folder.\\nPlease configure Zebra Printer in Settings to physically print.")
        return

    if not WIN32_PRINT_AVAILABLE:
        if not silent: messagebox.showerror("Error", "win32print module not found. Printing disabled.")
        return

    try:
        hPrinter = win32print.OpenPrinter(printer_name)
        try:
            hJob = win32print.StartDocPrinter(hPrinter, 1, ("Pick Ticket", None, "RAW"))
            try:
                win32print.StartPagePrinter(hPrinter)
                win32print.WritePrinter(hPrinter, zpl_string)
                win32print.EndPagePrinter(hPrinter)
            finally:
                win32print.EndDocPrinter(hPrinter)
            if not silent: messagebox.showinfo("Success", f"Pick ticket sent to printer '{printer_name}'")
        finally:
            win32print.ClosePrinter(hPrinter)
    except Exception as e:
        if not silent: messagebox.showerror("Printer Error", f"Failed to send to printer:\\n{e}")


class TraceabilityApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"HI-LEX ACT - Sub-Process Traceability System v{APP_VERSION}")
        self.geometry("1280x800")
        try:
            taskbar_logo = resource_path(os.path.join("assets", "taskbar_logo.png"))
            if os.path.exists(taskbar_logo):
                self.iconphoto(True, tk.PhotoImage(file=taskbar_logo))
        except: pass
        try:
            self.state('zoomed')
        except:
            pass
        self.configure(bg=BG_COLOR)
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.schedule_daily_snapshot()
        
        # Modern Windows 11 Title Bar Color (White background, Blue text)
        try:
            import ctypes
            self.update_idletasks()
            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            # DWMWA_CAPTION_COLOR = 35
            ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 35, ctypes.byref(ctypes.c_int(0x00FFFFFF)), 4)
            # DWMWA_TEXT_COLOR = 36 (HI-LEX Blue #005A8C -> 0x008C5A00)
            ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 36, ctypes.byref(ctypes.c_int(0x008C5A00)), 4)
        except Exception:
            pass
        
        self.option_add('*insertBackground', 'white')
        self.option_add('*Listbox.background', SURFACE_COLOR)
        self.option_add('*Listbox.foreground', TEXT_COLOR)
        self.option_add('*Listbox.selectBackground', ACCENT_COLOR)
        self.option_add('*Listbox.selectForeground', 'black')
        self.option_add('*TCombobox*Listbox.background', SURFACE_COLOR)
        self.option_add('*TCombobox*Listbox.foreground', TEXT_COLOR)
        self.option_add('*TCombobox*Listbox.selectBackground', ACCENT_COLOR)
        self.option_add('*TCombobox*Listbox.selectForeground', 'black')
        
        # Configure Styles
        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        self.style.configure('.', background=BG_COLOR, foreground=TEXT_COLOR, fieldbackground=SURFACE_COLOR)
        self.style.configure('TNotebook', background=BG_COLOR)
        self.style.configure('TNotebook.Tab', background=SURFACE_COLOR, foreground=TEXT_COLOR, padding=[20, 8], font=("Segoe UI", 11, "bold"))
        self.style.map('TNotebook.Tab', background=[('selected', ACCENT_COLOR)])
        self.style.configure('TFrame', background=BG_COLOR)
        self.style.configure('TLabel', background=SURFACE_COLOR, foreground=TEXT_COLOR)
        self.style.configure('TCheckbutton', background=SURFACE_COLOR, foreground=TEXT_COLOR)
        
        self.style.configure('TCombobox', fieldbackground=SURFACE_COLOR, background=SURFACE_COLOR, foreground=TEXT_COLOR)
        self.style.map('TCombobox', 
                       fieldbackground=[('readonly', SURFACE_COLOR)], 
                       selectbackground=[('readonly', ACCENT_COLOR)],
                       selectforeground=[('readonly', '#000000')])
        
        self.style.configure('TEntry', fieldbackground=SURFACE_COLOR, foreground=TEXT_COLOR, insertcolor='white')
        self.style.configure('TButton', background=SURFACE_COLOR, foreground=TEXT_COLOR, font=HMI_FONT_M, padding=5)
        self.style.map('TButton', background=[('active', BORDER_COLOR)])
        
        self.style.configure('Success.TButton', background=SUCCESS_COLOR, foreground="#000000", font=HMI_FONT_M, padding=5)
        self.style.map('Success.TButton', background=[('disabled', '#4A5568'), ('active', STATUS_RUNNING)], foreground=[('disabled', '#A0AEC0')])
        
        self.style.configure('Primary.TButton', background=ACCENT_COLOR, foreground="#000000", font=HMI_FONT_M, padding=5)
        self.style.map('Primary.TButton', background=[('disabled', '#4A5568'), ('active', "#0096C7")], foreground=[('disabled', '#A0AEC0')])
        
        self.style.configure('Secondary.TButton', background=STATUS_IDLE, foreground="#FFFFFF", font=HMI_FONT_M, padding=5)
        self.style.map('Secondary.TButton', background=[('disabled', '#4A5568'), ('active', BORDER_COLOR)], foreground=[('disabled', '#A0AEC0')])
        
        self.style.configure('Header.TButton', background="#005A8C", foreground="#FFFFFF", font=HMI_FONT_M, padding=5)
        self.style.map('Header.TButton', background=[('disabled', '#4A5568'), ('active', "#00456B")], foreground=[('disabled', '#A0AEC0')])
        
        self.style.configure('Warning.TButton', background=WARNING_COLOR, foreground="#000000", font=HMI_FONT_M, padding=5)
        self.style.map('Warning.TButton', background=[('disabled', '#4A5568'), ('active', "#D97706")], foreground=[('disabled', '#A0AEC0')])
        
        self.style.configure('Danger.TButton', background=ERROR_COLOR, foreground="#000000", font=HMI_FONT_M, padding=5)
        self.style.map('Danger.TButton', background=[('disabled', '#4A5568'), ('active', "#B91C1C")], foreground=[('disabled', '#A0AEC0')])
        
        self.style.configure('Treeview', background=SURFACE_COLOR, foreground=TEXT_COLOR, fieldbackground=SURFACE_COLOR, rowheight=30, font=HMI_FONT_S)
        self.style.configure('Treeview.Heading', background=BORDER_COLOR, foreground=TEXT_COLOR, font=HMI_FONT_M)
        self.style.map('Treeview.Heading',
                       background=[('active', ACCENT_COLOR)],
                       foreground=[('active', '#000000')])
        
        self.app_user_id = ""
        self.app_user_shift = ""
        self.is_admin = False
        
        self.build_ui()
        self.after(100, self.prompt_login)
        self.update_clock()
        self.update_stats()
        self.populate_sf_combobox()
        self.refresh_recent_treeview()
        self.refresh_records_treeview()
        self.update_sub_batch_preview()

    def prompt_login(self):
        login_win = tk.Toplevel(self)
        login_win.title("HI-LEX Login")
        center_window(login_win, 400, 300)
        login_win.configure(bg=BG_COLOR)
        login_win.transient(self)
        login_win.grab_set()
        
        def on_login_close():
            login_win.destroy()
            self.destroy()
            
        login_win.protocol("WM_DELETE_WINDOW", on_login_close)
        
        tk.Label(login_win, text="Operator Login", font=HMI_FONT_L, bg=BG_COLOR, fg=TEXT_COLOR).pack(pady=15)
        
        frame = tk.Frame(login_win, bg=BG_COLOR)
        frame.pack(pady=10)
        
        tk.Label(frame, text="Operator ID:", bg=BG_COLOR, fg=TEXT_COLOR).grid(row=0, column=0, padx=10, pady=10, sticky="e")
        vcmd_login = (login_win.register(lambda P: len(P) <= 7 and all(c.isalnum() or c == '-' for c in P)), '%P')
        entry_id = ttk.Entry(frame, width=20, validate="key", validatecommand=vcmd_login)
        entry_id.grid(row=0, column=1, padx=10, pady=10)
        entry_id.focus()
        
        tk.Label(frame, text="Password:", bg=BG_COLOR, fg=TEXT_COLOR).grid(row=1, column=0, padx=10, pady=10, sticky="e")
        entry_pass = ttk.Entry(frame, width=20, show="*")
        entry_pass.grid(row=1, column=1, padx=10, pady=10)
        
        tk.Label(frame, text="Shift:", bg=BG_COLOR, fg=TEXT_COLOR).grid(row=2, column=0, padx=10, pady=10, sticky="e")
        cb_shift = ttk.Combobox(frame, values=["A", "B", "C"], state="readonly", width=18)
        cb_shift.grid(row=2, column=1, padx=10, pady=10)
        cb_shift.set("A")
        
        def do_login(event=None):
            uid = entry_id.get().strip()
            upass = entry_pass.get()
            ushift = cb_shift.get().strip()
            
            if not uid or not upass or not ushift:
                messagebox.showerror("Error", "Please enter Operator ID, Password, and Shift.", parent=login_win)
                return
                
            try:
                conn = get_db_connection()
                c = conn.cursor()
                c.execute("SELECT role FROM auth WHERE id = ? AND password = ?", (uid, hash_password(upass)))
                result = c.fetchone()
                conn.close()
            except Exception as e:
                messagebox.showerror("DB Error", f"Failed to connect to authentication database:\n{e}", parent=login_win)
                return
                
            if not result:
                messagebox.showerror("Error", "Invalid Operator ID or Password.", parent=login_win)
                return
                
            role = result[0]
            self.is_admin = False
            display_name = uid
            
            if role in ("Supervisor", "Shift Leader", "Manager"):
                self.is_admin = True
                if uid == "mg90":
                    display_name = "Taketo Oi (Manager)"
                else:
                    display_name = role
                
            self.app_user_id = uid
            self.app_user_shift = ushift
            self.lbl_header_user.config(text=f"User: {display_name} | Shift: {ushift}")
            
            try:
                conn = get_db_connection()
                c = conn.cursor()
                timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                c.execute("INSERT INTO system_access_logs (event_type, user_id, shift, timestamp) VALUES (?, ?, ?, ?)",
                          ("LOGIN", uid, ushift, timestamp))
                conn.commit()
                conn.close()
            except Exception as e:
                pass
            
            if self.is_admin:
                if hasattr(self, 'lbl_header_warning'):
                    self.lbl_header_warning.config(text=f" You are {display_name}, Please don't forget to logout!")
                    self.lbl_header_warning.pack(side=tk.RIGHT, padx=20)
                if hasattr(self, 'notebook') and hasattr(self, 'tab3'):
                    self.notebook.add(self.tab3, text="KPIs")
                    if hasattr(self, 'refresh_kpis'):
                        self.refresh_kpis()
                
            login_win.destroy()
            # Initialize default shift for manual entry
            self.clear_form()
            
        ttk.Button(login_win, text="Login", style="Success.TButton", command=do_login).pack(pady=15)
        login_win.bind("<Return>", do_login)

    def create_card(self, parent, title, fg_color=TEXT_COLOR):
        card = tk.Frame(parent, bg=SURFACE_COLOR)
        card.pack(fill=tk.X, pady=5, padx=5)
        
        title_lbl = tk.Label(card, text=title, bg=SURFACE_COLOR, fg=fg_color, font=HMI_FONT_M, anchor="w", padx=10, pady=5)
        title_lbl.pack(fill=tk.X)
        
        tk.Frame(card, bg=BORDER_COLOR, height=1).pack(fill=tk.X)
        
        content = tk.Frame(card, bg=SURFACE_COLOR, padx=10, pady=10)
        content.pack(fill=tk.BOTH, expand=True)
        return card, content

    def open_user_manager(self):
        top = tk.Toplevel(self)
        top.title("User Management (Admin)")
        center_window(top, 800, 500)
        top.resizable(False, False)
        
        top.update_idletasks()
        x = (top.winfo_screenwidth() // 2) - (800 // 2)
        y = (top.winfo_screenheight() // 2) - (500 // 2)
        top.geometry(f"+{x}+{y}")
        
        top.configure(bg=BG_COLOR)
        top.transient(self)
        top.grab_set()
        
        main_frame = tk.Frame(top, bg=BG_COLOR)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        left_frame = tk.Frame(main_frame, bg=BG_COLOR)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 20))
        
        tk.Label(left_frame, text="Add New User", bg=BG_COLOR, fg=TEXT_COLOR, font=HMI_FONT_L).pack(pady=(0, 20))
        
        tk.Label(left_frame, text="Operator ID / Name:", bg=BG_COLOR, fg=TEXT_COLOR, font=HMI_FONT_S).pack(anchor="w")
        ent_id = ttk.Entry(left_frame, width=30)
        ent_id.pack(pady=(0, 15))
        
        tk.Label(left_frame, text="Password:", bg=BG_COLOR, fg=TEXT_COLOR, font=HMI_FONT_S).pack(anchor="w")
        ent_pass = ttk.Entry(left_frame, width=30, show="*")
        ent_pass.pack(pady=(0, 15))
        
        tk.Label(left_frame, text="Role:", bg=BG_COLOR, fg=TEXT_COLOR, font=HMI_FONT_S).pack(anchor="w")
        cb_role = ttk.Combobox(left_frame, values=["Operator", "Quality OP", "Shift Leader", "Supervisor", "Manager", "Admin"], state="readonly", width=28)
        cb_role.set("Operator")
        cb_role.pack(pady=(0, 20))
        
        right_frame = tk.Frame(main_frame, bg=BG_COLOR)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        tk.Label(right_frame, text="Current Users", bg=BG_COLOR, fg=TEXT_COLOR, font=HMI_FONT_L).pack(pady=(0, 15))
        
        cols = ("ID", "Role")
        tree = ttk.Treeview(right_frame, columns=cols, show="headings")
        for col in cols:
            tree.heading(col, text=col)
        tree.column("ID", width=200)
        tree.column("Role", width=150)
        
        sb = ttk.Scrollbar(right_frame, orient="vertical", command=tree.yview)
        tree.configure(yscroll=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        tree.pack(fill=tk.BOTH, expand=True)
        
        def load_users():
            for item in tree.get_children():
                tree.delete(item)
            try:
                conn = get_db_connection()
                c = conn.cursor()
                c.execute("SELECT id, role FROM auth ORDER BY role")
                for row in c.fetchall():
                    tree.insert("", "end", values=row)
                conn.close()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load users: {e}", parent=top)
                
        def add_user():
            uid = ent_id.get().strip()
            upass = ent_pass.get().strip()
            role = cb_role.get()
            if not uid or not upass:
                messagebox.showerror("Error", "ID and Password are required.", parent=top)
                return
            try:
                conn = get_db_connection()
                c = conn.cursor()
                c.execute("INSERT INTO auth (id, password, role) VALUES (?, ?, ?)", (uid, hash_password(upass), role))
                conn.commit()
                conn.close()
                ent_id.delete(0, tk.END)
                ent_pass.delete(0, tk.END)
                load_users()
                messagebox.showinfo("Success", "User added successfully.", parent=top)
            except sqlite3.IntegrityError:
                messagebox.showerror("Error", "User ID already exists.", parent=top)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to add user: {e}", parent=top)
                
        def delete_user():
            sel = tree.selection()
            if not sel:
                return
            uid = tree.item(sel[0])["values"][0]
            if uid in ["999", "998", "111", "mg90", self.app_user_id]:
                messagebox.showwarning("Restricted", "Cannot delete default or active users.", parent=top)
                return
            if messagebox.askyesno("Confirm Delete", f"Delete user '{uid}'?", parent=top):
                try:
                    conn = get_db_connection()
                    c = conn.cursor()
                    c.execute("DELETE FROM auth WHERE id=?", (uid,))
                    conn.commit()
                    conn.close()
                    load_users()
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to delete user: {e}", parent=top)

        btn_add = ttk.Button(left_frame, text="Add User", style="Success.TButton", command=add_user)
        btn_add.pack(fill=tk.X, pady=5)
        
        btn_del = ttk.Button(right_frame, text="Delete Selected", style="Danger.TButton", command=delete_user)
        btn_del.pack(side=tk.RIGHT, pady=10)
        
        load_users()

    def open_settings(self):
        top = tk.Toplevel(self)
        top.title("Settings")
        w = 600 if self.is_admin else 450
        h = 400 if self.is_admin else 200
        center_window(top, w, h)
        
        top.configure(bg=BG_COLOR)
        top.transient(self)
        top.grab_set()
        
        tk.Label(top, text="Printer Settings", bg=BG_COLOR, fg=TEXT_COLOR, font=HMI_FONT_L).pack(pady=(10, 5))
        
        frame = tk.Frame(top, bg=BG_COLOR)
        frame.pack(pady=10)
        
        tk.Label(frame, text="Zebra Printer:", bg=BG_COLOR, fg=TEXT_COLOR, font=HMI_FONT_M).pack(side=tk.LEFT)
        
        printers = []
        if WIN32_PRINT_AVAILABLE:
            try:
                flags = 2 | 4
                printers = [p[2] for p in win32print.EnumPrinters(flags)]
            except Exception as e:
                pass
        
        cb = ttk.Combobox(frame, values=printers, state="readonly", width=35)
        cb.pack(side=tk.LEFT, padx=10)
        
        config = {}
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r") as f:
                    config = json.load(f)
            except: pass
            
        if config.get("zebra_printer") in printers:
            cb.set(config.get("zebra_printer"))
        elif printers:
            cb.set(printers[0])
            
        def save():
            config["zebra_printer"] = cb.get()
            with open(CONFIG_FILE, "w") as f:
                json.dump(config, f)
            top.destroy()
            messagebox.showinfo("Saved", "Settings saved successfully.")
            
        ttk.Button(top, text="Save Settings", style="Success.TButton", command=save).pack(pady=10)
        
        if self.is_admin or self.app_user_role in ["Supervisor", "Manager"]:
            tk.Frame(top, bg=BORDER_COLOR, height=2).pack(fill=tk.X, padx=20, pady=10)
            tk.Label(top, text="Management Tools", bg=BG_COLOR, fg=TEXT_COLOR, font=HMI_FONT_L).pack(pady=(5, 5))
            
            admin_frame = tk.Frame(top, bg=BG_COLOR)
            admin_frame.pack(fill=tk.X, padx=20, pady=10)
            
            ttk.Button(admin_frame, text="Manage Targets", style="Success.TButton", command=self.open_targets_manager).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
            if self.is_admin:
                ttk.Button(admin_frame, text="Manage Products", style="Primary.TButton", command=self.open_product_manager).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
                ttk.Button(admin_frame, text="Audit Logs", style="Primary.TButton", command=self.open_logs_manager).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
                ttk.Button(admin_frame, text="Manage Users", style="Warning.TButton", command=self.open_user_manager).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

    def open_targets_manager(self):
        top = tk.Toplevel(self)
        top.title("Production Targets Management")
        center_window(top, 800, 600)
        top.configure(bg=BG_COLOR)
        top.transient(self)
        top.grab_set()
        
        main_frame = tk.Frame(top, bg=BG_COLOR)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        form_frame = tk.LabelFrame(main_frame, text="Add New Target", bg=BG_COLOR, fg=TEXT_COLOR)
        form_frame.pack(fill=tk.X, pady=(0, 20))
        
        tk.Label(form_frame, text="Product PN:", bg=BG_COLOR, fg=TEXT_COLOR).grid(row=0, column=0, padx=5, pady=5)
        var_pn = tk.StringVar()
        cb_pn = ttk.Combobox(form_frame, textvariable=var_pn, values=list(SF_DATA.keys()), state="normal")
        cb_pn.grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(form_frame, text="Target Qty:", bg=BG_COLOR, fg=TEXT_COLOR).grid(row=0, column=2, padx=5, pady=5)
        var_qty = tk.StringVar()
        entry_qty = ttk.Entry(form_frame, textvariable=var_qty)
        entry_qty.grid(row=0, column=3, padx=5, pady=5)
        
        def refresh_tree():
            for item in tree.get_children(): tree.delete(item)
            conn = get_db_connection()
            c = conn.cursor()
            c.execute("SELECT id, product_pn, target_qty, effective_date FROM shift_targets ORDER BY id DESC")
            for row in c.fetchall():
                tree.insert("", "end", values=row)
            conn.close()
            
        def add_target():
            pn = var_pn.get().strip()
            shift = "All"
            station = "All"
            qty_str = var_qty.get()
            
            if not pn or not qty_str.isdigit():
                messagebox.showerror("Error", "Invalid PN or Quantity")
                return
                
            qty = int(qty_str)
            now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            conn = get_db_connection()
            c = conn.cursor()
            c.execute("INSERT INTO shift_targets (product_pn, shift, station, target_qty, effective_date) VALUES (?, ?, ?, ?, ?)",
                      (pn, shift, station, qty, now_str))
            conn.commit()
            conn.close()
            
            var_qty.set("")
            refresh_tree()
            if hasattr(self, 'refresh_kpis'):
                self.refresh_kpis()
            
        def delete_target():
            selected = tree.selection()
            if not selected: return
            item = tree.item(selected[0])
            tid = item['values'][0]
            
            if messagebox.askyesno("Confirm", "Delete selected target?"):
                conn = get_db_connection()
                c = conn.cursor()
                c.execute("DELETE FROM shift_targets WHERE id=?", (tid,))
                conn.commit()
                conn.close()
                refresh_tree()
                if hasattr(self, 'refresh_kpis'):
                    self.refresh_kpis()
                
        ttk.Button(form_frame, text="Add Target", style="Success.TButton", command=add_target).grid(row=0, column=4, padx=10, pady=5)
        
        cols = ("ID", "PN", "Target Qty", "Effective Date")
        tree = ttk.Treeview(main_frame, columns=cols, show="headings")
        for c in cols: tree.heading(c, text=c)
        tree.pack(fill=tk.BOTH, expand=True, pady=10)
        
        ttk.Button(main_frame, text="Delete Selected", style="Danger.TButton", command=delete_target).pack(pady=5)
        
        refresh_tree()


    def open_product_manager(self):
        top = tk.Toplevel(self)
        top.title("Product Management (Admin)")
        center_window(top, 900, 600)
        top.resizable(False, False)
        
        top.update_idletasks()
        x = (top.winfo_screenwidth() // 2) - (900 // 2)
        y = (top.winfo_screenheight() // 2) - (600 // 2)
        top.geometry(f"+{x}+{y}")
        
        top.configure(bg=BG_COLOR)
        top.transient(self)
        top.grab_set()
        
        main_frame = tk.Frame(top, bg=BG_COLOR)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        left_frame = tk.Frame(main_frame, bg=BG_COLOR, width=300)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        
        right_frame = tk.Frame(main_frame, bg=BG_COLOR)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        tk.Label(left_frame, text="Existing Products (SF_PN)", bg=BG_COLOR, fg=TEXT_COLOR, font=HMI_FONT_M).pack(pady=(0, 10))
        
        listbox = tk.Listbox(left_frame, font=HMI_FONT_S, width=35, bg=SURFACE_COLOR, fg=TEXT_COLOR, selectbackground=ACCENT_COLOR)
        listbox.pack(fill=tk.BOTH, expand=True)
        
        def refresh_list():
            listbox.delete(0, tk.END)
            for pn in SF_DATA.keys():
                listbox.insert(tk.END, pn)
                
        refresh_list()
        
        tk.Label(right_frame, text="Product Details", bg=BG_COLOR, fg=TEXT_COLOR, font=HMI_FONT_L).pack(pady=(0, 10))
        
        form_frame = tk.Frame(right_frame, bg=SURFACE_COLOR, padx=20, pady=20)
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        top_form = tk.Frame(form_frame, bg=SURFACE_COLOR)
        top_form.pack(pady=10)
        
        tk.Label(top_form, text="SF Part Number (Key):", bg=SURFACE_COLOR, fg=TEXT_COLOR, font=HMI_FONT_M).grid(row=0, column=0, sticky="e", pady=10, padx=5)
        var_sf_pn = tk.StringVar()
        entry_sf_pn = ttk.Entry(top_form, textvariable=var_sf_pn, width=30, font=HMI_FONT_M)
        entry_sf_pn.grid(row=0, column=1, sticky="w", pady=10)
        
        tk.Label(top_form, text="Part Name (SF):", bg=SURFACE_COLOR, fg=TEXT_COLOR, font=HMI_FONT_M).grid(row=1, column=0, sticky="e", pady=10, padx=5)
        var_part_name = tk.StringVar()
        entry_part_name = ttk.Entry(top_form, textvariable=var_part_name, width=30, font=HMI_FONT_M)
        entry_part_name.grid(row=1, column=1, sticky="w", pady=10)
        
        tk.Frame(form_frame, height=2, bg=BORDER_COLOR).pack(fill=tk.X, pady=15)
        
        bottom_form = tk.Frame(form_frame, bg=SURFACE_COLOR)
        bottom_form.pack(pady=5)
        
        tk.Label(bottom_form, text="Raw Materials (Up to 4)", bg=SURFACE_COLOR, fg=TEXT_MUTED, font=HMI_FONT_S).grid(row=0, column=0, columnspan=4, sticky="w", pady=5)
        
        rm_vars = []
        rm_entries = []
        for i in range(4):
            tk.Label(bottom_form, text=f"RM {i+1} PN:", bg=SURFACE_COLOR, fg=TEXT_COLOR).grid(row=1+i, column=0, sticky="e", pady=5, padx=5)
            v_pn = tk.StringVar()
            e1 = ttk.Entry(bottom_form, textvariable=v_pn, width=25)
            e1.grid(row=1+i, column=1, sticky="w", pady=5)
            
            tk.Label(bottom_form, text=f"RM {i+1} Name:", bg=SURFACE_COLOR, fg=TEXT_COLOR).grid(row=1+i, column=2, sticky="e", pady=5, padx=5)
            v_name = tk.StringVar()
            e2 = ttk.Entry(bottom_form, textvariable=v_name, width=25)
            e2.grid(row=1+i, column=3, sticky="w", pady=5)
            
            rm_vars.append((v_pn, v_name))
            rm_entries.append((e1, e2))
            
        def set_form_state(state):
            entry_sf_pn.config(state=state)
            entry_part_name.config(state=state)
            for e1, e2 in rm_entries:
                e1.config(state=state)
                e2.config(state=state)
                
        def unlock_form():
            if messagebox.askyesno("Confirm Unlock", "Are you sure you want to unlock and edit this existing product?", parent=top):
                set_form_state("normal")
                btn_unlock.config(state="disabled")
                btn_save.config(state="normal")
                btn_delete.config(state="normal")
                btn_clear.config(state="normal")
            
        def clear_form():
            set_form_state("normal")
            var_sf_pn.set("")
            var_part_name.set("")
            for v_pn, v_name in rm_vars:
                v_pn.set("")
                v_name.set("")
            listbox.selection_clear(0, tk.END)
            btn_unlock.config(state="disabled")
            btn_save.config(state="normal")
            btn_delete.config(state="disabled")
            btn_clear.config(state="normal")
            
        def on_select(evt):
            if not listbox.curselection(): return
            sel_pn = listbox.get(listbox.curselection())
            var_sf_pn.set(sel_pn)
            val = SF_DATA[sel_pn]
            var_part_name.set(val[0])
            rms = val[1]
            for i in range(4):
                if i < len(rms):
                    rm_vars[i][0].set(rms[i][0])
                    rm_vars[i][1].set(rms[i][1])
                else:
                    rm_vars[i][0].set("")
                    rm_vars[i][1].set("")
                    
            set_form_state("disabled")
            btn_unlock.config(state="normal")
            btn_save.config(state="disabled")
            btn_delete.config(state="disabled")
            btn_clear.config(state="normal")
                    
        listbox.bind("<<ListboxSelect>>", on_select)
        
        def save_product():
            pn = var_sf_pn.get().strip()
            name = var_part_name.get().strip()
            if not pn or not name:
                messagebox.showerror("Error", "SF Part Number and Name are required.", parent=top)
                return
                
            is_new = pn not in SF_DATA
            msg = f"Are you sure you want to add the new product '{pn}'?" if is_new else f"Are you sure you want to update the existing product '{pn}'?"
            if not messagebox.askyesno("Confirm Save", msg, parent=top):
                return
            
            rms = []
            for v_pn, v_name in rm_vars:
                r_pn = v_pn.get().strip()
                r_name = v_name.get().strip()
                if r_pn and r_name:
                    rms.append((r_pn, r_name))
            
            SF_DATA[pn] = (name, rms)
            
            try:
                timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                conn = get_db_connection()
                c = conn.cursor()
                c.execute("INSERT OR REPLACE INTO products (pn_sf, name_sf, rms_json) VALUES (?, ?, ?)", (pn, name, json.dumps(rms)))
                
                action = "ADD" if is_new else "UPDATE"
                details = f"Name: {name}, RMs: {len(rms)}"
                c.execute("INSERT INTO product_audit_trail (action, pn_sf, details, user_id, shift, timestamp) VALUES (?, ?, ?, ?, ?, ?)",
                          (action, pn, details, getattr(self, 'app_user_id', 'Unknown'), getattr(self, 'app_user_shift', 'Unknown'), timestamp))
                conn.commit()
                conn.close()
            except Exception as e:
                messagebox.showerror("DB Error", f"Failed to update database:\n{e}", parent=top)
                
            self.populate_sf_combobox()
            if hasattr(self, 'pl_cb_sf_pn'):
                self.pl_cb_sf_pn['values'] = list(SF_DATA.keys())
            refresh_list()
            messagebox.showinfo("Success", f"Product '{pn}' saved successfully.", parent=top)
            
            # Auto-lock form and allow creating new product
            set_form_state("disabled")
            btn_unlock.config(state="normal")
            btn_save.config(state="disabled")
            btn_delete.config(state="disabled")
            btn_clear.config(state="normal")
            
        def delete_product():
            pn = var_sf_pn.get().strip()
            if not pn: return
            if pn in SF_DATA:
                if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete product '{pn}'?", parent=top):
                    del SF_DATA[pn]
                    
                    try:
                        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        conn = get_db_connection()
                        c = conn.cursor()
                        c.execute("DELETE FROM products WHERE pn_sf = ?", (pn,))
                        c.execute("INSERT INTO product_audit_trail (action, pn_sf, details, user_id, shift, timestamp) VALUES (?, ?, ?, ?, ?, ?)",
                                  ("DELETE", pn, "", getattr(self, 'app_user_id', 'Unknown'), getattr(self, 'app_user_shift', 'Unknown'), timestamp))
                        conn.commit()
                        conn.close()
                    except Exception as e:
                        messagebox.showerror("DB Error", f"Failed to delete from database:\n{e}", parent=top)
                        
                    self.populate_sf_combobox()
                    if hasattr(self, 'pl_cb_sf_pn'):
                        self.pl_cb_sf_pn['values'] = list(SF_DATA.keys())
                    refresh_list()
                    clear_form()
                    messagebox.showinfo("Deleted", f"Product '{pn}' deleted.", parent=top)
            
        btn_frame = tk.Frame(right_frame, bg=BG_COLOR)
        btn_frame.pack(fill=tk.X, pady=15)
        
        btn_clear = ttk.Button(btn_frame, text="Add New Product ", style="Secondary.TButton", command=clear_form)
        btn_clear.pack(side=tk.LEFT, padx=5)
        
        btn_unlock = ttk.Button(btn_frame, text="Unlock ", style="Danger.TButton", command=unlock_form)
        btn_unlock.pack(side=tk.LEFT, padx=5)
        btn_unlock.config(state="disabled")
        
        btn_delete = ttk.Button(btn_frame, text="Delete Product", style="Warning.TButton", command=delete_product)
        btn_delete.pack(side=tk.RIGHT, padx=5)
        btn_delete.config(state="disabled")
        
        btn_save = ttk.Button(btn_frame, text="Save Product", style="Success.TButton", command=save_product)
        btn_save.pack(side=tk.RIGHT, padx=5)

    def open_logs_manager(self):
        top = tk.Toplevel(self)
        top.title("System Audit Logs")
        center_window(top, 1000, 600)
        top.resizable(False, False)
        
        top.update_idletasks()
        x = (top.winfo_screenwidth() // 2) - (1000 // 2)
        y = (top.winfo_screenheight() // 2) - (600 // 2)
        top.geometry(f"+{x}+{y}")
        
        top.configure(bg=BG_COLOR)
        top.transient(self)
        top.grab_set()
        
        main_frame = tk.Frame(top, bg=BG_COLOR)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        tk.Label(main_frame, text="System Audit Logs", bg=BG_COLOR, fg=TEXT_COLOR, font=HMI_FONT_L).pack(pady=(0, 15))
        
        log_notebook = ttk.Notebook(main_frame)
        log_notebook.pack(fill=tk.BOTH, expand=True)
        
        tab_products = ttk.Frame(log_notebook)
        tab_access = ttk.Frame(log_notebook)
        
        log_notebook.add(tab_products, text="Product Modifications")
        log_notebook.add(tab_access, text="System Access Logs")
        
        cols_prod = ("ID", "Action", "Part Number", "Details", "User ID", "Shift", "Timestamp")
        tree_prod = ttk.Treeview(tab_products, columns=cols_prod, show="headings")
        for col in cols_prod:
            tree_prod.heading(col, text=col)
            tree_prod.column(col, width=120)
        tree_prod.column("Details", width=250)
        tree_prod.column("ID", width=50)
        
        sb_prod = ttk.Scrollbar(tab_products, orient="vertical", command=tree_prod.yview)
        tree_prod.configure(yscroll=sb_prod.set)
        sb_prod.pack(side=tk.RIGHT, fill=tk.Y)
        tree_prod.pack(fill=tk.BOTH, expand=True)
        
        cols_access = ("ID", "Event", "User ID", "Shift", "Timestamp")
        tree_access = ttk.Treeview(tab_access, columns=cols_access, show="headings")
        for col in cols_access:
            tree_access.heading(col, text=col)
            tree_access.column(col, width=150)
        tree_access.column("ID", width=50)
        
        sb_access = ttk.Scrollbar(tab_access, orient="vertical", command=tree_access.yview)
        tree_access.configure(yscroll=sb_access.set)
        sb_access.pack(side=tk.RIGHT, fill=tk.Y)
        tree_access.pack(fill=tk.BOTH, expand=True)
        
        def refresh_logs():
            for item in tree_prod.get_children():
                tree_prod.delete(item)
            for item in tree_access.get_children():
                tree_access.delete(item)
                
            try:
                conn = get_db_connection()
                c = conn.cursor()
                c.execute("SELECT id, action, pn_sf, details, user_id, shift, timestamp FROM product_audit_trail ORDER BY id DESC")
                for row in c.fetchall():
                    tree_prod.insert("", "end", values=row)
                    
                try:
                    c.execute("SELECT id, event_type, user_id, shift, timestamp FROM system_access_logs ORDER BY id DESC")
                    for row in c.fetchall():
                        tree_access.insert("", "end", values=row)
                except sqlite3.OperationalError:
                    pass
                conn.close()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to fetch logs: {e}", parent=top)
                
        btn_frame = tk.Frame(main_frame, bg=BG_COLOR)
        btn_frame.pack(fill=tk.X, pady=(10,0))
        ttk.Button(btn_frame, text="Refresh", style="Secondary.TButton", command=refresh_logs).pack(side=tk.RIGHT)
        
        refresh_logs()

    def build_ui(self):
        HEADER_BG = "#FFFFFF"
        HEADER_FG = "#005A8C"
        
        self.header_frame = tk.Frame(self, bg=HEADER_BG, height=70)
        self.header_frame.pack(side=tk.TOP, fill=tk.X)
        self.header_frame.pack_propagate(False)
        
        try:
            logo_path = resource_path(os.path.join("assets", "logo_en.png"))
            if os.path.exists(logo_path):
                try:
                    from PIL import Image, ImageTk
                    img = Image.open(logo_path)
                    
                    ratio = 50 / img.height
                    new_size = (int(img.width * ratio), 50)
                    if hasattr(Image, "Resampling"):
                        res_filter = Image.Resampling.LANCZOS
                    else:
                        res_filter = Image.ANTIALIAS
                    img = img.resize(new_size, res_filter)
                    self.logo_img = ImageTk.PhotoImage(img)
                    tk.Label(self.header_frame, image=self.logo_img, bg=HEADER_BG).pack(side=tk.LEFT, padx=(20, 10))
                except Exception as pil_e:
                    self.logo_img = tk.PhotoImage(file=logo_path)
                    tk.Label(self.header_frame, image=self.logo_img, bg=HEADER_BG).pack(side=tk.LEFT, padx=(20, 10))
            
            tk.Label(self.header_frame, text="SUB-PROCESS TRACEABILITY", bg=HEADER_BG, fg=HEADER_FG, font=HMI_FONT_L).pack(side=tk.LEFT)
        except Exception as e:
            tk.Label(self.header_frame, text="HI-LEX ACT - SUB-PROCESS TRACEABILITY", bg=HEADER_BG, fg=HEADER_FG, font=HMI_FONT_L).pack(side=tk.LEFT, padx=20)
        
        self.settings_btn = ttk.Button(self.header_frame, text="Settings", style="Header.TButton", command=self.open_settings)
        self.settings_btn.pack(side=tk.RIGHT, padx=20)
        
        def do_logout():
            self.perform_logout(force=False)
                
        self.logout_btn = ttk.Button(self.header_frame, text="Logout", style="Danger.TButton", command=do_logout)
        self.logout_btn.pack(side=tk.RIGHT, padx=5)
        
        self.lbl_header_user = tk.Label(self.header_frame, text="User: - | Shift: -", bg=HEADER_BG, fg=TEXT_MUTED, font=HMI_FONT_S)
        self.lbl_header_user.pack(side=tk.RIGHT, padx=10)
        
        self.lbl_header_warning = tk.Label(self.header_frame, text=" Please don't forget to logout!", bg=HEADER_BG, fg=ERROR_COLOR, font=HMI_FONT_M)
        
        self.lbl_quality_alert = tk.Label(self.header_frame, text=" QUALITY ALERT: High Defect Rate!", bg=HEADER_BG, fg=ERROR_COLOR, font=HMI_FONT_M)
        
        #tk.Label(self.header_frame, text="Developed by Naoufal El Hlou", bg=HEADER_BG, fg="#94A3B8", font=("Segoe UI", 9, "italic")).pack(side=tk.RIGHT, padx=10)
        
        try:
            settings_path = resource_path(os.path.join("assets", "settings_icon.png"))
            if os.path.exists(settings_path):
                from PIL import Image, ImageTk
                s_img = Image.open(settings_path).convert("RGBA")
                data = s_img.getdata()
                new_data = []
                for item in data:
                    if item[3] > 10:
                        new_data.append((255, 255, 255, item[3]))
                    else:
                        new_data.append((255, 255, 255, 0))
                s_img.putdata(new_data)
                
                if hasattr(Image, "Resampling"):
                    res_filter = Image.Resampling.LANCZOS
                else:
                    res_filter = Image.ANTIALIAS
                s_img = s_img.resize((20, 20), res_filter)
                self.settings_icon = ImageTk.PhotoImage(s_img)
                self.settings_btn.config(image=self.settings_icon, compound=tk.LEFT, text=" Settings")
        except Exception as e:
            print("Could not load settings icon:", e)
        
        # PanedWindow
        self.paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        self.paned.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Sidebar
        self.sidebar = tk.Frame(self.paned, bg=SURFACE_COLOR, width=180)
        self.sidebar.pack_propagate(False)
        self.paned.add(self.sidebar, weight=0)
        
        self.lbl_clock = tk.Label(self.sidebar, text="00:00", bg=SURFACE_COLOR, fg=ACCENT_COLOR, font=("Consolas", 28, "bold"))
        self.lbl_clock.pack(pady=(20,0))
        self.lbl_date = tk.Label(self.sidebar, text="YYYY-MM-DD", bg=SURFACE_COLOR, fg=TEXT_MUTED, font=HMI_FONT_S)
        self.lbl_date.pack(pady=(0, 20))

        
        tk.Label(self.sidebar, text="SHIFT OVERVIEW", bg=SURFACE_COLOR, fg=TEXT_COLOR, font=HMI_FONT_M).pack()
        
        def create_side_kpi(parent, title):
            f = tk.Frame(parent, bg=BG_COLOR, highlightbackground=BORDER_COLOR, highlightthickness=1)
            f.pack(fill=tk.X, padx=15, pady=5)
            lbl_title = tk.Label(f, text=title, bg=BG_COLOR, fg=TEXT_MUTED, font=HMI_FONT_S)
            lbl_title.pack(pady=(5,0))
            lbl_val = tk.Label(f, text="0", bg=BG_COLOR, fg=TEXT_COLOR, font=HMI_FONT_L)
            lbl_val.pack(pady=(0,5))
            return lbl_title, lbl_val
            
        self.lbl_title_recs, self.side_stat_recs = create_side_kpi(self.sidebar, "Records (Shift -)")
        self.lbl_title_qty, self.side_stat_qty = create_side_kpi(self.sidebar, "Qty (Shift -)")
        _, self.side_stat_today = create_side_kpi(self.sidebar, "Total Today")
        
        tk.Label(self.sidebar, text="RECENT PNs", bg=SURFACE_COLOR, fg=TEXT_COLOR, font=HMI_FONT_M).pack(pady=(10,0))
        self.recent_pns_listbox = tk.Listbox(self.sidebar, bg=BG_COLOR, fg=TEXT_COLOR, bd=0, relief="flat", height=8, font=HMI_FONT_S)
        self.recent_pns_listbox.pack(fill=tk.X, padx=15, pady=10)
        
        ttk.Button(self.sidebar, text="Open Excel", style="Secondary.TButton", command=self.open_excel).pack(fill=tk.X, padx=10, pady=2)
        ttk.Button(self.sidebar, text="Print Last Slip", style="Warning.TButton", command=self.print_last_slip).pack(fill=tk.X, padx=10, pady=2)
        
        # Notebook
        self.notebook = ttk.Notebook(self.paned)
        self.paned.add(self.notebook, weight=1)
        
        self.tab1 = ttk.Frame(self.notebook)
        self.tab5 = ttk.Frame(self.notebook)
        self.tab4 = ttk.Frame(self.notebook)
        self.tab2 = ttk.Frame(self.notebook)
        self.tab_inventory = ttk.Frame(self.notebook)
        self.tab3 = ttk.Frame(self.notebook)
        
        self.notebook.add(self.tab1, text="New Entry")
        self.notebook.add(self.tab5, text="Consume to Line")
        self.notebook.add(self.tab4, text="Search & Reprint")
        self.notebook.add(self.tab2, text="Records")
        self.notebook.add(self.tab_inventory, text="Live Inventory")
        self.notebook.add(self.tab3, text="KPIs")
        self.notebook.hide(self.tab3)
        
        self.build_tab1()
        self.build_tab5()
        self.build_tab4()
        self.build_tab2()
        self.build_tab_inventory()
        self.build_tab3()
        
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        self._check_quality_rate()

    def perform_logout(self, force=False):
        if not force:
            if not messagebox.askyesno("Confirm Logout", "Are you sure you want to log out?", parent=self):
                return
                
        if hasattr(self, 'app_user_id') and self.app_user_id:
            try:
                conn = get_db_connection()
                c = conn.cursor()
                timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                c.execute("INSERT INTO system_access_logs (event_type, user_id, shift, timestamp) VALUES (?, ?, ?, ?)",
                          ("LOGOUT", self.app_user_id, self.app_user_shift, timestamp))
                conn.commit()
                conn.close()
            except Exception as e:
                pass

        self.app_user_id = ""
        self.app_user_shift = ""
        self.is_admin = False
        self.lbl_header_user.config(text="User: - | Shift: -")
        if hasattr(self, 'lbl_header_warning'):
            self.lbl_header_warning.pack_forget()
        if hasattr(self, 'notebook') and hasattr(self, 'tab3'):
            self.notebook.hide(self.tab3)
        self.prompt_login()
        
    def _on_tab_changed(self, event):
        try:
            selected_tab = self.notebook.tab(self.notebook.select(), "text")
            if selected_tab == "Records":
                self.refresh_records_treeview()
            elif selected_tab == "Live Inventory":
                self.refresh_inventory()
            elif selected_tab == "KPIs":
                self.refresh_kpis()
        except Exception:
            pass
            
    def _check_quality_rate(self):
        try:
            conn = get_db_connection()
            c = conn.cursor()
            today = datetime.datetime.now().strftime("%Y-%m-%d")
            
            c.execute("SELECT COUNT(*) FROM quality_defects WHERE reported_at LIKE ?", (f"{today}%",))
            krow = c.fetchone()
            total_def_today = krow[0] or 0
            
            c.execute("SELECT SUM(quantity) FROM records WHERE dt_sp LIKE ?", (f"{today}%",))
            prod_res = c.fetchone()
            total_prod = prod_res[0] or 0
            
            conn.close()
            
            if total_prod > 0:
                rate = (total_def_today / total_prod) * 100
                if rate > 3.0:
                    self.lbl_quality_alert.pack(side=tk.LEFT, padx=20)
                else:
                    self.lbl_quality_alert.pack_forget()
            else:
                self.lbl_quality_alert.pack_forget()
        except Exception:
            pass
            
        try:
            self.after(60000, self._check_quality_rate)
        except Exception:
            pass
    def build_tab1(self):
        # Container Split
        top_panel = tk.Frame(self.tab1, bg=BG_COLOR)
        top_panel.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        
        bottom_panel = tk.Frame(self.tab1, bg=BG_COLOR)
        bottom_panel.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)

        # 2-Column Layout
        self.t1_left = tk.Frame(top_panel, bg=BG_COLOR, width=400)
        self.t1_left.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=5)
        self.t1_left.pack_propagate(False) # Keep fixed width for left panel
        
        separator = ttk.Separator(top_panel, orient='vertical')
        separator.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=10)
        
        self.t1_right = tk.Frame(top_panel, bg=BG_COLOR)
        self.t1_right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=5)

        # ---------------- LEFT PANEL ----------------
        # Shift Target Tracker
        _, tracker_frame = self.create_card(self.t1_left, "Shift Target Tracker", fg_color=ACCENT_COLOR)
        
        tracker_frame.config(height=260)
        tracker_frame.pack_propagate(False)
        
        self.tracker_canvas = tk.Canvas(tracker_frame, bg=SURFACE_COLOR, highlightthickness=0)
        tracker_scrollbar = ttk.Scrollbar(tracker_frame, orient="vertical", command=self.tracker_canvas.yview)
        
        self.tracker_scroll_frame = tk.Frame(self.tracker_canvas, bg=SURFACE_COLOR)
        
        self.tracker_scroll_frame.bind(
            "<Configure>",
            lambda e: self.tracker_canvas.configure(scrollregion=self.tracker_canvas.bbox("all"))
        )
        
        self.tracker_canvas.create_window((0, 0), window=self.tracker_scroll_frame, anchor="nw", width=330)
        self.tracker_canvas.configure(yscrollcommand=tracker_scrollbar.set)
        
        self.tracker_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2, pady=5)
        tracker_scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=5)
        
        style = ttk.Style()
        style.configure("TProgressbar", thickness=15)
        style.configure("Safe.Horizontal.TProgressbar", background=SUCCESS_COLOR, troughcolor=BG_COLOR, thickness=15)
        style.configure("Warn.Horizontal.TProgressbar", background=WARNING_COLOR, troughcolor=BG_COLOR, thickness=15)
        style.configure("Danger.Horizontal.TProgressbar", background=ERROR_COLOR, troughcolor=BG_COLOR, thickness=15)
        
        self.lbl_scroll_indicator = tk.Label(self.t1_left, text=" Scroll down for more refs ", bg=BG_COLOR, fg=TEXT_MUTED, font=HMI_FONT_S)
        # Quick Search
        _, search_frame = self.create_card(self.t1_left, "Quick Search")
        self.sv_search = tk.StringVar()
        self.sv_search.trace_add("write", self.filter_sf_combobox)
        search_entry = ttk.Entry(search_frame, textvariable=self.sv_search, width=20)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        self.lbl_search_count = tk.Label(search_frame, text="30", fg=ACCENT_COLOR, bg=BG_COLOR)
        self.lbl_search_count.pack(side=tk.LEFT, padx=5)
        # ttk.Button(search_frame, text="❌​", width=3, command=lambda: self.sv_search.set("")).pack(side=tk.LEFT)
        
        # Status Label
        self.lbl_status = tk.Label(self.t1_left, text="", bg=BG_COLOR, fg=SUCCESS_COLOR, font=HMI_FONT_M)
        self.lbl_status.pack(pady=10)
        
        btn_frame = tk.Frame(self.t1_left, bg=BG_COLOR)
        btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=10, padx=10)
        
        ttk.Button(btn_frame, text="Save Record", style="Success.TButton", command=self.save_record).pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)
        
        sub_btn = tk.Frame(btn_frame, bg=BG_COLOR)
        sub_btn.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)
        
        ttk.Button(sub_btn, text="Same as Last", style="Primary.TButton", command=self.same_as_last).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(sub_btn, text="Clear Form", style="Secondary.TButton", command=self.clear_form).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))
        
        # ---------------- BOTTOM PANEL ----------------
        # Recent Records Tree
        _, tree_frame = self.create_card(bottom_panel, "Last 5 Records (Double-click to print)")
        
        cols = ("SB_ID", "SF_PN", "Qty", "Shift", "Station", "DateTime")
        self.tree_recent = ttk.Treeview(tree_frame, columns=cols, show="headings", height=5)
        for c in cols:
            self.tree_recent.heading(c, text=c)
            self.tree_recent.column(c, width=100)
        self.tree_recent.pack(fill=tk.X)
        self.tree_recent.bind("<Double-1>", self.on_recent_double_click)
        
        self.tree_recent.tag_configure("even", background="#141A20")
        self.tree_recent.tag_configure("odd", background="#1B232C")
        

        # ---------------- RIGHT PANEL (FORM) ----------------
        self.form_frame = tk.Frame(self.t1_right, bg=BG_COLOR)
        self.form_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Section 1: Part Reference
        _, lf1 = self.create_card(self.form_frame, "Part Reference")
        
        ttk.Label(lf1, text="Scan Main RM Barcode", width=32).grid(row=0, column=0, sticky="w", pady=2)
        self.var_scan_rm_t1 = tk.StringVar()
        self.entry_scan_rm_t1 = ttk.Entry(lf1, textvariable=self.var_scan_rm_t1, width=30)
        self.entry_scan_rm_t1.grid(row=0, column=1, sticky="w", pady=2, padx=5)
        self.entry_scan_rm_t1.bind("<Return>", self.on_rm_scanned_t1)
        
        self.lbl_scan_rm_status = tk.Label(lf1, text="", bg=SURFACE_COLOR, font=HMI_FONT_S)
        self.lbl_scan_rm_status.grid(row=0, column=2, columnspan=2, sticky="w", padx=10)
        
        ttk.Label(lf1, text="FULL PN Semi fini ", width=32).grid(row=1, column=0, sticky="w", pady=2)
        self.cb_sf_pn = ttk.Combobox(lf1, state="readonly", width=30)
        self.cb_sf_pn.grid(row=1, column=1, sticky="w", pady=2, padx=5)
        self.cb_sf_pn.bind("<<ComboboxSelected>>", self.on_sf_selected)
        
        ttk.Label(lf1, text="PART NAME (SF)").grid(row=1, column=2, sticky="w", pady=2, padx=10)
        self.var_part_sf = tk.StringVar()
        ttk.Entry(lf1, textvariable=self.var_part_sf, state="disabled", width=30).grid(row=1, column=3, sticky="w", pady=2)
        
        # Dynamic RM Container T1
        self.rm_container_t1 = tk.Frame(lf1, bg=SURFACE_COLOR)
        self.rm_container_t1.grid(row=2, column=0, columnspan=4, sticky="w", pady=5)
        self.rm_vars_t1 = []
        
        # Section 2: Batch & Qty
        _, lf2 = self.create_card(self.form_frame, "Batch Numbers & Quantity")
        
        ttk.Label(lf2, text="Batch No. 1 / 2 / 3", width=32).grid(row=0, column=0, sticky="w", pady=2)
        b_frame = tk.Frame(lf2, bg=SURFACE_COLOR)
        b_frame.grid(row=0, column=1, sticky="w", pady=2, padx=5)
        self.var_b1 = tk.StringVar()
        self.var_b2 = tk.StringVar()
        self.var_b3 = tk.StringVar()
        self.var_b1.trace_add("write", lambda *args: self.var_b1.set(self.var_b1.get().upper()))
        self.var_b2.trace_add("write", lambda *args: self.var_b2.set(self.var_b2.get().upper()))
        self.var_b3.trace_add("write", lambda *args: self.var_b3.set(self.var_b3.get().upper()))
        ttk.Entry(b_frame, textvariable=self.var_b1, width=15).pack(side=tk.LEFT, padx=2)
        ttk.Entry(b_frame, textvariable=self.var_b2, width=15).pack(side=tk.LEFT, padx=2)
        ttk.Entry(b_frame, textvariable=self.var_b3, width=15).pack(side=tk.LEFT, padx=2)
        
        ttk.Label(lf2, text="Quantity ", width=32).grid(row=1, column=0, sticky="w", pady=2)
        q_frame = tk.Frame(lf2, bg=SURFACE_COLOR)
        q_frame.grid(row=1, column=1, sticky="w", pady=2, padx=5)
        self.var_qty = tk.StringVar()
        qty_entry = ttk.Entry(q_frame, textvariable=self.var_qty, width=10, font=HMI_FONT_M)
        qty_entry.pack(side=tk.LEFT, padx=2)
        ttk.Label(q_frame, text="pcs").pack(side=tk.LEFT, padx=5)
        
        # Section 3: Operation Details
        _, lf3 = self.create_card(self.form_frame, "Operation Details")
        
        ttk.Label(lf3, text="Work in Sub-Process by Shift ", width=32).grid(row=0, column=0, sticky="w", pady=2)
        self.cb_shift_sp = ttk.Combobox(lf3, values=["A", "B", "C"], state="readonly", width=5)
        self.cb_shift_sp.grid(row=0, column=1, sticky="w", pady=2, padx=5)
        self.cb_shift_sp.set("")
        self.cb_shift_sp.bind("<<ComboboxSelected>>", lambda e: [self.update_sub_batch_preview(), self.update_stats()])
        
        ttk.Label(lf3, text="Op ID ").grid(row=0, column=2, sticky="w", pady=2, padx=10)
        self.var_op_id = tk.StringVar()
        vcmd_op = (self.register(lambda P: len(P) <= 7 and all(c.isalnum() or c == '-' for c in P)), '%P')
        ttk.Entry(lf3, textvariable=self.var_op_id, width=15, validate="key", validatecommand=vcmd_op).grid(row=0, column=3, sticky="w", pady=2)
        
        ttk.Label(lf3, text="Station ").grid(row=0, column=4, sticky="w", pady=2, padx=10)
        self.cb_station = ttk.Combobox(lf3, values=["S06", "S07", "S10", "S11"], state="readonly", width=5)
        self.cb_station.grid(row=0, column=5, sticky="w", pady=2)
        self.cb_station.bind("<<ComboboxSelected>>", lambda e: self.update_sub_batch_preview())
        
        # Section 4: Date & Time
        _, lf4 = self.create_card(self.form_frame, "Date & Time")
        
        def create_dt_picker(parent, label_text):
            frame = tk.Frame(parent, bg=SURFACE_COLOR)
            ttk.Label(frame, text=label_text, width=32).pack(side=tk.LEFT, padx=5)
            de = DateEntry(frame, width=12, background=ACCENT_COLOR, foreground='white', borderwidth=2)
            de.pack(side=tk.LEFT, padx=2)
            h_spin = tk.Spinbox(frame, from_=0, to=23, wrap=True, width=3, format="%02.0f", bg=SURFACE_COLOR, fg=TEXT_COLOR)
            h_spin.pack(side=tk.LEFT, padx=2)
            ttk.Label(frame, text=":").pack(side=tk.LEFT)
            m_spin = tk.Spinbox(frame, from_=0, to=59, wrap=True, width=3, format="%02.0f", bg=SURFACE_COLOR, fg=TEXT_COLOR)
            m_spin.pack(side=tk.LEFT, padx=2)
            
            def set_now():
                now = datetime.datetime.now()
                de.set_date(now.date())
                h_spin.delete(0, "end")
                h_spin.insert(0, f"{now.hour:02d}")
                m_spin.delete(0, "end")
                m_spin.insert(0, f"{now.minute:02d}")
                
            live_var = tk.BooleanVar(value=True)
            cb_live = ttk.Checkbutton(frame, text="Live", variable=live_var)
            cb_live.pack(side=tk.LEFT, padx=5)
            
            def auto_update():
                if live_var.get():
                    set_now()
                parent.after(1000, auto_update)
                
            def stop_live(*args): live_var.set(False)
            de.bind("<<DateEntrySelected>>", stop_live)
            h_spin.bind("<Button-1>", stop_live)
            h_spin.bind("<Key>", stop_live)
            m_spin.bind("<Button-1>", stop_live)
            m_spin.bind("<Key>", stop_live)
            
            auto_update()
            return frame, de, h_spin, m_spin

        self.f_dt_sp, self.de_sp, self.h_sp, self.m_sp = create_dt_picker(lf4, "Sub-Process Work Date/Time")
        self.f_dt_sp.grid(row=0, column=0, sticky="w", pady=5)
        
        # Section 5: Additional Info
        _, lf5 = self.create_card(self.form_frame, "Additional Info")
        
        ttk.Label(lf5, text="Remarks", width=32).grid(row=0, column=0, sticky="nw", pady=2)
        self.txt_remarks = tk.Text(lf5, height=3, width=50, bg=SURFACE_COLOR, fg=TEXT_COLOR)
        self.txt_remarks.grid(row=0, column=1, sticky="w", pady=2, padx=5)
        
    def build_tab5(self):
        # Consume to Line UI
        main_frame = tk.Frame(self.tab5, bg=BG_COLOR)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        tk.Label(main_frame, text="Consume Box to Production Line", bg=BG_COLOR, fg=WARNING_COLOR, font=HMI_FONT_L).pack(pady=(0, 20))
        
        _, card = self.create_card(main_frame, "Scan Box Label")
        
        ttk.Label(card, text="Scan Sub-Batch ID (SB_ID) ", font=HMI_FONT_M).pack(pady=10)
        self.var_consume_scan = tk.StringVar()
        entry_scan = ttk.Entry(card, textvariable=self.var_consume_scan, width=40, font=HMI_FONT_L)
        entry_scan.pack(pady=10)
        entry_scan.bind("<Return>", self.on_consume_scan)
        
        shift_frame = tk.Frame(card, bg=SURFACE_COLOR)
        shift_frame.pack(pady=5)
        
        self.var_manual_shift = tk.BooleanVar(value=False)
        def toggle_manual():
            if self.var_manual_shift.get():
                self.cb_consume_shift.config(state="readonly")
            else:
                self.cb_consume_shift.config(state="disabled")
                self.cb_consume_shift.set(getattr(self, 'app_user_shift', ''))
                
        cb_manual = ttk.Checkbutton(shift_frame, text="Manual Shift", variable=self.var_manual_shift, command=toggle_manual)
        cb_manual.pack(side=tk.LEFT, padx=5)
        
        self.cb_consume_shift = ttk.Combobox(shift_frame, values=["A", "B", "C"], state="disabled", width=5, font=HMI_FONT_M)
        self.cb_consume_shift.pack(side=tk.LEFT, padx=5)
        
        self.lbl_consume_status = tk.Label(card, text="", bg=SURFACE_COLOR, font=HMI_FONT_M)
        self.lbl_consume_status.pack(pady=20)

    def on_consume_scan(self, event=None):
        scan_data = self.var_consume_scan.get().strip()
        if hasattr(self, 'var_manual_shift') and self.var_manual_shift.get():
            shift = self.cb_consume_shift.get()
        else:
            shift = getattr(self, 'app_user_shift', '')
            
        if not scan_data: return
        if not shift:
            self.lbl_consume_status.config(text="Please select shift first.", fg=ERROR_COLOR)
            return
            
        sb_id = scan_data
        if scan_data.startswith("{") and scan_data.endswith("}"):
            try:
                import json
                data = json.loads(scan_data)
                sb_id = data.get("sub_batch_id", "")
                if not sb_id:
                    dt_sp = data.get("sub_process_datetime", "")
                    station = data.get("station", "")
                    shift_sp = data.get("sub_process_shift", "")
                    if dt_sp and station and shift_sp:
                        sb_id = f"SB{dt_sp.replace('-', '').replace(':', '').replace(' ', '')}{station}{shift_sp}"
            except Exception:
                pass
                
        if not sb_id:
            self.lbl_consume_status.config(text="Invalid barcode format.", fg=ERROR_COLOR)
            return

        dt_line = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        
        conn = get_db_connection()
        c = conn.cursor()
        
        # Check Quality Quarantine Status
        try:
            c.execute("SELECT is_quarantined FROM quality_defects WHERE sub_batch_id=?", (sb_id,))
            q_row = c.fetchone()
            if q_row and q_row[0] == 1:
                self.lbl_consume_status.config(text="")
                messagebox.showerror("Quarantined Box", f" BLOCKED!\n\nBox {sb_id} is quarantined by Quality Department and CANNOT be consumed to the line.", parent=self)
                conn.close()
                self.var_consume_scan.set("")
                return
        except sqlite3.OperationalError:
            pass

        c.execute("SELECT id, status, dt_sp, pn_sf FROM records WHERE sub_batch_id = ?", (sb_id,))
        row = c.fetchone()
        
        if not row:
            self.lbl_consume_status.config(text="Error: Sub-Batch ID not found in database!", fg=ERROR_COLOR)
            conn.close()
            return
            
        if row[1] == 'Consumed':
            self.lbl_consume_status.config(text=f"Warning: Box {sb_id} is already Consumed!", fg=WARNING_COLOR)
            conn.close()
            self.var_consume_scan.set("")
            return
            
        dt_sp = row[2]
        pn_sf = row[3]
        
        # FIFO Check
        c.execute("SELECT COUNT(*) FROM records WHERE pn_sf=? AND status='In Rack' AND dt_sp < ?", (pn_sf, dt_sp))
        older_count = c.fetchone()[0]
        
        if older_count > 0:
            msg = f"FIFO WARNING: There are {older_count} OLDER box(es) of {pn_sf} in the rack.\n\nAre you sure you want to consume this newer box?"
            if not messagebox.askyesno("FIFO Alert", msg, parent=self):
                conn.close()
                self.var_consume_scan.set("")
                self.lbl_consume_status.config(text="Consume cancelled due to FIFO violation.", fg=WARNING_COLOR)
                return
                
        c.execute("UPDATE records SET status = 'Consumed', dt_line = ?, shift_line = ? WHERE sub_batch_id = ?", (dt_line, shift, sb_id))
        conn.commit()
        conn.close()
        
        self.update_excel_record(sb_id, dt_line, shift)
        
        self.lbl_consume_status.config(text=f"Success! {sb_id} consumed at {dt_line}", fg=SUCCESS_COLOR)
        self.var_consume_scan.set("")
        self.refresh_records_treeview()
        self.after(5000, lambda: self.lbl_consume_status.config(text=""))

    def update_excel_record(self, sb_id, dt_line, shift_line):
        if not os.path.exists(EXCEL_FILE): return
        try:
            from openpyxl import load_workbook
            wb = load_workbook(EXCEL_FILE)
            if "Sub-process fill by TL" not in wb.sheetnames: return
            ws = wb["Sub-process fill by TL"]
            
            for row in range(2, ws.max_row + 1):
                cell_sb_id = ws.cell(row=row, column=1).value
                if cell_sb_id == sb_id:
                    ws.cell(row=row, column=14).value = dt_line
                    ws.cell(row=row, column=15).value = shift_line
            wb.save(EXCEL_FILE)
        except Exception as e:
            print("Error updating excel:", e)
            
    def build_tab4(self):
        # Container Split
        main_frame = tk.Frame(self.tab4, bg=BG_COLOR)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        tk.Label(main_frame, text="Search & Reprint Label", bg=BG_COLOR, fg=ACCENT_COLOR, font=HMI_FONT_L).pack(pady=(0, 20))
        
        _, search_card = self.create_card(main_frame, "Search Record")
        ttk.Label(search_card, text="Scan QR Code or Type SB_ID ", font=HMI_FONT_M).pack(pady=10)
        self.var_reprint_scan = tk.StringVar()
        entry_scan = ttk.Entry(search_card, textvariable=self.var_reprint_scan, width=40, font=HMI_FONT_L)
        entry_scan.pack(pady=10)
        entry_scan.bind("<Return>", self.on_reprint_scan)
        
        self.lbl_reprint_status = tk.Label(search_card, text="", bg=SURFACE_COLOR, font=HMI_FONT_M)
        self.lbl_reprint_status.pack(pady=10)
        
        _, details_card = self.create_card(main_frame, "Record Details")
        self.txt_reprint_details = tk.Text(details_card, height=12, width=80, bg=BG_COLOR, fg=TEXT_COLOR, font=HMI_FONT_M, state="disabled")
        self.txt_reprint_details.pack(pady=10, padx=10)
        
        self.btn_do_reprint = ttk.Button(details_card, text="Print Duplicate Label", style="Warning.TButton", command=self.do_reprint_action)
        self.btn_do_reprint.pack(pady=10)
        self.btn_do_reprint.config(state="disabled")

    def on_reprint_scan(self, event=None):
        raw_input = self.var_reprint_scan.get().strip()
        if not raw_input: return
        
        sb_id = raw_input
        # If it looks like a JSON QR Code, extract the SB_ID
        if raw_input.startswith('{') and raw_input.endswith('}'):
            try:
                import json
                data = json.loads(raw_input)
                sb_id = data.get("sub_batch_id", "")
            except json.JSONDecodeError:
                pass
                
        if not sb_id:
            self.lbl_reprint_status.config(text="Error: Could not determine Sub-Batch ID!", fg=ERROR_COLOR)
            self.txt_reprint_details.config(state="normal")
            self.txt_reprint_details.delete("1.0", tk.END)
            self.txt_reprint_details.config(state="disabled")
            self.btn_do_reprint.config(state="disabled")
            self.var_reprint_scan.set("")
            return
            
        self.var_reprint_scan.set(sb_id)
        
        conn = get_db_connection()
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute("SELECT * FROM records WHERE sub_batch_id = ?", (sb_id,))
        row = c.fetchone()
        conn.close()
        
        if not row:
            self.lbl_reprint_status.config(text="Error: Record not found in Database!", fg=ERROR_COLOR)
            self.txt_reprint_details.config(state="normal")
            self.txt_reprint_details.delete("1.0", tk.END)
            self.txt_reprint_details.config(state="disabled")
            self.btn_do_reprint.config(state="disabled")
            return
            
        self.lbl_reprint_status.config(text="Record authenticated successfully.", fg=SUCCESS_COLOR)
        
        details = (
            f"Sub-Batch ID: {row['sub_batch_id']}\n"
            f"SF PN: {row['pn_sf']} ({row['part_sf']})\n"
            f"Quantity: {row['quantity']}\n"
            f"Status: {row['status']}\n\n"
            f"--- RM Components ---\n"
            f"RM 1: {row['rm1_pn']}  |  Batch: {row['batch1']}\n"
        )
        if row['rm2_pn']: details += f"RM 2: {row['rm2_pn']}  |  Batch: {row['batch2']}\n"
        if row['rm3_pn']: details += f"RM 3: {row['rm3_pn']}  |  Batch: {row['batch3']}\n"
        if row['rm4_pn']: details += f"RM 4: {row['rm4_pn']}\n"
        details += (
            f"\n--- Process Info ---\n"
            f"Operator ID: {row['op_id']}\n"
            f"Station: {row['station']}\n"
            f"Sub-Process Time: {row['dt_sp']} (Shift: {row['shift_sp']})\n"
        )
        if row['dt_line']:
            details += f"Prod Line Time: {row['dt_line']} (Shift: {row['shift_line']})\n"
        if row['remarks']:
            details += f"Remarks: {row['remarks']}\n"
        
        self.txt_reprint_details.config(state="normal")
        self.txt_reprint_details.delete("1.0", tk.END)
        self.txt_reprint_details.insert(tk.END, details)
        self.txt_reprint_details.config(state="disabled")
        
        self.btn_do_reprint.config(state="normal")
        self.current_reprint_sb_id = sb_id
        
    def do_reprint_action(self):
        if hasattr(self, 'current_reprint_sb_id'):
            self.do_print(self.current_reprint_sb_id)
            self.lbl_reprint_status.config(text="Label sent to printer.", fg=SUCCESS_COLOR)
            self.after(3000, lambda: getattr(self, 'lbl_reprint_status', tk.Label()).config(text=""))
            
    def on_rm_scanned_t1(self, event=None):
        scanned_rm = self.var_scan_rm_t1.get().strip()
        if hasattr(self, 'lbl_scan_rm_status'):
            self.lbl_scan_rm_status.config(text="")
        if not scanned_rm: return
        
        matched_sf_pn = None
        for sf_pn, (sf_name, rm_list) in SF_DATA.items():
            if rm_list and rm_list[0][0] == scanned_rm:
                matched_sf_pn = sf_pn
                break
                
        if matched_sf_pn:
            self.cb_sf_pn.config(state="readonly")
            self.cb_sf_pn.set(matched_sf_pn)
            if hasattr(self, 'lbl_scan_rm_status'):
                self.lbl_scan_rm_status.config(text="Match Found!", fg=SUCCESS_COLOR)
                self.after(3000, lambda: getattr(self, 'lbl_scan_rm_status', tk.Label()).config(text=""))
            self.on_sf_selected(None)
            self.cb_sf_pn.config(state="disabled")
        else:
            if hasattr(self, 'lbl_scan_rm_status'):
                self.lbl_scan_rm_status.config(text="Mismatch! Invalid RM Barcode.", fg=ERROR_COLOR)
                self.after(4000, lambda: getattr(self, 'lbl_scan_rm_status', tk.Label()).config(text=""))
            self.var_scan_rm_t1.set("")
        


    def build_tab2(self):
        filter_frame = tk.Frame(self.tab2, bg=SURFACE_COLOR)
        filter_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(filter_frame, text="Search:").pack(side=tk.LEFT, padx=2)
        self.var_rec_search = tk.StringVar()
        ttk.Entry(filter_frame, textvariable=self.var_rec_search, width=20).pack(side=tk.LEFT, padx=2)
        
        ttk.Label(filter_frame, text="Shift:").pack(side=tk.LEFT, padx=(10,2))
        self.cb_rec_shift = ttk.Combobox(filter_frame, values=["All", "A", "B", "C"], state="readonly", width=5)
        self.cb_rec_shift.set("All")
        self.cb_rec_shift.pack(side=tk.LEFT, padx=2)
        
        ttk.Label(filter_frame, text="Station:").pack(side=tk.LEFT, padx=(10,2))
        self.cb_rec_station = ttk.Combobox(filter_frame, values=["All", "S06", "S07", "S10", "S11"], state="readonly", width=5)
        self.cb_rec_station.set("All")
        self.cb_rec_station.pack(side=tk.LEFT, padx=2)
        
        ttk.Button(filter_frame, text="Search", style="Primary.TButton", command=self.refresh_records_treeview).pack(side=tk.LEFT, padx=10)
        ttk.Button(filter_frame, text="All", style="Secondary.TButton", command=self.reset_records_filter).pack(side=tk.LEFT, padx=2)
        ttk.Button(filter_frame, text="Print Selected", style="Warning.TButton", command=self.print_selected_record).pack(side=tk.RIGHT, padx=5)
        
        self.lbl_rec_count = tk.Label(filter_frame, text="0 records", bg=SURFACE_COLOR, fg=TEXT_COLOR)
        self.lbl_rec_count.pack(side=tk.RIGHT, padx=10)
        
        tree_frame = tk.Frame(self.tab2, bg=BG_COLOR)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        cols = ("SB_ID", "SF_PN", "SF_Name", "Qty", "Shift", "Station", "Op_ID", "DateTime", "Status", "Reprints")
        self.tree_records = ttk.Treeview(tree_frame, columns=cols, show="headings")
        for c in cols:
            self.tree_records.heading(c, text=c, command=lambda _c=c: self.sort_treeview(self.tree_records, _c, False))
            self.tree_records.column(c, width=90)
            
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree_records.xview)
        self.tree_records.configure(xscrollcommand=hsb.set)
        
        self.tree_records.grid(row=0, column=0, sticky="nsew")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        
        self.tree_records.tag_configure("pending", background=WARNING_COLOR, foreground="black")
        self.tree_records.tag_configure("even", background="#141A20")
        self.tree_records.tag_configure("odd", background="#1B232C")
        self.tree_records.tag_configure("reprint_highlight", background="#B45309", foreground="white") # Orange background for reprinted labels
        self.tree_records.bind("<Double-1>", lambda e: self.print_selected_record())
        
        self.rec_context_menu = tk.Menu(self.tree_records, tearoff=0, bg=SURFACE_COLOR, fg=TEXT_COLOR, activebackground=ACCENT_COLOR, activeforeground=TEXT_COLOR)
        self.rec_context_menu.add_command(label="View Operator Scorecard", command=self.open_operator_scorecard)

        def on_rec_right_click(event):
            user_role = getattr(self, 'app_user_role', 'Operator')
            if user_role not in ["Supervisor", "Manager", "Admin"]:
                return
            item = self.tree_records.identify_row(event.y)
            if item:
                self.tree_records.selection_set(item)
                self.rec_context_menu.tk_popup(event.x_root, event.y_root)

        self.tree_records.bind("<Button-3>", on_rec_right_click)
        self.page_frame = tk.Frame(self.tab2, bg=BG_COLOR)
        self.page_frame.pack(fill=tk.X, pady=5)

    def get_dashboard_data(self):
        conn = get_db_connection()
        c = conn.cursor()
        
        now = datetime.datetime.now()
        today_date = now.date()
        if now.hour < 6: today_date -= datetime.timedelta(days=1)
        
        start_today_str = f"{today_date.strftime('%Y-%m-%d')} 06:00:00"
        
        start_week_date = today_date - datetime.timedelta(days=today_date.weekday())
        start_week_str = f"{start_week_date.strftime('%Y-%m-%d')} 06:00:00"
        
        start_month_date = today_date.replace(day=1)
        start_month_str = f"{start_month_date.strftime('%Y-%m-%d')} 06:00:00"
        
        def get_stats(start_dt):
            c.execute("SELECT shift_sp, SUM(quantity) FROM records WHERE dt_sp >= ? GROUP BY shift_sp", (start_dt,))
            rows = c.fetchall()
            stats = {'A': 0, 'B': 0, 'C': 0}
            for r in rows:
                if r[0] in stats:
                    stats[r[0]] = r[1] if r[1] else 0
            total = sum(stats.values())
            return total, stats['A'], stats['B'], stats['C']
            
        today_stats = get_stats(start_today_str)
        week_stats = get_stats(start_week_str)
        month_stats = get_stats(start_month_str)
        
        conn.close()
        return today_stats, week_stats, month_stats


    def update_clock(self):
        try:
            now = datetime.datetime.now()
            self.lbl_clock.config(text=now.strftime("%H:%M:%S"))
            self.lbl_date.config(text=now.strftime("%Y-%m-%d"))
            self._check_shift_end_reports(now)
        except Exception:
            pass
        self.after(1000, self.update_clock)

    def _check_shift_end_reports(self, now):
        try:
            wd = now.weekday() # 0=Mon, 4=Fri, 5=Sat, 6=Sun
            h = now.hour
            m = now.minute
            
            trigger = False
            shift_label = ""
            start_dt = ""
            end_dt = now.strftime("%Y-%m-%d %H:%M:00")
            
            if wd in [0, 1, 2, 3, 5]:
                if h == 14 and m == 0:
                    trigger = True; shift_label = "Morning Shift"
                    start_dt = now.strftime("%Y-%m-%d 06:00:00")
                elif h == 22 and m == 0:
                    trigger = True; shift_label = "Evening Shift"
                    start_dt = now.strftime("%Y-%m-%d 14:00:00")
                elif h == 6 and m == 0:
                    trigger = True; shift_label = "Night Shift"
                    if wd == 0:
                        start_dt = (now - datetime.timedelta(hours=8)).strftime("%Y-%m-%d 22:00:00")
                    elif wd == 5:
                        start_dt = now.strftime("%Y-%m-%d 00:00:00") # Sat 06:00 (Friday night ends)
                    else:
                        start_dt = (now - datetime.timedelta(hours=8)).strftime("%Y-%m-%d 22:00:00")
                        
            elif wd == 4: # Friday
                if h == 6 and m == 0:
                    trigger = True; shift_label = "Night Shift"
                    start_dt = now.strftime("%Y-%m-%d 00:00:00")
                elif h == 12 and m == 0:
                    trigger = True; shift_label = "Morning Shift"
                    start_dt = now.strftime("%Y-%m-%d 06:00:00")
                elif h == 18 and m == 30:
                    trigger = True; shift_label = "Evening Shift"
                    start_dt = now.strftime("%Y-%m-%d 12:00:00")
                    
            if trigger:
                if hasattr(self, 'app_user_id') and self.app_user_id:
                    print(f"Shift ended! Forcing logout for {self.app_user_id}...")
                    self.perform_logout(force=True)

                key = f"{now.strftime('%Y-%m-%d')}_{h:02d}{m:02d}"
                
                last_key_file = os.path.join(DATA_DIR, "last_report.json")
                persisted_key = ""
                try:
                    if os.path.exists(last_key_file):
                        with open(last_key_file, "r") as f:
                            persisted_key = json.load(f).get("last_pdf_report_key", "")
                except Exception:
                    pass
                
                if persisted_key != key and getattr(self, 'last_pdf_report_key', '') != key:
                    self.last_pdf_report_key = key
                    try:
                        with open(last_key_file, "w") as f:
                            json.dump({"last_pdf_report_key": key}, f)
                    except Exception:
                        pass
                        
                    import threading
                    try:
                        import report_generator
                        threading.Thread(target=report_generator.generate_shift_pdf_report, args=(start_dt, end_dt, shift_label), daemon=True).start()
                    except Exception as e:
                        print(f"Failed to trigger PDF report generator: {e}")
        except Exception as e:
            print(f"Report Trigger Error: {e}")

    def update_stats(self):
        PROD_RATES = {
            "MOCK-A01-10001-X-SUB": 240
        }
        DEFAULT_RATE = 120  # default hourly rate if PN is not in PROD_RATES

        now = datetime.datetime.now()
        # Production day starts at 06:00 and ends at 05:59 the next day
        if now.hour < 6:
            start_dt = (now - datetime.timedelta(days=1)).strftime("%Y-%m-%d 06:00:00")
            end_dt = now.strftime("%Y-%m-%d 05:59:59")
        else:
            start_dt = now.strftime("%Y-%m-%d 06:00:00")
            end_dt = (now + datetime.timedelta(days=1)).strftime("%Y-%m-%d 05:59:59")
            
        shift_val = getattr(self, 'app_user_shift', "")
        if not shift_val:
            shift_val = "-"
            
        current_pn = ""
        if hasattr(self, 'cb_sf_pn'):
            current_pn = self.cb_sf_pn.get()
        
        conn = get_db_connection()
        c = conn.cursor()
        
        # Shift & PN stats
        if current_pn:
            c.execute("SELECT COUNT(*), SUM(quantity) FROM records WHERE dt_sp >= ? AND dt_sp <= ? AND shift_sp = ? AND pn_sf = ?", (start_dt, end_dt, shift_val, current_pn))
        else:
            c.execute("SELECT COUNT(*), SUM(quantity) FROM records WHERE dt_sp >= ? AND dt_sp <= ? AND shift_sp = ?", (start_dt, end_dt, shift_val))
            
        res_shift = c.fetchone()
        shift_count = res_shift[0] if res_shift[0] else 0
        shift_qty = res_shift[1] if res_shift[1] else 0
        
        # Today stats (Overall)
        c.execute("SELECT COUNT(*) FROM records WHERE dt_sp >= ? AND dt_sp <= ?", (start_dt, end_dt))
        today_count = c.fetchone()[0]
        
        # Get all PNs produced in this shift
        c.execute("SELECT pn_sf, SUM(quantity) FROM records WHERE dt_sp >= ? AND dt_sp <= ? AND shift_sp = ? GROUP BY pn_sf", (start_dt, end_dt, shift_val))
        shift_pns_data = {row[0]: (row[1] or 0) for row in c.fetchall() if row[0]}
        
        conn.close()
        
        # Calculate Work Mins
        wd = now.weekday()
        if wd == 6:  # Sunday
            work_mins = 0
        elif wd == 4: # Friday
            work_mins = 290
        else:
            work_mins = 440
            
        # Update Shift Target Tracker on Tab 1
        if hasattr(self, 'tracker_scroll_frame'):
            for widget in self.tracker_scroll_frame.winfo_children():
                widget.destroy()
                
            conn = get_db_connection()
            c = conn.cursor()
            c.execute("SELECT product_pn, target_qty FROM shift_targets ORDER BY id ASC")
            db_targets = {}
            for row in c.fetchall():
                db_targets[row[0]] = row[1]
            conn.close()
            all_pns_to_track = set([p for p, q in shift_pns_data.items() if q > 0])
            if current_pn:
                all_pns_to_track.add(current_pn)
                
            sorted_pns = sorted(list(all_pns_to_track), key=lambda x: (x not in db_targets, x))
            
            for pn in sorted_pns:
                qty = shift_pns_data.get(pn, 0)
                tgt = db_targets.get(pn, 0)
                pct = min(100, (qty / tgt) * 100) if tgt > 0 else 0
                
                f = tk.Frame(self.tracker_scroll_frame, bg=SURFACE_COLOR)
                f.pack(fill=tk.X, pady=(0, 10))
                
                part_name = SF_DATA.get(pn, ("", []))[0]
                pn_display = f"{pn} ({part_name})" if part_name else pn
                lbl = tk.Label(f, text=f"{pn_display}\nProduced: {qty:,} / Target: {tgt:,}", bg=SURFACE_COLOR, fg=TEXT_COLOR, font=HMI_FONT_S, justify=tk.LEFT)
                lbl.pack(anchor="w")
                
                if pct >= 80:
                    pb_style = "Safe.Horizontal.TProgressbar"
                elif pct >= 40:
                    pb_style = "Warn.Horizontal.TProgressbar"
                else:
                    pb_style = "Danger.Horizontal.TProgressbar"
                
                pb = ttk.Progressbar(f, style=pb_style, orient="horizontal", mode="determinate")
                pb.pack(fill=tk.X, pady=(2, 0))
                pb['value'] = pct
                
                # Bind mousewheel to children too
                def _on_mw(e, cvs=self.tracker_canvas): 
                    cvs.yview_scroll(int(-1*(e.delta/120)), "units")
                    return "break"
                    
                lbl.bind("<MouseWheel>", _on_mw)
                pb.bind("<MouseWheel>", _on_mw)
                f.bind("<MouseWheel>", _on_mw)

            def _on_canvas_mw(e):
                self.tracker_canvas.yview_scroll(int(-1*(e.delta/120)), "units")
                return "break"
                
            self.tracker_canvas.bind("<MouseWheel>", _on_canvas_mw)
            self.tracker_scroll_frame.bind("<MouseWheel>", _on_canvas_mw)
            
            self.tracker_scroll_frame.update_idletasks()
            self.tracker_canvas.config(scrollregion=self.tracker_canvas.bbox("all"))
            
            if len(sorted_pns) > 4:
                self.lbl_scroll_indicator.pack(pady=(0, 5), before=self.t1_left.winfo_children()[2])
            else:
                self.lbl_scroll_indicator.pack_forget()
        
        self.lbl_title_recs.config(text=f"Records (Shift {shift_val})")
        self.side_stat_recs.config(text=str(shift_count))
        self.lbl_title_qty.config(text=f"Qty (Shift {shift_val})")
        self.side_stat_qty.config(text=str(shift_qty))
        self.side_stat_today.config(text=str(today_count))

    def populate_sf_combobox(self):
        sf_list = list(SF_DATA.keys())
        self.cb_sf_pn['values'] = sf_list
        self.lbl_search_count.config(text=f"{len(sf_list)} matches", fg=SUCCESS_COLOR)

    def filter_sf_combobox(self, *args):
        search_term = self.sv_search.get().lower()
        filtered = [pn for pn in SF_DATA.keys() if search_term in pn.lower()]
        self.cb_sf_pn['values'] = filtered
        if filtered:
            self.lbl_search_count.config(text=f"{len(filtered)} matches", fg=SUCCESS_COLOR)
            self.cb_sf_pn.set(filtered[0])
            self.on_sf_selected(None)
        else:
            self.lbl_search_count.config(text="0 matches", fg="red")
            self.cb_sf_pn.set("")

    def on_sf_selected(self, event):
        sf_pn = self.cb_sf_pn.get()
        for widget in getattr(self, 'rm_container_t1', tk.Frame()).winfo_children():
            widget.destroy()
        self.rm_vars_t1 = []

        if sf_pn in SF_DATA:
            sf_name, rm_list = SF_DATA[sf_pn]
            self.var_part_sf.set(sf_name)
            
            for idx, (rm_id, rm_name) in enumerate(rm_list):
                ttk.Label(self.rm_container_t1, text=f"RM {idx+1} Ref ", font=HMI_FONT_S).grid(row=idx, column=0, sticky="w", padx=5, pady=2)
                cb_var = tk.StringVar(value=rm_id)
                cb = ttk.Entry(self.rm_container_t1, textvariable=cb_var, state="readonly", width=25)
                cb.grid(row=idx, column=1, sticky="w", padx=5, pady=2)
                
                ttk.Label(self.rm_container_t1, text=f"RM {idx+1} Name", font=HMI_FONT_S).grid(row=idx, column=2, sticky="w", padx=15, pady=2)
                name_var = tk.StringVar(value=rm_name)
                ttk.Entry(self.rm_container_t1, textvariable=name_var, state="readonly", width=25).grid(row=idx, column=3, sticky="w", padx=5, pady=2)
                
                self.rm_vars_t1.append((cb_var, name_var))

    def update_sub_batch_preview(self):
        self.update_stats()
        pass

    def get_dt_string(self, de, h, m):
        d = de.get_date()
        return f"{d.strftime('%Y-%m-%d')} {h.get()}:{m.get()}"

    def clear_form(self):
        if hasattr(self, 'var_scan_rm_t1'):
            self.var_scan_rm_t1.set("")
        self.cb_sf_pn.config(state="readonly")
        self.cb_sf_pn.set("")
        self.var_part_sf.set("")
        for widget in getattr(self, "rm_container_t1", tk.Frame()).winfo_children():
            widget.destroy()
        self.rm_vars_t1 = []
        self.var_b1.set("")
        self.var_b2.set("")
        self.var_b3.set("")
        self.var_qty.set("")
        self.cb_shift_sp.set("")
        self.var_op_id.set("")
        self.cb_station.set("")
        
        now = datetime.datetime.now()
        self.de_sp.set_date(now.date())
        self.h_sp.delete(0, "end"); self.h_sp.insert(0, f"{now.hour:02d}")
        self.m_sp.delete(0, "end"); self.m_sp.insert(0, f"{now.minute:02d}")
        
        self.txt_remarks.delete("1.0", tk.END)
        if hasattr(self, 'cb_consume_shift') and hasattr(self, 'var_manual_shift'):
            if not self.var_manual_shift.get():
                self.cb_consume_shift.set(getattr(self, 'app_user_shift', ''))
        self.update_sub_batch_preview()

    def same_as_last(self):
        if hasattr(self, 'var_scan_rm_t1'):
            self.var_scan_rm_t1.set("")
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("SELECT * FROM records ORDER BY id DESC LIMIT 1")
        rec = c.fetchone()
        conn.close()
        if rec:
            self.cb_sf_pn.set(rec[2])
            self.on_sf_selected(None)
            self.var_qty.set(rec[15])
            self.cb_shift_sp.set(rec[16])
            if hasattr(self, 'var_op_id'):
                self.var_op_id.set(rec[17])
            self.cb_station.set(rec[18])
            self.update_sub_batch_preview()
            self.lbl_status.config(text="Pre-filled with last record data.", fg=SUCCESS_COLOR)
            self.after(3000, lambda: self.lbl_status.config(text=""))

    def background_excel_save(self, excel_data):
        try:
            save_to_excel(excel_data)
            self.export_kpis_to_excel()
        except Exception as e:
            print(f"Background Excel Save Error: {e}")

    def save_record(self):
        sf_pn = self.cb_sf_pn.get()
        qty_str = self.var_qty.get()
        op_id = self.var_op_id.get()
        station = self.cb_station.get()
        shift_sp = self.cb_shift_sp.get()
        
        if not all([sf_pn, qty_str, op_id, station, shift_sp]) or not self.rm_vars_t1:
            messagebox.showerror("Error", "Please fill all required fields ().")
            return
            
        if not any([self.var_b1.get().strip(), self.var_b2.get().strip(), self.var_b3.get().strip()]):
            messagebox.showerror("Error", "Please provide at least one Batch Number.")
            return
            
        rm_pns = [""] * 4
        rm_names = [""] * 4
        for idx, (cb_var, name_var) in enumerate(self.rm_vars_t1):
            if not cb_var.get():
                messagebox.showerror("Error", "Please fill all RM Reference fields.")
                return
            if idx < 4:
                rm_pns[idx] = cb_var.get()
                rm_names[idx] = name_var.get()

        try:
            qty = int(qty_str)
        except ValueError:
            messagebox.showerror("Error", "Quantity must be a valid integer.")
            return

        shift_sp = self.cb_shift_sp.get()
        dt_sp = self.get_dt_string(self.de_sp, self.h_sp, self.m_sp)
        dt_line = ""
        shift_line = ""
        
        conn = get_db_connection()
        c = conn.cursor()
        
        base_sb_id = f"SB{dt_sp.replace('-', '').replace(':', '').replace(' ', '')}{station}{shift_sp[0] if shift_sp else ''}"
        c.execute("SELECT COUNT(*) FROM records WHERE sub_batch_id LIKE ?", (base_sb_id + "%",))
        count = c.fetchone()[0]
        sb_id = f"{base_sb_id}{count+1:02d}"

        created_at = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        registered_by = f"{getattr(self, 'app_user_id', '')}"
        
        data = (
            sb_id, sf_pn, self.var_part_sf.get(), 
            rm_pns[0], rm_names[0], rm_pns[1], rm_names[1],
            rm_pns[2], rm_names[2], rm_pns[3], rm_names[3],
            self.var_b1.get(), self.var_b2.get(), self.var_b3.get(), qty,
            shift_sp, op_id, station, dt_sp, dt_line,
            shift_line, self.txt_remarks.get("1.0", tk.END).strip(),
            'In Rack', created_at, registered_by
        )

        try:
            c.execute('''
                INSERT INTO records (
                    sub_batch_id, pn_sf, part_sf, 
                    rm1_pn, rm1_name, rm2_pn, rm2_name,
                    rm3_pn, rm3_name, rm4_pn, rm4_name,
                    batch1, batch2, batch3, quantity,
                    shift_sp, op_id, station, dt_sp, dt_line,
                    shift_line, remarks, status, created_at, registered_by
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            ''', data)
            conn.commit()
        except sqlite3.IntegrityError:
            conn.rollback()
            conn.close()
            messagebox.showerror("Error", "Duplicate Sub-Batch ID detected due to concurrent scans. Please try saving again.")
            return
        except Exception as e:
            conn.rollback()
            conn.close()
            messagebox.showerror("Error", f"Failed to save record: {e}")
            return
            
        conn.close()
        
        excel_data = [
            sb_id, sf_pn, self.var_part_sf.get(), 
            rm_pns[0], rm_names[0], rm_pns[1], rm_names[1],
            rm_pns[2], rm_names[2], rm_pns[3], rm_names[3],
            self.var_b1.get(), self.var_b2.get(), self.var_b3.get(), qty,
            shift_sp, op_id, station, dt_sp, dt_line,
            shift_line, self.txt_remarks.get("1.0", tk.END).strip(),
            registered_by
        ]
        
        threading.Thread(target=self.background_excel_save, args=(excel_data,), daemon=True).start()

        self.lbl_status.config(text=f"Saved successfully: {sb_id}", fg=SUCCESS_COLOR)
        self.after(7000, lambda: self.lbl_status.config(text=""))
        
        self.update_stats()
        self.refresh_recent_treeview()
        self.refresh_records_treeview()
        self.update_sub_batch_preview()
        
        # Auto-print without preview
        try:
            conn_print = get_db_connection()
            conn_print.row_factory = sqlite3.Row
            c_print = conn_print.cursor()
            c_print.execute("SELECT * FROM records WHERE sub_batch_id=?", (sb_id,))
            row_print = c_print.fetchone()
            conn_print.close()
            if row_print:
                print_html_slip(dict(row_print), silent=True)
        except Exception as e:
            print("Auto-print failed:", e)
        
        # update recent PNs sidebar
        pns = self.recent_pns_listbox.get(0, tk.END)
        if sf_pn not in pns:
            self.recent_pns_listbox.insert(0, sf_pn)
            if self.recent_pns_listbox.size() > 8:
                self.recent_pns_listbox.delete(tk.END)
                
        self.clear_form()

    def refresh_recent_treeview(self):
        for item in self.tree_recent.get_children():
            self.tree_recent.delete(item)
            
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("SELECT sub_batch_id, pn_sf, quantity, shift_sp, station, dt_sp FROM records ORDER BY id DESC LIMIT 5")
        rows = c.fetchall()
        conn.close()
        
        for i, row in enumerate(rows):
            tag = "even" if i % 2 == 0 else "odd"
            self.tree_recent.insert("", "end", values=row, tags=(tag,))

    def refresh_records_treeview(self, keep_page=False):
        if not keep_page:
            self.current_page = 1

        for item in self.tree_records.get_children():
            self.tree_records.delete(item)
            
        query = "SELECT sub_batch_id, pn_sf, part_sf, quantity, shift_sp, station, op_id, dt_sp, status, reprint_count FROM records WHERE 1=1"
        params = []
        
        search = self.var_rec_search.get()
        if search:
            query += " AND (sub_batch_id LIKE ? OR pn_sf LIKE ? OR part_sf LIKE ? OR op_id LIKE ?)"
            params.extend([f"%{search}%", f"%{search}%", f"%{search}%", f"%{search}%"])
            
        shift = self.cb_rec_shift.get()
        if shift != "All":
            query += " AND shift_sp = ?"
            params.append(shift)
            
        station = self.cb_rec_station.get()
        if station != "All":
            query += " AND station = ?"
            params.append(station)
            
        query += " ORDER BY id DESC"
        
        conn = get_db_connection()
        c = conn.cursor()
        c.execute(query, params)
        rows = c.fetchall()
        conn.close()
        
        total_count = len(rows)
        self.records_per_page = 20
        self.total_pages = (total_count + self.records_per_page - 1) // self.records_per_page if total_count > 0 else 1
        
        self.current_page = getattr(self, 'current_page', 1)
        if self.current_page > self.total_pages:
            self.current_page = self.total_pages
            
        start_idx = (self.current_page - 1) * self.records_per_page
        end_idx = start_idx + self.records_per_page
        display_rows = rows[start_idx:end_idx]
        
        for i, row in enumerate(display_rows):
            status_val = row[8]
            reprint_count = row[9] if row[9] is not None else 0
            
            tags = ()
            if reprint_count >= 1:
                tags = ("reprint_highlight",)
            elif status_val == 'pending':
                tags = ("pending",)
            else:
                tags = ("even" if i % 2 == 0 else "odd",)
            self.tree_records.insert("", "end", values=row, tags=tags)
            
        start_display = start_idx + 1 if total_count > 0 else 0
        end_display = min(end_idx, total_count)
        self.lbl_rec_count.config(text=f"Showing {start_display}-{end_display} of {total_count} records")
        
        self.render_pagination()

    def render_pagination(self):
        for widget in getattr(self, 'page_frame', tk.Frame()).winfo_children():
            widget.destroy()
            
        if not hasattr(self, 'page_frame') or not hasattr(self, 'total_pages') or self.total_pages <= 1:
            return
            
        btn_frame = tk.Frame(self.page_frame, bg=BG_COLOR)
        btn_frame.pack(anchor="center")
        
        def go_page(p):
            self.current_page = p
            self.refresh_records_treeview(keep_page=True)
            
        if self.current_page > 1:
            ttk.Button(btn_frame, text="< Prev", command=lambda: go_page(self.current_page - 1), width=6).pack(side=tk.LEFT, padx=2)
            
        start_p = max(1, self.current_page - 3)
        end_p = min(self.total_pages, self.current_page + 3)
        
        if start_p > 1:
            ttk.Button(btn_frame, text="1", command=lambda: go_page(1), width=3).pack(side=tk.LEFT, padx=1)
            if start_p > 2:
                tk.Label(btn_frame, text="...", bg=BG_COLOR, fg=TEXT_COLOR).pack(side=tk.LEFT, padx=1)
                
        for p in range(start_p, end_p + 1):
            if p == self.current_page:
                lbl = tk.Label(btn_frame, text=str(p), bg=ACCENT_COLOR, fg="white", font=("Segoe UI", 10, "bold"), width=3)
                lbl.pack(side=tk.LEFT, padx=2)
            else:
                ttk.Button(btn_frame, text=str(p), command=lambda p=p: go_page(p), width=3).pack(side=tk.LEFT, padx=1)
                
        if end_p < self.total_pages:
            if end_p < self.total_pages - 1:
                tk.Label(btn_frame, text="...", bg=BG_COLOR, fg=TEXT_COLOR).pack(side=tk.LEFT, padx=1)
            ttk.Button(btn_frame, text=str(self.total_pages), command=lambda: go_page(self.total_pages), width=3).pack(side=tk.LEFT, padx=1)
            
        if self.current_page < self.total_pages:
            ttk.Button(btn_frame, text="Next >", command=lambda: go_page(self.current_page + 1), width=6).pack(side=tk.LEFT, padx=2)

    def reset_records_filter(self):
        self.var_rec_search.set("")
        self.cb_rec_shift.set("All")
        self.cb_rec_station.set("All")
        self.refresh_records_treeview()

    def sort_treeview(self, tree, col, reverse):
        l = [(tree.set(k, col), k) for k in tree.get_children('')]
        try:
            l.sort(key=lambda t: int(t[0]), reverse=reverse)
        except ValueError:
            l.sort(reverse=reverse)
            
        for index, (val, k) in enumerate(l):
            tree.move(k, '', index)
        tree.heading(col, command=lambda: self.sort_treeview(tree, col, not reverse))

    def on_recent_double_click(self, event):
        item = self.tree_recent.selection()
        if not item: return
        sb_id = self.tree_recent.item(item[0], "values")[0]
        self.do_print(sb_id)

    def open_operator_scorecard(self):
        selected = self.tree_records.selection()
        if not selected: return
        item = self.tree_records.item(selected[0])
        op_id = item['values'][6]
        
        top = tk.Toplevel(self)
        top.title(f"Operator Scorecard: {op_id}")
        center_window(top, 600, 400)
        top.configure(bg=BG_COLOR)
        top.transient(self)
        
        # Header
        header = tk.Frame(top, bg=SURFACE_COLOR)
        header.pack(fill=tk.X, padx=10, pady=10)
        
        tk.Label(header, text="Performance Scorecard", font=("Inter", 14, "bold"), bg=SURFACE_COLOR, fg=TEXT_COLOR).pack(side=tk.LEFT, padx=10, pady=10)
        tk.Label(header, text=f"Operator: {op_id}", font=("Inter", 12), bg=SURFACE_COLOR, fg=ACCENT_COLOR).pack(side=tk.RIGHT, padx=10, pady=10)
        
        # Filter
        filter_frame = tk.Frame(top, bg=BG_COLOR)
        filter_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(filter_frame, text="Time Window:", bg=BG_COLOR, fg=TEXT_COLOR).pack(side=tk.LEFT)
        var_window = tk.StringVar(value="Today")
        
        def refresh_scorecard(*args):
            window = var_window.get()
            now = datetime.datetime.now()
            today_date = now.date()
            if now.hour < 6: today_date -= datetime.timedelta(days=1)
            
            if window == "Today":
                start_dt = f"{today_date.strftime('%Y-%m-%d')} 06:00:00"
            elif window == "This Week":
                start_week = today_date - datetime.timedelta(days=today_date.weekday())
                start_dt = f"{start_week.strftime('%Y-%m-%d')} 06:00:00"
            else: # This Month
                start_month = today_date.replace(day=1)
                start_dt = f"{start_month.strftime('%Y-%m-%d')} 06:00:00"
                
            conn = get_db_connection()
            c = conn.cursor()
            
            # 1. Total Output & Shift Avg
            c.execute("SELECT op_id, SUM(quantity) FROM records WHERE dt_sp >= ? GROUP BY op_id", (start_dt,))
            op_data = c.fetchall()
            
            total_output = 0
            shift_total = 0
            op_count = len(op_data)
            for r in op_data:
                q = r[1] or 0
                shift_total += q
                if r[0] == str(op_id):
                    total_output = q
                    
            shift_avg = (shift_total / op_count) if op_count > 0 else 0
            
            # 2. Quality Rate & Quarantine Events
            c.execute("SELECT qty_defective, is_quarantined FROM quality_defects WHERE produced_by_op = ? AND reported_at >= ?", (str(op_id), start_dt))
            defects = c.fetchall()
            total_defects = sum(d[0] for d in defects if d[0])
            quarantine_events = sum(1 for d in defects if d[1] == 1)
            
            qual_rate = ((total_output - total_defects) / total_output * 100) if total_output > 0 else (100 if total_output > 0 else 0)
            
            conn.close()
            
            # Update UI
            lbl_out_val.config(text=f"{total_output}")
            lbl_avg_val.config(text=f"{shift_avg:.1f}")
            lbl_qual_val.config(text=f"{qual_rate:.1f}%" if total_output > 0 else "N/A")
            lbl_quar_val.config(text=f"{quarantine_events}")
            
        cb_window = ttk.Combobox(filter_frame, textvariable=var_window, values=["Today", "This Week", "This Month"], state="readonly", width=15)
        cb_window.pack(side=tk.LEFT, padx=10)
        cb_window.bind("<<ComboboxSelected>>", refresh_scorecard)
        
        # Metrics Grid
        metrics_frame = tk.Frame(top, bg=BG_COLOR)
        metrics_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        def create_metric_card(parent, title, row, col, fg=TEXT_COLOR):
            f = tk.Frame(parent, bg=SURFACE_COLOR, bd=1, relief="ridge")
            f.grid(row=row, column=col, sticky="nsew", padx=10, pady=10)
            tk.Label(f, text=title, font=("Inter", 10), bg=SURFACE_COLOR, fg=TEXT_MUTED).pack(pady=(15,5))
            lbl_val = tk.Label(f, text="0", font=("Inter", 24, "bold"), bg=SURFACE_COLOR, fg=fg)
            lbl_val.pack(pady=(0,15))
            return lbl_val
            
        metrics_frame.columnconfigure(0, weight=1)
        metrics_frame.columnconfigure(1, weight=1)
        metrics_frame.rowconfigure(0, weight=1)
        metrics_frame.rowconfigure(1, weight=1)
        
        lbl_out_val = create_metric_card(metrics_frame, "Total Output", 0, 0, SUCCESS_COLOR)
        lbl_avg_val = create_metric_card(metrics_frame, "Shift Average (All Ops)", 0, 1, ACCENT_COLOR)
        lbl_qual_val = create_metric_card(metrics_frame, "Quality Rate", 1, 0, SUCCESS_COLOR)
        lbl_quar_val = create_metric_card(metrics_frame, "Quarantine Events", 1, 1, WARNING_COLOR)
        
        refresh_scorecard()

    def print_selected_record(self):
        item = self.tree_records.selection()
        if not item: return
        sb_id = self.tree_records.item(item[0], "values")[0]
        self.do_print(sb_id)

    def print_last_slip(self):
        conn = get_db_connection()
        c = conn.cursor()
        c.execute("SELECT sub_batch_id FROM records ORDER BY id DESC LIMIT 1")
        row = c.fetchone()
        conn.close()
        if row:
            self.do_print(row[0])
        else:
            messagebox.showinfo("Info", "No records found.")

    def do_print(self, sb_id):
        conn = get_db_connection()
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        c.execute("SELECT * FROM records WHERE sub_batch_id=?", (sb_id,))
        row = c.fetchone()
        conn.close()
        
        if row:
            record_data = dict(row)
            
            # Update reprint audit trail silently
            try:
                conn_update = get_db_connection()
                c_update = conn_update.cursor()
                now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                c_update.execute('''UPDATE records 
                                  SET reprint_count = reprint_count + 1, 
                                      last_reprinted_at = ?, 
                                      last_reprinted_by = ? 
                                  WHERE sub_batch_id = ?''', 
                               (now_str, getattr(self, 'app_user_id', ''), sb_id))
                conn_update.commit()
                conn_update.close()
                self.refresh_records_treeview() # Update UI to show new reprint count
                
                record_data['reprint_count'] = int(record_data.get('reprint_count') or 0) + 1
                record_data['last_reprinted_by'] = getattr(self, 'app_user_id', '')
                record_data['last_reprinted_at'] = now_str
            except Exception as e:
                print(f"Failed to update reprint audit trail: {e}")
                
            # Print silently
            print_html_slip(record_data, silent=True)
            
        else:
            messagebox.showerror("Error", f"Record not found: {sb_id}")

    def build_tab_inventory(self):
        main_frame = tk.Frame(self.tab_inventory, bg=BG_COLOR)
        main_frame.pack(fill=tk.BOTH, expand=True)
        tk.Label(main_frame, text="Live Inventory Dashboard", bg=BG_COLOR, fg=ACCENT_COLOR, font=HMI_FONT_L).pack(pady=(10, 5))
        self.stats_frame = tk.Frame(main_frame, bg=BG_COLOR)
        self.stats_frame.pack(fill=tk.X, padx=20, pady=5)
        self.lbl_stat_total = self.create_stat_card(self.stats_frame, "Total Boxes in Rack", "0", ACCENT_COLOR)
        self.lbl_stat_oldest = self.create_stat_card(self.stats_frame, "Oldest Box Age", "0 days", ERROR_COLOR)
        self.lbl_stat_low = self.create_stat_card(self.stats_frame, "Low WIP Alerts", "0", WARNING_COLOR)
        
        # Control Panel Row
        control_frame = tk.Frame(main_frame, bg=SURFACE_COLOR, bd=1, relief="solid")
        control_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Filter
        filter_frame = tk.Frame(control_frame, bg=SURFACE_COLOR)
        filter_frame.pack(side=tk.LEFT, padx=20, pady=10)
        tk.Label(filter_frame, text="Filter by PN:", font=HMI_FONT_S, bg=SURFACE_COLOR, fg=TEXT_MUTED).pack(side=tk.LEFT)
        self.inv_search_var = tk.StringVar()
        self.inv_search_var.trace_add('write', lambda *args: self.refresh_inventory())
        tk.Entry(filter_frame, textvariable=self.inv_search_var, font=HMI_FONT_M, bg=BG_COLOR, fg=TEXT_COLOR).pack(side=tk.LEFT, padx=10)
        
        # Pick List Generator
        pick_frame = tk.Frame(control_frame, bg=SURFACE_COLOR)
        pick_frame.pack(side=tk.RIGHT, padx=20, pady=10)
        tk.Label(pick_frame, text="Pick-List Generator | PN:", font=HMI_FONT_S, bg=SURFACE_COLOR, fg=TEXT_MUTED).pack(side=tk.LEFT)
        self.pick_pn_var = tk.StringVar()
        self.pick_pn_combo = ttk.Combobox(pick_frame, textvariable=self.pick_pn_var, font=HMI_FONT_M, width=18, state="readonly")
        self.pick_pn_combo.pack(side=tk.LEFT, padx=5)
        tk.Label(pick_frame, text="Qty:", font=HMI_FONT_S, bg=SURFACE_COLOR, fg=TEXT_MUTED).pack(side=tk.LEFT)
        self.pick_qty_var = tk.StringVar()
        tk.Entry(pick_frame, textvariable=self.pick_qty_var, font=HMI_FONT_M, width=8, bg=BG_COLOR, fg=TEXT_COLOR).pack(side=tk.LEFT, padx=5)
        ttk.Button(pick_frame, text="Generate Ticket", style="Accent.TButton", command=self.generate_pick_list).pack(side=tk.LEFT, padx=10)
        
        # PanedWindow for the tables
        paned = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        left_pane = tk.Frame(paned, bg=BG_COLOR, width=450)
        paned.add(left_pane, weight=2)
        
        right_pane = tk.Frame(paned, bg=BG_COLOR, width=850)
        paned.add(right_pane, weight=3)
        
        table_paned = ttk.PanedWindow(left_pane, orient=tk.VERTICAL)
        table_paned.pack(fill=tk.BOTH, expand=True, padx=(0, 10))

        # Row 1: Summary Table
        row1 = tk.Frame(table_paned, bg=BG_COLOR)
        table_paned.add(row1, weight=1)

        _, agg_card = self.create_card(row1, "Summary by Part Number (Right-Click to Set Min Threshold)")
        cols_agg = ("PN", "Part Name", "Total Qty", "Min Qty", "Box Count")
        self.tree_inv_agg = ttk.Treeview(agg_card, columns=cols_agg, show="headings", height=8)
        for col in cols_agg: self.tree_inv_agg.heading(col, text=col)
        self.tree_inv_agg.column("PN", width=120); self.tree_inv_agg.column("Part Name", width=120); self.tree_inv_agg.column("Total Qty", width=60); self.tree_inv_agg.column("Min Qty", width=60); self.tree_inv_agg.column("Box Count", width=60)
        self.tree_inv_agg.pack(fill=tk.BOTH, expand=True, pady=10, padx=10)
        
        self.tree_inv_agg.tag_configure("low_wip", background="#FFEBEE", foreground="#C62828")
        
        def set_min_threshold(event):
            item = self.tree_inv_agg.identify_row(event.y)
            if not item: return
            self.tree_inv_agg.selection_set(item)
            pn = self.tree_inv_agg.item(item, "values")[0]
            qty_str = simpledialog.askstring("Set Threshold", f"Enter minimum required quantity for {pn}:", parent=self)
            if qty_str and qty_str.isdigit():
                try:
                    conn = get_db_connection()
                    c = conn.cursor()
                    c.execute("INSERT OR REPLACE INTO part_thresholds (pn_sf, min_qty) VALUES (?, ?)", (pn, int(qty_str)))
                    conn.commit()
                    conn.close()
                    self.refresh_inventory()
                except Exception as e:
                    messagebox.showerror("Error", str(e))
        self.tree_inv_agg.bind("<Button-3>", set_min_threshold)
        
        # Row 2: Details Table
        row2 = tk.Frame(table_paned, bg=BG_COLOR)
        table_paned.add(row2, weight=1)
        
        _, det_card = self.create_card(row2, "Detailed Rack Content (FIFO)")
        cols_det = ("SB_ID", "PN", "Qty", "Age", "Status")
        self.tree_inv_det = ttk.Treeview(det_card, columns=cols_det, show="headings", height=10)
        for col in cols_det: self.tree_inv_det.heading(col, text=col)
        self.tree_inv_det.column("SB_ID", width=120); self.tree_inv_det.column("PN", width=120); self.tree_inv_det.column("Qty", width=50); self.tree_inv_det.column("Age", width=80); self.tree_inv_det.column("Status", width=60)
        
        self.tree_inv_det.tag_configure("age_green", background="#E8F5E9", foreground="#2E7D32")
        self.tree_inv_det.tag_configure("age_yellow", background="#FFFDE7", foreground="#F9A825")
        self.tree_inv_det.tag_configure("age_red", background="#FFEBEE", foreground="#C62828")
        
        scrolldet = ttk.Scrollbar(det_card, orient=tk.VERTICAL, command=self.tree_inv_det.yview)
        self.tree_inv_det.configure(yscroll=scrolldet.set)
        scrolldet.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_inv_det.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Right Pane: WIP Chart
        _, chart1_card = self.create_card(right_pane, "WIP Level by Part Number")
        
        if MATPLOTLIB_AVAILABLE:
            self.fig_wip = Figure(figsize=(5, 5), dpi=100, facecolor=SURFACE_COLOR)
            self.ax_wip = self.fig_wip.add_subplot(111)
            self.style_ax(self.ax_wip, "")
            self.lbl_chart_wip = tk.Label(chart1_card, bg=SURFACE_COLOR)
            self.lbl_chart_wip.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def create_stat_card(self, parent, title, initial_value, color):
        frame = tk.Frame(parent, bg=SURFACE_COLOR, bd=1, relief="solid", padx=20, pady=10)
        frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10)
        tk.Label(frame, text=title, font=HMI_FONT_S, bg=SURFACE_COLOR, fg=TEXT_MUTED).pack()
        lbl = tk.Label(frame, text=initial_value, font=HMI_FONT_XL, bg=SURFACE_COLOR, fg=color)
        lbl.pack()
        return lbl

    def generate_pick_list(self):
        pn = self.pick_pn_var.get().strip()
        qty_str = self.pick_qty_var.get().strip()
        if not pn or not qty_str.isdigit():
            messagebox.showerror("Error", "Please enter a valid Part Number and numeric Quantity.")
            return
        target_qty = int(qty_str)
        
        conn = get_db_connection()
        c = conn.cursor()
        c.execute('''SELECT sub_batch_id, quantity, dt_sp 
                     FROM records 
                     WHERE status="In Rack" AND pn_sf=? 
                     ORDER BY dt_sp ASC''', (pn,))
        rows = c.fetchall()
        conn.close()
        
        if not rows:
            messagebox.showinfo("Pick List", f"No inventory found in rack for {pn}")
            return
            
        picked = []
        accumulated = 0
        for r in rows:
            picked.append((r[0], r[1], r[2]))
            accumulated += r[1]
            if accumulated >= target_qty:
                break
                
        if accumulated < target_qty:
            messagebox.showwarning("Shortage", f"Only {accumulated} units available in the rack. Cannot fulfill {target_qty}.")
            
        # Display the Pick List
        popup = tk.Toplevel(self)
        popup.title(f"FIFO Pick Ticket - {pn}")
        center_window(popup, 500, 600)
        popup.configure(bg=BG_COLOR)
        popup.transient(self)
        popup.grab_set()
        
        tk.Label(popup, text=f"FIFO Pick Ticket", font=HMI_FONT_L, bg=BG_COLOR, fg=ACCENT_COLOR).pack(pady=10)
        tk.Label(popup, text=f"Part: {pn} | Target: {target_qty} | Actual: {accumulated}", font=HMI_FONT_M, bg=BG_COLOR, fg=TEXT_COLOR).pack(pady=5)
        
        frame = tk.Frame(popup, bg=SURFACE_COLOR)
        frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        for i, (sb, q, dt) in enumerate(picked):
            tk.Label(frame, text=f"{i+1}. {sb} (Qty: {q})", font=HMI_FONT_M, bg=SURFACE_COLOR, fg=TEXT_COLOR, anchor="w").pack(fill=tk.X, padx=10, pady=2)
            tk.Label(frame, text=f"   Stored: {dt}", font=HMI_FONT_S, bg=SURFACE_COLOR, fg=TEXT_MUTED, anchor="w").pack(fill=tk.X, padx=10)

        btn_frame = tk.Frame(popup, bg=BG_COLOR)
        btn_frame.pack(pady=20)
        
        def print_ticket():
            zpl_data = generate_pick_ticket_zpl(pn, target_qty, accumulated, picked)
            execute_ticket_print(zpl_data, pn)
            
        ttk.Button(btn_frame, text="Print Ticket", style="Accent.TButton", command=print_ticket).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="Close", style="Secondary.TButton", command=popup.destroy).pack(side=tk.LEFT, padx=10)

    def refresh_inventory(self):
        search_q = getattr(self, 'inv_search_var', tk.StringVar()).get().strip().lower()
        
        conn = get_db_connection()
        c = conn.cursor()
        
        # Aggregated with threshold
        c.execute('''SELECT r.pn_sf, r.part_sf, SUM(r.quantity), COUNT(r.id), COALESCE(t.min_qty, 0)
                     FROM records r
                     LEFT JOIN part_thresholds t ON r.pn_sf = t.pn_sf
                     WHERE r.status="In Rack"
                     GROUP BY r.pn_sf ORDER BY SUM(r.quantity) DESC''')
        agg_rows = c.fetchall()
        
        # Detailed
        c.execute('''SELECT sub_batch_id, pn_sf, quantity, dt_sp, status 
                     FROM records WHERE status="In Rack" ORDER BY pn_sf, dt_sp ASC''')
        det_rows = c.fetchall()
        conn.close()
        
        for item in self.tree_inv_agg.get_children(): self.tree_inv_agg.delete(item)
        for item in self.tree_inv_det.get_children(): self.tree_inv_det.delete(item)
        
        pn_labels = []
        qty_values = []
        min_values = []
        part_labels = []
        full_pns = []
        low_wip_count = 0
        total_boxes = 0
        
        for row in agg_rows:
            pn, part, total_qty, box_count, min_qty = row
            if search_q and search_q not in pn.lower() and search_q not in part.lower():
                continue
            
            tag = ""
            if total_qty < min_qty:
                tag = "low_wip"
                low_wip_count += 1
                
            self.tree_inv_agg.insert("", "end", values=(pn, part, total_qty, min_qty, box_count), tags=(tag,))
            pn_labels.append(pn[:15])
            part_labels.append(part[:25])
            full_pns.append(pn)
            qty_values.append(total_qty or 0)
            min_values.append(min_qty or 0)
            total_boxes += box_count
            
        now = datetime.datetime.now()
        age_counts = {"< 2 Days": 0, "2-7 Days": 0, "> 7 Days": 0}
        max_age_days = 0
        max_age_hours = 0
        
        # Track oldest box per PN to add star
        pn_oldest_dt = {}
        for row in det_rows:
            pn = row[1]
            dt_str = row[3]
            try:
                try:
                    dt_obj = datetime.datetime.strptime(dt_str, "%Y-%m-%d %H:%M:%S")
                except ValueError:
                    dt_obj = datetime.datetime.strptime(dt_str, "%Y-%m-%d %H:%M")
                if pn not in pn_oldest_dt or dt_obj < pn_oldest_dt[pn]:
                    pn_oldest_dt[pn] = dt_obj
            except: pass
            
        for row in det_rows:
            pn = row[1]
            if search_q and search_q not in pn.lower():
                continue
                
            dt_str = row[3]
            try:
                try:
                    dt_obj = datetime.datetime.strptime(dt_str, "%Y-%m-%d %H:%M:%S")
                except ValueError:
                    dt_obj = datetime.datetime.strptime(dt_str, "%Y-%m-%d %H:%M")
                diff = now - dt_obj
                hours = diff.total_seconds() / 3600
                days = diff.days
                
                if days > max_age_days or (days == max_age_days and int(hours % 24) > max_age_hours):
                    max_age_days = days
                    max_age_hours = int(hours % 24)
                
                age_str = f"{days}d {int(hours % 24)}h"
                if dt_obj == pn_oldest_dt.get(pn):
                    age_str += " ⭐"
                
                if days < 2:
                    tag = "age_green"
                    age_counts["< 2 Days"] += 1
                elif days <= 7:
                    tag = "age_yellow"
                    age_counts["2-7 Days"] += 1
                else:
                    tag = "age_red"
                    age_counts["> 7 Days"] += 1
            except Exception:
                age_str = "Unknown"
                tag = "age_green"
                age_counts["< 2 Days"] += 1
                
            vals = (row[0], pn, row[2], age_str, row[4])
            self.tree_inv_det.insert("", "end", values=vals, tags=(tag,))
            
        # Update Quick Stats
        if hasattr(self, 'lbl_stat_total'):
            self.lbl_stat_total.config(text=str(total_boxes))
            self.lbl_stat_oldest.config(text=f"{max_age_days}d {max_age_hours}h")
            self.lbl_stat_low.config(text=str(low_wip_count))
            
        # Update Charts
        if MATPLOTLIB_AVAILABLE and hasattr(self, 'ax_wip'):
            self.ax_wip.clear()
            self.style_ax(self.ax_wip, "")
            if part_labels:
                top_part = part_labels[:8]
                top_qty = qty_values[:8]
                top_min = min_values[:8]
                
                x = np.arange(len(top_part))
                width = 0.35
                
                rects1 = self.ax_wip.bar(x - width/2, top_qty, width, color='#38C958')
                rects2 = self.ax_wip.bar(x + width/2, top_min, width, color='#FFA000')
                
                for rect in rects1 + rects2:
                    height = rect.get_height()
                    if height > 0:
                        self.ax_wip.annotate(f"{int(height)}",
                                             xy=(rect.get_x() + rect.get_width() / 2, height),
                                             xytext=(0, 3),
                                             textcoords="offset points",
                                             ha='center', va='bottom', color=TEXT_COLOR, fontsize=8)
                
                self.ax_wip.set_xticks(x)
                self.ax_wip.set_xticklabels(top_part, rotation=20, ha='right', fontsize=8)
                self.ax_wip.tick_params(axis='y', labelsize=8)
            else:
                self.ax_wip.text(0.5, 0.5, "No Data", color=TEXT_MUTED, ha="center")
            self.fig_wip.subplots_adjust(left=0.35, right=0.95, top=0.90, bottom=0.35)
            self.render_chart_to_label(self.fig_wip, self.lbl_chart_wip)

        if hasattr(self, 'pick_pn_combo'):
            self.pick_pn_combo['values'] = full_pns


    def render_chart_to_label(self, fig, label_widget):
        if not MATPLOTLIB_AVAILABLE: return
        canvas = FigureCanvasAgg(fig)
        canvas.draw()
        rgba = np.asarray(canvas.buffer_rgba())
        img = Image.fromarray(rgba)
        photo = ImageTk.PhotoImage(img)
        label_widget.configure(image=photo)
        label_widget.image = photo

    def style_ax(self, ax, title):
        ax.set_facecolor(SURFACE_COLOR)
        ax.tick_params(colors=TEXT_COLOR)
        ax.set_title(title, color=TEXT_COLOR, pad=15)
        for spine in ax.spines.values():
            spine.set_color(BORDER_COLOR)

    def build_tab3(self):
        main_frame = tk.Frame(self.tab3, bg=BG_COLOR)
        main_frame.pack(fill=tk.BOTH, expand=True)

        if not MATPLOTLIB_AVAILABLE:
            tk.Label(main_frame, text=" Matplotlib is required to view KPIs.\nPlease run 'pip install matplotlib' and restart.", bg=BG_COLOR, fg=ERROR_COLOR, font=HMI_FONT_L).pack(pady=50, padx=50)
            return

        # Top Cards Row
        self.kpi_cards_frame = tk.Frame(main_frame, bg=BG_COLOR)
        self.kpi_cards_frame.pack(fill=tk.X, pady=10, padx=20)
        
        self.lbl_kpi_total = tk.Label(self.kpi_cards_frame, text="Daily Total: 0", bg=SURFACE_COLOR, fg=TEXT_COLOR, font=HMI_FONT_L, padx=20, pady=10)
        self.lbl_kpi_total.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=10)
        
        self.lbl_kpi_top_op = tk.Label(self.kpi_cards_frame, text="Top Operator: -", bg=SURFACE_COLOR, fg=TEXT_COLOR, font=HMI_FONT_L, padx=20, pady=10)
        self.lbl_kpi_top_op.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=10)
        
        self.lbl_kpi_oee = tk.Label(self.kpi_cards_frame, text="Quality: - | Perf: -", bg=SURFACE_COLOR, fg=TEXT_COLOR, font=HMI_FONT_L, padx=20, pady=10)
        self.lbl_kpi_oee.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=10)
        
        controls_frame = tk.Frame(main_frame, bg=BG_COLOR)
        controls_frame.pack(pady=5)

        tk.Label(controls_frame, text="Select Date:", bg=BG_COLOR, fg=TEXT_COLOR, font=HMI_FONT_M).pack(side=tk.LEFT, padx=10)
        self.kpi_date_entry = DateEntry(controls_frame, width=12, background=ACCENT_COLOR, foreground='white', borderwidth=2)
        self.kpi_date_entry.pack(side=tk.LEFT, padx=10)

        tk.Label(controls_frame, text="Shift:", bg=BG_COLOR, fg=TEXT_COLOR, font=HMI_FONT_M).pack(side=tk.LEFT, padx=10)
        self.cb_kpi_shift = ttk.Combobox(controls_frame, values=["All", "A", "B", "C"], state="readonly", width=8, font=HMI_FONT_M)
        self.cb_kpi_shift.pack(side=tk.LEFT, padx=5)

        user_role = getattr(self, 'app_user_role', 'Operator')
        if user_role in ["Admin", "Supervisor", "Quality OP"]:
            self.cb_kpi_shift.set("All")
        else:
            self.cb_kpi_shift.set(getattr(self, 'app_user_shift', 'A'))

        ttk.Button(controls_frame, text="Refresh KPIs", style="Primary.TButton", command=self.refresh_kpis).pack(side=tk.LEFT, padx=10)

        # Notebook for Charts
        self.kpi_notebook = ttk.Notebook(main_frame)
        self.kpi_notebook.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        self.kpi_tab_daily = tk.Frame(self.kpi_notebook, bg=BG_COLOR)
        self.kpi_tab_trends = tk.Frame(self.kpi_notebook, bg=BG_COLOR)
        
        self.kpi_notebook.add(self.kpi_tab_daily, text="Daily Production KPIs")
        self.kpi_notebook.add(self.kpi_tab_trends, text="30-Day WIP Trends")

        # Charts Area (Daily)
        self.charts_frame_1 = tk.Frame(self.kpi_tab_daily, bg=BG_COLOR)
        self.charts_frame_1.pack(fill=tk.BOTH, expand=True, pady=10, padx=20)
        
        self.fig_hourly = Figure(figsize=(5, 3), dpi=100)
        self.ax_hourly = self.fig_hourly.add_subplot(111)
        self.lbl_chart_hourly = tk.Label(self.charts_frame_1, bg=BG_COLOR)
        self.lbl_chart_hourly.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10)
        
        self.fig_shift = Figure(figsize=(5, 3), dpi=100)
        self.ax_shift = self.fig_shift.add_subplot(111)
        self.lbl_chart_shift = tk.Label(self.charts_frame_1, bg=BG_COLOR)
        self.lbl_chart_shift.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10)

        self.charts_frame_2 = tk.Frame(self.kpi_tab_daily, bg=BG_COLOR)
        self.charts_frame_2.pack(fill=tk.BOTH, expand=True, pady=10, padx=20)
        
        self.fig_op = Figure(figsize=(5, 3), dpi=100)
        self.ax_op = self.fig_op.add_subplot(111)
        self.lbl_chart_op = tk.Label(self.charts_frame_2, bg=BG_COLOR)
        self.lbl_chart_op.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10)
        
        self.fig_pn = Figure(figsize=(5, 3), dpi=100)
        self.ax_pn = self.fig_pn.add_subplot(111)
        self.lbl_chart_pn = tk.Label(self.charts_frame_2, bg=BG_COLOR)
        self.lbl_chart_pn.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10)
        
        # 30-Day Trend Chart Area (Trends)
        tk.Label(self.kpi_tab_trends, text="WIP Trend Analysis (Last 30 Days)", bg=BG_COLOR, fg=ACCENT_COLOR, font=HMI_FONT_L).pack(pady=(20, 5))
        self.charts_frame_3 = tk.Frame(self.kpi_tab_trends, bg=BG_COLOR)
        self.charts_frame_3.pack(fill=tk.BOTH, expand=True, pady=10, padx=20)
        
        self.fig_trend = Figure(figsize=(10, 3), dpi=100)
        self.ax_trend = self.fig_trend.add_subplot(111)
        self.lbl_chart_trend = tk.Label(self.charts_frame_3, bg=BG_COLOR)
        self.lbl_chart_trend.pack(fill=tk.BOTH, expand=True, padx=10)
        
        # Style all figures
        for fig, ax, title in [(self.fig_hourly, self.ax_hourly, "Hourly Production"),
                               (self.fig_shift, self.ax_shift, "Shift Output"),
                               (self.fig_op, self.ax_op, "Operator Mix"),
                               (self.fig_pn, self.ax_pn, "Product Mix"),
                               (self.fig_trend, self.ax_trend, "30-Day WIP Accumulation by Part Number")]:
            fig.patch.set_facecolor(BG_COLOR)
            self.style_ax(ax, title)
            
        self.render_chart_to_label(self.fig_hourly, self.lbl_chart_hourly)
        self.render_chart_to_label(self.fig_shift, self.lbl_chart_shift)
        self.render_chart_to_label(self.fig_op, self.lbl_chart_op)
        self.render_chart_to_label(self.fig_pn, self.lbl_chart_pn)
        self.render_chart_to_label(self.fig_trend, self.lbl_chart_trend)
            
    def refresh_kpis(self):
        if not MATPLOTLIB_AVAILABLE:
            return
        try:
            conn = get_db_connection()
            c = conn.cursor()
            
            if hasattr(self, 'kpi_date_entry'):
                selected_date = self.kpi_date_entry.get_date()
            else:
                selected_date = datetime.datetime.now().date()
                
            start_dt = selected_date.strftime("%Y-%m-%d 06:00:00")
            end_dt = (selected_date + datetime.timedelta(days=1)).strftime("%Y-%m-%d 05:59:59")
            
            shift_filter = getattr(self, 'cb_kpi_shift', None)
            selected_shift = shift_filter.get() if shift_filter else "All"
            
            if selected_shift == "All":
                shift_cond = ""
                shift_params = []
            else:
                shift_cond = " AND shift_sp = ?"
                shift_params = [selected_shift]
            
            # 1. Total Daily Production
            c.execute(f"SELECT SUM(quantity) FROM records WHERE dt_sp >= ? AND dt_sp <= ?{shift_cond}", [start_dt, end_dt] + shift_params)
            total_daily = c.fetchone()[0] or 0
            self.lbl_kpi_total.config(text=f"Daily Total: {total_daily}")
            
            # 2. Top Operator
            c.execute(f"SELECT op_id, SUM(quantity) as q FROM records WHERE dt_sp >= ? AND dt_sp <= ?{shift_cond} GROUP BY op_id ORDER BY q DESC LIMIT 1", [start_dt, end_dt] + shift_params)
            top_op = c.fetchone()
            top_op_txt = f"{top_op[0]} ({top_op[1]})" if top_op else "-"
            self.lbl_kpi_top_op.config(text=f"Top Operator: {top_op_txt}")
            
            # OEE Calculations
            c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='quality_defects'")
            has_qual = c.fetchone()
            defective_qty = 0
            if has_qual:
                c.execute(f"SELECT SUM(qty_defective) FROM quality_defects WHERE reported_at >= ? AND reported_at <= ?{shift_cond}", [start_dt, end_dt] + shift_params)
                res_def = c.fetchone()
                defective_qty = res_def[0] or 0
                
            qual_rate = ((total_daily - defective_qty) / total_daily * 100) if total_daily > 0 else 0
            
            # Fetch Targets
            c.execute("SELECT product_pn, target_qty FROM shift_targets ORDER BY id ASC")
            pn_targets_dict = {}
            for row in c.fetchall():
                pn, tqty = row
                pn_targets_dict[pn] = tqty
                
            total_target = sum(pn_targets_dict.values())
                    
            perf_rate = (total_daily / total_target * 100) if total_target > 0 else 0
            self.lbl_kpi_oee.config(text=f"Quality: {qual_rate:.1f}% | Perf: {perf_rate:.1f}%")
            
            # 3. Hourly Production
            self.ax_hourly.clear()
            self.style_ax(self.ax_hourly, "Hourly Production")
            
            c.execute(f"SELECT substr(dt_sp, 12, 2) as hr, SUM(quantity) FROM records WHERE dt_sp >= ? AND dt_sp <= ?{shift_cond} GROUP BY hr", [start_dt, end_dt] + shift_params)
            hourly_data = {row[0]: (row[1] or 0) for row in c.fetchall() if row[0] is not None}
            hrs = [f"{h:02d}" for h in range(24)]
            qts = [hourly_data.get(h, 0) for h in hrs]
            
            self.ax_hourly.plot(hrs, qts, color=ACCENT_COLOR, marker='o')
            max_qt = max(qts) if qts else 0
            self.ax_hourly.set_ylim(bottom=0, top=max(max_qt * 1.15, 1))
            for i, v in enumerate(qts):
                if v > 0:
                    self.ax_hourly.text(hrs[i], v + (max_qt*0.02 if max_qt>0 else 1), str(v), color=TEXT_COLOR, ha='center', va='bottom', fontsize=8)
            
            self.ax_hourly.grid(True, color=BORDER_COLOR, alpha=0.3)
            self.ax_hourly.set_xticks(hrs[::2])
            self.render_chart_to_label(self.fig_hourly, self.lbl_chart_hourly)
            
            # 4. PN Comparison (Target vs Actual)
            self.ax_shift.clear()
            self.style_ax(self.ax_shift, "Actual vs Target per PN")
            
            c.execute(f"SELECT pn_sf, SUM(quantity) FROM records WHERE dt_sp >= ? AND dt_sp <= ?{shift_cond} GROUP BY pn_sf", [start_dt, end_dt] + shift_params)
            pn_actuals_dict = {row[0]: (row[1] or 0) for row in c.fetchall() if row[0] is not None}
            
            pns = list(set(pn_actuals_dict.keys()).union(set(pn_targets_dict.keys())))
            pns = sorted([p for p in pns if pn_actuals_dict.get(p, 0) > 0])
            
            sqts = [pn_actuals_dict.get(p, 0) for p in pns]
            tqts = [pn_targets_dict.get(p, 0) for p in pns]
            
            # Use part name from SF_DATA if available, else short part number
            pn_labels = []
            for p in pns:
                if p in SF_DATA and SF_DATA[p][0]:
                    pn_labels.append(SF_DATA[p][0])
                else:
                    pn_labels.append(p.split('-')[-1] if '-' in p else p[:8])
            
            x = range(len(pns))
            width = 0.2
            
            if len(pns) > 0:
                bars1 = self.ax_shift.bar([i - width/2 for i in x], sqts, width, color=SUCCESS_COLOR, label='Actual')
                bars2 = self.ax_shift.bar([i + width/2 for i in x], tqts, width, color=WARNING_COLOR, label='Target')
                
                max_sqt = max(max(sqts) if sqts else 0, max(tqts) if tqts else 0)
                self.ax_shift.set_ylim(bottom=0, top=max(max_sqt * 1.15, 1))
                
                for bars in [bars1, bars2]:
                    for bar in bars:
                        yval = bar.get_height()
                        if yval > 0:
                            self.ax_shift.text(bar.get_x() + bar.get_width()/2.0, yval + (max_sqt*0.02 if max_sqt>0 else 1), int(yval), va='bottom', ha='center', color=TEXT_COLOR, fontsize=8)
                            
                self.ax_shift.set_xticks(x)
                self.ax_shift.set_xticklabels(pn_labels, rotation=15, fontsize=8)
                self.ax_shift.legend(loc='center left', bbox_to_anchor=(1.02, 0.5), ncol=1, fontsize=8, facecolor=BG_COLOR, edgecolor=BORDER_COLOR, labelcolor=TEXT_COLOR)
            else:
                self.ax_shift.text(0.5, 0.5, "No Data", ha='center', va='center', color=TEXT_MUTED, transform=self.ax_shift.transAxes)
            self.fig_shift.subplots_adjust(bottom=0.25, top=0.82, right=0.75)
            self.render_chart_to_label(self.fig_shift, self.lbl_chart_shift)
            
            # 5. Production by Operator Pie
            self.ax_op.clear()
            self.style_ax(self.ax_op, "Operator Mix")
            
            c.execute(f"SELECT op_id, SUM(quantity) FROM records WHERE dt_sp >= ? AND dt_sp <= ?{shift_cond} GROUP BY op_id", [start_dt, end_dt] + shift_params)
            op_data = c.fetchall()
            if op_data:
                ops = [row[0] for row in op_data]
                o_qts = [row[1] for row in op_data]
                self.ax_op.pie(o_qts, labels=ops, autopct='%1.1f%%', textprops={'color': TEXT_COLOR})
            self.render_chart_to_label(self.fig_op, self.lbl_chart_op)
            
            # 6. Top PNs Pie
            self.ax_pn.clear()
            self.style_ax(self.ax_pn, "Product Mix")
            
            c.execute(f"SELECT pn_sf, SUM(quantity) FROM records WHERE dt_sp >= ? AND dt_sp <= ?{shift_cond} GROUP BY pn_sf ORDER BY SUM(quantity) DESC LIMIT 5", [start_dt, end_dt] + shift_params)
            pn_data = c.fetchall()
            if pn_data:
                pns_labels = []
                for row in pn_data:
                    p = row[0]
                    if p in SF_DATA and SF_DATA[p][0]:
                        pns_labels.append(SF_DATA[p][0])
                    else:
                        pns_labels.append(p.split('-')[-1] if '-' in p else p[:8])
                p_qts = [row[1] for row in pn_data]
                self.ax_pn.pie(p_qts, labels=pns_labels, autopct='%1.1f%%', textprops={'color': TEXT_COLOR})
            self.render_chart_to_label(self.fig_pn, self.lbl_chart_pn)

            # 7. 30-Day WIP Trend Chart
            self.ax_trend.clear()
            self.style_ax(self.ax_trend, "30-Day WIP Accumulation by Part Number")
            thirty_days_ago = (datetime.datetime.now() - datetime.timedelta(days=30)).strftime("%Y-%m-%d")
            c.execute('''SELECT snapshot_date, pn_sf, total_qty_in_rack 
                         FROM inventory_snapshots 
                         WHERE snapshot_date >= ?
                         ORDER BY snapshot_date ASC''', (thirty_days_ago,))
            trend_data = c.fetchall()
            
            if trend_data:
                pn_series = {}
                dates_set = set()
                for r in trend_data:
                    dt, pn, qty = r
                    if pn not in pn_series: pn_series[pn] = {}
                    pn_series[pn][dt] = qty
                    dates_set.add(dt)
                
                sorted_dates = sorted(list(dates_set))
                x_labels = [d[5:] for d in sorted_dates]
                
                for pn, data_dict in pn_series.items():
                    y_vals = [data_dict.get(d, 0) for d in sorted_dates]
                    self.ax_trend.plot(x_labels, y_vals, marker='o', markersize=4, label=pn[:15])
                    
                self.ax_trend.legend(loc='center left', bbox_to_anchor=(1.02, 0.5), ncol=1, fontsize=8, facecolor=BG_COLOR, edgecolor=BORDER_COLOR, labelcolor=TEXT_COLOR)
                self.ax_trend.tick_params(axis='x', rotation=45, labelsize=8)
            else:
                self.ax_trend.text(0.5, 0.5, "No Snapshot Data Yet\n(Snapshots are taken nightly at 23:50)", ha='center', va='center', color=TEXT_MUTED, transform=self.ax_trend.transAxes)
                
            self.fig_trend.subplots_adjust(bottom=0.3, top=0.85, right=0.80)
            self.render_chart_to_label(self.fig_trend, self.lbl_chart_trend)

            
            conn.close()
        except Exception as e:
            print("KPI Refresh Error:", e)

    def export_kpis_to_excel(self):
        try:
            conn = get_db_connection()
            c = conn.cursor()
            
            # Production day logic (06:00 to 05:59)
            prod_date_sql = "CASE WHEN CAST(substr(dt_sp, 12, 2) AS INTEGER) < 6 THEN date(substr(dt_sp, 1, 10), '-1 day') ELSE substr(dt_sp, 1, 10) END"
            
            c.execute(f"SELECT DISTINCT {prod_date_sql} as prod_date FROM records ORDER BY prod_date DESC")
            dates = [r[0] for r in c.fetchall() if r[0]]
            
            c.execute("SELECT SUM(quantity) FROM records")
            grand_total = c.fetchone()[0] or 0

            if not os.path.exists(EXCEL_FILE):
                wb = Workbook()
                ws = wb.active
                ws.title = "Advanced KPI Reports"
            else:
                wb = load_workbook(EXCEL_FILE)
                if "Advanced KPI Reports" in wb.sheetnames:
                    ws = wb["Advanced KPI Reports"]
                    ws.delete_rows(1, ws.max_row)
                else:
                    ws = wb.create_sheet("Advanced KPI Reports")
                    
                # Clean up old sheet if exists
                if "KPI Reports" in wb.sheetnames:
                    del wb["KPI Reports"]
            
            # Define Styles
            header_font = Font(name='Segoe UI', size=16, bold=True, color="FFFFFF")
            header_fill = PatternFill(start_color="1B232C", end_color="1B232C", fill_type="solid")
            
            date_font = Font(name='Segoe UI', size=14, bold=True, color="FFFFFF")
            date_fill = PatternFill(start_color="0078D7", end_color="0078D7", fill_type="solid")
            
            sub_header_font = Font(name='Segoe UI', size=11, bold=True, color="FFFFFF")
            sub_header_fill = PatternFill(start_color="28A745", end_color="28A745", fill_type="solid")
            
            data_font = Font(name='Segoe UI', size=11, color="000000")
            bold_data_font = Font(name='Segoe UI', size=11, bold=True, color="000000")
            
            align_center = Alignment(horizontal="center", vertical="center")
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            
            # Main Header
            ws.merge_cells('A1:K2')
            top_cell = ws.cell(row=1, column=1, value=f"KPI REPORT | GENERATED: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')} | TOTAL ALL-TIME PRODUCTION: {grand_total}")
            top_cell.font = header_font
            top_cell.fill = header_fill
            top_cell.alignment = align_center
            
            current_row = 4
            
            for d in dates:
                c.execute(f"SELECT SUM(quantity) FROM records WHERE {prod_date_sql} = ?", (d,))
                d_total = c.fetchone()[0] or 0
                
                c.execute(f"SELECT shift_sp, SUM(quantity) FROM records WHERE {prod_date_sql} = ? GROUP BY shift_sp ORDER BY shift_sp", (d,))
                shifts = c.fetchall()
                
                c.execute(f"SELECT op_id, SUM(quantity) as q FROM records WHERE {prod_date_sql} = ? GROUP BY op_id ORDER BY q DESC LIMIT 5", (d,))
                ops = c.fetchall()
                
                c.execute(f"SELECT pn_sf, SUM(quantity) as q FROM records WHERE {prod_date_sql} = ? GROUP BY pn_sf ORDER BY q DESC LIMIT 5", (d,))
                pns = c.fetchall()
                
                c.execute(f"SELECT station, SUM(quantity) as q FROM records WHERE {prod_date_sql} = ? GROUP BY station ORDER BY q DESC LIMIT 5", (d,))
                stations = c.fetchall()
                
                # Date Header
                ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=11)
                d_cell = ws.cell(row=current_row, column=1, value=f"PRODUCTION DATE: {d}   |   TOTAL DAILY OUTPUT: {d_total}")
                d_cell.font = date_font
                d_cell.fill = date_fill
                d_cell.alignment = align_center
                current_row += 1
                
                # Grid Headers
                headers = [
                    (1, "SHIFT", "OUTPUT"),
                    (4, "TOP 5 OPERATORS", "OUTPUT"),
                    (7, "TOP 5 PRODUCTS", "OUTPUT"),
                    (10, "STATIONS", "OUTPUT")
                ]
                
                for col_start, t1, t2 in headers:
                    ws.cell(row=current_row, column=col_start, value=t1).font = sub_header_font
                    ws.cell(row=current_row, column=col_start).fill = sub_header_fill
                    ws.cell(row=current_row, column=col_start).alignment = align_center
                    ws.cell(row=current_row, column=col_start).border = thin_border
                    
                    ws.cell(row=current_row, column=col_start+1, value=t2).font = sub_header_font
                    ws.cell(row=current_row, column=col_start+1).fill = sub_header_fill
                    ws.cell(row=current_row, column=col_start+1).alignment = align_center
                    ws.cell(row=current_row, column=col_start+1).border = thin_border
                
                current_row += 1
                
                # Grid Data (5 rows)
                start_data_row = current_row
                for i in range(5):
                    # Shift
                    if i < len(shifts):
                        ws.cell(row=start_data_row+i, column=1, value=f"Shift {shifts[i][0]}").font = data_font
                        ws.cell(row=start_data_row+i, column=2, value=shifts[i][1]).font = bold_data_font
                    else:
                        ws.cell(row=start_data_row+i, column=1, value="").font = data_font
                        ws.cell(row=start_data_row+i, column=2, value="").font = data_font
                        
                    # Operator
                    if i < len(ops):
                        ws.cell(row=start_data_row+i, column=4, value=ops[i][0]).font = data_font
                        ws.cell(row=start_data_row+i, column=5, value=ops[i][1]).font = bold_data_font
                    else:
                        ws.cell(row=start_data_row+i, column=4, value="")
                        ws.cell(row=start_data_row+i, column=5, value="")
                        
                    # Product
                    if i < len(pns):
                        ws.cell(row=start_data_row+i, column=7, value=pns[i][0]).font = data_font
                        ws.cell(row=start_data_row+i, column=8, value=pns[i][1]).font = bold_data_font
                    else:
                        ws.cell(row=start_data_row+i, column=7, value="")
                        ws.cell(row=start_data_row+i, column=8, value="")
                        
                    # Station
                    if i < len(stations):
                        ws.cell(row=start_data_row+i, column=10, value=stations[i][0]).font = data_font
                        ws.cell(row=start_data_row+i, column=11, value=stations[i][1]).font = bold_data_font
                    else:
                        ws.cell(row=start_data_row+i, column=10, value="")
                        ws.cell(row=start_data_row+i, column=11, value="")
                        
                    # Apply Borders
                    for c_idx in [1,2,4,5,7,8,10,11]:
                        ws.cell(row=start_data_row+i, column=c_idx).border = thin_border
                        ws.cell(row=start_data_row+i, column=c_idx).alignment = align_center

                current_row += 6

            ws.column_dimensions['A'].width = 15
            ws.column_dimensions['B'].width = 15
            ws.column_dimensions['C'].width = 3
            ws.column_dimensions['D'].width = 25
            ws.column_dimensions['E'].width = 15
            ws.column_dimensions['F'].width = 3
            ws.column_dimensions['G'].width = 35
            ws.column_dimensions['H'].width = 15
            ws.column_dimensions['I'].width = 3
            ws.column_dimensions['J'].width = 15
            ws.column_dimensions['K'].width = 15
            
            ws.sheet_view.showGridLines = False

            conn.close()
            wb.save(EXCEL_FILE)
        except Exception as e:
            print(f"Background KPI Export Error: {e}")

    def open_excel(self):
        if os.path.exists(EXCEL_FILE):
            os.startfile(EXCEL_FILE)
        else:
            messagebox.showinfo("Info", "Excel file not created yet. Save a record first.")

    def take_snapshot(self, date_str):
        conn = get_db_connection()
        c = conn.cursor()
        
        c.execute('''SELECT pn_sf, SUM(quantity), COUNT(id)
                     FROM records WHERE status="In Rack" GROUP BY pn_sf''')
        agg_rows = c.fetchall()
        
        c.execute('''SELECT pn_sf, dt_sp FROM records WHERE status="In Rack"''')
        det_rows = c.fetchall()
        
        now = datetime.datetime.now()
        oldest_dt_per_pn = {}
        for row in det_rows:
            pn, dt_str = row
            try:
                try:
                    dt_obj = datetime.datetime.strptime(dt_str, "%Y-%m-%d %H:%M:%S")
                except ValueError:
                    dt_obj = datetime.datetime.strptime(dt_str, "%Y-%m-%d %H:%M")
                if pn not in oldest_dt_per_pn or dt_obj < oldest_dt_per_pn[pn]:
                    oldest_dt_per_pn[pn] = dt_obj
            except: pass
            
        for row in agg_rows:
            pn, total_qty, box_count = row
            oldest_age_hours = 0.0
            if pn in oldest_dt_per_pn:
                oldest_age_hours = (now - oldest_dt_per_pn[pn]).total_seconds() / 3600.0
                
            c.execute('''INSERT OR REPLACE INTO inventory_snapshots 
                         (snapshot_date, pn_sf, boxes_in_rack, total_qty_in_rack, oldest_box_age_hours)
                         VALUES (?, ?, ?, ?, ?)''', 
                         (date_str, pn, box_count, total_qty, oldest_age_hours))
        conn.commit()
        conn.close()
        print(f"[{datetime.datetime.now()}] Automated Inventory Snapshot Taken for {date_str}.")
        
    def schedule_daily_snapshot(self):
        now = datetime.datetime.now()
        yesterday = now - datetime.timedelta(days=1)
        yesterday_str = yesterday.strftime("%Y-%m-%d")
        
        try:
            conn = get_db_connection()
            c = conn.cursor()
            c.execute("SELECT COUNT(*) FROM inventory_snapshots WHERE snapshot_date=?", (yesterday_str,))
            count = c.fetchone()[0]
            conn.close()
            
            if count == 0:
                self.take_snapshot(yesterday_str)
        except Exception as e:
            print("Error checking/taking daily snapshot:", e)

        self.after(60000, self.schedule_daily_snapshot)

    def on_closing(self):
        msg = "Are you sure you want to close the dashboard?\n\n(Automated Reports & Snapshots will continue running in the background)"
        if messagebox.askyesno("Confirm Exit", msg):
            self.withdraw_to_tray()
    def withdraw_to_tray(self):
        self.withdraw()
        try:
            import pystray
            from PIL import Image
            import threading

            def show_window(icon, item):
                icon.stop()
                self.after(0, self.deiconify)

            def quit_window(icon, item):
                icon.stop()
                self.after(0, self.destroy)

            menu = pystray.Menu(
                pystray.MenuItem('Restore Dashboard', show_window, default=True),
                pystray.MenuItem('Exit Completely', quit_window)
            )

            icon_path = resource_path(os.path.join("assets", "taskbar_logo.png"))
            if os.path.exists(icon_path):
                image = Image.open(icon_path)
            else:
                image = Image.new('RGB', (64, 64), color='red')

            icon = pystray.Icon("traceability", image, "Sub-Process Traceability", menu)
            threading.Thread(target=icon.run, daemon=True).start()

        except ImportError:
            self.destroy()

if __name__ == "__main__":
    init_db()
    app = TraceabilityApp()
    app.mainloop()
