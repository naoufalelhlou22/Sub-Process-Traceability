"""
Printing services for the Sub-Process Traceability application.
Handles ZPL label generation and Zebra printer communication.
"""
import os
import json
import datetime
from tkinter import messagebox

from config import CONFIG_FILE, persistent_path
from database import logger

try:
  import win32print
  WIN32_PRINT_AVAILABLE = True
except ImportError:
  WIN32_PRINT_AVAILABLE = False

# ── ZPL Label Generation ─────────────────────────────────────────────────────

def generate_zpl(record_data):
  def format_val(val, default="-"):
    return str(val) if val else default

  zpl = []
  zpl.append("^XA")
  zpl.append("^PW520")
  zpl.append("^LH0,0")
  zpl.append("^CI28")
  
  def draw_layout(offset_y):
    zpl.append(f"^FO30,{20+offset_y}^A0N,24,24^FB460,1,0,C^FDSUB-PROCESS RECORD\\&^FS")
    zpl.append(f"^FO30,{55+offset_y}^GB460,3,3^FS")
    
    zpl.append(f"^FO30,{70+offset_y}^GB460,40,40^FS")
    zpl.append(f"^FO30,{80+offset_y}^A0N,24,24^FR^FB460,1,0,C^FD{format_val(record_data.get('sub_batch_id'))}\\&^FS")
    
    sb_id_val = format_val(record_data.get('sub_batch_id'))
    zpl.append(f"^BY2^FO50,{120+offset_y}^BCN,50,N,N,N,A^FD{sb_id_val}^FS")
    
    y = 185 + offset_y
    def add_row(label, value, font_size=18, y_inc=20):
      nonlocal y
      zpl.append(f"^FO30,{y}^A0N,{font_size},{font_size}^FD{label}^FS")
      zpl.append(f"^FO30,{y}^A0N,{font_size},{font_size}^FB460,1,0,R^FD{value}\\&^FS")
      y += y_inc

    add_row("PN Semi fini", format_val(record_data.get('pn_sf')))
    add_row("Part Name (SF)", format_val(record_data.get('part_sf')))
        
    y += 3
    zpl.append(f"^FO30,{y}^GB460,1,1^FS")
    y += 8
    
    add_row("Batch No. 1", format_val(record_data.get('batch1')))
    add_row("Batch No. 2", format_val(record_data.get('batch2')))
    add_row("Batch No. 3", format_val(record_data.get('batch3')))
    add_row("Quantity", f"{format_val(record_data.get('quantity'))} pcs", font_size=20, y_inc=25)
    
    zpl.append(f"^FO30,{y}^GB460,1,1^FS")
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
    zpl.append(f"^FO30,{y}^GB460,2,2^FS")
    y += 8
    
    reprints = int(record_data.get('reprint_count') or 0)
    if reprints > 0:
      last_by = record_data.get('last_reprinted_by', 'Unknown')
      last_at = record_data.get('last_reprinted_at', '')[:16] if record_data.get('last_reprinted_at') else ''
      audit_text = f"** REPRINT #{reprints} by {last_by} at {last_at} **"
      zpl.append(f"^FO30,{y}^A0N,14,14^FB460,1,0,C^FD{audit_text}\\&^FS")
    else:
      zpl.append(f"^FO30,{y}^A0N,14,14^FB460,1,0,C^FD-- Original Print --\\&^FS")
    y += 18
    
  draw_layout(0)
  
  zpl.append(f"^FO30,728^A0N,20,20^FB460,1,0,C^FD- - - - - - - - - - - -\\&^FS")
  
  draw_layout(982)

  zpl.append("^PQ1")
  zpl.append("^XZ")
  return "\n".join(zpl).encode("utf-8")

# ── Print Execution ───────────────────────────────────────────────────────────

def execute_zpl_print(zpl_string, record_data, silent=False):
  config = {}
  if os.path.exists(CONFIG_FILE):
    try:
      with open(CONFIG_FILE, "r") as f:
        config = json.load(f)
    except Exception as e:
      logger.warning("Config load error: %s", e)
      
  printer_name = config.get("zebra_printer", "")
  
  try:
    qr_dir = persistent_path("qr")
    os.makedirs(qr_dir, exist_ok=True)
    sb_id_safe = str(record_data.get('sub_batch_id', 'unknown')).replace(":", "").replace("-", "")
    file_path = os.path.join(qr_dir, f"{sb_id_safe}.zpl")
    with open(file_path, "wb") as f:
      f.write(zpl_string)
  except Exception as e:
    logger.warning("Could not save ZPL file: %s", e)
    
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

# ── Pick Ticket ZPL ───────────────────────────────────────────────────────────

def generate_pick_ticket_zpl(pn, target_qty, accumulated, picked):
  zpl = []
  zpl.append("^XA")
  zpl.append("^PW480")
  zpl.append("^LH20,0")
  zpl.append("^CI28")
  
  zpl.append("^FO0,30^A0N,30,30^FB440,1,0,C^FDFIFO PICK TICKET\\&^FS")
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
    except Exception as e:
      logger.warning("Config load error: %s", e)

  printer_name = config.get("zebra_printer", "")

  try:
    tickets_dir = persistent_path("tickets")
    os.makedirs(tickets_dir, exist_ok=True)
    safe_pn = str(pn).replace(":", "").replace("-", "")
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    file_path = os.path.join(tickets_dir, f"TICKET_{safe_pn}_{timestamp}.zpl")
    with open(file_path, "wb") as f:
      f.write(zpl_string)
  except Exception as e:
    logger.warning("Could not save Pick Ticket ZPL file: %s", e)

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

# ── Startup Registration ─────────────────────────────────────────────────────

def add_to_startup():
  try:
    import winreg
    import sys
    
    key = winreg.HKEY_CURRENT_USER
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    
    with winreg.OpenKey(key, key_path, 0, winreg.KEY_ALL_ACCESS) as reg_key:
      app_name = "SubProcessTraceability"
      
      if getattr(sys, 'frozen', False):
        exe_path = f'"{sys.executable}"'
      else:
        exe_path = f'"{sys.executable}" "{os.path.abspath(sys.argv[0])}"'
        
      winreg.SetValueEx(reg_key, app_name, 0, winreg.REG_SZ, exe_path)
  except Exception as e:
    logger.warning("Could not add to startup: %s", e)
