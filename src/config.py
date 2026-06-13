"""
Configuration constants for the Sub-Process Traceability application.
Centralizes all paths, theme colors, fonts, and default product data.
"""
import os
import sys

# ── Path Helpers ──────────────────────────────────────────────────────────────

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

# ── Data Directories ─────────────────────────────────────────────────────────

DATA_DIR = persistent_path("data")
os.makedirs(DATA_DIR, exist_ok=True)

# Migration logic for old data files
_base = os.path.abspath(os.path.join(DATA_DIR, ".."))
for _f in ["traceability_config.json", "traceability.db", "production_data.xlsx"]:
  _old = os.path.join(_base, _f)
  _new = os.path.join(DATA_DIR, _f)
  if os.path.exists(_old) and not os.path.exists(_new):
    try:
      import shutil
      shutil.move(_old, _new)
    except Exception:
      pass

# ── File Paths ────────────────────────────────────────────────────────────────

CONFIG_FILE = os.path.join(DATA_DIR, "traceability_config.json")
DB_FILE = os.path.join(DATA_DIR, "traceability.db")
EXCEL_FILE = os.path.join(DATA_DIR, "production_data.xlsx")


APP_VERSION = "1.1.0"

# ── Premium Industrial HMI Theme ─────────────────────────────────────────────

BG_COLOR = "#121212"    # Main Background
SURFACE_COLOR = "#1E1E1E" # Panels & Cards

ACCENT_COLOR = "#00B4D8"  # Primary Action
SUCCESS_COLOR = "#2DC653" # Running / Completed
WARNING_COLOR = "#F59E0B" # Warning
ERROR_COLOR = "#DC2626"  # Error / Stop

BORDER_COLOR = "#334155"  # Borders

TEXT_COLOR = "#F8FAFC"   # Main Text
TEXT_MUTED = "#94A3B8"   # Secondary Text

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

# ── Default Hardcoded Product Data (Fallback) ────────────────────────────────

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
