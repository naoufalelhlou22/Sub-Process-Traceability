"""
Database access layer for the Sub-Process Traceability application.
Handles connections, authentication, schema initialization, and migrations.
"""
import sqlite3
import os
import json
import hashlib
import binascii
import logging
import logging.handlers
from contextlib import contextmanager

from config import DATA_DIR, DB_FILE, SF_DATA_DEFAULT

# ── Logging ───────────────────────────────────────────────────────────────────

LOG_DIR = os.path.join(DATA_DIR, "logs")
os.makedirs(LOG_DIR, exist_ok=True)

_log_formatter = logging.Formatter(
  "%(asctime)s [%(levelname)s] %(name)s: %(message)s",
  datefmt="%Y-%m-%d %H:%M:%S"
)

_file_handler = logging.handlers.RotatingFileHandler(
  os.path.join(LOG_DIR, "app.log"),
  maxBytes=5 * 1024 * 1024,  # 5 MB
  backupCount=3,
  encoding="utf-8"
)
_file_handler.setFormatter(_log_formatter)
_file_handler.setLevel(logging.DEBUG)

_console_handler = logging.StreamHandler()
_console_handler.setFormatter(_log_formatter)
_console_handler.setLevel(logging.INFO)

logger = logging.getLogger("TraceabilityApp")
logger.setLevel(logging.DEBUG)
logger.addHandler(_file_handler)
logger.addHandler(_console_handler)

# ── DB Connection ─────────────────────────────────────────────────────────────

def get_db_connection():
  conn = sqlite3.connect(DB_FILE, check_same_thread=False, timeout=10)
  conn.execute('PRAGMA journal_mode=WAL')
  return conn

@contextmanager
def db_connection():
  """Context manager for safe DB access. Usage: with db_connection() as (conn, c): ..."""
  conn = get_db_connection()
  try:
    c = conn.cursor()
    yield conn, c
  except Exception:
    conn.rollback()
    raise
  finally:
    conn.close()

# ── Password Hashing ─────────────────────────────────────────────────────────

def hash_password(password: str, salt: bytes = None) -> str:
  if salt is None:
    salt = os.urandom(32)
  key = hashlib.pbkdf2_hmac("sha256", password.encode(), salt, 310_000)
  return binascii.hexlify(salt).decode() + ":" + binascii.hexlify(key).decode()

def verify_password(password: str, stored: str) -> bool:
  try:
    if ":" not in stored:
      return False
    salt_hex, key_hex = stored.split(":")
    salt = binascii.unhexlify(salt_hex)
    
    # 1. Normal check (new schema)
    expected = hash_password(password, salt)
    if expected == stored:
      return True
      
    # 2. Migration fallback check (double-hashed from old static salt schema)
    old_salt = b"subproc_trace_salt_2026"
    old_hash_obj = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), old_salt, 100000)
    old_hash_str = binascii.hexlify(old_hash_obj).decode("utf-8")
    
    expected_old = hash_password(old_hash_str, salt)
    if expected_old == stored:
      return True
      
    return False
  except Exception:
    return False

def migrate_passwords():
  with db_connection() as (conn, c):
    c.execute("SELECT id, password FROM auth")
    rows = c.fetchall()
    for row in rows:
      uid, pwd = row
      if pwd and ":" not in pwd:
        hashed = hash_password(pwd)
        c.execute("UPDATE auth SET password = ? WHERE id = ?", (hashed, uid))
    conn.commit()

# ── Product Data ──────────────────────────────────────────────────────────────

SF_DATA = {}

def load_sf_data():
  global SF_DATA
  SF_DATA.clear()
  try:
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("SELECT pn_sf, name_sf, rms_json, std_box_qty, std_hourly_target FROM products")
    rows = c.fetchall()
    for r in rows:
      pn = r[0]
      name = r[1]
      try:
        rms = json.loads(r[2])
      except (json.JSONDecodeError, TypeError) as e:
        logger.debug("Could not parse rms_json for %s: %s", pn, e)
        rms = []
      std_box_qty = r[3]
      std_hourly_target = r[4] if len(r) > 4 else None
      SF_DATA[pn] = (name, rms, std_box_qty, std_hourly_target)
    conn.close()
    
    if not SF_DATA:
      SF_DATA.update(SF_DATA_DEFAULT)
      conn = get_db_connection()
      c = conn.cursor()
      for pn, val in SF_DATA.items():
        c.execute("INSERT INTO products (pn_sf, name_sf, rms_json, std_box_qty, std_hourly_target) VALUES (?, ?, ?, NULL, NULL)", (pn, val[0], json.dumps(val[1])))
        SF_DATA[pn] = (val[0], val[1], None, None)
      conn.commit()
      conn.close()
      
  except Exception as e:
    logger.error("Error loading SF_DATA from DB: %s", e)
    SF_DATA.update(SF_DATA_DEFAULT)

# ── Schema Initialization ────────────────────────────────────────────────────

def init_db():
  try:
    from quality_app import init_quality_db
    init_quality_db()
  except ImportError:
    pass
    
  conn = get_db_connection()
  c = conn.cursor()
  
  c.execute('''
    CREATE TABLE IF NOT EXISTS products (
      pn_sf TEXT PRIMARY KEY,
      name_sf TEXT,
      rms_json TEXT,
      std_box_qty INTEGER,
      std_hourly_target INTEGER
    )
  ''')
  
  try:
      c.execute("ALTER TABLE products ADD COLUMN std_box_qty INTEGER")
  except sqlite3.OperationalError:
      pass
      
  try:
      c.execute("ALTER TABLE products ADD COLUMN std_hourly_target INTEGER")
  except sqlite3.OperationalError:
      pass
  
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

  c.execute('''
    CREATE TABLE IF NOT EXISTS schema_version (
      version INTEGER PRIMARY KEY,
      applied_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  ''')
  c.execute("INSERT OR IGNORE INTO schema_version (version) VALUES (1)")
  
  try:
    default_users = [
      ('admin', os.environ.get('ADMIN_PASS', 'admin'), 'Manager')
    ]
    for uid, plain_pwd, role in default_users:
      c.execute("INSERT OR IGNORE INTO auth (id, password, role) VALUES (?, ?, ?)", (uid, hash_password(plain_pwd), role))
    
    conn.commit()
  except Exception as e:
    logger.error("Auth seed error: %s", e)
  
  migrate_passwords()
    
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
    
  c.execute('''CREATE TABLE IF NOT EXISTS downtime_logs (
           id INTEGER PRIMARY KEY AUTOINCREMENT,
           sub_batch_id TEXT,
           station TEXT,
           shift TEXT,
           op_id TEXT,
           duration_min REAL,
           reason TEXT,
           created_at TEXT
         )''')
         
  # Migration: add op_id to downtime_logs if it doesn't exist
  try:
      c.execute("ALTER TABLE downtime_logs ADD COLUMN op_id TEXT")
  except sqlite3.OperationalError:
      pass

  conn.commit()
  conn.close()
  
  load_sf_data()
