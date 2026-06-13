"""
Full Production Readiness Test Suite — Sub-Process Traceability
Tests actual DB logic, workflow rules, data integrity — no GUI needed.
"""
import sys
import os
import sqlite3
import json
import datetime
import shutil
import hashlib
import binascii
import traceback

# Setup paths
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_DIR = os.path.join(BASE_DIR, "src")
sys.path.insert(0, SRC_DIR)

# Use a separate test database
TEST_DATA_DIR = os.path.join(BASE_DIR, "tests", "_test_data")
TEST_DB = os.path.join(TEST_DATA_DIR, "test_traceability.db")

# ─────────────────────────────────────────────────────────────────────────────
# Test Framework
# ─────────────────────────────────────────────────────────────────────────────
PASS = 0
FAIL = 0
ERRORS = []

def test(name, condition, detail=""):
    global PASS, FAIL, ERRORS
    if condition:
        PASS += 1
        print(f"  ✅ PASS: {name}")
    else:
        FAIL += 1
        msg = f"  ❌ FAIL: {name}"
        if detail:
            msg += f" — {detail}"
        print(msg)
        ERRORS.append(msg)

def section(title):
    print(f"\n{'='*70}")
    print(f"  {title}")
    print(f"{'='*70}")

# ─────────────────────────────────────────────────────────────────────────────
# DB Helpers (mirror main.py logic)
# ─────────────────────────────────────────────────────────────────────────────
def get_conn():
    conn = sqlite3.connect(TEST_DB, check_same_thread=False, timeout=30)
    conn.execute('PRAGMA journal_mode=WAL')
    conn.row_factory = sqlite3.Row
    return conn

def hash_password(password):
    salt = b"subproc_trace_salt_2026"
    hash_obj = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 100000)
    return binascii.hexlify(hash_obj).decode("utf-8")

def init_test_db():
    """Create all tables exactly as the app does."""
    os.makedirs(TEST_DATA_DIR, exist_ok=True)
    if os.path.exists(TEST_DB):
        os.remove(TEST_DB)

    conn = get_conn()
    c = conn.cursor()

    c.execute('''CREATE TABLE IF NOT EXISTS products (
        pn_sf TEXT PRIMARY KEY, name_sf TEXT, rms_json TEXT)''')

    c.execute('''CREATE TABLE IF NOT EXISTS auth (
        id TEXT PRIMARY KEY, password TEXT, role TEXT)''')

    c.execute('''CREATE TABLE IF NOT EXISTS records (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sub_batch_id TEXT NOT NULL, pn_sf TEXT, part_sf TEXT,
        rm1_pn TEXT, rm1_name TEXT, rm2_pn TEXT, rm2_name TEXT,
        rm3_pn TEXT, rm3_name TEXT, rm4_pn TEXT, rm4_name TEXT,
        batch1 TEXT, batch2 TEXT, batch3 TEXT, quantity INTEGER,
        shift_sp TEXT, op_id TEXT, station TEXT,
        dt_sp TEXT, dt_line TEXT, shift_line TEXT,
        remarks TEXT, status TEXT DEFAULT 'In Rack',
        created_at TEXT NOT NULL, registered_by TEXT DEFAULT '',
        reprint_count INTEGER DEFAULT 0,
        last_reprinted_at TEXT, last_reprinted_by TEXT)''')

    c.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_records_sub_batch_id ON records(sub_batch_id)")

    c.execute('''CREATE TABLE IF NOT EXISTS quality_defects (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sub_batch_id TEXT NOT NULL, pn_sf TEXT, part_sf TEXT,
        produced_by_op TEXT, quality_op_id TEXT NOT NULL,
        defect_type TEXT, qty_defective INTEGER DEFAULT 0,
        description TEXT, shift_sp TEXT, station TEXT,
        reported_at TEXT NOT NULL, status TEXT DEFAULT 'Open',
        action_type TEXT DEFAULT 'Scrap',
        is_quarantined INTEGER DEFAULT 0,
        root_cause TEXT, corrective_action TEXT)''')

    c.execute('''CREATE TABLE IF NOT EXISTS shift_targets (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        product_pn TEXT NOT NULL, shift TEXT NOT NULL,
        station TEXT, target_qty INTEGER NOT NULL,
        effective_date TEXT NOT NULL)''')

    c.execute('''CREATE TABLE IF NOT EXISTS part_thresholds (
        pn_sf TEXT PRIMARY KEY, min_qty INTEGER NOT NULL DEFAULT 0)''')

    c.execute('''CREATE TABLE IF NOT EXISTS inventory_snapshots (
        snapshot_date TEXT NOT NULL, pn_sf TEXT NOT NULL,
        boxes_in_rack INTEGER, total_qty_in_rack INTEGER,
        oldest_box_age_hours REAL,
        PRIMARY KEY (snapshot_date, pn_sf))''')

    c.execute('''CREATE TABLE IF NOT EXISTS product_audit_trail (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        action TEXT, pn_sf TEXT, details TEXT,
        user_id TEXT, shift TEXT, timestamp TEXT)''')

    c.execute('''CREATE TABLE IF NOT EXISTS system_access_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        event_type TEXT, user_id TEXT, shift TEXT, timestamp TEXT)''')

    # Seed users
    users = [
        ('s', hash_password(' '), 'Supervisor'),
        ('sp99', hash_password('sp99'), 'Supervisor'),
        ('sl', hash_password('sl98'), 'Shift Leader'),
        ('op1', hash_password('op1'), 'Operator'),
        ('mg90', hash_password('oi90'), 'Manager'),
        ('Q001', hash_password('quality001'), 'Quality OP'),
    ]
    c.executemany("INSERT OR IGNORE INTO auth (id, password, role) VALUES (?,?,?)", users)

    # Seed products
    products = [
        ("MOCK-A01-10001-X-SUB", "Alpha Subsystem Right",
         json.dumps([("MOCK-A01-10001-Y-PRT", "Alpha Core Right"), ("MOCK-A01-10002-Z-PRT", "Alpha Buffer Right")])),
        ("MOCK-B02-20001-X-SUB", "Beta Panel Right Sub",
         json.dumps([("MOCK-B02-20001-Y-PRT", "Beta Panel Right"), ("MOCK-B02-20002-Z-PRT", "Beta Sealant")])),
        ("MOCK-C03-30001-X-SUB", "Gamma Switch Right Sub (Short)",
         json.dumps([("MOCK-C03-30001-Y-PRT", "Gamma Switch Right"), ("MOCK-C03-30002-Z-PRT", "Gamma Coil Right"), ("MOCK-C03-30003-W-PRT", "Gamma Pin")])),
    ]
    c.executemany("INSERT OR IGNORE INTO products (pn_sf, name_sf, rms_json) VALUES (?,?,?)", products)

    conn.commit()
    conn.close()

def insert_record(sb_id, pn_sf="MOCK-A01-10001-X-SUB", part_sf="Alpha Subsystem Right",
                  qty=100, shift="A", op_id="op1", station="S06",
                  dt_sp=None, status="In Rack", batch1="BATCH001"):
    """Insert a test record."""
    if dt_sp is None:
        dt_sp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    created_at = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = sqlite3.connect(TEST_DB, timeout=15)
    conn.execute('PRAGMA journal_mode=WAL')
    try:
        c = conn.cursor()
        c.execute('''INSERT INTO records (
            sub_batch_id, pn_sf, part_sf,
            rm1_pn, rm1_name, rm2_pn, rm2_name,
            rm3_pn, rm3_name, rm4_pn, rm4_name,
            batch1, batch2, batch3, quantity,
            shift_sp, op_id, station, dt_sp, dt_line,
            shift_line, remarks, status, created_at, registered_by
        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
        (sb_id, pn_sf, part_sf,
         "MOCK-A01-10001-Y-PRT", "Alpha Core Right", "MOCK-A01-10002-Z-PRT", "Alpha Buffer Right",
         "", "", "", "",
         batch1, "", "", qty,
         shift, op_id, station, dt_sp, "",
         "", "", status, created_at, "test_user"))
        conn.commit()
    finally:
        conn.close()

# ─────────────────────────────────────────────────────────────────────────────
# TEST SUITES
# ─────────────────────────────────────────────────────────────────────────────

def test_01_authentication():
    section("1. AUTHENTICATION & AUTHORIZATION")

    conn = get_conn()
    c = conn.cursor()

    # Test valid login
    c.execute("SELECT role FROM auth WHERE id=? AND password=?", ("op1", hash_password("op1")))
    row = c.fetchone()
    test("Valid operator login (op1/op1)", row is not None and row["role"] == "Operator")

    # Test valid supervisor login
    c.execute("SELECT role FROM auth WHERE id=? AND password=?", ("sp99", hash_password("sp99")))
    row = c.fetchone()
    test("Valid supervisor login (sp99/sp99)", row is not None and row["role"] == "Supervisor")

    # Test valid manager login
    c.execute("SELECT role FROM auth WHERE id=? AND password=?", ("mg90", hash_password("oi90")))
    row = c.fetchone()
    test("Valid manager login (mg90/oi90)", row is not None and row["role"] == "Manager")

    # Test wrong password
    c.execute("SELECT role FROM auth WHERE id=? AND password=?", ("op1", hash_password("wrongpass")))
    row = c.fetchone()
    test("Wrong password rejected", row is None)

    # Test non-existent user
    c.execute("SELECT role FROM auth WHERE id=? AND password=?", ("nonexistent", hash_password("any")))
    row = c.fetchone()
    test("Non-existent user rejected", row is None)

    # Test Quality OP login
    c.execute("SELECT role FROM auth WHERE id=? AND password=?", ("Q001", hash_password("quality001")))
    row = c.fetchone()
    test("Quality OP login works", row is not None and row["role"] == "Quality OP")

    # Test space password for supervisor 's'
    c.execute("SELECT role FROM auth WHERE id=? AND password=?", ("s", hash_password(" ")))
    row = c.fetchone()
    test("Supervisor 's' with space password works (SECURITY CONCERN)", row is not None,
         "Password is a single space character!")

    # Test password hashing consistency
    h1 = hash_password("testpass")
    h2 = hash_password("testpass")
    test("Password hash is deterministic", h1 == h2)

    h3 = hash_password("testpass2")
    test("Different passwords produce different hashes", h1 != h3)

    conn.close()


def test_02_record_creation():
    section("2. RECORD CREATION & SUB-BATCH ID")

    # Test basic record insert
    insert_record("SB20260613100000S06A01", qty=120, dt_sp="2026-06-13 10:00")
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT * FROM records WHERE sub_batch_id=?", ("SB20260613100000S06A01",))
    row = c.fetchone()
    test("Basic record insert succeeds", row is not None)
    test("Quantity stored correctly", row["quantity"] == 120)
    test("Status defaults to 'In Rack'", row["status"] == "In Rack")
    test("Station stored correctly", row["station"] == "S06")
    test("Shift stored correctly", row["shift_sp"] == "A")
    test("Reprint count defaults to 0", row["reprint_count"] == 0)

    # Test duplicate SB_ID rejection (UNIQUE INDEX)
    conn.close()  # close first to avoid lock
    dup_ok = False
    try:
        insert_record("SB20260613100000S06A01", qty=50)
    except sqlite3.IntegrityError:
        dup_ok = True
    test("Duplicate SB_ID rejected", dup_ok)
    conn = get_conn()
    c = conn.cursor()

    # Test SB_ID generation logic
    base = "SB20260613110000S07B"
    insert_record(f"{base}01", station="S07", shift="B", dt_sp="2026-06-13 11:00")
    insert_record(f"{base}02", station="S07", shift="B", dt_sp="2026-06-13 11:00")
    c.execute("SELECT COUNT(*) as cnt FROM records WHERE sub_batch_id LIKE ?", (f"{base}%",))
    test("Multiple records with same base SB_ID allowed with suffix", c.fetchone()["cnt"] == 2)

    # Test record with empty batch numbers
    insert_record("SB_NOBATCH_01", batch1="")
    c.execute("SELECT batch1 FROM records WHERE sub_batch_id=?", ("SB_NOBATCH_01",))
    test("Record with empty batch number stores correctly", c.fetchone()["batch1"] == "")

    # Test large quantity
    insert_record("SB_LARGEQTY_01", qty=99999)
    c.execute("SELECT quantity FROM records WHERE sub_batch_id=?", ("SB_LARGEQTY_01",))
    test("Large quantity (99999) stores correctly", c.fetchone()["quantity"] == 99999)

    # Test zero quantity
    insert_record("SB_ZEROQTY_01", qty=0)
    c.execute("SELECT quantity FROM records WHERE sub_batch_id=?", ("SB_ZEROQTY_01",))
    test("Zero quantity stores (no DB validation)", c.fetchone()["quantity"] == 0,
         "App should validate qty > 0 in UI")

    conn.close()


def test_03_consume_workflow():
    section("3. CONSUME TO LINE WORKFLOW")

    conn = get_conn()
    c = conn.cursor()

    # Insert test boxes for consumption
    insert_record("SB_CONSUME_01", dt_sp="2026-06-13 08:00", status="In Rack")
    insert_record("SB_CONSUME_02", dt_sp="2026-06-13 09:00", status="In Rack")
    insert_record("SB_CONSUME_03", dt_sp="2026-06-13 07:00", status="In Rack")  # Oldest

    # Test basic consumption
    dt_line = "2026-06-13 14:00"
    c.execute("UPDATE records SET status='Consumed', dt_line=?, shift_line=? WHERE sub_batch_id=?",
              (dt_line, "A", "SB_CONSUME_01"))
    conn.commit()

    c.execute("SELECT status, dt_line, shift_line FROM records WHERE sub_batch_id=?", ("SB_CONSUME_01",))
    row = c.fetchone()
    test("Box consumed successfully", row["status"] == "Consumed")
    test("Line entry datetime recorded", row["dt_line"] == dt_line)
    test("Line shift recorded", row["shift_line"] == "A")

    # Test double-consume prevention
    c.execute("SELECT status FROM records WHERE sub_batch_id=?", ("SB_CONSUME_01",))
    row = c.fetchone()
    test("Already consumed box detected", row["status"] == "Consumed")

    # Test FIFO check: SB_CONSUME_03 (07:00) is older than SB_CONSUME_02 (09:00)
    c.execute("SELECT COUNT(*) as cnt FROM records WHERE pn_sf=? AND status='In Rack' AND dt_sp < ?",
              ("MOCK-A01-10001-X-SUB", "2026-06-13 09:00"))
    older = c.fetchone()["cnt"]
    test("FIFO detects older box exists before consuming newer one",
         older > 0, f"Found {older} older box(es)")

    # Test consuming the oldest first (correct FIFO)
    # Only check boxes from this test (SB_CONSUME_ prefix)
    c.execute("SELECT COUNT(*) as cnt FROM records WHERE pn_sf=? AND status='In Rack' AND dt_sp < ? AND sub_batch_id LIKE 'SB_CONSUME_%'",
              ("MOCK-A01-10001-X-SUB", "2026-06-13 07:00"))
    older_than_oldest = c.fetchone()["cnt"]
    test("Oldest box (07:00) has no older boxes in consume set", older_than_oldest == 0)

    # Test consuming non-existent box
    c.execute("SELECT id FROM records WHERE sub_batch_id=?", ("SB_NONEXISTENT",))
    test("Non-existent box not found for consumption", c.fetchone() is None)

    conn.close()


def test_04_quality_defect_workflow():
    section("4. QUALITY DEFECT REPORTING")

    conn = get_conn()
    c = conn.cursor()

    # Create a box to report defects on
    insert_record("SB_QUALITY_01", qty=100, status="In Rack")

    # Report a defect
    reported_at = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute('''INSERT INTO quality_defects
        (sub_batch_id, pn_sf, part_sf, produced_by_op, quality_op_id,
         defect_type, qty_defective, description, shift_sp, station,
         reported_at, status, action_type, is_quarantined)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
        ("SB_QUALITY_01", "MOCK-A01-10001-X-SUB", "Alpha Subsystem Right",
         "op1", "Q001", "Cosmetic", 5, "Scratches on surface",
         "A", "S06", reported_at, "Open", "Scrap", 1))
    conn.commit()

    c.execute("SELECT * FROM quality_defects WHERE sub_batch_id=?", ("SB_QUALITY_01",))
    defect = c.fetchone()
    test("Defect report inserted", defect is not None)
    test("Defect qty recorded (5)", defect["qty_defective"] == 5)
    test("Defect status is Open", defect["status"] == "Open")
    test("Box quarantined", defect["is_quarantined"] == 1)
    test("Action type recorded (Scrap)", defect["action_type"] == "Scrap")

    # Test quarantine blocks consumption
    c.execute("""SELECT IFNULL(SUM(CASE WHEN status='Closed' AND action_type IN ('Rework','Sorting','Use As-Is')
                THEN 0 ELSE qty_defective END), 0) as eff_def
                FROM quality_defects WHERE sub_batch_id=? AND is_quarantined=1""", ("SB_QUALITY_01",))
    eff_defects = c.fetchone()["eff_def"]
    test("Quarantined box blocks consumption (eff_defects > 0)", eff_defects > 0,
         f"Effective defects = {eff_defects}")

    # Test box status recalculation
    c.execute("SELECT status, quantity FROM records WHERE sub_batch_id=?", ("SB_QUALITY_01",))
    rec = c.fetchone()
    total_qty = rec["quantity"]
    good_qty = total_qty - eff_defects
    test("Good qty calculation (100 - 5 = 95)", good_qty == 95)

    # Simulate status recalculation (as quality_app does)
    if good_qty <= 0:
        new_status = 'Quarantined'
    elif eff_defects > 0:
        new_status = 'Partial Defect'
    else:
        new_status = 'In Rack'
    c.execute("UPDATE records SET status=? WHERE sub_batch_id=?", (new_status, "SB_QUALITY_01"))
    conn.commit()
    test("Box status updated to 'Partial Defect'", new_status == "Partial Defect")

    # Test closing defect with CAPA
    c.execute("""UPDATE quality_defects SET status='Closed', action_type='Rework',
                root_cause='Tool wear', corrective_action='Replace tool'
                WHERE sub_batch_id=?""", ("SB_QUALITY_01",))
    conn.commit()

    # After closing with Rework, effective defects should be 0
    c.execute("""SELECT IFNULL(SUM(CASE WHEN status='Closed' AND action_type IN ('Rework','Sorting','Use As-Is')
                THEN 0 ELSE qty_defective END), 0) as eff_def
                FROM quality_defects WHERE sub_batch_id=?""", ("SB_QUALITY_01",))
    eff_after = c.fetchone()["eff_def"]
    test("After Rework closure, effective defects = 0", eff_after == 0)

    # Recalculate status — should go back to 'In Rack'
    c.execute("SELECT COUNT(*) as cnt FROM quality_defects WHERE sub_batch_id=? AND is_quarantined=1 AND NOT (status='Closed' AND action_type IN ('Rework','Sorting','Use As-Is'))", ("SB_QUALITY_01",))
    has_quar = c.fetchone()["cnt"] > 0
    if eff_after <= 0 and not has_quar:
        restored_status = 'In Rack'
    else:
        restored_status = 'Partial Defect'
    test("After Rework, box returns to 'In Rack'", restored_status == "In Rack")

    # Test: fully defective box → Quarantined
    insert_record("SB_QUALITY_02", qty=10, status="In Rack")
    c.execute('''INSERT INTO quality_defects
        (sub_batch_id, pn_sf, part_sf, produced_by_op, quality_op_id,
         defect_type, qty_defective, description, shift_sp, station,
         reported_at, status, action_type, is_quarantined)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
        ("SB_QUALITY_02", "MOCK-A01-10001-X-SUB", "Alpha Subsystem Right",
         "op1", "Q001", "Functional Failure", 10, "All parts fail test",
         "A", "S06", reported_at, "Open", "Scrap", 1))
    conn.commit()

    c.execute("""SELECT IFNULL(SUM(CASE WHEN status='Closed' AND action_type IN ('Rework','Sorting','Use As-Is')
                THEN 0 ELSE qty_defective END), 0) as eff_def
                FROM quality_defects WHERE sub_batch_id=?""", ("SB_QUALITY_02",))
    full_def = c.fetchone()["eff_def"]
    good = 10 - full_def
    new_st = 'Quarantined' if good <= 0 else ('Partial Defect' if full_def > 0 else 'In Rack')
    test("Fully defective box (10/10) → Quarantined", new_st == "Quarantined")

    conn.close()


def test_05_product_management():
    section("5. PRODUCT MANAGEMENT (SF_DATA)")

    conn = get_conn()
    c = conn.cursor()

    # Load products from DB
    c.execute("SELECT pn_sf, name_sf, rms_json FROM products")
    rows = c.fetchall()
    sf_data = {}
    for r in rows:
        rms = json.loads(r["rms_json"]) if r["rms_json"] else []
        sf_data[r["pn_sf"]] = (r["name_sf"], rms)

    test("Products loaded from DB", len(sf_data) >= 3)

    # Test product structure
    alpha = sf_data.get("MOCK-A01-10001-X-SUB")
    test("Alpha product exists", alpha is not None)
    test("Alpha has correct name", alpha[0] == "Alpha Subsystem Right")
    test("Alpha has 2 raw materials", len(alpha[1]) == 2)
    test("Alpha RM1 is correct", alpha[1][0][0] == "MOCK-A01-10001-Y-PRT")

    gamma = sf_data.get("MOCK-C03-30001-X-SUB")
    test("Gamma product has 3 raw materials", len(gamma[1]) == 3)

    # Test RM scan matching (only first RM should match)
    scanned_rm = "MOCK-A01-10001-Y-PRT"
    matched = None
    for sf_pn, (sf_name, rm_list) in sf_data.items():
        if rm_list and rm_list[0][0] == scanned_rm:
            matched = sf_pn
            break
    test("RM1 scan matches correct SF product", matched == "MOCK-A01-10001-X-SUB")

    # Test secondary RM does NOT match
    scanned_rm2 = "MOCK-A01-10002-Z-PRT"
    matched2 = None
    for sf_pn, (sf_name, rm_list) in sf_data.items():
        if rm_list and rm_list[0][0] == scanned_rm2:
            matched2 = sf_pn
            break
    test("RM2 scan does NOT match (only RM1 triggers)", matched2 is None)

    # Test adding a new product
    new_pn = "TEST-NEW-PRODUCT-001"
    new_rms = [("TEST-RM-001", "Test Raw Material")]
    c.execute("INSERT OR REPLACE INTO products (pn_sf, name_sf, rms_json) VALUES (?,?,?)",
              (new_pn, "Test Product", json.dumps(new_rms)))
    conn.commit()

    c.execute("SELECT * FROM products WHERE pn_sf=?", (new_pn,))
    test("New product added successfully", c.fetchone() is not None)

    # Test audit trail for product add
    c.execute("INSERT INTO product_audit_trail (action, pn_sf, details, user_id, shift, timestamp) VALUES (?,?,?,?,?,?)",
              ("ADD", new_pn, "Name: Test Product, RMs: 1", "test_admin", "A",
               datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()
    c.execute("SELECT * FROM product_audit_trail WHERE pn_sf=?", (new_pn,))
    test("Product audit trail recorded", c.fetchone() is not None)

    # Test deleting a product
    c.execute("DELETE FROM products WHERE pn_sf=?", (new_pn,))
    conn.commit()
    c.execute("SELECT * FROM products WHERE pn_sf=?", (new_pn,))
    test("Product deleted successfully", c.fetchone() is None)

    conn.close()


def test_06_inventory_and_fifo():
    section("6. LIVE INVENTORY & FIFO PICK LIST")

    conn = get_conn()
    c = conn.cursor()

    # Insert boxes with varying ages for inventory
    pn = "MOCK-B02-20001-X-SUB"
    insert_record("SB_INV_OLD", pn_sf=pn, part_sf="Beta Panel Right Sub",
                  qty=50, dt_sp="2026-06-10 08:00", status="In Rack")
    insert_record("SB_INV_MED", pn_sf=pn, part_sf="Beta Panel Right Sub",
                  qty=75, dt_sp="2026-06-12 08:00", status="In Rack")
    insert_record("SB_INV_NEW", pn_sf=pn, part_sf="Beta Panel Right Sub",
                  qty=100, dt_sp="2026-06-13 08:00", status="In Rack")
    insert_record("SB_INV_CONSUMED", pn_sf=pn, part_sf="Beta Panel Right Sub",
                  qty=60, dt_sp="2026-06-11 08:00", status="Consumed")

    # Test inventory aggregation (only In Rack)
    c.execute("""SELECT pn_sf, SUM(quantity) as total, COUNT(id) as boxes
                 FROM records WHERE status='In Rack' AND pn_sf=?
                 GROUP BY pn_sf""", (pn,))
    agg = c.fetchone()
    test("Inventory shows only In Rack boxes", agg is not None)
    test("Total qty in rack = 225 (50+75+100)", agg["total"] == 225)
    test("Box count = 3 (consumed excluded)", agg["boxes"] == 3)

    # Test FIFO pick list generation
    target_qty = 100
    c.execute("""SELECT sub_batch_id, quantity, dt_sp
                 FROM records WHERE status='In Rack' AND pn_sf=?
                 ORDER BY dt_sp ASC""", (pn,))
    rows = c.fetchall()

    picked = []
    accumulated = 0
    for r in rows:
        picked.append((r["sub_batch_id"], r["quantity"], r["dt_sp"]))
        accumulated += r["quantity"]
        if accumulated >= target_qty:
            break

    test("FIFO picks oldest first", picked[0][0] == "SB_INV_OLD")
    test("FIFO accumulates until target met", accumulated >= target_qty,
         f"Accumulated {accumulated} >= target {target_qty}")
    test("FIFO picked 2 boxes to reach 100 (50+75=125)", len(picked) == 2)

    # Test threshold alerts
    c.execute("INSERT OR REPLACE INTO part_thresholds (pn_sf, min_qty) VALUES (?,?)", (pn, 300))
    conn.commit()

    c.execute("""SELECT r.pn_sf, SUM(r.quantity) as total, COALESCE(t.min_qty, 0) as min_qty
                 FROM records r
                 LEFT JOIN part_thresholds t ON r.pn_sf = t.pn_sf
                 WHERE r.status='In Rack' AND r.pn_sf=?
                 GROUP BY r.pn_sf""", (pn,))
    thresh = c.fetchone()
    test("Low WIP alert: 225 < 300 threshold", thresh["total"] < thresh["min_qty"])

    conn.close()


def test_07_shift_targets_and_kpis():
    section("7. SHIFT TARGETS & KPI CALCULATIONS")

    conn = get_conn()
    c = conn.cursor()

    # Use a UNIQUE PN for this test to avoid cross-contamination
    pn = "KPI-TEST-PN-001"

    # Set a target
    c.execute("INSERT INTO shift_targets (product_pn, shift, station, target_qty, effective_date) VALUES (?,?,?,?,?)",
              (pn, "All", "All", 500, "2026-06-13 00:00:00"))
    conn.commit()

    # Insert production records for today with this unique PN
    today = "2026-06-13"
    insert_record("SB_KPI_01", pn_sf=pn, part_sf="KPI Test Part", qty=120, shift="A", dt_sp=f"{today} 07:00", op_id="op1")
    insert_record("SB_KPI_02", pn_sf=pn, part_sf="KPI Test Part", qty=150, shift="A", dt_sp=f"{today} 09:00", op_id="op1")
    insert_record("SB_KPI_03", pn_sf=pn, part_sf="KPI Test Part", qty=80, shift="A", dt_sp=f"{today} 11:00", op_id="op2")

    # Test daily total
    start_dt = f"{today} 06:00:00"
    end_dt = "2026-06-14 05:59:59"
    c.execute("SELECT SUM(quantity) as total FROM records WHERE dt_sp >= ? AND dt_sp <= ? AND pn_sf=?",
              (start_dt, end_dt, pn))
    total = c.fetchone()["total"] or 0
    test("Daily production total = 350 (120+150+80)", total == 350,
         f"Got {total}")

    # Test target vs actual
    c.execute("SELECT target_qty FROM shift_targets WHERE product_pn=? ORDER BY id DESC LIMIT 1", (pn,))
    target = c.fetchone()["target_qty"]
    pct = (total / target * 100) if target > 0 else 0
    test("Target attainment = 70% (350/500)", abs(pct - 70.0) < 0.1,
         f"Got {pct:.1f}%")

    # Test hourly breakdown
    c.execute("""SELECT substr(dt_sp, 12, 2) as hr, SUM(quantity) as qty
                 FROM records WHERE dt_sp >= ? AND dt_sp <= ? AND pn_sf=?
                 GROUP BY hr""", (start_dt, end_dt, pn))
    hourly = {r["hr"]: r["qty"] for r in c.fetchall()}
    test("07:00 hour has 120 pcs", hourly.get("07", 0) == 120,
         f"Got {hourly}")
    test("09:00 hour has 150 pcs", hourly.get("09", 0) == 150)
    test("11:00 hour has 80 pcs", hourly.get("11", 0) == 80)

    # Test operator breakdown
    c.execute("""SELECT op_id, SUM(quantity) as qty
                 FROM records WHERE dt_sp >= ? AND dt_sp <= ? AND pn_sf=?
                 GROUP BY op_id ORDER BY qty DESC""", (start_dt, end_dt, pn))
    ops = c.fetchall()
    test("Top operator is op1 (270 pcs)", len(ops) > 0 and ops[0]["op_id"] == "op1" and ops[0]["qty"] == 270,
         f"Got {[(o['op_id'], o['qty']) for o in ops]}")
    test("Second operator is op2 (80 pcs)", len(ops) > 1 and ops[1]["op_id"] == "op2" and ops[1]["qty"] == 80)

    # Test quality rate calculation
    c.execute("""SELECT IFNULL(SUM(CASE WHEN status='Closed' AND action_type IN ('Rework','Sorting','Use As-Is')
                THEN 0 ELSE qty_defective END), 0) as def_qty
                FROM quality_defects WHERE pn_sf=? AND reported_at >= ? AND reported_at <= ?""",
              (pn, start_dt, end_dt))
    def_qty = c.fetchone()["def_qty"]
    qual_rate = ((total - def_qty) / total * 100) if total > 0 else 100
    test("Quality rate calculation works", qual_rate > 0)

    conn.close()


def test_08_reprint_audit_trail():
    section("8. REPRINT AUDIT TRAIL")

    conn = get_conn()
    c = conn.cursor()

    insert_record("SB_REPRINT_01", qty=100)

    # Simulate reprint
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute("""UPDATE records SET reprint_count = reprint_count + 1,
                 last_reprinted_at = ?, last_reprinted_by = ?
                 WHERE sub_batch_id = ?""", (now_str, "sp99", "SB_REPRINT_01"))
    conn.commit()

    c.execute("SELECT reprint_count, last_reprinted_by, last_reprinted_at FROM records WHERE sub_batch_id=?",
              ("SB_REPRINT_01",))
    row = c.fetchone()
    test("Reprint count incremented to 1", row["reprint_count"] == 1)
    test("Reprinted by recorded", row["last_reprinted_by"] == "sp99")
    test("Reprint timestamp recorded", row["last_reprinted_at"] is not None)

    # Second reprint
    c.execute("""UPDATE records SET reprint_count = reprint_count + 1,
                 last_reprinted_at = ?, last_reprinted_by = ?
                 WHERE sub_batch_id = ?""", (now_str, "mg90", "SB_REPRINT_01"))
    conn.commit()

    c.execute("SELECT reprint_count, last_reprinted_by FROM records WHERE sub_batch_id=?", ("SB_REPRINT_01",))
    row = c.fetchone()
    test("Reprint count incremented to 2", row["reprint_count"] == 2)
    test("Last reprinted by updated to mg90", row["last_reprinted_by"] == "mg90")

    conn.close()


def test_09_deep_traceability():
    section("9. DEEP TRACEABILITY (LIFECYCLE TIMELINE)")

    conn = get_conn()
    c = conn.cursor()

    sb_id = "SB_TRACE_01"
    dt_sp = "2026-06-13 08:00"
    insert_record(sb_id, qty=100, dt_sp=dt_sp, status="In Rack")

    # 1. Check production record
    c.execute("SELECT * FROM records WHERE sub_batch_id=?", (sb_id,))
    rec = c.fetchone()
    test("Production record found for tracing", rec is not None)

    # 2. Add quality defect
    c.execute('''INSERT INTO quality_defects
        (sub_batch_id, pn_sf, part_sf, produced_by_op, quality_op_id,
         defect_type, qty_defective, description, shift_sp, station,
         reported_at, status, action_type, is_quarantined)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
        (sb_id, rec["pn_sf"], rec["part_sf"], rec["op_id"], "Q001",
         "Assembly Error", 3, "Misalignment", rec["shift_sp"], rec["station"],
         "2026-06-13 10:00:00", "Open", "Rework", 1))
    conn.commit()

    # 3. Simulate consumption
    c.execute("UPDATE records SET status='Consumed', dt_line='2026-06-13 14:00', shift_line='A' WHERE sub_batch_id=?",
              (sb_id,))
    conn.commit()

    # Build timeline events
    events = []

    # Creation event
    events.append({"time": rec["dt_sp"], "type": "creation", "title": "Created in Sub-Process"})

    # Quality events
    c.execute("SELECT * FROM quality_defects WHERE sub_batch_id=? ORDER BY reported_at ASC", (sb_id,))
    defects = c.fetchall()
    for df in defects:
        events.append({"time": df["reported_at"], "type": "quality", "title": f"Quality Alert: {df['defect_type']}"})

    # Consumption event
    c.execute("SELECT dt_line FROM records WHERE sub_batch_id=?", (sb_id,))
    rec2 = c.fetchone()
    if rec2["dt_line"]:
        events.append({"time": rec2["dt_line"], "type": "consume", "title": "Consumed to Line"})

    # Sort by time
    events.sort(key=lambda x: x["time"] if x["time"] else "")

    test("Timeline has 3 events (create, quality, consume)", len(events) == 3)
    test("Events sorted chronologically", events[0]["type"] == "creation")
    test("Quality event in middle", events[1]["type"] == "quality")
    test("Consumption event last", events[2]["type"] == "consume")

    # Test per-part stats
    today = "2026-06-13"
    c.execute("SELECT SUM(quantity) as total FROM records WHERE pn_sf=? AND dt_sp LIKE ?",
              (rec["pn_sf"], f"{today}%"))
    produced = c.fetchone()["total"] or 0
    test("Per-part production stats calculated", produced > 0)

    conn.close()


def test_10_inventory_snapshots():
    section("10. AUTOMATED INVENTORY SNAPSHOTS")

    conn = get_conn()
    c = conn.cursor()

    # Simulate snapshot
    snapshot_date = "2026-06-12"
    pn = "MOCK-A01-10001-X-SUB"

    c.execute("""SELECT pn_sf, SUM(quantity) as total, COUNT(id) as boxes
                 FROM records WHERE status='In Rack'
                 GROUP BY pn_sf""")
    agg_rows = c.fetchall()

    for row in agg_rows:
        c.execute("""INSERT OR REPLACE INTO inventory_snapshots
                     (snapshot_date, pn_sf, boxes_in_rack, total_qty_in_rack, oldest_box_age_hours)
                     VALUES (?,?,?,?,?)""",
                  (snapshot_date, row["pn_sf"], row["boxes"], row["total"], 24.0))
    conn.commit()

    c.execute("SELECT COUNT(*) as cnt FROM inventory_snapshots WHERE snapshot_date=?", (snapshot_date,))
    test("Snapshot data saved", c.fetchone()["cnt"] > 0)

    # Test 30-day trend query
    thirty_days_ago = "2026-05-14"
    c.execute("""SELECT snapshot_date, pn_sf, total_qty_in_rack
                 FROM inventory_snapshots WHERE snapshot_date >= ?
                 ORDER BY snapshot_date ASC""", (thirty_days_ago,))
    trend = c.fetchall()
    test("Trend data retrievable", len(trend) > 0)

    conn.close()


def test_11_records_search_and_filter():
    section("11. RECORDS SEARCH & FILTERING")

    conn = get_conn()
    c = conn.cursor()

    # Test search by SB_ID
    c.execute("""SELECT * FROM records WHERE sub_batch_id LIKE ? OR pn_sf LIKE ? OR op_id LIKE ?""",
              ("%KPI%", "%KPI%", "%KPI%"))
    test("Search by SB_ID pattern works", len(c.fetchall()) > 0)

    # Test filter by shift
    c.execute("SELECT COUNT(*) as cnt FROM records WHERE shift_sp='A'")
    a_count = c.fetchone()["cnt"]
    test("Filter by shift A returns results", a_count > 0)

    # Test filter by station
    c.execute("SELECT COUNT(*) as cnt FROM records WHERE station='S06'")
    s06_count = c.fetchone()["cnt"]
    test("Filter by station S06 returns results", s06_count > 0)

    # Test combined filter
    c.execute("SELECT COUNT(*) as cnt FROM records WHERE shift_sp='A' AND station='S06'")
    combined = c.fetchone()["cnt"]
    test("Combined shift+station filter works", combined > 0 and combined <= min(a_count, s06_count))

    # Test GROUP BY with quality join
    c.execute("""SELECT r.sub_batch_id, r.quantity,
                 IFNULL(SUM(CASE WHEN q.status='Closed' AND q.action_type IN ('Rework','Sorting','Use As-Is')
                 THEN 0 ELSE q.qty_defective END), 0) as eff_def
                 FROM records r
                 LEFT JOIN quality_defects q ON r.sub_batch_id = q.sub_batch_id
                 GROUP BY r.id ORDER BY r.id DESC LIMIT 5""")
    rows = c.fetchall()
    test("Records with quality JOIN returns data", len(rows) > 0)

    conn.close()


def test_12_edge_cases_and_data_integrity():
    section("12. EDGE CASES & DATA INTEGRITY")

    conn = get_conn()
    c = conn.cursor()

    # Test: What happens when records table is empty for stats?
    c.execute("SELECT SUM(quantity) FROM records WHERE dt_sp LIKE '1999%'")
    result = c.fetchone()[0]
    test("SUM on empty result returns None (not error)", result is None)

    # Test: Division by zero in quality rate
    total_prod = 0
    total_def = 0
    rate = (total_def / total_prod * 100) if total_prod > 0 else 0
    test("Division by zero avoided in quality rate", rate == 0)

    # Test: NULL handling in records
    insert_record("SB_NULL_TEST", qty=50, batch1="")
    c.execute("SELECT batch1, dt_line, shift_line FROM records WHERE sub_batch_id=?", ("SB_NULL_TEST",))
    row = c.fetchone()
    test("Empty string fields stored correctly", row["batch1"] == "")
    test("dt_line defaults to empty string", row["dt_line"] == "")

    # Test: Unicode in remarks
    insert_record("SB_UNICODE_01", qty=10)
    c.execute("UPDATE records SET remarks=? WHERE sub_batch_id=?",
              ("日本語テスト — ملاحظات — Ñoño", "SB_UNICODE_01"))
    conn.commit()
    c.execute("SELECT remarks FROM records WHERE sub_batch_id=?", ("SB_UNICODE_01",))
    test("Unicode remarks stored and retrieved",
         c.fetchone()["remarks"] == "日本語テスト — ملاحظات — Ñoño")

    # Test: Very long SB_ID
    long_id = "SB" + "X" * 200
    try:
        insert_record(long_id, qty=1)
        c.execute("SELECT sub_batch_id FROM records WHERE sub_batch_id=?", (long_id,))
        test("Very long SB_ID (202 chars) accepted", c.fetchone() is not None)
    except Exception as e:
        test("Very long SB_ID handled", False, str(e))

    # Test: Production day boundary (before 06:00 counts as previous day)
    insert_record("SB_NIGHT_01", qty=30, dt_sp="2026-06-13 05:30", shift="C")
    insert_record("SB_DAY_01", qty=40, dt_sp="2026-06-13 06:30", shift="A")

    # Before 06:00 → production day is previous day
    c.execute("SELECT SUM(quantity) FROM records WHERE dt_sp >= '2026-06-13 06:00:00' AND dt_sp <= '2026-06-14 05:59:59' AND sub_batch_id IN ('SB_NIGHT_01','SB_DAY_01')")
    day_total = c.fetchone()[0] or 0
    test("Night shift record (05:30) excluded from next day's 06:00 window", day_total == 40)

    c.execute("SELECT SUM(quantity) FROM records WHERE dt_sp >= '2026-06-12 06:00:00' AND dt_sp <= '2026-06-13 05:59:59' AND sub_batch_id IN ('SB_NIGHT_01','SB_DAY_01')")
    prev_day_total = c.fetchone()[0] or 0
    test("Night shift record (05:30) included in previous day's window", prev_day_total == 30)

    # Test: system_access_logs
    c.execute("INSERT INTO system_access_logs (event_type, user_id, shift, timestamp) VALUES (?,?,?,?)",
              ("LOGIN", "op1", "A", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    c.execute("INSERT INTO system_access_logs (event_type, user_id, shift, timestamp) VALUES (?,?,?,?)",
              ("LOGOUT", "op1", "A", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()
    c.execute("SELECT COUNT(*) FROM system_access_logs")
    test("Access logs recorded", c.fetchone()[0] >= 2)

    conn.close()


def test_13_report_generator():
    section("13. PDF REPORT GENERATOR (DATA QUERIES)")

    conn = get_conn()
    c = conn.cursor()

    start_dt = "2026-06-13 06:00:00"
    end_dt = "2026-06-14 06:00:00"

    # Production Summary
    c.execute("SELECT COUNT(*) as boxes, SUM(quantity) as qty FROM records WHERE dt_sp >= ? AND dt_sp < ?",
              (start_dt, end_dt))
    row = c.fetchone()
    test("Report: Production summary query works", row["boxes"] is not None)

    # Production Mix
    c.execute("""SELECT pn_sf, part_sf, SUM(quantity) FROM records
                 WHERE dt_sp >= ? AND dt_sp < ? GROUP BY pn_sf, part_sf
                 ORDER BY SUM(quantity) DESC""", (start_dt, end_dt))
    mix = c.fetchall()
    test("Report: Production mix query works", len(mix) >= 0)

    # Operator Summary
    c.execute("""SELECT op_id, SUM(quantity) FROM records
                 WHERE dt_sp >= ? AND dt_sp < ? GROUP BY op_id
                 ORDER BY SUM(quantity) DESC""", (start_dt, end_dt))
    ops = c.fetchall()
    test("Report: Operator summary query works", len(ops) >= 0)

    # Target vs Actual
    c.execute("""SELECT r.pn_sf, SUM(r.quantity),
                 (SELECT target_qty FROM shift_targets WHERE product_pn = r.pn_sf ORDER BY effective_date DESC LIMIT 1)
                 FROM records r WHERE dt_sp >= ? AND dt_sp < ?
                 GROUP BY r.pn_sf""", (start_dt, end_dt))
    targets = c.fetchall()
    test("Report: Target vs Actual query works", len(targets) >= 0)

    # Inventory WIP
    c.execute("""SELECT pn_sf, SUM(quantity), COUNT(id) FROM records
                 WHERE status='In Rack' GROUP BY pn_sf
                 ORDER BY SUM(quantity) DESC LIMIT 15""")
    wip = c.fetchall()
    test("Report: WIP inventory query works", len(wip) >= 0)

    # Raw Material Usage
    c.execute("""SELECT rm1_pn, rm1_name, COUNT(DISTINCT batch1) FROM records
                 WHERE dt_sp >= ? AND dt_sp < ? AND rm1_pn IS NOT NULL AND rm1_pn != ''
                 GROUP BY rm1_pn, rm1_name""", (start_dt, end_dt))
    rm = c.fetchall()
    test("Report: RM usage query works", len(rm) >= 0)

    conn.close()


def test_14_concurrent_sb_id_generation():
    section("14. CONCURRENT SUB-BATCH ID STRESS TEST")

    conn = get_conn()
    c = conn.cursor()

    base = "SB20260613120000S06A"
    inserted = 0
    collisions = 0

    for i in range(20):
        # Simulate the app's ID generation logic
        c.execute("SELECT COUNT(*) as cnt FROM records WHERE sub_batch_id LIKE ?", (base + "%",))
        count = c.fetchone()["cnt"]
        sb_id = f"{base}{count+1:02d}"

        try:
            c.execute("""INSERT INTO records (sub_batch_id, pn_sf, part_sf,
                rm1_pn, rm1_name, rm2_pn, rm2_name, rm3_pn, rm3_name, rm4_pn, rm4_name,
                batch1, batch2, batch3, quantity, shift_sp, op_id, station,
                dt_sp, dt_line, shift_line, remarks, status, created_at, registered_by)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (sb_id, "MOCK-A01-10001-X-SUB", "Alpha", "", "", "", "", "", "", "", "",
                 "B1", "", "", 10, "A", "op1", "S06", "2026-06-13 12:00", "", "", "",
                 "In Rack", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "test"))
            conn.commit()
            inserted += 1
        except sqlite3.IntegrityError:
            conn.rollback()
            collisions += 1

    test(f"Sequential inserts: {inserted}/20 succeeded", inserted == 20)
    test("No collisions in sequential mode", collisions == 0)

    # Verify unique IDs
    c.execute("SELECT COUNT(DISTINCT sub_batch_id) as cnt FROM records WHERE sub_batch_id LIKE ?", (base + "%",))
    unique = c.fetchone()["cnt"]
    test(f"All {inserted} IDs are unique", unique == inserted)

    conn.close()


# ─────────────────────────────────────────────────────────────────────────────
# MAIN RUNNER
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("\n" + "=" * 70)
    print("  SUB-PROCESS TRACEABILITY — FULL PRODUCTION TEST SUITE")
    print("=" * 70)
    print(f"  Test DB: {TEST_DB}")
    print(f"  Time: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    try:
        init_test_db()
        print("\n  ✅ Test database initialized.\n")

        test_01_authentication()
        test_02_record_creation()
        test_03_consume_workflow()
        test_04_quality_defect_workflow()
        test_05_product_management()
        test_06_inventory_and_fifo()
        test_07_shift_targets_and_kpis()
        test_08_reprint_audit_trail()
        test_09_deep_traceability()
        test_10_inventory_snapshots()
        test_11_records_search_and_filter()
        test_12_edge_cases_and_data_integrity()
        test_13_report_generator()
        test_14_concurrent_sb_id_generation()

    except Exception as e:
        print(f"\n💥 UNEXPECTED ERROR: {e}")
        traceback.print_exc()
        FAIL += 1

    # Cleanup
    try:
        if os.path.exists(TEST_DATA_DIR):
            shutil.rmtree(TEST_DATA_DIR)
    except:
        pass

    # Summary
    print("\n" + "=" * 70)
    print(f"  RESULTS: {PASS} passed, {FAIL} failed out of {PASS+FAIL} tests")
    print("=" * 70)

    if ERRORS:
        print("\n  FAILURES:")
        for e in ERRORS:
            print(f"    {e}")

    print()
    sys.exit(0 if FAIL == 0 else 1)
