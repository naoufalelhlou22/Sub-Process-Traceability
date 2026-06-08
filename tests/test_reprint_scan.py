import sys
import os
import json
import sqlite3
import tkinter as tk
from tkinter import messagebox
import importlib.util
import time

print("Loading Sub-Process Traceability application...")

base_dir = os.path.dirname(os.path.dirname(__file__))
main_script = os.path.join(base_dir, "src", "main.py")

# Dynamically import the main module
spec = importlib.util.spec_from_file_location("main", main_script)
app_module = importlib.util.module_from_spec(spec)
sys.modules["main"] = app_module
spec.loader.exec_module(app_module)

App = app_module.TraceabilityApp

# --- Mocking to bypass login and dialogs for automated testing ---
def mock_prompt_login(self):
    self.app_user_id = "TEST-ADMIN"
    self.app_user_shift = "A"
    self.is_admin = True
    try:
        self.lbl_header_user.config(text="User: TEST-ADMIN | Shift: A")
    except AttributeError:
        pass

App.prompt_login = mock_prompt_login

def mock_showerror(title, message, parent=None):
    print(f"      -> [POPUP INTERCEPTED] Error: {message}")
messagebox.showerror = mock_showerror

def mock_showinfo(title, message, parent=None):
    print(f"      -> [POPUP INTERCEPTED] Info: {message}")
messagebox.showinfo = mock_showinfo

# -----------------------------------------------------------------

def get_latest_record():
    db_path = os.path.join(base_dir, "data", "traceability.db")
    if not os.path.exists(db_path):
        return None
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    c.execute("SELECT * FROM records ORDER BY id DESC LIMIT 1")
    row = c.fetchone()
    conn.close()
    return dict(row) if row else None

def run_test():
    print("Initializing App UI...")
    app_module.init_db()
    app = App()
    app.deiconify()
    app.update()
    
    print("\nNavigating to Tab 4 (Search & Reprint)...")
    app.notebook.select(2) # Tab 4 (Search & Reprint) is at index 2
    app.update()
    time.sleep(1)
    
    def test_reprint_scan(description, payload, should_succeed):
        print(f"\n=========================================")
        print(f"[*] TEST: {description}")
        print(f"[*] Scanning Payload: '{payload}'")
        
        # Reset fields before test
        app.var_reprint_scan.set("")
        app.txt_reprint_details.config(state="normal")
        app.txt_reprint_details.delete("1.0", tk.END)
        app.txt_reprint_details.config(state="disabled")
        app.btn_do_reprint.config(state="disabled")
        if hasattr(app, 'lbl_reprint_status'):
            app.lbl_reprint_status.config(text="")
        app.update()
        
        # Simulate scanning the payload
        app.var_reprint_scan.set(payload)
        app.update()
        time.sleep(0.5)
        
        # Trigger the scan event logic
        app.on_reprint_scan(None)
        app.update()
        time.sleep(1)
        
        # Read the resulting state
        details = app.txt_reprint_details.get("1.0", tk.END).strip()
        btn_state = str(app.btn_do_reprint['state'])
        
        if details and btn_state == "normal":
            if should_succeed:
                print("   [PASS] Details populated and Print Button enabled!")
            else:
                print("   [FAIL] Expected failure, but the test succeeded!")
        else:
            if not should_succeed:
                print("   [PASS] Rejected input as expected.")
            else:
                print("   [FAIL] Failed to process a valid payload!")

    record = get_latest_record()
    if not record:
        print("\n[WARN] No records found in DB. Inserting a real MOCK record into the database for Tests 1 & 2.")
        
        db_path = os.path.join(base_dir, "data", "traceability.db")
        os.makedirs(os.path.dirname(db_path), exist_ok=True)
        
        # Connect and create table if not exists (just in case)
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        c.execute('''
            CREATE TABLE IF NOT EXISTS records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sub_batch_id TEXT, pn_sf TEXT, part_sf TEXT,
                rm1_pn TEXT, rm1_name TEXT, rm2_pn TEXT, rm2_name TEXT,
                rm3_pn TEXT, rm3_name TEXT, rm4_pn TEXT, rm4_name TEXT,
                batch1 TEXT, batch2 TEXT, batch3 TEXT,
                quantity INTEGER, dt_sp TEXT, shift_sp TEXT,
                dt_line TEXT, shift_line TEXT,
                op_id TEXT, station TEXT, remarks TEXT, status TEXT
            )
        ''')
        
        # Insert a real valid MOCK record
        c.execute('''
            INSERT INTO records (
                sub_batch_id, pn_sf, part_sf,
                rm1_pn, rm1_name, rm2_pn, rm2_name,
                rm3_pn, rm3_name, rm4_pn, rm4_name,
                batch1, batch2, batch3,
                quantity, dt_sp, shift_sp,
                dt_line, shift_line,
                op_id, station, remarks, status
            ) VALUES (
                'SB202606081630S06A', 'MOCK-A01-10001-X-SUB', 'Alpha Subsystem Right',
                'MOCK-A01-10001-Y-PRT', 'Alpha Core Right', 'MOCK-A01-10002-Z-PRT', 'Alpha Buffer Right',
                '', '', '', '',
                'HLAT', '', '',
                120, '2026-06-08 16:30', 'A',
                '2026-06-08 16:31', '',
                '232', 'S06', '', 'Consumed'
            )
        ''')
        conn.commit()
        conn.close()
        
        # Fetch the newly inserted record
        record = get_latest_record()

    # Test 1: Valid JSON from DB (QR Code)
    qr_data = {
        "sub_batch_id": record.get("sub_batch_id", ""),
        "full_pn_sf": record.get("pn_sf", ""),
        "part_sf": record.get("part_sf", ""),
        "rm1_pn": record.get("rm1_pn", ""),
        "rm2_pn": record.get("rm2_pn", ""),
        "rm3_pn": record.get("rm3_pn", ""),
        "rm4_pn": record.get("rm4_pn", ""),
        "batch_no_1": record.get("batch1", ""),
        "batch_no_2": record.get("batch2", ""),
        "batch_no_3": record.get("batch3", ""),
        "quantity": record.get("quantity", 0),
        "sub_process_shift": record.get("shift_sp", ""),
        "op_id": record.get("op_id", ""),
        "station": record.get("station", ""),
        "sub_process_datetime": record.get("dt_sp", "")
    }
    valid_json = json.dumps(qr_data)
    test_reprint_scan("Valid QR JSON Payload", valid_json, should_succeed=True)
    time.sleep(2)
    
    # Test 2: Literal SB_ID (Raw String input) should now be accepted again!
    raw_sb_id = record.get("sub_batch_id", "")
    test_reprint_scan("Raw SB_ID String", raw_sb_id, should_succeed=True)
    time.sleep(2)

    # Test 3: Invalid JSON missing SB_ID
    invalid_json = '{"rm1_pn": "MOCK-123", "quantity": 50}'
    test_reprint_scan("JSON missing sub_batch_id", invalid_json, should_succeed=False)
    time.sleep(2)
    
    # Test 4: Broken/Malformed JSON
    broken_json = '{"sub_batch_id": "SB2026'
    test_reprint_scan("Broken/Malformed JSON", broken_json, should_succeed=False)
    time.sleep(2)
    
    # Test 5: Fake SB_ID
    fake_id = "SB99999999S00X"
    test_reprint_scan("Non-existent SB_ID", fake_id, should_succeed=False)

    print("\n=========================================")
    print("All tests completed.")
    print("The app will close automatically in 5 seconds...")
    
    app.after(5000, app.destroy)
    app.mainloop()

if __name__ == "__main__":
    run_test()
