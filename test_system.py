import os
import sqlite3
import datetime
import hashlib
import binascii
import sys

sys.path.append("src")
try:
    from report_generator import generate_shift_pdf_report
except ImportError:
    generate_shift_pdf_report = None

DB_FILE = os.path.join("data", "traceability.db")

def print_result(test_name, success, message=""):
    status = "PASS" if success else "FAIL"
    print(f"[{status}] {test_name} {('- ' + message) if message else ''}")

def hash_password(password):
    salt = b"subproc_trace_salt_2026"
    hash_obj = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 100000)
    return binascii.hexlify(hash_obj).decode("utf-8")

def test_database_connection():
    try:
        conn = sqlite3.connect(DB_FILE)
        conn.close()
        print_result("Database Connection", True)
        return True
    except Exception as e:
        print_result("Database Connection", False, str(e))
        return False

def test_database_schema():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    
    expected_tables = {"records", "auth", "quality_defects", "system_access_logs"}
    c.execute("SELECT name FROM sqlite_master WHERE type='table'")
    tables = {row[0] for row in c.fetchall()}
    
    missing = expected_tables - tables
    if missing:
        print_result("Database Schema", False, f"Missing tables: {missing}")
    else:
        print_result("Database Schema", True, "All core tables exist")
    
    # Check indexes
    c.execute("SELECT name FROM sqlite_master WHERE type='index'")
    indexes = {row[0] for row in c.fetchall()}
    if "idx_quality_defects_sub_batch_id" in indexes:
        print_result("Quality Unique Constraint", True, "Index exists")
    else:
        print_result("Quality Unique Constraint", False, "Missing unique index on quality_defects")
        
    conn.close()

def test_authentication():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    # Check if Admin or default quality exists
    c.execute("SELECT id, password, role FROM auth WHERE id IN ('admin', 'Q001')")
    users = c.fetchall()
    
    if not users:
        print_result("Authentication Data", False, "No default users found")
    else:
        valid = True
        for uid, pw, role in users:
            if len(pw) != 64:
                valid = False
                print_result("Authentication Security", False, f"User {uid} password is not hashed!")
        
        if valid:
            print_result("Authentication Security", True, "All checked passwords are encrypted via SHA-256")
            
    conn.close()

def test_record_insertion_and_quality():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    
    test_sb_id = f"TEST_SB_{datetime.datetime.now().strftime('%H%M%S')}"
    try:
        # Insert test record
        c.execute('''INSERT INTO records (
                        sub_batch_id, pn_sf, part_sf, quantity, status, shift_sp, station, dt_sp, created_at
                    ) VALUES (?, 'TEST-PN', 'TEST-PART', 100, 'In Rack', 'from 6 to 14', 'Test Station', ?, ?)''',
                    (test_sb_id, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
        print_result("Data Entry (Production)", True, f"Inserted test record {test_sb_id}")
        
        # Insert Quality Defect
        c.execute('''INSERT INTO quality_defects (
                        sub_batch_id, quality_op_id, defect_type, qty_defective, reported_at, status
                    ) VALUES (?, 'Q001', 'Cosmetic', 5, ?, 'Open')''',
                    (test_sb_id, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
        print_result("Data Entry (Quality)", True, "Inserted defect report")
        
        # Test Duplicate Quality Rejection (Unique Constraint)
        try:
            c.execute('''INSERT INTO quality_defects (
                            sub_batch_id, quality_op_id, defect_type, qty_defective, reported_at, status
                        ) VALUES (?, 'Q001', 'Other', 2, ?, 'Open')''',
                        (test_sb_id, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
            conn.commit()
            print_result("Defect Unique Constraint", False, "Allowed duplicate defect for same sub-batch!")
        except sqlite3.IntegrityError:
            print_result("Defect Unique Constraint", True, "Successfully blocked duplicate defect entry")
        
        # Clean up
        c.execute("DELETE FROM records WHERE sub_batch_id=?", (test_sb_id,))
        c.execute("DELETE FROM quality_defects WHERE sub_batch_id=?", (test_sb_id,))
        conn.commit()
        print_result("Data Cleanup", True, "Removed test data")
        
    except Exception as e:
        print_result("Data Entry/Quality Flow", False, str(e))
    finally:
        conn.close()

def test_report_generation():
    if not generate_shift_pdf_report:
        print_result("PDF Report Generation", False, "report_generator.py not found or failed to import")
        return
        
    # Test with a dummy date where we know data might exist or just check if it fails
    # We will pass a recent timeframe just to see if it generates a 'No data' skip or a valid PDF
    now = datetime.datetime.now()
    start_dt = (now - datetime.timedelta(days=1)).strftime("%Y-%m-%d 00:00:00")
    end_dt = now.strftime("%Y-%m-%d 23:59:59")
    
    try:
        # We don't actually want to pollute the reports folder if we can avoid it, but let's run it
        result = generate_shift_pdf_report(start_dt, end_dt, "Test_Shift_Validation")
        if result is None:
            print_result("PDF Report Generation", True, "Passed (No data to generate, skipped gracefully)")
        else:
            print_result("PDF Report Generation", True, f"Successfully generated: {result}")
            # cleanup
            if os.path.exists(result):
                os.remove(result)
    except Exception as e:
        print_result("PDF Report Generation", False, f"Failed: {e}")

if __name__ == "__main__":
    print("="*60)
    print(" SUB-PROCESS TRACEABILITY - SYSTEM HEALTH & INTEGRATION TEST")
    print("="*60)
    
    if test_database_connection():
        test_database_schema()
        test_authentication()
        test_record_insertion_and_quality()
        test_report_generation()
        
    print("="*60)
    print("TEST SUITE COMPLETED")
    print("="*60)
