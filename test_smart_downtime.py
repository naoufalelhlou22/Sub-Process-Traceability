import sqlite3
import datetime

def test_smart_downtime():
    print("=== Testing Smart Downtime Logic ===")
    conn = sqlite3.connect(':memory:')
    c = conn.cursor()
    
    # Setup tables
    c.execute('''CREATE TABLE records (
        sub_batch_id TEXT, pn_sf TEXT, quantity REAL, shift_sp TEXT, op_id TEXT, created_at TEXT
    )''')
    c.execute('''CREATE TABLE shift_targets (
        id INTEGER PRIMARY KEY, product_pn TEXT, target_qty REAL
    )''')
    c.execute('''CREATE TABLE downtime_logs (
        sub_batch_id TEXT, station TEXT, shift TEXT, op_id TEXT, duration_min REAL, reason TEXT, created_at TEXT
    )''')
    
    # Define our test time: We are currently at 15:05
    # So the "previous hour" is 14:00:00 to 14:59:59
    current_time = datetime.datetime(2026, 6, 13, 15, 5, 0)
    op_id = "op1"
    shift_sp = "Morning"
    
    prev_hour_end = current_time.replace(minute=0, second=0, microsecond=0) - datetime.timedelta(seconds=1)
    prev_hour_start = prev_hour_end.replace(minute=0, second=0, microsecond=0)
    
    print(f"Current Time: {current_time}")
    print(f"Evaluating Previous Hour: {prev_hour_start} to {prev_hour_end}\n")
    
    # --- Scenario 1: Same Ref ---
    print("--- Scenario 1: Single Product (Same Ref) ---")
    c.execute("DELETE FROM records")
    c.execute("DELETE FROM shift_targets")
    c.execute("DELETE FROM downtime_logs")
    
    # Target: 240 / hour
    c.execute("INSERT INTO shift_targets (product_pn, target_qty) VALUES ('RefA', 240)")
    
    # Production: 200 pieces (took them 60 mins, but standard time is 50 mins. Deficit = 10 mins)
    c.execute("INSERT INTO records (pn_sf, quantity, shift_sp, op_id, created_at) VALUES ('RefA', 200, ?, ?, ?)",
              (shift_sp, op_id, "2026-06-13 14:30:00"))
              
    evaluate_logic(c, op_id, shift_sp, prev_hour_start, prev_hour_end)
    
    # --- Scenario 2: Different Refs (Mixed) ---
    print("\n--- Scenario 2: Mixed Products (Different Refs) ---")
    c.execute("DELETE FROM records")
    c.execute("DELETE FROM shift_targets")
    c.execute("DELETE FROM downtime_logs")
    
    # Targets: RefA=240/h, RefB=300/h
    c.execute("INSERT INTO shift_targets (product_pn, target_qty) VALUES ('RefA', 240)")
    c.execute("INSERT INTO shift_targets (product_pn, target_qty) VALUES ('RefB', 300)")
    
    # Production: 120 of RefA (Earns 30 mins) + 100 of RefB (Earns 20 mins) = Total 50 mins earned. Deficit = 10 mins
    c.execute("INSERT INTO records (pn_sf, quantity, shift_sp, op_id, created_at) VALUES ('RefA', 120, ?, ?, ?)",
              (shift_sp, op_id, "2026-06-13 14:20:00"))
    c.execute("INSERT INTO records (pn_sf, quantity, shift_sp, op_id, created_at) VALUES ('RefB', 100, ?, ?, ?)",
              (shift_sp, op_id, "2026-06-13 14:50:00"))
              
    evaluate_logic(c, op_id, shift_sp, prev_hour_start, prev_hour_end)
    
    # --- Scenario 3: Overperforming Mixed ---
    print("\n--- Scenario 3: Overperforming Mixed Products ---")
    c.execute("DELETE FROM records")
    c.execute("DELETE FROM shift_targets")
    c.execute("DELETE FROM downtime_logs")
    
    c.execute("INSERT INTO shift_targets (product_pn, target_qty) VALUES ('RefA', 240)")
    c.execute("INSERT INTO shift_targets (product_pn, target_qty) VALUES ('RefB', 300)")
    
    # Production: 120 of RefA (30 mins) + 200 of RefB (40 mins) = 70 mins earned. Deficit = -10 mins
    c.execute("INSERT INTO records (pn_sf, quantity, shift_sp, op_id, created_at) VALUES ('RefA', 120, ?, ?, ?)",
              (shift_sp, op_id, "2026-06-13 14:20:00"))
    c.execute("INSERT INTO records (pn_sf, quantity, shift_sp, op_id, created_at) VALUES ('RefB', 200, ?, ?, ?)",
              (shift_sp, op_id, "2026-06-13 14:50:00"))
              
    evaluate_logic(c, op_id, shift_sp, prev_hour_start, prev_hour_end)
    
    # --- Scenario 4: Missing Target ---
    print("\n--- Scenario 4: Missing Target on One Product ---")
    c.execute("DELETE FROM records")
    c.execute("DELETE FROM shift_targets")
    c.execute("DELETE FROM downtime_logs")
    
    c.execute("INSERT INTO shift_targets (product_pn, target_qty) VALUES ('RefA', 240)")
    # No target for RefC!
    
    c.execute("INSERT INTO records (pn_sf, quantity, shift_sp, op_id, created_at) VALUES ('RefA', 120, ?, ?, ?)",
              (shift_sp, op_id, "2026-06-13 14:20:00"))
    c.execute("INSERT INTO records (pn_sf, quantity, shift_sp, op_id, created_at) VALUES ('RefC', 50, ?, ?, ?)",
              (shift_sp, op_id, "2026-06-13 14:50:00"))
              
    evaluate_logic(c, op_id, shift_sp, prev_hour_start, prev_hour_end)

def evaluate_logic(c, op_id, shift_sp, prev_hour_start, prev_hour_end):
    c.execute("SELECT pn_sf, SUM(quantity) FROM records WHERE op_id=? AND shift_sp=? AND created_at >= ? AND created_at <= ? GROUP BY pn_sf",
              (op_id, shift_sp, prev_hour_start.strftime("%Y-%m-%d %H:%M:%S"), prev_hour_end.strftime("%Y-%m-%d %H:%M:%S")))
    prod_rows = c.fetchall()
    
    if prod_rows:
        earned_minutes = 0
        has_missing_target = False
        
        print(f"Products Produced in Hour:")
        for row in prod_rows:
            pn = row[0]
            qty = row[1]
            print(f"  - {pn}: {qty} pieces")
            
            c.execute("SELECT target_qty FROM shift_targets WHERE product_pn=? ORDER BY id DESC LIMIT 1", (pn,))
            target_row = c.fetchone()
            
            if target_row:
                db_target = target_row[0]
                hourly_target = (db_target / 8.0) if db_target > 500 else db_target
                earned = (qty / hourly_target) * 60
                earned_minutes += earned
                print(f"    Target: {hourly_target}/hr -> Earned Mins: {earned:.2f} mins")
            else:
                print(f"    Target: MISSING! Cannot evaluate.")
                has_missing_target = True
                break
                
        if not has_missing_target:
            c.execute("SELECT SUM(duration_min) FROM downtime_logs WHERE op_id=? AND shift=? AND created_at >= ? AND created_at <= ?", 
                      (op_id, shift_sp, prev_hour_start.strftime("%Y-%m-%d %H:%M:%S"), prev_hour_end.strftime("%Y-%m-%d %H:%M:%S")))
            logged_dt = c.fetchone()[0] or 0
            
            effective_elapsed = 60 - logged_dt
            deficit_mins = effective_elapsed - earned_minutes
            
            print(f"> Total Earned Mins: {earned_minutes:.2f}")
            print(f"> Elapsed Mins (minus logged DT): {effective_elapsed}")
            print(f"> Deficit Mins: {deficit_mins:.2f}")
            
            if deficit_mins > 3.75:
                print(">> [RESULT]: Deficit > 3.75. PROMPT DOWNTIME TRIGGERED!")
            else:
                print(">> [RESULT]: Deficit OK. NO PROMPT.")
        else:
            print(">> [RESULT]: Aborted check due to missing target.")
            
if __name__ == "__main__":
    test_smart_downtime()
