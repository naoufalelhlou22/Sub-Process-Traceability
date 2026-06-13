import sqlite3
import datetime

# Connect to the database
conn = sqlite3.connect('data/traceability.db')
c = conn.cursor()

# 1. Update the Product to have a Std Target/Hr/Op of 100
target_pn = "MOCK-A01-10001-X-SUB"
c.execute("UPDATE products SET std_hourly_target = 100 WHERE pn_sf = ?", (target_pn,))

# 2. Calculate the "Previous Hour" exactly as the main app does
current_time = datetime.datetime.now()
prev_hour_end = current_time.replace(minute=0, second=0, microsecond=0) - datetime.timedelta(seconds=1)
prev_hour_start = prev_hour_end.replace(minute=0, second=0, microsecond=0)

# Pick a random minute in the previous hour to inject the production records
simulated_time = prev_hour_start + datetime.timedelta(minutes=30)
simulated_time_str = simulated_time.strftime("%Y-%m-%d %H:%M:%S")

# Clean up any previous test records for these dummy operators in the previous hour
c.execute("DELETE FROM records WHERE op_id IN ('TEST_OP_WEAK', 'TEST_OP_GOOD') AND created_at >= ?", (prev_hour_start.strftime("%Y-%m-%d %H:%M:%S"),))

# 3. Inject Record for Operator 1: "TEST_OP_WEAK" (Produced only 50/100) -> Should trigger 30 min Downtime
c.execute("""
    INSERT INTO records (sub_batch_id, pn_sf, quantity, op_id, shift_sp, station, created_at, status) 
    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
""", ("SIM_WEAK_BATCH", target_pn, 50, "TEST_OP_WEAK", "Shift 1", "Station 1", simulated_time_str, "In Rack"))

# 4. Inject Record for Operator 2: "TEST_OP_GOOD" (Produced 120/100) -> Should NOT trigger Downtime
c.execute("""
    INSERT INTO records (sub_batch_id, pn_sf, quantity, op_id, shift_sp, station, created_at, status) 
    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
""", ("SIM_GOOD_BATCH", target_pn, 120, "TEST_OP_GOOD", "Shift 1", "Station 1", simulated_time_str, "In Rack"))

conn.commit()
conn.close()

print("="*60)
print("✅ SIMULATION DATA INJECTED SUCCESSFULLY!")
print("="*60)
print(f"Product: {target_pn} (Target set to 100/hr)")
print(f"Time injected: {simulated_time_str} (Previous Hour)")
print("\n🟢 NEXT STEPS:")
print("1. Open the application (I will launch it for you).")
print("2. Try scanning ANY product right now using Operator ID: TEST_OP_WEAK")
print("   -> The system will STOP you and demand 30 minutes of Downtime justification!")
print("3. Try scanning ANY product right now using Operator ID: TEST_OP_GOOD")
print("   -> The system will let it pass seamlessly because they produced 120/100.")
print("="*60)
