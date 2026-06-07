import sys
import tkinter as tk
from traceability_v3 import TraceabilityApp

print("Starting test...")

root = tk.Tk()
root.withdraw()

app = TraceabilityApp()
app.withdraw()

# 1. Provide fake JSON
test_json = '{"full_pn_sf": "AT-V07-11313-A-SUB", "part_name_sf": "SUB BODY N R", "rm1_pn": "AT-V07-11313-P-WIE", "rm1_name": "BODY R", "batch_no_1": "TEST-123", "quantity": 150, "op_id": "888", "station": "S06", "sub_process_shift": "A"}'

# 2. Set the scan input
app.var_scan_input.set(test_json)

# 3. Simulate pressing Enter on the scan entry
app.on_scan_enter(None)

# 4. Read back the fields
print('--- Test Results ---')
print('PN Semi fini:', app.cb_sf_pn.get())
print('Batch 1:', app.var_b1.get())
print('Quantity:', app.var_qty.get())
print('Op ID:', app.var_op_id.get())
print('Station:', app.cb_station.get())

if app.cb_sf_pn.get() == 'AT-V07-11313-A-SUB' and app.var_b1.get() == 'TEST-123':
    print('SUCCESS: The QR Scan Logic successfully populated the form!')
else:
    print('FAILED: Fields were not populated correctly.')

app.destroy()
