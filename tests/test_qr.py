import sys
import tkinter as tk
from tkinter import messagebox
import importlib.util
import time

print("Loading Sub-Process Traceability application...")

# Load the module dynamically since the filename contains spaces and hyphens
module_name = "Sub-Process Traceability"
file_path = "Sub-Process Traceability.py"
spec = importlib.util.spec_from_file_location(module_name, file_path)
app_module = importlib.util.module_from_spec(spec)
sys.modules[module_name] = app_module
spec.loader.exec_module(app_module)

TraceabilityApp = app_module.TraceabilityApp

# Override messagebox to prevent the test from halting when an error popup appears
def mock_showerror(title, message):
    print(f"      -> [POPUP INTERCEPTED] {title}: {message}")
messagebox.showerror = mock_showerror

def run_test():
    print("Initializing App UI...")
    app = TraceabilityApp()
    # Bring the window to the front
    app.deiconify()
    app.update()
    
    def test_qr(description, qr_data, should_succeed):
        print(f"\n=========================================")
        print(f"[*] TEST: {description}")
        print(f"[*] Scanning QR Code Data: '{qr_data}'")
        
        # Clear the form first
        app.clear_form()
        app.update()
        
        # Set the scan input
        app.var_scan_input.set(qr_data)
        app.update()
        time.sleep(1)
        
        # Simulate pressing Enter on the scan entry
        app.on_scan_enter(None)
        app.update()
        time.sleep(1)
        
        # Read back the fields
        sf_pn = app.cb_sf_pn.get()
        b1 = app.var_b1.get()
        
        if sf_pn:
            print(f"      -> Form Populated with SF PN: {sf_pn}")
            if should_succeed:
                print("   [PASS] Successfully populated the correct data.")
            else:
                print("   [FAIL] Populated data when it should have been rejected!")
        else:
            print(f"      -> Form remains empty.")
            if not should_succeed:
                print("   [PASS] Successfully rejected the invalid QR Code.")
            else:
                print("   [FAIL] Failed to populate data for a valid QR Code!")

    # Test 1: Valid QR Code (Proper JSON)
    valid_json = '{"full_pn_sf": "AT-V07-11313-A-SUB", "part_name_sf": "SUB BODY N R", "rm1_pn": "AT-V07-11313-P-WIE", "rm1_name": "BODY R", "batch_no_1": "TEST-123", "quantity": 150, "op_id": "888", "station": "S06", "sub_process_shift": "A", "sub_process_datetime": "2024-05-18 14:30:00", "production_line_entry_datetime": "2024-05-18 15:00", "production_line_shift": "B", "remarks": "Urgent order"}'
    test_qr("Valid QR Code (JSON format)", valid_json, should_succeed=True)
    
    print("\nWaiting 10 seconds before next test...")
    for i in range(10):
        app.update()
        time.sleep(1)
    
    # Test 2: Missing 'rm1_pn' (Should Reject)
    missing_rm1_json = '{"full_pn_sf": "AT-V07-11314-A-SUB", "part_name_sf": "SUB BODY N L", "batch_no_1": "TEST-456"}'
    test_qr("Valid JSON but missing rm1_pn", missing_rm1_json, should_succeed=False)
    
    print("\nWaiting 10 seconds before next test...")
    for i in range(10):
        app.update()
        time.sleep(1)
    
    # Test 3: Invalid / Malformed QR Code
    invalid_qr = "THIS IS A BROKEN QR CODE TRASH DATA"
    test_qr("Invalid/Fake QR Code", invalid_qr, should_succeed=False)
    
    print("\n=========================================")
    print("All tests completed.")
    print("The app will close automatically in 5 seconds...")
    
    app.after(5000, app.destroy)
    app.mainloop()

if __name__ == "__main__":
    run_test()
