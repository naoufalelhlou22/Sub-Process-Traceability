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
    
    # 1. Switch to the 'Print Label' tab (index 1)
    app.notebook.select(1)
    app.update()
    
    def test_barcode(description, barcode, should_succeed):
        print(f"\n=========================================")
        print(f"[*] TEST: {description}")
        print(f"[*] Scanning Barcode: '{barcode}'")
        
        # Insert barcode into the scan entry
        app.pl_var_scan_rm.set(barcode)
        app.update()
        time.sleep(1)
        
        # Trigger the Enter key event handler
        app.pl_on_rm_scanned(None)
        app.update()
        time.sleep(1)
        
        # Check the results
        matched_sf_pn = getattr(app, 'pl_var_sf_pn', tk.StringVar()).get()
        matched_sf_name = app.pl_var_part_sf.get()
        
        if matched_sf_pn:
            print(f"      -> Matched SF PN: {matched_sf_pn} ({matched_sf_name})")
            if should_succeed:
                print("[PASS] Successfully populated the correct data.")
            else:
                print("[FAIL] Populated data when it should have been rejected!")
        else:
            print(f"-> Matched SF PN: (Empty)")
            if not should_succeed:
                print("[PASS] Successfully rejected the invalid barcode.")
            else:
                print("[FAIL] Failed to populate data for a valid barcode!")

    # Test 1: Valid Primary Component Barcode
    test_barcode("Valid Primary Component", "AT-V07-11313-P-WIE", should_succeed=True)
    
    # Test 2: Invalid/Secondary Component Barcode (Bumper Rubber)
    test_barcode("Secondary Component (Should Reject)", "AT-V07-11323-E-WIE", should_succeed=False)
    
    # Test 3: Completely Fake Barcode
    test_barcode("Completely Fake Barcode (Should Reject)", "INVALID-12345", should_succeed=False)
    
    print("\n=========================================")
    print("All tests completed.")
    print("The app will close automatically in 5 seconds...")
    
    app.after(5000, app.destroy)
    app.mainloop()

if __name__ == "__main__":
    run_test()
