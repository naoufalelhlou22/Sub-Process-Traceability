import sys
import os
import tkinter as tk
from tkinter import messagebox
import importlib.util
import time

print("Loading Sub-Process Traceability application...")

base_dir = os.path.dirname(os.path.dirname(__file__))
main_script = os.path.join(base_dir, "src", "main.py")

spec = importlib.util.spec_from_file_location("main", main_script)
app_module = importlib.util.module_from_spec(spec)
sys.modules["main"] = app_module
spec.loader.exec_module(app_module)

App = app_module.TraceabilityApp

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
    print(f"      -> [POPUP INTERCEPTED] {title}: {message}")
messagebox.showerror = mock_showerror

def run_test():
    print("Initializing App UI...")
    app_module.init_db()
    app = App()
    app.deiconify()
    app.update()
    
    print("Switching to New Entry Tab...")
    app.notebook.select(0) # New Entry
    app.update()
    time.sleep(1)
    
    def test_barcode(description, barcode, should_succeed):
        print(f"\n=========================================")
        print(f"[*] TEST: {description}")
        print(f"[*] Scanning Barcode: '{barcode}'")
        
        # Clear
        app.clear_form()
        app.update()
        
        # Insert barcode
        app.var_scan_rm_t1.set(barcode)
        app.update()
        time.sleep(0.5)
        
        # Trigger
        app.on_rm_scanned_t1(None)
        app.update()
        time.sleep(1)
        
        # Check results
        matched_sf_pn = app.cb_sf_pn.get()
        matched_sf_name = app.var_part_sf.get()
        
        print(f"DEBUG SF_DATA size: {len(app_module.SF_DATA)}")
        if not matched_sf_pn:
            print("DEBUG: Checking SF_DATA keys and first RM:")
            for k, v in list(app_module.SF_DATA.items())[:2]:
                print(f"   {k} -> RM1: {v[1][0][0] if v[1] else 'None'}")
        
        if matched_sf_pn:
            print(f"      -> Matched SF PN: {matched_sf_pn} ({matched_sf_name})")
            if should_succeed:
                print("   [PASS] Successfully populated the correct data.")
            else:
                print("   [FAIL] Populated data when it should have been rejected!")
        else:
            print(f"      -> Matched SF PN: (Empty)")
            if not should_succeed:
                print("   [PASS] Successfully rejected the invalid barcode.")
            else:
                print("   [FAIL] Failed to populate data for a valid barcode!")

    # Test 1: Valid Primary Component
    test_barcode("Valid Primary Component", "MOCK-A01-10001-Y-PRT", should_succeed=True)
    time.sleep(2)
    
    # Test 2: Invalid/Secondary Component
    test_barcode("Secondary Component (Should Reject)", "MOCK-A01-10002-Z-PRT", should_succeed=False)
    time.sleep(2)
    
    # Test 3: Completely Fake Barcode
    test_barcode("Completely Fake Barcode (Should Reject)", "INVALID-12345", should_succeed=False)

    print("\n=========================================")
    print("All tests completed.")
    print("The app will close automatically in 5 seconds...")
    
    app.after(5000, app.destroy)
    app.mainloop()

if __name__ == "__main__":
    run_test()
