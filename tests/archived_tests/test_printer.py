import tkinter as tk
from tkinter import ttk, messagebox
import win32print
import os

# Your ZPL data
ZPL_DATA = r"""^XA
^PW520
^LH0,0
^CI28
^FO30,20^A0N,24,24^FB460,1,0,C^FDSUB-PROCESS RECORD^FS
^FO30,55^GB460,3,3^FS
^FO30,70^GB460,40,40^FS
^FO30,80^A0N,24,24^FR^FB460,1,0,C^FDSB202606112038S06A01\&^FS
^BY2^FO50,120^BCN,50,N,N,N,A^FDSB202606112038S06A01^FS
^FO30,185^A0N,18,18^FDPN Semi fini^FS
^FO30,185^A0N,18,18^FB460,1,0,R^FDMOCK-A01-10003-X-SUB\&^FS
^FO30,205^A0N,18,18^FDPart Name (SF)^FS
^FO30,205^A0N,18,18^FB460,1,0,R^FDAlpha Subsystem Aux Right\&^FS
^FO30,228^GB460,1,1^FS
^FO30,236^A0N,18,18^FDBatch No. 1^FS
^FO30,236^A0N,18,18^FB460,1,0,R^FD-\&^FS
^FO30,256^A0N,18,18^FDBatch No. 2^FS
^FO30,256^A0N,18,18^FB460,1,0,R^FDHUIK\&^FS
^FO30,276^A0N,18,18^FDBatch No. 3^FS
^FO30,276^A0N,18,18^FB460,1,0,R^FD-\&^FS
^FO30,296^A0N,20,20^FDQuantity^FS
^FO30,296^A0N,20,20^FB460,1,0,R^FD5000 pcs\&^FS
^FO30,321^GB460,1,1^FS
^FO30,329^A0N,18,18^FDShift SP^FS
^FO30,329^A0N,18,18^FB460,1,0,R^FDA\&^FS
^FO30,349^A0N,18,18^FDOp ID^FS
^FO30,349^A0N,18,18^FB460,1,0,R^FD111\&^FS
^FO30,369^A0N,18,18^FDStation^FS
^FO30,369^A0N,18,18^FB460,1,0,R^FDS06\&^FS
^FO30,389^A0N,18,18^FDSP Date/Time^FS
^FO30,389^A0N,18,18^FB460,1,0,R^FD2026-05-11 20:38\&^FS
^FO30,409^A0N,18,18^FDLine Entry^FS
^FO30,409^A0N,18,18^FB460,1,0,R^FD-\&^FS
^FO30,429^A0N,18,18^FDPrinted At^FS
^FO30,429^A0N,18,18^FB460,1,0,R^FD2026-05-11 20:41\&^FS
^FO30,452^GB460,2,2^FS
^FO30,460^A0N,14,14^FB460,1,0,C^FD** REPRINT #1 by OP ID: 1 at 2026-06-11 20:41 **\&^FS
^FO160,478^BQN,2,2^FDMA,{"sub_batch_id": "SB202606112038S06A01", "full_pn_sf": "MOCK-A01-10003-X-SUB", "part_name_sf": "Alpha Subsystem Aux Right", "rm1_pn": "MOCK-A01-10003-Y-PRT", "rm1_name": "Alpha Core (Auxiliary) Right", "rm2_pn": "MOCK-A01-10002-Z-PRT", "rm2_name": "Alpha Buffer Right", "rm3_pn": "", "rm3_name": "", "rm4_pn": "", "rm4_name": "", "batch_no_1": "", "batch_no_2": "HUIK", "batch_no_3": "", "quantity": 5000, "sub_process_shift": "A", "op_id": "111", "station": "S06", "sub_process_datetime": "2026-06-11 20:38", "production_line_entry_datetime": "", "production_line_shift": "", "remarks": ""}^FS
^FO30,728^A0N,20,20^FB460,1,0,C^FD- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -^FS
^FO30,818^A0N,24,24^FB460,1,0,C^FDSUB-PROCESS RECORD^FS
^FO30,853^GB460,3,3^FS
^FO30,868^GB460,40,40^FS
^FO30,878^A0N,24,24^FR^FB460,1,0,C^FDSB202606112038S06A01\&^FS
^BY2^FO50,918^BCN,50,N,N,N,A^FDSB202606112038S06A01^FS
^FO30,983^A0N,18,18^FDPN Semi fini^FS
^FO30,983^A0N,18,18^FB460,1,0,R^FDMOCK-A01-10003-X-SUB\&^FS
^FO30,1003^A0N,18,18^FDPart Name (SF)^FS
^FO30,1003^A0N,18,18^FB460,1,0,R^FDAlpha Subsystem Aux Right\&^FS
^FO30,1026^GB460,1,1^FS
^FO30,1034^A0N,18,18^FDBatch No. 1^FS
^FO30,1034^A0N,18,18^FB460,1,0,R^FD-\&^FS
^FO30,1054^A0N,18,18^FDBatch No. 2^FS
^FO30,1054^A0N,18,18^FB460,1,0,R^FDHUIK\&^FS
^FO30,1074^A0N,18,18^FDBatch No. 3^FS
^FO30,1074^A0N,18,18^FB460,1,0,R^FD-\&^FS
^FO30,1094^A0N,20,20^FDQuantity^FS
^FO30,1094^A0N,20,20^FB460,1,0,R^FD5000 pcs\&^FS
^FO30,1119^GB460,1,1^FS
^FO30,1127^A0N,18,18^FDShift SP^FS
^FO30,1127^A0N,18,18^FB460,1,0,R^FDA\&^FS
^FO30,1147^A0N,18,18^FDOp ID^FS
^FO30,1147^A0N,18,18^FB460,1,0,R^FD111\&^FS
^FO30,1167^A0N,18,18^FDStation^FS
^FO30,1167^A0N,18,18^FB460,1,0,R^FDS06\&^FS
^FO30,1187^A0N,18,18^FDSP Date/Time^FS
^FO30,1187^A0N,18,18^FB460,1,0,R^FD2026-05-05 20:38\&^FS
^FO30,1207^A0N,18,18^FDLine Entry^FS
^FO30,1207^A0N,18,18^FB460,1,0,R^FD-\&^FS
^FO30,1227^A0N,18,18^FDPrinted At^FS
^FO30,1227^A0N,18,18^FB460,1,0,R^FD2026-05-05 20:41\&^FS
^FO30,1250^GB460,2,2^FS
^FO30,1258^A0N,14,14^FB460,1,0,C^FD** REPRINT #1 by OP ID: 1 at 2026-05-11 20:41 **\&^FS
^FO160,1276^BQN,2,2^FDMA,{"sub_batch_id": "SB202606112038S06A01", "full_pn_sf": "MOCK-A01-10003-X-SUB", "part_name_sf": "Alpha Subsystem Aux Right", "rm1_pn": "MOCK-A01-10003-Y-PRT", "rm1_name": "Alpha Core (Auxiliary) Right", "rm2_pn": "MOCK-A01-10002-Z-PRT", "rm2_name": "Alpha Buffer Right", "rm3_pn": "", "rm3_name": "", "rm4_pn": "", "rm4_name": "", "batch_no_1": "", "batch_no_2": "HUIK", "batch_no_3": "", "quantity": 5000, "sub_process_shift": "A", "op_id": "111", "station": "S06", "sub_process_datetime": "2026-06-11 20:38", "production_line_entry_datetime": "", "production_line_shift": "", "remarks": ""}^FS
^PQ1
^XZ"""

def get_printers():
    printers = [printer[2] for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
    return printers

def save_and_print():
    selected_printer = printer_combo.get()
    if not selected_printer:
        messagebox.showwarning("Warning", "Please select a printer first!")
        return
    
    # 1. Save to file
    try:
        with open("label_data.zpl", "w", encoding="utf-8") as f:
            f.write(ZPL_DATA)
    except Exception as e:
        messagebox.showerror("File Error", f"Could not save file: {str(e)}")
        return

    # 2. Print
    try:
        hPrinter = win32print.OpenPrinter(selected_printer)
        try:
            hJob = win32print.StartDocPrinter(hPrinter, 1, ("Label Print", None, "RAW"))
            win32print.StartPagePrinter(hPrinter)
            win32print.WritePrinter(hPrinter, ZPL_DATA.encode('utf-8'))
            win32print.EndPagePrinter(hPrinter)
            win32print.EndDocPrinter(hPrinter)
            messagebox.showinfo("Success", f"Label saved and sent to {selected_printer}")
        finally:
            win32print.ClosePrinter(hPrinter)
    except Exception as e:
        messagebox.showerror("Print Error", f"Failed to print: {str(e)}")

# GUI setup
root = tk.Tk()
root.title("ZPL Printer Selector")
root.geometry("350x200")

tk.Label(root, text="Select Thermal Printer:").pack(pady=10)

printer_combo = ttk.Combobox(root, values=get_printers(), width=40)
printer_combo.pack(pady=5)

print_btn = tk.Button(root, text="Save & Print Label", command=save_and_print, bg="blue", fg="white")
print_btn.pack(pady=20)

root.mainloop()