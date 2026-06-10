import os
import sqlite3
import datetime
from fpdf import FPDF

import sys

def persistent_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    return os.path.join(base_path, relative_path)

DB_FILE = os.path.join(persistent_path("data"), "traceability.db")
def get_db_connection():
    conn = sqlite3.connect(DB_FILE, check_same_thread=False, timeout=10)
    conn.execute('PRAGMA journal_mode=WAL')
    return conn

class PDFReport(FPDF):
    def __init__(self, shift_label, start_dt, end_dt):
        super().__init__()
        self.shift_label = shift_label
        self.start_dt = start_dt
        self.end_dt = end_dt
        
    def header(self):
        self.set_font('helvetica', 'B', 15)
        self.cell(0, 10, 'Sub-Process Traceability', border=0, align='C', new_x="LMARGIN", new_y="NEXT")
        self.set_font('helvetica', 'B', 12)
        self.cell(0, 10, f'End of Shift Report ({self.shift_label})', border=0, align='C', new_x="LMARGIN", new_y="NEXT")
        
        self.set_font('helvetica', '', 10)
        self.cell(0, 5, f'Shift Period: {self.start_dt} to {self.end_dt}', align='C', new_x="LMARGIN", new_y="NEXT")
        self.cell(0, 5, f'Generated On: {datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}', align='C', new_x="LMARGIN", new_y="NEXT")
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font('helvetica', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}/{{nb}}', align='C')


def generate_shift_pdf_report(start_dt, end_dt, shift_label):
    conn = get_db_connection()
    c = conn.cursor()
    
    # B. Production Summary
    c.execute('''SELECT COUNT(*), SUM(quantity) 
                 FROM records 
                 WHERE dt_sp >= ? AND dt_sp < ?''', (start_dt, end_dt))
    total_boxes, total_qty = c.fetchone()
    total_boxes = total_boxes or 0
    total_qty = total_qty or 0
    
    # C. Production Mix
    c.execute('''SELECT pn_sf, part_sf, SUM(quantity) 
                 FROM records 
                 WHERE dt_sp >= ? AND dt_sp < ?
                 GROUP BY pn_sf, part_sf ORDER BY SUM(quantity) DESC''', (start_dt, end_dt))
    production_mix = c.fetchall()
    
    # D. Quality Summary
    c.execute('''SELECT defect_type, COUNT(*), SUM(qty_defective)
                 FROM quality_defects
                 WHERE reported_at >= ? AND reported_at < ?
                 GROUP BY defect_type''', (start_dt, end_dt))
    defects = c.fetchall()
    
    # E. Operator Summary
    c.execute('''SELECT op_id, SUM(quantity) 
                 FROM records 
                 WHERE dt_sp >= ? AND dt_sp < ?
                 GROUP BY op_id ORDER BY SUM(quantity) DESC''', (start_dt, end_dt))
    operators = c.fetchall()
    
    conn.close()
    
    if total_boxes == 0 and len(defects) == 0:
        print(f"No data for shift {start_dt} to {end_dt}. Skipping PDF generation.")
        return None
        
    pdf = PDFReport(shift_label, start_dt, end_dt)
    pdf.add_page()
    
    # Summary
    pdf.set_font("helvetica", "B", 12)
    pdf.cell(0, 8, "1. Production Summary", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("helvetica", "", 10)
    pdf.cell(0, 6, f"Total Boxes Produced: {total_boxes}", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 6, f"Total Quantity Produced: {total_qty}", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)
    
    # Production Mix Table
    pdf.set_font("helvetica", "B", 12)
    pdf.cell(0, 8, "2. Production Mix", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("helvetica", "B", 10)
    
    pdf.cell(60, 8, "Part Number", border=1)
    pdf.cell(100, 8, "Part Name", border=1)
    pdf.cell(30, 8, "Total Qty", border=1, new_x="LMARGIN", new_y="NEXT")
    
    pdf.set_font("helvetica", "", 10)
    for row in production_mix:
        pn, name, qty = row
        pdf.cell(60, 8, str(pn)[:25], border=1)
        pdf.cell(100, 8, str(name)[:45], border=1)
        pdf.cell(30, 8, str(qty), border=1, new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)
    
    # Quality Summary
    pdf.set_font("helvetica", "B", 12)
    pdf.cell(0, 8, "3. Quality & Defects Summary", new_x="LMARGIN", new_y="NEXT")
    
    if defects:
        pdf.set_font("helvetica", "B", 10)
        pdf.cell(90, 8, "Defect Type", border=1)
        pdf.cell(50, 8, "Incidents", border=1)
        pdf.cell(50, 8, "Total Defective Qty", border=1, new_x="LMARGIN", new_y="NEXT")
        
        pdf.set_font("helvetica", "", 10)
        for row in defects:
            dtype, count, qty = row
            pdf.cell(90, 8, str(dtype)[:40], border=1)
            pdf.cell(50, 8, str(count), border=1)
            pdf.cell(50, 8, str(qty), border=1, new_x="LMARGIN", new_y="NEXT")
    else:
        pdf.set_font("helvetica", "I", 10)
        pdf.cell(0, 6, "No defects reported during this shift.", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)
    
    # Operator Summary
    pdf.set_font("helvetica", "B", 12)
    pdf.cell(0, 8, "4. Operator Output Summary", new_x="LMARGIN", new_y="NEXT")
    
    if operators:
        pdf.set_font("helvetica", "B", 10)
        pdf.cell(60, 8, "Operator ID", border=1)
        pdf.cell(60, 8, "Total Qty Produced", border=1, new_x="LMARGIN", new_y="NEXT")
        
        pdf.set_font("helvetica", "", 10)
        for row in operators:
            op_id, qty = row
            pdf.cell(60, 8, str(op_id), border=1)
            pdf.cell(60, 8, str(qty), border=1, new_x="LMARGIN", new_y="NEXT")
    
    # Output file
    try:
        start_date = start_dt.split(" ")[0]
        report_dir = os.path.join(persistent_path("reports"), start_date)
        os.makedirs(report_dir, exist_ok=True)
        
        shift_safe = shift_label.replace(":", "-").replace(" ", "_").replace("/", "-")
        filename = f"Shift_Report_{shift_safe}.pdf"
        filepath = os.path.join(report_dir, filename)
        
        pdf.output(filepath)
        print(f"Generated PDF Report: {filepath}")
        return filepath
    except Exception as e:
        print(f"Failed to generate PDF: {e}")
        return None
