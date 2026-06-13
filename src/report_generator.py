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
    def __init__(self, start_dt_str, end_dt_str):
        super().__init__()
        self.start_dt = start_dt_str
        self.end_dt = end_dt_str
        self.brand_primary = (0, 71, 143)
        self.text_dark = (50, 50, 50)
        
    def header(self):
        logo_path = persistent_path(os.path.join("assets", "logo_en.png"))
        if os.path.exists(logo_path):
            self.image(logo_path, 10, 8, 45)
            
        self.set_y(10)
        self.set_font('helvetica', 'B', 18)
        self.set_text_color(*self.brand_primary)
        self.cell(0, 8, 'Sub-Process Traceability', border=0, align='R', new_x="LMARGIN", new_y="NEXT")
        
        self.set_font('helvetica', 'B', 12)
        self.set_text_color(100, 100, 100)
        self.cell(0, 6, 'Production Report', border=0, align='R', new_x="LMARGIN", new_y="NEXT")
        
        self.set_y(32)
        self.set_font('helvetica', '', 10)
        self.set_text_color(*self.text_dark)
        self.set_fill_color(245, 245, 245)
        self.cell(0, 8, f' Report Period: {self.start_dt} to {self.end_dt}  |  Generated: {datetime.datetime.now().strftime("%Y-%m-%d %H:%M")}', border=0, align='L', fill=True, new_x="LMARGIN", new_y="NEXT")
        
        self.ln(5)
        self.set_draw_color(*self.brand_primary)
        self.set_line_width(0.5)
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_draw_color(200, 200, 200)
        self.set_line_width(0.2)
        self.line(10, self.get_y(), 200, self.get_y())
        self.set_y(-12)
        self.set_font('helvetica', 'I', 8)
        self.set_text_color(120, 120, 120)
        self.cell(100, 10, 'HI-LEX ACT - Sub-Process Traceability', align='L')
        self.set_x(10)
        self.cell(0, 10, f'Page {self.page_no()}/{{nb}}', align='R')

def check_section_break(pdf, required_height):
    if pdf.get_y() + required_height > 270:
        pdf.add_page()

def print_table_header(pdf, cols):
    pdf.set_font("helvetica", "B", 10)
    pdf.set_fill_color(0, 71, 143)
    pdf.set_text_color(255, 255, 255)
    pdf.set_draw_color(255, 255, 255)
    for width, label in cols:
        pdf.cell(width, 8, label, border=1, align='C', fill=True)
    pdf.ln()
    pdf.set_text_color(50, 50, 50)
    pdf.set_draw_color(200, 200, 200)

def generate_shift_pdf_report(target_date, shift_db_value):
    shift_label = f"Shift {shift_db_value}"
    conn = get_db_connection()
    c = conn.cursor()

    c.execute('''SELECT COUNT(*), SUM(quantity) 
                 FROM records 
                 WHERE CASE WHEN CAST(substr(dt_sp, 12, 2) AS INTEGER) < 6 THEN date(substr(dt_sp, 1, 10), '-1 day') ELSE substr(dt_sp, 1, 10) END = ? AND shift_sp = ?''', (target_date, shift_db_value))
    total_boxes, total_qty = c.fetchone()
    total_boxes = total_boxes or 0
    total_qty = total_qty or 0
    
    c.execute('''SELECT pn_sf, part_sf, SUM(quantity) 
                 FROM records 
                 WHERE CASE WHEN CAST(substr(dt_sp, 12, 2) AS INTEGER) < 6 THEN date(substr(dt_sp, 1, 10), '-1 day') ELSE substr(dt_sp, 1, 10) END = ? AND shift_sp = ?
                 GROUP BY pn_sf, part_sf ORDER BY SUM(quantity) DESC''', (target_date, shift_db_value))
    production_mix = c.fetchall()
    
    c.execute('''SELECT defect_type, COUNT(*), SUM(qty_defective)
                 FROM quality_defects
                 WHERE CASE WHEN CAST(substr(reported_at, 12, 2) AS INTEGER) < 6 THEN date(substr(reported_at, 1, 10), '-1 day') ELSE substr(reported_at, 1, 10) END = ? AND shift = ?
                 GROUP BY defect_type''', (target_date, shift_db_value))
    defects = c.fetchall()
    
    c.execute('''SELECT op_id, SUM(quantity) 
                 FROM records 
                 WHERE CASE WHEN CAST(substr(dt_sp, 12, 2) AS INTEGER) < 6 THEN date(substr(dt_sp, 1, 10), '-1 day') ELSE substr(dt_sp, 1, 10) END = ? AND shift_sp = ?
                 GROUP BY op_id ORDER BY SUM(quantity) DESC''', (target_date, shift_db_value))
    operators = c.fetchall()
    
    c.execute('''SELECT r.pn_sf, SUM(r.quantity), 
                 (SELECT target_qty FROM shift_targets WHERE product_pn = r.pn_sf ORDER BY effective_date DESC LIMIT 1)
                 FROM records r 
                 WHERE CASE WHEN CAST(substr(dt_sp, 12, 2) AS INTEGER) < 6 THEN date(substr(dt_sp, 1, 10), '-1 day') ELSE substr(dt_sp, 1, 10) END = ? AND shift_sp = ?
                 GROUP BY r.pn_sf ORDER BY SUM(r.quantity) DESC''', (target_date, shift_db_value))
    targets_actual = c.fetchall()
    
    c.execute('''SELECT pn_sf, SUM(quantity), COUNT(id)
                 FROM records
                 WHERE status="In Rack"
                 GROUP BY pn_sf ORDER BY SUM(quantity) DESC LIMIT 15''')
    inventory_wip = c.fetchall()
    
    c.execute('''SELECT rm1_pn, rm1_name, COUNT(DISTINCT batch1)
                 FROM records
                 WHERE CASE WHEN CAST(substr(dt_sp, 12, 2) AS INTEGER) < 6 THEN date(substr(dt_sp, 1, 10), '-1 day') ELSE substr(dt_sp, 1, 10) END = ? AND shift_sp = ? AND rm1_pn IS NOT NULL AND rm1_pn != ""
                 GROUP BY rm1_pn, rm1_name
                 ORDER BY COUNT(DISTINCT batch1) DESC''', (target_date, shift_db_value))
    rm_usage = c.fetchall()
    
    c.execute('''SELECT op_id, GROUP_CONCAT(DISTINCT reason), SUM(duration_min)
                 FROM downtime_logs
                 WHERE CASE WHEN CAST(substr(created_at, 12, 2) AS INTEGER) < 6 THEN date(substr(created_at, 1, 10), '-1 day') ELSE substr(created_at, 1, 10) END = ? AND shift = ?
                 GROUP BY op_id
                 ORDER BY SUM(duration_min) DESC''', (target_date, shift_db_value))
    downtime_summary = c.fetchall()
    
    conn.close()
    
    # if total_boxes == 0 and len(defects) == 0 and len(downtime_summary) == 0:
    #     print(f"No data for {target_date} {shift_label}. Skipping PDF generation.")
    #     return None
        
    start_dt_str = f"{target_date} {shift_label} Start"
    end_dt_str = f"{target_date} {shift_label} End"
    return _build_pdf(target_date, shift_label, start_dt_str, end_dt_str, total_boxes, total_qty, production_mix, defects, operators, targets_actual, inventory_wip, rm_usage, downtime_summary)


def generate_daily_pdf_report(start_dt, end_dt):
    target_date = start_dt.strftime("%Y-%m-%d")
    shift_label = "All Shifts"
    
    conn = get_db_connection()
    c = conn.cursor()

    c.execute('''SELECT COUNT(*), SUM(quantity) 
                 FROM records 
                 WHERE CASE WHEN CAST(substr(dt_sp, 12, 2) AS INTEGER) < 6 THEN date(substr(dt_sp, 1, 10), '-1 day') ELSE substr(dt_sp, 1, 10) END = ?''', (target_date,))
    total_boxes, total_qty = c.fetchone()
    total_boxes = total_boxes or 0
    total_qty = total_qty or 0
    
    c.execute('''SELECT pn_sf, part_sf, SUM(quantity) 
                 FROM records 
                 WHERE CASE WHEN CAST(substr(dt_sp, 12, 2) AS INTEGER) < 6 THEN date(substr(dt_sp, 1, 10), '-1 day') ELSE substr(dt_sp, 1, 10) END = ?
                 GROUP BY pn_sf, part_sf ORDER BY SUM(quantity) DESC''', (target_date,))
    production_mix = c.fetchall()
    
    c.execute('''SELECT defect_type, COUNT(*), SUM(qty_defective)
                 FROM quality_defects
                 WHERE CASE WHEN CAST(substr(reported_at, 12, 2) AS INTEGER) < 6 THEN date(substr(reported_at, 1, 10), '-1 day') ELSE substr(reported_at, 1, 10) END = ?
                 GROUP BY defect_type''', (target_date,))
    defects = c.fetchall()
    
    c.execute('''SELECT op_id, SUM(quantity) 
                 FROM records 
                 WHERE CASE WHEN CAST(substr(dt_sp, 12, 2) AS INTEGER) < 6 THEN date(substr(dt_sp, 1, 10), '-1 day') ELSE substr(dt_sp, 1, 10) END = ?
                 GROUP BY op_id ORDER BY SUM(quantity) DESC''', (target_date,))
    operators = c.fetchall()
    
    c.execute('''SELECT r.pn_sf, SUM(r.quantity), 
                 (SELECT target_qty FROM shift_targets WHERE product_pn = r.pn_sf ORDER BY effective_date DESC LIMIT 1)
                 FROM records r 
                 WHERE CASE WHEN CAST(substr(dt_sp, 12, 2) AS INTEGER) < 6 THEN date(substr(dt_sp, 1, 10), '-1 day') ELSE substr(dt_sp, 1, 10) END = ?
                 GROUP BY r.pn_sf ORDER BY SUM(r.quantity) DESC''', (target_date,))
    targets_actual = c.fetchall()
    
    c.execute('''SELECT pn_sf, SUM(quantity), COUNT(id)
                 FROM records
                 WHERE status="In Rack"
                 GROUP BY pn_sf ORDER BY SUM(quantity) DESC LIMIT 15''')
    inventory_wip = c.fetchall()
    
    c.execute('''SELECT rm1_pn, rm1_name, COUNT(DISTINCT batch1)
                 FROM records
                 WHERE CASE WHEN CAST(substr(dt_sp, 12, 2) AS INTEGER) < 6 THEN date(substr(dt_sp, 1, 10), '-1 day') ELSE substr(dt_sp, 1, 10) END = ? AND rm1_pn IS NOT NULL AND rm1_pn != ""
                 GROUP BY rm1_pn, rm1_name
                 ORDER BY COUNT(DISTINCT batch1) DESC''', (target_date,))
    rm_usage = c.fetchall()
    
    c.execute('''SELECT op_id, GROUP_CONCAT(DISTINCT reason), SUM(duration_min)
                 FROM downtime_logs
                 WHERE CASE WHEN CAST(substr(created_at, 12, 2) AS INTEGER) < 6 THEN date(substr(created_at, 1, 10), '-1 day') ELSE substr(created_at, 1, 10) END = ?
                 GROUP BY op_id
                 ORDER BY SUM(duration_min) DESC''', (target_date,))
    downtime_summary = c.fetchall()
    
    conn.close()
    
    start_dt_str = start_dt.strftime("%Y-%m-%d %H:%M")
    end_dt_str = end_dt.strftime("%Y-%m-%d %H:%M")
    
    # if total_boxes == 0 and len(defects) == 0 and len(downtime_summary) == 0:
    #     print(f"No data for {target_date}. Skipping daily PDF generation.")
    #     return None
        
    return _build_pdf(target_date, "Daily", start_dt_str, end_dt_str, total_boxes, total_qty, production_mix, defects, operators, targets_actual, inventory_wip, rm_usage, downtime_summary)


def _build_pdf(target_date, report_type_label, start_dt_str, end_dt_str, total_boxes, total_qty, production_mix, defects, operators, targets_actual, inventory_wip, rm_usage, downtime_summary):
    pdf = PDFReport(start_dt_str, end_dt_str)
    pdf.add_page()
    pdf.set_text_color(50, 50, 50)
    
    check_section_break(pdf, 30)
    pdf.set_font("helvetica", "B", 12)
    pdf.set_text_color(0, 71, 143)
    pdf.cell(0, 8, "1. Production Summary", new_x="LMARGIN", new_y="NEXT")
    pdf.set_text_color(50, 50, 50)
    pdf.set_font("helvetica", "", 10)
    pdf.cell(0, 6, f"Total Boxes Produced: {total_boxes}", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 6, f"Total Quantity Produced: {total_qty}", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)
    
    check_section_break(pdf, 40)
    pdf.set_font("helvetica", "B", 12)
    pdf.set_text_color(0, 71, 143)
    pdf.cell(0, 8, "2. Production Mix", new_x="LMARGIN", new_y="NEXT")
    
    cols = [(60, "Part Number"), (100, "Part Name"), (30, "Total Qty")]
    print_table_header(pdf, cols)
    
    pdf.set_font("helvetica", "", 10)
    for i, row in enumerate(production_mix):
        if pdf.get_y() > 265:
            pdf.add_page()
            print_table_header(pdf, cols)
            pdf.set_font("helvetica", "", 10)
            
        pdf.set_fill_color(245, 245, 245) if i % 2 == 0 else pdf.set_fill_color(255, 255, 255)
        pn, name, qty = row
        pdf.cell(60, 8, str(pn)[:25], border=1, align='C', fill=True)
        pdf.cell(100, 8, str(name)[:45], border=1, align='C', fill=True)
        pdf.cell(30, 8, str(qty), border=1, align='C', fill=True, new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)
    
    check_section_break(pdf, 40)
    pdf.set_font("helvetica", "B", 12)
    pdf.set_text_color(0, 71, 143)
    pdf.cell(0, 8, "3. Quality & Defects Summary", new_x="LMARGIN", new_y="NEXT")
    
    if defects:
        cols = [(90, "Defect Type"), (50, "Incidents"), (50, "Total Defective Qty")]
        print_table_header(pdf, cols)
        
        pdf.set_font("helvetica", "", 10)
        for i, row in enumerate(defects):
            if pdf.get_y() > 265:
                pdf.add_page()
                print_table_header(pdf, cols)
                pdf.set_font("helvetica", "", 10)
                
            pdf.set_fill_color(245, 245, 245) if i % 2 == 0 else pdf.set_fill_color(255, 255, 255)
            dtype, count, qty = row
            pdf.cell(90, 8, str(dtype)[:40], border=1, align='C', fill=True)
            pdf.cell(50, 8, str(count), border=1, align='C', fill=True)
            pdf.cell(50, 8, str(qty), border=1, align='C', fill=True, new_x="LMARGIN", new_y="NEXT")
    else:
        pdf.set_text_color(50, 50, 50)
        pdf.set_font("helvetica", "I", 10)
        pdf.cell(0, 6, "No defects reported during this day.", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)
    
    check_section_break(pdf, 40)
    pdf.set_font("helvetica", "B", 12)
    pdf.set_text_color(0, 71, 143)
    pdf.cell(0, 8, "4. Operator Output Summary", new_x="LMARGIN", new_y="NEXT")
    
    if operators:
        cols = [(95, "Operator ID"), (95, "Total Qty Produced")]
        print_table_header(pdf, cols)
        
        pdf.set_font("helvetica", "", 10)
        for i, row in enumerate(operators):
            if pdf.get_y() > 265:
                pdf.add_page()
                print_table_header(pdf, cols)
                pdf.set_font("helvetica", "", 10)
                
            pdf.set_fill_color(245, 245, 245) if i % 2 == 0 else pdf.set_fill_color(255, 255, 255)
            op_id, qty = row
            pdf.cell(95, 8, str(op_id), border=1, align='C', fill=True)
            pdf.cell(95, 8, str(qty), border=1, align='C', fill=True, new_x="LMARGIN", new_y="NEXT")
    
    check_section_break(pdf, 40)
    pdf.ln(5)
    pdf.set_font("helvetica", "B", 12)
    pdf.set_text_color(0, 71, 143)
    pdf.cell(0, 8, "5. Target vs Actual Performance", new_x="LMARGIN", new_y="NEXT")
    
    if targets_actual:
        cols = [(90, "Part Number"), (50, "Target Qty"), (50, "Actual Qty")]
        print_table_header(pdf, cols)
        pdf.set_font("helvetica", "", 10)
        for i, row in enumerate(targets_actual):
            if pdf.get_y() > 265:
                pdf.add_page()
                print_table_header(pdf, cols)
                pdf.set_font("helvetica", "", 10)
            pdf.set_fill_color(245, 245, 245) if i % 2 == 0 else pdf.set_fill_color(255, 255, 255)
            pn, act_qty, tgt_qty = row
            tgt_val = tgt_qty if tgt_qty is not None else "N/A"
            pdf.cell(90, 8, str(pn)[:35], border=1, align='C', fill=True)
            pdf.cell(50, 8, str(tgt_val), border=1, align='C', fill=True)
            pdf.cell(50, 8, str(act_qty), border=1, align='C', fill=True, new_x="LMARGIN", new_y="NEXT")
    else:
        pdf.set_text_color(50, 50, 50)
        pdf.set_font("helvetica", "I", 10)
        pdf.cell(0, 6, "No target data available for this day.", new_x="LMARGIN", new_y="NEXT")
            
    check_section_break(pdf, 40)
    pdf.ln(5)
    pdf.set_font("helvetica", "B", 12)
    pdf.set_text_color(0, 71, 143)
    pdf.cell(0, 8, "6. Current WIP Inventory (In Rack)", new_x="LMARGIN", new_y="NEXT")
    
    if inventory_wip:
        cols = [(90, "Part Number"), (50, "Total Qty in Rack"), (50, "Box Count")]
        print_table_header(pdf, cols)
        pdf.set_font("helvetica", "", 10)
        for i, row in enumerate(inventory_wip):
            if pdf.get_y() > 265:
                pdf.add_page()
                print_table_header(pdf, cols)
                pdf.set_font("helvetica", "", 10)
            pdf.set_fill_color(245, 245, 245) if i % 2 == 0 else pdf.set_fill_color(255, 255, 255)
            pn, qty, count = row
            pdf.cell(90, 8, str(pn)[:35], border=1, align='C', fill=True)
            pdf.cell(50, 8, str(qty), border=1, align='C', fill=True)
            pdf.cell(50, 8, str(count), border=1, align='C', fill=True, new_x="LMARGIN", new_y="NEXT")
    else:
        pdf.set_text_color(50, 50, 50)
        pdf.set_font("helvetica", "I", 10)
        pdf.cell(0, 6, "No inventory WIP currently in racks.", new_x="LMARGIN", new_y="NEXT")

    check_section_break(pdf, 40)
    pdf.ln(5)
    pdf.set_font("helvetica", "B", 12)
    pdf.set_text_color(0, 71, 143)
    pdf.cell(0, 8, "7. Raw Material Usage (Traceability)", new_x="LMARGIN", new_y="NEXT")
    
    if rm_usage:
        cols = [(60, "RM Part Number"), (90, "RM Name"), (40, "Batches Consumed")]
        print_table_header(pdf, cols)
        pdf.set_font("helvetica", "", 10)
        for i, row in enumerate(rm_usage):
            if pdf.get_y() > 265:
                pdf.add_page()
                print_table_header(pdf, cols)
                pdf.set_font("helvetica", "", 10)
            pdf.set_fill_color(245, 245, 245) if i % 2 == 0 else pdf.set_fill_color(255, 255, 255)
            rm_pn, rm_name, batches = row
            pdf.cell(60, 8, str(rm_pn)[:25], border=1, align='C', fill=True)
            pdf.cell(90, 8, str(rm_name)[:40], border=1, align='C', fill=True)
            pdf.cell(40, 8, str(batches), border=1, align='C', fill=True, new_x="LMARGIN", new_y="NEXT")
    else:
        pdf.set_text_color(50, 50, 50)
        pdf.set_font("helvetica", "I", 10)
        pdf.cell(0, 6, "No raw material consumption recorded.", new_x="LMARGIN", new_y="NEXT")

    check_section_break(pdf, 40)
    pdf.ln(5)
    pdf.set_font("helvetica", "B", 12)
    pdf.set_text_color(0, 71, 143)
    pdf.cell(0, 8, "8. Downtime Summary", new_x="LMARGIN", new_y="NEXT")
    
    if downtime_summary:
        cols = [(30, "Operator ID"), (110, "Downtime Reasons"), (50, "Total Duration (Mins)")]
        print_table_header(pdf, cols)
        pdf.set_font("helvetica", "", 10)
        for i, row in enumerate(downtime_summary):
            if pdf.get_y() > 265:
                pdf.add_page()
                print_table_header(pdf, cols)
                pdf.set_font("helvetica", "", 10)
            pdf.set_fill_color(245, 245, 245) if i % 2 == 0 else pdf.set_fill_color(255, 255, 255)
            op_id, reasons, duration = row
            op_id_str = str(op_id) if op_id else "-"
            reasons_str = str(reasons) if reasons else "-"
            if len(reasons_str) > 55:
                reasons_str = reasons_str[:52] + "..."
            
            pdf.cell(30, 8, op_id_str, border=1, align='C', fill=True)
            pdf.cell(110, 8, reasons_str, border=1, align='C', fill=True)
            pdf.cell(50, 8, str(duration), border=1, align='C', fill=True, new_x="LMARGIN", new_y="NEXT")
    else:
        pdf.set_text_color(50, 50, 50)
        pdf.set_font("helvetica", "I", 10)
        pdf.cell(0, 6, "No downtime recorded.", new_x="LMARGIN", new_y="NEXT")

    try:
        report_dir = os.path.join(persistent_path("reports"), target_date)
        os.makedirs(report_dir, exist_ok=True)
        
        safe_shift = report_type_label.replace(" ", "_")
        filename = f"Report_{target_date}_{safe_shift}.pdf"
        filepath = os.path.join(report_dir, filename)
        
        pdf.output(filepath)
        print(f"Generated PDF Report: {filepath}")
        return filepath
    except Exception as e:
        print(f"Failed to generate PDF: {e}")
        return None
