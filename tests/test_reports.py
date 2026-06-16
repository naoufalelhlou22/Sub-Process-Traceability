import os
import pytest
import datetime
from src import database
from src import report_generator

def test_generate_daily_pdf_report(mock_db_file, tmp_path, monkeypatch):
    # Monkeypatch the reports directory to be the temporary pytest path
    def mock_persistent_path(folder_name):
        return str(tmp_path / folder_name)
    
    monkeypatch.setattr('src.report_generator.persistent_path', mock_persistent_path)
    
    # Insert dummy data into records
    conn = database.get_db_connection()
    c = conn.cursor()
    c.execute(
        "INSERT INTO products (pn_sf, name_sf, rms_json, std_box_qty, std_hourly_target) VALUES (?, ?, ?, ?, ?)",
        ("TEST-PN", "Test Part", '[]', 50, 100)
    )
    database.load_sf_data()
    
    # Insert some dummy records to be aggregated
    test_dt = datetime.datetime.now()
    dt_str = test_dt.strftime("%Y-%m-%d %H:%M:%S")
    
    c.execute("""
        INSERT INTO records (sub_batch_id, pn_sf, quantity, op_id, shift_sp, station, dt_sp, status, created_at) 
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, ("B1", "TEST-PN", 50, "OP1", "Shift 1", "Station 1", dt_str, "Consumed", dt_str))
    c.execute("""
        INSERT INTO downtime_logs (created_at, shift, station, op_id, duration_min, reason) 
        VALUES (?, ?, ?, ?, ?, ?)
    """, (dt_str, "Shift 1", "Station 1", "OP1", 15, "001"))
    conn.commit()
    conn.close()

    start_dt = test_dt.replace(hour=0, minute=0, second=0)
    end_dt = test_dt.replace(hour=23, minute=59, second=59)
    
    # Generate report
    report_generator.generate_daily_pdf_report(start_dt, end_dt)
    
    # Verify PDF exists
    target_date = start_dt.strftime("%Y-%m-%d")
    reports_dir = tmp_path / "reports" / target_date
    expected_file = reports_dir / f"Report_{target_date}_Daily.pdf"
    
    assert expected_file.exists()
    assert expected_file.stat().st_size > 0
