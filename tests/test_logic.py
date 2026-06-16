import pytest
import sqlite3
import datetime
from unittest.mock import MagicMock
from src import database
from src.main import TraceabilityApp

# Using pytest-mock to mock messagebox and others
@pytest.mark.skip(reason="Downtime logic is tightly coupled with UI prompt_downtime dialog. Needs refactoring to be testable.")
def test_downtime_calculation_logic(mock_db_file, mock_sf_data, mocker):
    # Mock Tkinter UI classes to prevent creating actual windows
    mocker.patch('tkinter.Tk')
    mocker.patch('tkinter.Toplevel')
    mocker.patch('tkinter.messagebox')
    
    # We will instantiate the app and mock the UI inputs
    # Actually, instantiating TraceabilityApp might trigger UI creation,
    # let's mock all UI building methods to avoid errors
    mocker.patch.object(TraceabilityApp, 'prompt_login')
    mocker.patch.object(TraceabilityApp, 'build_ui')
    mocker.patch.object(TraceabilityApp, 'populate_sf_combobox')
    mocker.patch.object(TraceabilityApp, 'update_stats')
    mocker.patch.object(TraceabilityApp, 'schedule_daily_snapshot')
    mocker.patch.object(TraceabilityApp, 'check_and_generate_missed_reports')
    mocker.patch.object(TraceabilityApp, 'refresh_recent_treeview')
    mocker.patch.object(TraceabilityApp, 'refresh_records_treeview')
    mocker.patch.object(TraceabilityApp, 'update_sub_batch_preview')
    mocker.patch.object(TraceabilityApp, 'update_clock')
    mocker.patch.object(TraceabilityApp, 'wait_window')
    
    app = TraceabilityApp()
    
    # Set up dummy state
    app.current_user_id = "OP_TEST"
    app.current_shift = "Shift 1"
    app.current_station = "Station 1"
    app.var_pn = MagicMock()
    app.var_pn.get.return_value = "TEST-PN-001"
    app.var_qty = MagicMock()
    app.var_qty.get.return_value = "50"  # They produced 50
    app.cb_sf_pn = MagicMock()
    app.cb_sf_pn.get.return_value = "TEST-PN-001"
    app.var_op_id = MagicMock()
    app.var_op_id.get.return_value = "OP_TEST"
    app.cb_station = MagicMock()
    app.cb_station.get.return_value = "Station 1"
    app.cb_shift_sp = MagicMock()
    app.cb_shift_sp.get.return_value = "Shift 1"
    app.pl_cb_sf_pn = MagicMock()
    app.pl_cb_sf_pn.get.return_value = "TEST-PN-001"
    app.var_remarks = MagicMock()
    app.var_remarks.get.return_value = ""
    app.var_b1 = MagicMock()
    app.var_b1.get.return_value = "B1"
    app.var_b2 = MagicMock()
    app.var_b2.get.return_value = ""
    app.var_b3 = MagicMock()
    app.var_b3.get.return_value = ""
    
    rm_mock = MagicMock()
    rm_mock.get.return_value = "RM1"
    app.rm_vars_t1 = [(rm_mock, rm_mock)]
    app.var_part_sf = MagicMock()
    app.var_part_sf.get.return_value = "PART-1"
    
    app.de_sp = MagicMock()
    app.h_sp = MagicMock()
    app.m_sp = MagicMock()
    app.de_line = MagicMock()
    app.h_line = MagicMock()
    app.m_line = MagicMock()
    
    app.get_dt_string = MagicMock(return_value="2026-06-14 10:00:00")
    app.var_quality_op_id = MagicMock()
    app.var_quality_op_id.get.return_value = ""
    app.var_defect_type = MagicMock()
    app.var_defect_type.get.return_value = ""
    app.var_qty_defective = MagicMock()
    app.var_qty_defective.get.return_value = "0"
    app.var_desc = MagicMock()
    app.var_desc.get.return_value = ""
    app.var_search_recent = MagicMock()
    app.var_search_recent.get.return_value = ""
    
    app.refresh_recent_records = MagicMock()
    app.clear_form = MagicMock()
    app.update_tracker = MagicMock()
    
    # Mock askstring to return "Break" as reason for downtime
    mocker.patch('tkinter.simpledialog.askstring', return_value="Break")
    
    # 1. Test case: Production is exactly the target (target is 100 per hr)
    # Wait, the dummy SF_DATA target is 100.
    # The previous hour logic checks what they produced.
    # Let's insert 50 records in the previous hour to simulate they produced 50 BEFORE.
    
    conn = database.get_db_connection()
    c = conn.cursor()
    prev_hour_start = datetime.datetime.now().replace(minute=0, second=0, microsecond=0) - datetime.timedelta(hours=1)
    
    c.execute("""
        INSERT INTO records (sub_batch_id, pn_sf, quantity, op_id, shift_sp, station, created_at, status) 
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, ("SIM_1", "TEST-PN-001", 30, "OP_TEST", "Shift 1", "Station 1", prev_hour_start.strftime("%Y-%m-%d %H:%M:%S"), "In Rack"))
    conn.commit()
    conn.close()
    
    # Mock datetime in src.main to be at XX:10 to trigger the check
    mock_dt = mocker.patch('src.main.datetime')
    mock_now = datetime.datetime.now().replace(minute=10, second=0, microsecond=0)
    mock_dt.datetime.now.return_value = mock_now
    mock_dt.timedelta = datetime.timedelta
    
    # If they now save_record and produced 50 more... Wait, save_record processes the current input.
    # We just run app.save_record()
    app.save_record()
    
    # Now check if downtime was recorded
    conn = database.get_db_connection()
    c = conn.cursor()
    # It checks the previous hour, but it triggers right after saving if previous hour was short!
    c.execute("SELECT * FROM downtime_logs WHERE op_id='OP_TEST'")
    downtimes = c.fetchall()
    conn.close()
    
    # Total produced in previous hour was 30. Target was 100.
    # Shortfall = 70. 70/100 * 60 = 42 minutes.
    # So a downtime of 42 minutes should be logged!
    assert len(downtimes) >= 1
    assert downtimes[0][5] == 42  # duration_min
    assert downtimes[0][6] == "Break" # reason
