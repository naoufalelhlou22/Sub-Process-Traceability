import pytest
import sqlite3
from src import database

def test_hash_and_verify_password():
    password = "secure_password_123"
    hashed = database.hash_password(password)
    
    # Check that hashed password is not plaintext
    assert hashed != password
    # Check that it verifies correctly
    assert database.verify_password(password, hashed) is True
    # Check that a wrong password fails
    assert database.verify_password("wrong_password", hashed) is False

def test_init_db_creates_tables(mock_db_file):
    # init_db is called in the conftest.py fixture mock_db_file
    conn = database.get_db_connection()
    c = conn.cursor()
    
    # Verify records table exists
    c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='records'")
    assert c.fetchone() is not None
    
    # Verify products table exists
    c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='products'")
    assert c.fetchone() is not None
    
    # Verify auth table exists
    c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='auth'")
    assert c.fetchone() is not None
    
    conn.close()

def test_load_sf_data(mock_db_file):
    # Insert a dummy product
    conn = database.get_db_connection()
    c = conn.cursor()
    c.execute(
        "INSERT INTO products (pn_sf, name_sf, rms_json, std_box_qty, std_hourly_target) VALUES (?, ?, ?, ?, ?)",
        ("TEST-1", "Test Part", '[]', 50, 100)
    )
    conn.commit()
    conn.close()
    
    # Call load_sf_data
    database.load_sf_data()
    
    # Verify SF_DATA is populated correctly
    assert "TEST-1" in database.SF_DATA
    data = database.SF_DATA["TEST-1"]
    assert data[0] == "Test Part"
    assert data[1] == []  # parsed JSON
    assert data[2] == 50
    assert data[3] == 100
