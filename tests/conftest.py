import os
import pytest
import sqlite3

# Import database module to monkeypatch DB_FILE
from src import database
from src import config

@pytest.fixture(autouse=True)
def mock_db_file(tmp_path, monkeypatch):
    """
    Override the database file path to a temporary file for all tests.
    This ensures we don't write to the real production database.
    """
    db_path = tmp_path / "test_traceability.db"
    monkeypatch.setattr(database, 'DB_FILE', str(db_path))
    monkeypatch.setattr(config, 'DB_FILE', str(db_path))
    
    import src.quality_app as quality_app
    monkeypatch.setattr(quality_app, 'DB_FILE', str(db_path))
    
    # Initialize the database schema
    database.init_db()
    
    # Clear SF_DATA to ensure a clean state
    database.SF_DATA.clear()
    
    yield db_path

@pytest.fixture
def mock_sf_data(monkeypatch):
    """
    Provide dummy SF_DATA for logic tests that depend on product configuration.
    """
    dummy_sf_data = {
        "TEST-PN-001": ("Test Product 1", [{"pn": "RM-1", "name": "Raw Material 1"}], 50, 100),
        "TEST-PN-002": ("Test Product 2", [], 10, 50)
    }
    monkeypatch.setattr(database, 'SF_DATA', dummy_sf_data)
    return dummy_sf_data
