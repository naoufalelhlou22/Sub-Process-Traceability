import sys
import os
sys.path.append(os.path.dirname(__file__))
from main import get_db_connection

conn = get_db_connection()
c = conn.cursor()
c.execute("SELECT id, sub_batch_id, status, action_type, qty_defective, CASE WHEN status='Closed' AND action_type IN ('Rework', 'Sorting', 'Use As-Is') THEN 0 ELSE qty_defective END FROM quality_defects")
print(c.fetchall())
conn.close()
