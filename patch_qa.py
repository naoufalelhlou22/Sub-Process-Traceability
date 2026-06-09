import os
import re

with open('src/quality_app.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Add hash_password function
hash_func = '''
import hashlib
import binascii

def hash_password(password):
    salt = b"subproc_trace_salt_2026"
    hash_obj = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 100000)
    return binascii.hexlify(hash_obj).decode("utf-8")
'''

content = re.sub(r'(def get_db_connection.*?\n\s+return conn\n)', r'\1' + hash_func, content, flags=re.DOTALL, count=1)

# Update seed
seed_old = '''c.executemany(
            "INSERT OR IGNORE INTO auth (id, password, role) VALUES (?,?,?)",
            [('Q001', 'quality001', 'Quality OP')])'''
seed_new = '''c.executemany(
            "INSERT OR IGNORE INTO auth (id, password, role) VALUES (?,?,?)",
            [('Q001', hash_password('quality001'), 'Quality OP')])'''
content = content.replace(seed_old, seed_new)

# Update login verification
login_old = 'c.execute("SELECT role FROM auth WHERE id=? AND password=?",\n                                  (user_id, password))'
login_new = 'c.execute("SELECT role FROM auth WHERE id=? AND password=?",\n                                  (user_id, hash_password(password)))'
content = content.replace(login_old, login_new)

with open('src/quality_app.py', 'w', encoding='utf-8') as f:
    f.write(content)

print("Patch applied to quality_app.py")
