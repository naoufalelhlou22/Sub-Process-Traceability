import os
import re

with open('src/main.py', 'r', encoding='utf-8') as f:
    content = f.read()

helper = '''
import hashlib
import binascii

def hash_password(password):
    salt = b"subproc_trace_salt_2026"
    hash_obj = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 100000)
    return binascii.hexlify(hash_obj).decode("utf-8")

def migrate_passwords():
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("SELECT id, password FROM auth")
    rows = c.fetchall()
    for row in rows:
        uid, pwd = row
        if pwd and (len(pwd) != 64 or not all(ch in "0123456789abcdef" for ch in pwd)):
            hashed = hash_password(pwd)
            c.execute("UPDATE auth SET password = ? WHERE id = ?", (hashed, uid))
    conn.commit()
    conn.close()
'''

content = re.sub(r'(def get_db_connection.*?\n\s+return conn\n)', r'\1' + helper, content, flags=re.DOTALL, count=1)

content = content.replace('"SELECT role FROM auth WHERE id = ? AND password = ?", (uid, upass)', '"SELECT role FROM auth WHERE id = ? AND password = ?", (uid, hash_password(upass))')
content = content.replace('"INSERT INTO auth (id, password, role) VALUES (?, ?, ?)", (uid, upass, role)', '"INSERT INTO auth (id, password, role) VALUES (?, ?, ?)", (uid, hash_password(upass), role)')

seed_old = '''c.executemany("INSERT OR IGNORE INTO auth (id, password, role) VALUES (?, ?, ?)", [
        ("admin", "admin123", "Admin"),
        ("TL01", "pass123", "Supervisor"),
        ("OP01", "1234", "Operator")
    ])'''
seed_new = '''c.executemany("INSERT OR IGNORE INTO auth (id, password, role) VALUES (?, ?, ?)", [
        ("admin", hash_password("admin123"), "Admin"),
        ("TL01", hash_password("pass123"), "Supervisor"),
        ("OP01", hash_password("1234"), "Operator")
    ])'''
content = content.replace(seed_old, seed_new)

# Insert migrate_passwords at end of init_db
content = content.replace('    except Exception as e:\n        print("Auth seed error:", e)', '    except Exception as e:\n        print("Auth seed error:", e)\n    \n    migrate_passwords()')

with open('src/main.py', 'w', encoding='utf-8') as f:
    f.write(content)

print("Patch successful")
