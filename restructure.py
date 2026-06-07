import os
import shutil

base_dir = r"c:\Users\Mr_X\Desktop\Sub-Process Traceability"
os.chdir(base_dir)

dirs = ["src", "data", "docs", "scripts", "tests"]
for d in dirs:
    os.makedirs(d, exist_ok=True)

moves = {
    "Sub-Process Traceability.py": "src/main.py",
    "traceability.db": "data/traceability.db",
    "production_data.xlsx": "data/production_data.xlsx",
    "sf_data.json": "data/sf_data.json",
    "traceability_config.json": "data/traceability_config.json",
    "RELEASE_v1.0.0.md": "docs/RELEASE_v1.0.0.md",
    "RELEASE_v1.1.0.md": "docs/RELEASE_v1.1.0.md",
    "build_exe.bat": "scripts/build_exe.bat",
    "file_version_info.txt": "scripts/file_version_info.txt",
    "test_qr.py": "tests/test_qr.py",
    "test_scan.py": "tests/test_scan.py"
}

for src, dst in moves.items():
    if os.path.exists(src):
        if os.path.exists(dst):
            os.remove(dst)
        shutil.move(src, dst)
        print(f"Moved {src} to {dst}")
