# HI-LEX ACT - Sub-Process Traceability

A modern, industrial-grade desktop application built with Python and Tkinter, designed for tracking and managing sub-process manufacturing data, printing Zebra ZPL labels, and generating comprehensive Excel reports.

## Features

- **Modern UI/UX**: Dark mode industrial aesthetic with smooth interactions and responsive tables.
- **Traceability Management**: Track raw materials, batch numbers, shifts, and operators seamlessly.
- **Zebra ZPL Integration**: Direct printing of QA validation labels with embedded QR codes.
- **Excel Reporting**: Automated generation and formatting of `.xlsx` traceability sheets (`traceability_data.xlsx`).
- **Data Persistence**: Local SQLite database (`traceability_v3.db`) with fast search, filtering, and history retention.
- **Smart Validation**: Mandatory batch number checks and empty-field mitigation for clean Excel outputs.
- **Pagination/Reset UI**: Automatically resets the live view every 10 records for a clean workspace while keeping data safe in the database.

## Prerequisites

Make sure you have Python 3.8+ installed on your system.

## Installation & Setup

1. **Clone or Download the Project** to your local machine.
2. **Install Dependencies**:
   Open your terminal/command prompt in the project directory and run:
   ```bash
   pip install -r requirements.txt
   ```

## How to Run

To run the application in development mode:

```bash
python traceability_v3.py
```

## Zebra Printer Configuration

1. Connect your Zebra printer via USB or Network.
2. Click the **"Settings (⛭)"** button in the top right corner of the application.
3. Select your designated Zebra printer from the dropdown list.
4. Click **Save Settings**.

## Compiling to Standalone Executable (.exe)

If you wish to deploy the application on Windows without needing Python installed, you can compile it into a single `.exe` file.

1. Ensure PyInstaller is installed (included in `requirements.txt`).
2. Double-click the included `build_exe.bat` file.
3. Wait for the compilation process to finish.
4. Your standalone application will be generated in the `dist` folder as `traceability_v3.exe`.

*Note: The compilation script automatically embeds the `assets` directory (images/logos) directly into the executable.*

## Project Structure

- `traceability_v3.py` : Main application source code.
- `requirements.txt` : List of Python dependencies.
- `build_exe.bat` : Batch script for compiling the executable using PyInstaller.
- `assets/` : Contains UI graphics like the HI-LEX logo, taskbar icons, and settings icons.
- `traceability_config.json` : Automatically generated configuration file saving user printer preferences.
- `traceability_v3.db` : Automatically generated SQLite database for persistent records.
- `traceability_data.xlsx` : Automatically generated output reports.
