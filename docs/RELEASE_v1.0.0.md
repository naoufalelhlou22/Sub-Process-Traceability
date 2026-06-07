# Release Notes: Sub-Process Traceability App v1.0.0 🚀

**Release Date:** June 7, 2026  
**Developer:** Naoufal El Hlou  
**Platform:** Windows 10/11 (Standalone Executable)  

We are thrilled to announce the official release of **v1.0.0** of the HI-LEX ACT Sub-Process Traceability application. This release establishes a robust, modern foundation for tracking, managing, and printing sub-process manufacturing data.

## ✨ Key Features & Enhancements

### 🖥️ Modern User Interface
- **Industrial Dark Theme**: A sleek, high-contrast UI designed for shop-floor environments with clear color-coded statuses (`Success`, `Warning`, `Danger`).
- **Responsive Navigation**: Tab-based navigation separating "New Entry", "Records", and "Print Label" workflows for optimal operator efficiency.
- **Dynamic Live View**: The "Records" tab automatically clears and refreshes every 10 entries to maintain UI responsiveness without sacrificing underlying database history.

### 🏭 Traceability & Data Entry
- **Smart Data Binding**: Automatic matching of Semi-Finished (SF) part numbers to their respective Raw Materials (RM).
- **Batch Validation**: Smart enforcement ensuring at least one Batch Number is provided before a record can be saved or printed.
- **Comprehensive Fields**: Captures Operator IDs, Shift Data (A/B/C), Stations (S06, S07, S10, S11), and exact Sub-Process/Line Entry DateTimes.

### 🖨️ Zebra Printer Integration
- **Direct ZPL Printing**: Sends Zebra Programming Language (ZPL) commands directly to configured network/USB printers.
- **Double-Copy Automation**: Automatically executes a `^PQ2` command to print 2 copies of each traceability label per single print job.
- **Embedded QA QR Codes**: Generates high-density QR codes mapping all batch and RM/SF constraints for QA scanning.

### 📊 Excel & Database Reporting
- **Local Persistence**: Stores all operations securely in an offline SQLite database (`traceability_v3.db`).
- **Automated Excel Export**: Appends new records directly into `traceability_data.xlsx` with formatted cells, custom headers, and intelligent empty-cell handling (auto-fills missing data with `-`).

### 📦 Build & Deployment
- **Standalone Executable**: Built as a zero-dependency `Sub-Process Traceability.exe` file.
- **Embedded Assets**: UI icons, logos, and settings icons are natively bundled into the binary using `sys._MEIPASS`.
- **Automated Build Script**: Shipped with `build_exe.bat` to streamline compiling future modifications.

---
*Thank you for using the HI-LEX ACT Traceability system. For bug reports or feature requests, please contact the developer.*
