# Release Notes: Sub-Process Traceability App v1.1.0

**Release Date:** June 7, 2026  
**Developer:** Taketo Oi (Manager) / Naoufal El Hlou  
**Platform:** Windows 10/11 (Standalone Executable)  

We are excited to announce the release of **v1.1.0** of the HI-LEX ACT Sub-Process Traceability application. This major update brings a complete Role-Based Access Control (RBAC) system, an advanced Administration panel, and a powerful new visual Analytics & KPIs Dashboard!

## Key Features & Enhancements

### Authentication & Role-Based Access Control
- **Secure Login System**: A dedicated login screen to authenticate users before accessing the application.
- **Roles & Permissions**: Four distinct roles added (Manager, Supervisor, Shift Leader, Operator), each with specific permissions.
- **Dynamic UI**: Application tabs and settings intelligently hide or show depending on the logged-in user's role (e.g., only Managers/Supervisors can access Settings).
- **User Display**: The currently logged-in user and their role are displayed prominently at the top right of the application.

### Administration & Settings Hub
- **Centralized Settings Menu**: "Manage Products" and "Audit Logs" have been moved into a clean, dedicated Settings page.
- **User Management**: A brand-new "Manage Users" interface for Admins to seamlessly Add, Edit, and Delete system users and assign roles.
- **Categorized System Logs**: A comprehensive, tabbed System Logs viewer specifically for Admins, separating events into clear categories (Product Logs, User Activity, Login Logs, and Record Actions).

### Advanced Analytics & KPIs Dashboard
- **Real-Time Visualizations**: Integration with `matplotlib` to render beautiful, responsive charts directly inside the application.
- **Actionable Metrics**: Includes Total Daily Production, Top Performing Operator, Hourly Production (24-hour detailed line chart), Shift Output comparison, Operator Mix, and Product Mix pie charts.
- **Historical Data Filtering**: Added a Date Picker to the KPI dashboard, allowing supervisors to easily check production analytics for any past day.
- **Smart Data Labels**: Charts now feature smart data labels with calculated padding to avoid overlapping with titles, ensuring at-a-glance readability.

### Excel Integration Upgrade
- **Export KPIs to Excel**: A new one-click button added to the KPIs dashboard that generates a beautifully formatted "KPI Reports" sheet directly inside your `production_data.xlsx` file.

### Fixes & Optimizations
- **Data Query Accuracy**: Fixed bugs relating to empty queries; charts now accurately fetch data using `created_at` timestamps instead of older fields.
- **Responsive Layouts**: Fixed UI stretching and canvas scrolling issues on the KPI dashboard to ensure it perfectly fits the screen.
- **Pylance/IDE Fixes**: Suppressed false-positive `matplotlib` module warnings to keep the development environment clean.

---
*Thank you for using the HI-LEX ACT Traceability system. For bug reports or feature requests, please contact the developer.*
