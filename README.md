Automated Student Attendance Monitoring and Analytics System
A smart, QR-code based app for automated student attendance, built for colleges and higher education institutions. This system streamlines roll call, reduces errors, provides powerful analytics, and supports NEP 2020 digital learning transformation.

Features
QR Code-Based Attendance: Each class session generates a unique QR code for secure, proxy-proof attendance marking.

Automated Record-Keeping: Attendance data is securely logged and instantly available for review.

Real-Time Dashboard: Faculty view live attendance stats, absent/late alerts, and class progress.

Analytics & Reports: Visual summary of student attendance trends; highlights at-risk students; supports academic planning.

Supports Online & Offline Classes: Fully compatible with both physical and remote environments.

Cloud Ready: Easily upgradeable to cloud storage or database backends for multi-campus scalability.

User-Friendly UI: Clean web interface for faculty and students; requires minimal training.

Security & Integrity: Prevents proxy attendance and enforces session validity, time windows, and audit trails.

Installation
Clone the repository:

bash
git clone https://github.com/yourusername/attendance-analytics-system.git
cd attendance-analytics-system
Install dependencies:

bash
pip install flask openpyxl qrcode
Run the app:

bash
python app.py
The server will start on http://localhost:5000.

Usage
Faculty Login: Teachers log in to select their subject and start a session.

Generate QR: Each class generates a unique QR code displayed on the dashboard.

Mark Attendance: Students scan the QR code to record their attendance using mobile/web.

Dashboard & Analytics: Faculty view and download attendance records and reports.

File Structure
app.py — Main Flask application and backend logic.

Attendance.xlsx — Excel file for attendance data and analytics.

/templates/ — HTML interface files (login.html, dashboard.html, etc.).

/static/qr_codes/ — Directory for saved QR images.

README.md — This documentation file.

Customization & Upgrades
Cloud Backend: Integrate with SQL/NoSQL DBs for large-scale institutions.

Multi-modal Attendance: Upgrade for biometric/facial recognition authentication.

Mobile Integration: Adapt UI and endpoints for Android/iOS.

More Analytics: Add charts, engagement/behavioral stats, and alerts for low-attendance students.

License
MIT License. See LICENSE for details.

Credits
Developed by Gaurav Deshmukh for Smart India Hackathon
Developed for the Government of Punjab, Department of Higher Education under the Smart Education theme (Govt Problem Statement ID: 25016).

Contact
For feature requests or collaboration, open a GitHub issue or contact via gauravgdeshmukh5225@gmail.com.

