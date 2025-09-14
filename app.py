import os
import socket
import time
import secrets
from flask import Flask, render_template, request, jsonify, session, redirect, url_for
from openpyxl import Workbook, load_workbook
from datetime import datetime, date
import calendar
from functools import wraps
from typing import Dict, List, Tuple


app = Flask(__name__)
app.secret_key = secrets.token_hex(16)  # For session management

BASE_URL = os.environ.get('BASE_URL')  # Optional override for base URL

QR_REFRESH_INTERVAL = 90  # seconds
EXCEL_FILE = "Attendance.xlsx"
ATTENDANCE_THRESHOLD = 75  # Minimum attendance %

# Subject definitions with individual passwords
SUBJECTS = {
    'MATH01': {'name': 'Mathematics', 'password': 'math2025'},
    'PHY01': {'name': 'Physics', 'password': 'phys2025'},
    'CHEM01': {'name': 'Chemistry', 'password': 'chem2025'},
    'BIO01': {'name': 'Biology', 'password': 'bio2025'},
    'ENG01': {'name': 'English', 'password': 'eng2025'},
    'COMP01': {'name': 'Computer Science', 'password': 'comp2025'},
    'HIST01': {'name': 'History', 'password': 'hist2025'},
    'GEO01': {'name': 'Geography', 'password': 'geo2025'},
    'ECO01': {'name': 'Economics', 'password': 'eco2025'},
    'STAT01': {'name': 'Statistics', 'password': 'stat2025'}
}

# Store active sessions per subject code
active_sessions = {}

# Excel related functions (unchanged except they are called as before)
def init_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        master = wb.active
        master.title = "Master"
        master.append(["Date", "Subject", "Roll No", "Present", "Token ID", "Timestamp", "IP Address"])
        summary = wb.create_sheet("Summary")
        summary.append([
            "Week Ending", "Roll No", "Subject",
            "Classes Conducted", "Classes Attended",
            "Attendance Percentage", "Status"
        ])
        wb.save(EXCEL_FILE)

def get_or_create_subject_sheet(wb, subject_code: str, date_str: str):
    sheet_name = f"{subject_code}_{date_str}"
    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(sheet_name)
        ws.append(["Roll No", "Present", "Token ID", "Timestamp", "IP Address"])
        for roll_no in range(1, 81):
            ws.append([roll_no, "Absent", "", "", ""])
        wb.save(EXCEL_FILE)
        return True, ws
    else:
        return False, wb[sheet_name]

def update_summary_sheet(wb):
    today = date.today()
    if today.weekday() != calendar.SATURDAY:
        return
    summary = wb["Summary"]
    subject_sheets = [name for name in wb.sheetnames if any(name.startswith(code) for code in SUBJECTS.keys())]
    attendance_data = {}
    for sheet_name in subject_sheets:
        subject_code = sheet_name.split('_')[0]
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=2, values_only=True):
            roll_no = row[0]
            is_present = row[1] == "Present"
            if roll_no not in attendance_data:
                attendance_data[roll_no] = {code: {'attended': 0, 'total': 0} for code in SUBJECTS.keys()}
            attendance_data[roll_no][subject_code]['total'] += 1
            if is_present:
                attendance_data[roll_no][subject_code]['attended'] += 1
    week_ending = today.strftime("%Y-%m-%d")
    for roll_no, subjects in attendance_data.items():
        for subject_code, counts in subjects.items():
            if counts['total'] > 0:
                percentage = (counts['attended'] / counts['total']) * 100
                status = "Below Required" if percentage < ATTENDANCE_THRESHOLD else "Satisfactory"
                summary.append([
                    week_ending, roll_no, subject_code,
                    counts['total'], counts['attended'],
                    f"{percentage:.2f}%", status
                ])
    wb.save(EXCEL_FILE)

def get_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
    except:
        ip = "127.0.0.1"
    finally:
        s.close()
    return ip

def get_base_url():
    if BASE_URL:
        return BASE_URL.rstrip('/')
    ip = get_ip()
    return f"http://{ip}:5000"

# Token generation and QR code generation per subject
def generate_token(subject_code: str, length: int = 16) -> str:
    token = secrets.token_urlsafe(length)[:length]
    # Store or update session info per subject
    history = active_sessions.get(subject_code, {}).get('history', [])
    history.append({'token': token, 'generated_at': time.time()})
    active_sessions[subject_code] = {
        'token': token,
        'generated_at': time.time(),
        'history': history
    }
    return token

def generate_subject_qr(subject_code: str) -> Tuple[str, str]:
    if subject_code not in SUBJECTS:
        raise ValueError(f"Invalid subject code: {subject_code}")
    token = generate_token(subject_code)
    current_time = int(time.time())
    base = get_base_url()
    url = f"{base}/attend/{token}?t={current_time}"  # Timestamp prevents caching
    import qrcode
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4
    )
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    qr_dir = os.path.join("static", "qr_codes")
    os.makedirs(qr_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"qr_{subject_code}_{timestamp}.png"
    filepath = os.path.join(qr_dir, filename)
    latest_path = os.path.join(qr_dir, "latest.png")
    img.save(filepath)
    img.save(latest_path)  # Always update latest.png
    return url, "qr_codes/latest.png"

# Validate attendance session per subject_code
def validate_attendance_session(token: str, subject_code: str) -> Tuple[bool, str]:
    session_data = active_sessions.get(subject_code)
    if not session_data or session_data.get('token') != token:
        return False, "Invalid session token. Please scan the latest QR code."
    if not session_data.get('generated_at'):
        return False, "No active session. Please scan the latest QR code."
    if (time.time() - session_data['generated_at']) > QR_REFRESH_INTERVAL:
        return False, "Session expired. Please scan the latest QR code."
    return True, ""

# Mark attendance - note subject_code must be passed explicitly
def mark_attendance(wb, roll_no: int, token: str, ip_address: str, subject_code: str) -> Tuple[bool, str]:
    is_valid, message = validate_attendance_session(token, subject_code)
    if not is_valid:
        return False, message
    today_str = date.today().strftime("%Y-%m-%d")
    _, ws = get_or_create_subject_sheet(wb, subject_code, today_str)
    try:
        roll_no = int(roll_no)
        if roll_no < 1 or roll_no > 80:
            return False, "Invalid roll number. Please enter a number between 1 and 80."
    except ValueError:
        return False, "Please enter a valid roll number."
    for row in ws.iter_rows(min_row=2):
        if row[0].value == roll_no:
            if row[1].value == "Present":
                return False, "Attendance already marked for your roll number today."
            row[1].value = "Present"
            row[2].value = token
            row[3].value = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            row[4].value = ip_address
            # Also update master sheet
            master = wb["Master"]
            master.append([
                today_str,
                subject_code,
                roll_no,
                "Present",
                token,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                ip_address
            ])
            if date.today().weekday() == calendar.SATURDAY:
                update_summary_sheet(wb)
            wb.save(EXCEL_FILE)
            return True, f"Attendance marked successfully for {SUBJECTS[subject_code]['name']}!"
    return False, "Roll number not found."

# Login required decorator
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'subject_code' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

# Routes
@app.route("/")
def index():
    if 'subject_code' in session:
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        subject_code = request.form.get("subject_code")
        password = request.form.get("password")
        if subject_code in SUBJECTS and password == SUBJECTS[subject_code]['password']:
            session['subject_code'] = subject_code
            return redirect(url_for('dashboard'))
        return render_template("login.html", subjects=SUBJECTS,
                               error="Invalid subject code or password")
    return render_template("login.html", subjects=SUBJECTS)

@app.route("/logout")
def logout():
    session.pop('subject_code', None)
    return redirect(url_for('login'))

@app.route("/dashboard")
@login_required
def dashboard():
    init_excel()
    subject_code = session.get('subject_code')
    current_time = time.time()
    subject_session = active_sessions.get(subject_code)
    if (subject_session and subject_session.get('generated_at') is not None and
            (current_time - subject_session['generated_at']) < QR_REFRESH_INTERVAL):
        time_remaining = int(QR_REFRESH_INTERVAL - (current_time - subject_session['generated_at']))
        qr_url = url_for('attend', token=subject_session['token'], _external=True)
        qr_filename = 'qr_codes/latest.png'
    else:
        # No active session, generate one now:
        qr_url, qr_filename = generate_subject_qr(subject_code)
        subject_session = active_sessions.get(subject_code)
        time_remaining = QR_REFRESH_INTERVAL
    return render_template(
        "dashboard.html",
        subject_code=subject_code,
        subject_name=SUBJECTS[subject_code]['name'],
        time_remaining=time_remaining,
        qr_filename=qr_filename,
        qr_url=qr_url
    )

@app.route("/start_session")
@login_required
def start_session():
    subject_code = session.get('subject_code')
    if not subject_code or subject_code not in SUBJECTS:
        return jsonify({"error": "Invalid session"}), 400
    try:
        url, filename = generate_subject_qr(subject_code)
        # Build absolute static URL for the newly generated latest.png
        qr_static_url = url_for('static', filename=filename, _external=True)
        return jsonify({
            "status": "success",
            "subject_name": SUBJECTS[subject_code]['name'],
            "qr_url": url,
            "time_remaining": QR_REFRESH_INTERVAL,
            "qr_filename": filename,
            "qr_static": qr_static_url
        })
    except Exception as e:
        return jsonify({"error": f"Failed to generate QR code: {str(e)}"}), 500

@app.route("/attend/<token>", methods=["GET", "POST"])
def attend(token):
    # We must find subject_code from token by searching active_sessions (reverse lookup)
    subject_code = None
    for scode, sdata in active_sessions.items():
        if sdata.get('token') == token:
            subject_code = scode
            break
    if not subject_code:
        return render_template("error.html", error="Invalid or expired attendance session token.")

    is_valid, message = validate_attendance_session(token, subject_code)
    if not is_valid:
        return render_template("error.html", error=message)

    if request.method == "POST":
        roll_no = request.form.get("roll_no")
        ip_address = request.headers.get('X-Forwarded-For', request.remote_addr)
        if not roll_no:
            return render_template("attend.html", token=token, error="Please provide your roll number", subject_name=SUBJECTS[subject_code])
        wb = load_workbook(EXCEL_FILE)
        success, msg = mark_attendance(wb, roll_no, token, ip_address, subject_code)
        if success:
            return render_template("success.html", message=msg)
        else:
            return render_template("attend.html", token=token, error=msg, subject_name=SUBJECTS[subject_code])
    return render_template("attend.html", token=token, error=None, subject_name=SUBJECTS[subject_code])

# Run app
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
