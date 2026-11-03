import os
import re
import sqlite3
from datetime import datetime
from functools import wraps
from io import BytesIO

from flask import (
    Flask,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
    jsonify,
)
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from werkzeug.security import check_password_hash, generate_password_hash

from students_data import BRANCH, COLLEGE_NAME, STUDENTS, YEARS

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "Student Allotment for Webcasting.xlsx")
DB_PATH = os.path.join(BASE_DIR, "allotment.db")

ADMIN_USER = "admin"
ADMIN_PASS = "admin@chai$$1"

app = Flask(__name__)
app.secret_key = "webcasting-allotment-secret-key-change-me"


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    cur = conn.cursor()
    cur.executescript(
        """
        CREATE TABLE IF NOT EXISTS stations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            mandal TEXT NOT NULL,
            gram_panchayat TEXT NOT NULL,
            polling_station_no TEXT NOT NULL,
            building_name TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS master_students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            year_of_study TEXT NOT NULL,
            branch TEXT NOT NULL,
            contact_no TEXT NOT NULL,
            college TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS registrations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            roll_number TEXT NOT NULL UNIQUE,
            student_name TEXT NOT NULL,
            college TEXT NOT NULL,
            year_of_study TEXT NOT NULL,
            branch TEXT NOT NULL,
            contact_no TEXT NOT NULL,
            original_contact_no TEXT NOT NULL,
            contact_changed INTEGER NOT NULL DEFAULT 0,
            mandal TEXT NOT NULL,
            gram_panchayat TEXT NOT NULL,
            polling_station_no TEXT NOT NULL,
            building_name TEXT NOT NULL,
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS admins (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL
        );
        """
    )

    # Admin account (keep password in sync with ADMIN_PASS)
    password_hash = generate_password_hash(ADMIN_PASS)
    existing = cur.execute(
        "SELECT id FROM admins WHERE username = ?", (ADMIN_USER,)
    ).fetchone()
    if not existing:
        cur.execute(
            "INSERT INTO admins (username, password_hash) VALUES (?, ?)",
            (ADMIN_USER, password_hash),
        )
    else:
        cur.execute(
            "UPDATE admins SET password_hash = ? WHERE username = ?",
            (password_hash, ADMIN_USER),
        )

    # Seed master students
    count = cur.execute("SELECT COUNT(*) AS c FROM master_students").fetchone()["c"]
    if count == 0:
        for s in STUDENTS:
            cur.execute(
                """
                INSERT INTO master_students (name, year_of_study, branch, contact_no, college)
                VALUES (?, ?, ?, ?, ?)
                """,
                (s["name"], s["year"], BRANCH, s["contact"], COLLEGE_NAME),
            )

    # Seed polling stations from Excel
    station_count = cur.execute("SELECT COUNT(*) AS c FROM stations").fetchone()["c"]
    if station_count == 0 and os.path.exists(EXCEL_PATH):
        wb = load_workbook(EXCEL_PATH, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            mandal = row[8]
            gp = row[9]
            ps = row[10]
            building = row[11]
            if mandal and gp and ps is not None and building:
                cur.execute(
                    """
                    INSERT INTO stations (mandal, gram_panchayat, polling_station_no, building_name)
                    VALUES (?, ?, ?, ?)
                    """,
                    (str(mandal).strip(), str(gp).strip(), str(ps).strip(), str(building).strip()),
                )

    conn.commit()
    conn.close()


def admin_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("admin_logged_in"):
            return redirect(url_for("admin_login"))
        return view(*args, **kwargs)

    return wrapped


@app.route("/")
def index():
    conn = get_db()
    students = [
        dict(r)
        for r in conn.execute(
            "SELECT name, year_of_study, contact_no FROM master_students ORDER BY name"
        ).fetchall()
    ]
    stations = [
        dict(r)
        for r in conn.execute(
            """
            SELECT mandal, gram_panchayat, polling_station_no, building_name
            FROM stations
            ORDER BY mandal, gram_panchayat,
                     CAST(polling_station_no AS INTEGER), polling_station_no
            """
        ).fetchall()
    ]
    mandals = sorted({s["mandal"] for s in stations})
    conn.close()
    return render_template(
        "register.html",
        students=students,
        stations=stations,
        mandals=mandals,
        years=YEARS,
        college=COLLEGE_NAME,
        branch=BRANCH,
    )


@app.route("/api/gram-panchayats")
def api_gram_panchayats():
    mandal = request.args.get("mandal", "").strip()
    conn = get_db()
    rows = conn.execute(
        """
        SELECT DISTINCT gram_panchayat FROM stations
        WHERE mandal = ? ORDER BY gram_panchayat
        """,
        (mandal,),
    ).fetchall()
    conn.close()
    return jsonify([r["gram_panchayat"] for r in rows])


@app.route("/api/polling-stations")
def api_polling_stations():
    mandal = request.args.get("mandal", "").strip()
    gp = request.args.get("gram_panchayat", "").strip()
    conn = get_db()
    rows = conn.execute(
        """
        SELECT DISTINCT polling_station_no FROM stations
        WHERE mandal = ? AND gram_panchayat = ?
        ORDER BY CAST(polling_station_no AS INTEGER), polling_station_no
        """,
        (mandal, gp),
    ).fetchall()
    conn.close()
    return jsonify([r["polling_station_no"] for r in rows])


@app.route("/api/buildings")
def api_buildings():
    mandal = request.args.get("mandal", "").strip()
    gp = request.args.get("gram_panchayat", "").strip()
    ps = request.args.get("polling_station_no", "").strip()
    conn = get_db()
    rows = conn.execute(
        """
        SELECT DISTINCT building_name FROM stations
        WHERE mandal = ? AND gram_panchayat = ? AND polling_station_no = ?
        ORDER BY building_name
        """,
        (mandal, gp, ps),
    ).fetchall()
    conn.close()
    return jsonify([r["building_name"] for r in rows])


@app.route("/register", methods=["POST"])
def register():
    roll_number = request.form.get("roll_number", "").strip().upper()
    student_name = request.form.get("student_name", "").strip()
    year_of_study = request.form.get("year_of_study", "").strip()
    contact_no = request.form.get("contact_no", "").strip()
    mandal = request.form.get("mandal", "").strip()
    gram_panchayat = request.form.get("gram_panchayat", "").strip()
    polling_station_no = request.form.get("polling_station_no", "").strip()
    building_name = request.form.get("building_name", "").strip()

    if not re.fullmatch(r"[A-Za-z0-9]{10}", roll_number):
        flash("Roll number must be exactly 10 alphanumeric characters.", "error")
        return redirect(url_for("index"))

    if not all(
        [
            student_name,
            year_of_study,
            contact_no,
            mandal,
            gram_panchayat,
            polling_station_no,
            building_name,
        ]
    ):
        flash("Please fill all required fields.", "error")
        return redirect(url_for("index"))

    if year_of_study not in YEARS:
        flash("Invalid year of study.", "error")
        return redirect(url_for("index"))

    if not re.fullmatch(r"\d{10}", contact_no):
        flash("Contact number must be 10 digits.", "error")
        return redirect(url_for("index"))

    conn = get_db()
    master = conn.execute(
        "SELECT * FROM master_students WHERE name = ?", (student_name,)
    ).fetchone()
    if not master:
        conn.close()
        flash("Selected student is not in the master list.", "error")
        return redirect(url_for("index"))

    # Validate cascading location belongs together
    station = conn.execute(
        """
        SELECT id FROM stations
        WHERE mandal = ? AND gram_panchayat = ?
          AND polling_station_no = ? AND building_name = ?
        """,
        (mandal, gram_panchayat, polling_station_no, building_name),
    ).fetchone()
    if not station:
        conn.close()
        flash("Invalid polling station selection.", "error")
        return redirect(url_for("index"))

    original_contact = master["contact_no"]
    contact_changed = 1 if contact_no != original_contact else 0

    try:
        conn.execute(
            """
            INSERT INTO registrations (
                roll_number, student_name, college, year_of_study, branch,
                contact_no, original_contact_no, contact_changed,
                mandal, gram_panchayat, polling_station_no, building_name, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                roll_number,
                student_name,
                COLLEGE_NAME,
                year_of_study,
                BRANCH,
                contact_no,
                original_contact,
                contact_changed,
                mandal,
                gram_panchayat,
                polling_station_no,
                building_name,
                datetime.now().isoformat(timespec="seconds"),
            ),
        )
        conn.commit()
        flash("Registration saved successfully.", "success")
    except sqlite3.IntegrityError:
        flash("This roll number is already registered.", "error")
    finally:
        conn.close()

    return redirect(url_for("index"))


@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        conn = get_db()
        admin = conn.execute(
            "SELECT * FROM admins WHERE username = ?", (username,)
        ).fetchone()
        conn.close()
        if admin and check_password_hash(admin["password_hash"], password):
            session["admin_logged_in"] = True
            session["admin_user"] = username
            return redirect(url_for("admin_dashboard"))
        flash("Invalid username or password.", "error")
    return render_template("admin_login.html")


@app.route("/admin/logout")
def admin_logout():
    session.clear()
    return redirect(url_for("admin_login"))


@app.route("/admin")
@admin_required
def admin_dashboard():
    conn = get_db()
    rows = conn.execute(
        """
        SELECT * FROM registrations
        ORDER BY created_at DESC
        """
    ).fetchall()
    conn.close()
    return render_template("admin_dashboard.html", rows=rows)


@app.route("/admin/remove/<int:reg_id>", methods=["POST"])
@admin_required
def admin_remove(reg_id):
    conn = get_db()
    cur = conn.execute("DELETE FROM registrations WHERE id = ?", (reg_id,))
    conn.commit()
    conn.close()
    if cur.rowcount:
        flash("Registration removed.", "success")
    else:
        flash("Registration not found.", "error")
    return redirect(url_for("admin_dashboard"))


@app.route("/admin/export")
@admin_required
def admin_export():
    conn = get_db()
    rows = conn.execute(
        """
        SELECT * FROM registrations
        ORDER BY id ASC
        """
    ).fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Registrations"

    headers = [
        "S.No",
        "Roll Number",
        "Name of College",
        "Name of Student",
        "Year of Study",
        "Branch",
        "Contact No",
        "Original Contact No",
        "Contact Changed",
        "Name of the Mandal",
        "Gram Panchayat Name/s or Ward Numbers",
        "Polling Station No.",
        "Name of the Building and section or block proposed for the polling station",
        "Submitted At",
    ]
    ws.append(headers)

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    changed_fill = PatternFill("solid", fgColor="FFF2CC")  # yellow highlight
    thin = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )

    for col, _ in enumerate(headers, start=1):
        cell = ws.cell(1, col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        cell.border = thin

    for idx, row in enumerate(rows, start=1):
        values = [
            idx,
            row["roll_number"],
            row["college"],
            row["student_name"],
            row["year_of_study"],
            row["branch"],
            row["contact_no"],
            row["original_contact_no"],
            "Yes" if row["contact_changed"] else "No",
            row["mandal"],
            row["gram_panchayat"],
            row["polling_station_no"],
            row["building_name"],
            row["created_at"],
        ]
        ws.append(values)
        excel_row = idx + 1
        for col in range(1, len(headers) + 1):
            cell = ws.cell(excel_row, col)
            cell.border = thin
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if row["contact_changed"] and col in (7, 8, 9):
                cell.fill = changed_fill

    from openpyxl.utils import get_column_letter

    widths = [6, 14, 40, 32, 12, 10, 14, 16, 14, 16, 28, 12, 55, 20]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.row_dimensions[1].height = 35

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    filename = f"webcasting_registrations_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        buffer,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    init_db()
    app.run(debug=True, host="0.0.0.0", port=5000)
