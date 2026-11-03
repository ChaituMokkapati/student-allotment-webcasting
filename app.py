import os
import re
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
from openpyxl.utils import get_column_letter
from werkzeug.security import check_password_hash, generate_password_hash

from db import cast_int_expr, create_schema_sql, get_db, integrity_error
from students_data import BRANCH, COLLEGE_NAME, STUDENTS, YEARS

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "Student Allotment for Webcasting.xlsx")

ADMIN_USER = os.environ.get("ADMIN_USER", "admin")
ADMIN_PASS = os.environ.get("ADMIN_PASS", "admin@chai$$1")

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "webcasting-allotment-secret-key-change-me")


def row_get(row, key):
    if row is None:
        return None
    if isinstance(row, dict):
        return row.get(key)
    return row[key]


def init_db():
    conn = get_db()
    cur = None
    for stmt in create_schema_sql():
        cur = conn.execute(stmt)

    password_hash = generate_password_hash(ADMIN_PASS)
    existing = conn.execute(
        "SELECT id FROM admins WHERE username = ?", (ADMIN_USER,)
    ).fetchone()
    if not existing:
        conn.execute(
            "INSERT INTO admins (username, password_hash) VALUES (?, ?)",
            (ADMIN_USER, password_hash),
        )
    else:
        conn.execute(
            "UPDATE admins SET password_hash = ? WHERE username = ?",
            (password_hash, ADMIN_USER),
        )

    count_row = conn.execute("SELECT COUNT(*) AS c FROM master_students").fetchone()
    count = row_get(count_row, "c")
    if count == 0:
        for s in STUDENTS:
            conn.execute(
                """
                INSERT INTO master_students (name, year_of_study, branch, contact_no, college)
                VALUES (?, ?, ?, ?, ?)
                """,
                (s["name"], s["year"], BRANCH, s["contact"], COLLEGE_NAME),
            )

    station_row = conn.execute("SELECT COUNT(*) AS c FROM stations").fetchone()
    station_count = row_get(station_row, "c")
    if station_count == 0 and os.path.exists(EXCEL_PATH):
        wb = load_workbook(EXCEL_PATH, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            mandal = row[8]
            gp = row[9]
            ps = row[10]
            building = row[11]
            if mandal and gp and ps is not None and building:
                conn.execute(
                    """
                    INSERT INTO stations (mandal, gram_panchayat, polling_station_no, building_name)
                    VALUES (?, ?, ?, ?)
                    """,
                    (
                        str(mandal).strip(),
                        str(gp).strip(),
                        str(ps).strip(),
                        str(building).strip(),
                    ),
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
        dict(r) if not isinstance(r, dict) else r
        for r in conn.execute(
            "SELECT name, year_of_study, contact_no FROM master_students ORDER BY name"
        ).fetchall()
    ]
    cast_ps = cast_int_expr("polling_station_no")
    stations = [
        dict(r) if not isinstance(r, dict) else r
        for r in conn.execute(
            f"""
            SELECT mandal, gram_panchayat, polling_station_no, building_name
            FROM stations
            ORDER BY mandal, gram_panchayat, {cast_ps}, polling_station_no
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
    return jsonify([row_get(r, "gram_panchayat") for r in rows])


@app.route("/api/polling-stations")
def api_polling_stations():
    mandal = request.args.get("mandal", "").strip()
    gp = request.args.get("gram_panchayat", "").strip()
    cast_ps = cast_int_expr("polling_station_no")
    conn = get_db()
    rows = conn.execute(
        f"""
        SELECT DISTINCT polling_station_no FROM stations
        WHERE mandal = ? AND gram_panchayat = ?
        ORDER BY {cast_ps}, polling_station_no
        """,
        (mandal, gp),
    ).fetchall()
    conn.close()
    return jsonify([row_get(r, "polling_station_no") for r in rows])


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
    return jsonify([row_get(r, "building_name") for r in rows])


@app.route("/api/roll-available")
def roll_available():
    roll_number = request.args.get("roll", "").strip().upper()
    if not re.fullmatch(r"[A-Za-z0-9]{10}", roll_number):
        return jsonify({"ok": False, "available": False, "reason": "invalid"})
    conn = get_db()
    existing = conn.execute(
        "SELECT id FROM registrations WHERE UPPER(roll_number) = ?", (roll_number,)
    ).fetchone()
    conn.close()
    return jsonify({"ok": True, "available": existing is None, "roll": roll_number})


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
    duplicate = conn.execute(
        "SELECT id FROM registrations WHERE UPPER(roll_number) = ?", (roll_number,)
    ).fetchone()
    if duplicate:
        conn.close()
        flash("This roll number is already registered. Duplicate roll numbers are not allowed.", "error")
        return redirect(url_for("index"))

    master = conn.execute(
        "SELECT * FROM master_students WHERE name = ?", (student_name,)
    ).fetchone()
    if not master:
        conn.close()
        flash("Selected student is not in the master list.", "error")
        return redirect(url_for("index"))

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

    original_contact = row_get(master, "contact_no")
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
    except integrity_error():
        flash("This roll number is already registered. Duplicate roll numbers are not allowed.", "error")
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
        if admin and check_password_hash(row_get(admin, "password_hash"), password):
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
    rows = [
        dict(r) if not isinstance(r, dict) else r
        for r in conn.execute(
            """
            SELECT * FROM registrations
            ORDER BY created_at DESC
            """
        ).fetchall()
    ]
    conn.close()
    return render_template("admin_dashboard.html", rows=rows)


@app.route("/admin/remove/<int:reg_id>", methods=["POST"])
@admin_required
def admin_remove(reg_id):
    conn = get_db()
    existing = conn.execute(
        "SELECT id, roll_number, student_name FROM registrations WHERE id = ?",
        (reg_id,),
    ).fetchone()
    if not existing:
        conn.close()
        flash("Registration not found.", "error")
        return redirect(url_for("admin_dashboard"))

    roll_number = str(row_get(existing, "roll_number") or "").strip().upper()
    student_name = row_get(existing, "student_name") or "Student"

    conn.execute("DELETE FROM registrations WHERE id = ?", (reg_id,))
    conn.commit()

    # Confirm row is gone so the same roll can be registered again.
    still_there = conn.execute(
        "SELECT id FROM registrations WHERE id = ? OR UPPER(roll_number) = ?",
        (reg_id, roll_number),
    ).fetchone()
    conn.close()

    if still_there:
        flash("Could not remove registration. Please try again.", "error")
    else:
        flash(
            f"Removed {student_name} ({roll_number}). This roll number can be registered again.",
            "success",
        )
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
    changed_fill = PatternFill("solid", fgColor="FFF2CC")
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
        changed = row_get(row, "contact_changed")
        values = [
            idx,
            row_get(row, "roll_number"),
            row_get(row, "college"),
            row_get(row, "student_name"),
            row_get(row, "year_of_study"),
            row_get(row, "branch"),
            row_get(row, "contact_no"),
            row_get(row, "original_contact_no"),
            "Yes" if changed else "No",
            row_get(row, "mandal"),
            row_get(row, "gram_panchayat"),
            row_get(row, "polling_station_no"),
            row_get(row, "building_name"),
            row_get(row, "created_at"),
        ]
        ws.append(values)
        excel_row = idx + 1
        for col in range(1, len(headers) + 1):
            cell = ws.cell(excel_row, col)
            cell.border = thin
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if changed and col in (7, 8, 9):
                cell.fill = changed_fill

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


@app.route("/health")
def health():
    return {"ok": True}


# Initialize schema/seed on startup (gunicorn + local)
try:
    init_db()
except Exception as exc:  # noqa: BLE001
    print(f"DB init warning: {exc}")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(debug=os.environ.get("FLASK_DEBUG", "0") == "1", host="0.0.0.0", port=port)
