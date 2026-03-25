from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
from datetime import datetime
import os, io, tempfile

app = Flask(__name__)

# Allow requests from GitHub Pages frontend (set FRONTEND_ORIGIN env var on Render)
frontend_origin = os.environ.get("FRONTEND_ORIGIN", "*")
CORS(app, origins=[frontend_origin])

VALID_USER = os.environ.get("PAYROLL_USER", "admin")
VALID_PASS = os.environ.get("PAYROLL_PASS", "payroll@2026")

# ── Styles ────────────────────────────────────────────────────────
header_font = Font(bold=True, name="Arial")
data_font   = Font(name="Arial", size=10)
center      = Alignment(horizontal="center", vertical="center")
thin        = Border(left=Side(style="thin"), right=Side(style="thin"),
                     top=Side(style="thin"),  bottom=Side(style="thin"))
header_fill = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")
total_fill  = PatternFill("solid", start_color="E2EFDA", end_color="E2EFDA")
emp_fill    = PatternFill("solid", start_color="FCE4D6", end_color="FCE4D6")


# ── Auth ──────────────────────────────────────────────────────────
@app.route("/api/login", methods=["POST"])
def login():
    data = request.get_json()
    if data.get("username") == VALID_USER and data.get("password") == VALID_PASS:
        return jsonify({"success": True, "message": "Login successful"})
    return jsonify({"success": False, "message": "Invalid credentials"}), 401


# ── Parse Truein input sheet ──────────────────────────────────────
# Structure: row N = [EmployeeName, ..., ..., ..., ..., Designation]
#            row N+1 = ['DATE', 'STATUS', 'IN TIME', ...]
#            row N+2..M = data rows
def parse_input_sheet(filepath):
    df = pd.read_excel(filepath, header=None)
    employees = []

    for i in range(len(df) - 1):
        row      = df.iloc[i]
        next_row = df.iloc[i + 1]

        col0      = str(row[0]).strip()      if pd.notna(row[0])      else ""
        next_col0 = str(next_row[0]).strip().upper() if pd.notna(next_row[0]) else ""

        # Employee name row is immediately above the DATE header row
        if next_col0 == "DATE" and col0 and col0 != "DATE":
            name  = col0
            desig = str(row[5]).strip() if pd.notna(row[5]) else ""
            data  = {}
            j     = i + 2  # skip the DATE header row

            while j < len(df):
                drow       = df.iloc[j]
                date_val   = drow[0]
                status_val = drow[1]

                # Blank date cell — skip
                if pd.isna(date_val):
                    j += 1
                    continue

                # Non-datetime value that can't be parsed = next employee block
                if not isinstance(date_val, datetime):
                    try:
                        pd.to_datetime(date_val)
                    except Exception:
                        break

                # No status — skip (employee had no check-in data for this date)
                if pd.isna(status_val) or str(status_val).strip() == "":
                    j += 1
                    continue

                try:
                    if isinstance(date_val, datetime):
                        dt_key = date_val.strftime("%d-%b-%Y")
                    else:
                        dt_key = pd.to_datetime(date_val).strftime("%d-%b-%Y")
                except Exception:
                    j += 1
                    continue

                def cv(v):
                    if pd.isna(v) or str(v).strip() in ["-", "na", "NA", "nan"]:
                        return ""
                    if isinstance(v, datetime):
                        return v.strftime("%H:%M")
                    s = str(v).strip()
                    # Handle "HH:MM:SS" → "HH:MM"
                    if s.count(":") == 2:
                        return s[:5]
                    return s

                data[dt_key] = {
                    "status":  cv(status_val),
                    "in":      cv(drow[2]) if len(drow) > 2 else "",
                    "out":     cv(drow[3]) if len(drow) > 3 else "",
                    "worked":  cv(drow[4]) if len(drow) > 4 else "",
                    "ot":      cv(drow[5]) if len(drow) > 5 else "",
                }
                j += 1

            employees.append({"name": name, "designation": desig, "data": data})

    return employees


# ── Parse existing final_payroll.xlsx ────────────────────────────
# Structure: row = ['EMPLOYEE NAME', name, ..., ..., 'DESIGNATION', desig]
#            blank row
#            ['DATE', 'STATUS', ...]
#            data rows
#            ['TOTAL DAYS', n]  ← stop marker
def parse_existing_output(df):
    employees = []
    i = 0
    while i < len(df):
        row  = df.iloc[i]
        col0 = str(row[0]).strip() if pd.notna(row[0]) else ""
        col1 = str(row[1]).strip() if pd.notna(row[1]) else ""

        if col0 == "EMPLOYEE NAME" and col1:
            name  = col1
            desig = str(row[5]).strip() if pd.notna(row[5]) else ""
            data_start = i + 3  # skip blank row + DATE header
            data = {}
            j    = data_start

            while j < len(df):
                drow = df.iloc[j]
                d0   = str(drow[0]).strip() if pd.notna(drow[0]) else ""

                if d0 in ("TOTAL DAYS", "TOTAL OT", "", "EMPLOYEE NAME"):
                    break
                try:
                    datetime.strptime(d0, "%d-%b-%Y")
                    date_str = d0
                except ValueError:
                    j += 1
                    continue

                def cv(v):
                    if pd.isna(v) or str(v).strip() in ["-", "na", "NA", "nan"]:
                        return ""
                    return str(v).strip()

                data[date_str] = {
                    "status":  cv(drow[1]),
                    "in":      cv(drow[2]),
                    "out":     cv(drow[3]),
                    "worked":  cv(drow[4]),
                    "ot":      cv(drow[5]) if len(drow) > 5 else "",
                }
                j += 1

            employees.append({"name": name, "designation": desig, "data": data})
            i = j
        else:
            i += 1

    return employees


# ── Write one employee block, return next available row ──────────
def write_employee_block(ws, start_row, emp):
    # Header row
    for col, val in [(1, "EMPLOYEE NAME"), (2, emp["name"]),
                     (5, "DESIGNATION"),   (6, emp["designation"])]:
        c = ws.cell(start_row, col, val)
        c.font  = Font(bold=True, name="Arial")
        c.fill  = emp_fill
        c.alignment = center

    # Column headers
    header_row = start_row + 2
    for col, h in enumerate(["DATE","STATUS","IN TIME","OUT TIME","DUTY HRS","OT HRS"], 1):
        c = ws.cell(header_row, col, h)
        c.font = header_font; c.fill = header_fill
        c.alignment = center; c.border = thin

    # Data rows sorted by date
    data_start   = header_row + 1
    sorted_dates = sorted(emp["data"].keys(),
                          key=lambda d: datetime.strptime(d, "%d-%b-%Y"))

    for idx, date_str in enumerate(sorted_dates):
        d = emp["data"][date_str]
        for col, val in enumerate([date_str, d["status"], d["in"],
                                    d["out"], d["worked"], d["ot"]], 1):
            c = ws.cell(data_start + idx, col, val)
            c.font = data_font; c.border = thin; c.alignment = center

    # Totals
    total_row  = data_start + len(sorted_dates) + 1
    total_days = sum(1 for d in emp["data"].values() if d["status"] == "PR")
    total_ot_m = 0
    for d in emp["data"].values():
        ot = d.get("ot", "")
        if ot:
            try:
                h, m = map(int, ot.split(":"))
                total_ot_m += h * 60 + m
            except Exception:
                pass

    for label, val, row in [("TOTAL DAYS", total_days,            total_row),
                             ("TOTAL OT",  round(total_ot_m/60,2), total_row+1)]:
        for col, v in [(1, label), (2, val)]:
            c = ws.cell(row, col, v)
            c.font = Font(bold=True, name="Arial")
            c.fill = total_fill; c.border = thin

    return total_row + 2


def build_workbook(employees):
    wb = Workbook()
    ws = wb.active
    ws.title = "Payroll"
    for i, w in enumerate([14, 12, 10, 10, 10, 10], 1):
        ws.column_dimensions[chr(64 + i)].width = w

    current_row = 1
    for emp in employees:
        current_row = write_employee_block(ws, current_row, emp)
        current_row += 1  # blank row between employees
    return wb


# ── New Month ─────────────────────────────────────────────────────
@app.route("/api/new-month", methods=["POST"])
def new_month():
    if "input_file" not in request.files:
        return jsonify({"error": "No input file provided"}), 400

    f   = request.files["input_file"]
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    f.save(tmp.name); tmp.close()

    try:
        employees = parse_input_sheet(tmp.name)
        if not employees:
            return jsonify({"error": "No employee data found in input file"}), 400

        wb  = build_workbook(employees)
        out = io.BytesIO()
        wb.save(out); out.seek(0)
        return send_file(out, as_attachment=True,
                         download_name="final_payroll.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        os.unlink(tmp.name)


# ── Existing Month ────────────────────────────────────────────────
@app.route("/api/existing-month", methods=["POST"])
def existing_month():
    if "input_file" not in request.files or "output_file" not in request.files:
        return jsonify({"error": "Both input and output files are required"}), 400

    tmp_in  = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp_out = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    request.files["input_file"].save(tmp_in.name);   tmp_in.close()
    request.files["output_file"].save(tmp_out.name); tmp_out.close()

    try:
        new_employees  = parse_input_sheet(tmp_in.name)
        if not new_employees:
            return jsonify({"error": "No employee data found in input file"}), 400

        df_existing    = pd.read_excel(tmp_out.name, header=None)
        existing_emps  = parse_existing_output(df_existing)
        new_data_map   = {e["name"]: e for e in new_employees}

        # Update existing employees with new dates (only append missing dates)
        for emp in existing_emps:
            if emp["name"] in new_data_map:
                for date_str, day_data in new_data_map[emp["name"]]["data"].items():
                    if date_str not in emp["data"]:
                        emp["data"][date_str] = day_data

        # Add brand-new employees not in existing file
        existing_names = {e["name"] for e in existing_emps}
        for emp in new_employees:
            if emp["name"] not in existing_names and emp["data"]:
                existing_emps.append(emp)

        wb  = build_workbook(existing_emps)
        out = io.BytesIO()
        wb.save(out); out.seek(0)
        return send_file(out, as_attachment=True,
                         download_name="final_payroll.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        os.unlink(tmp_in.name)
        os.unlink(tmp_out.name)


if __name__ == "__main__":
    print(f"Starting Payroll Server — user: {VALID_USER}")
    app.run(debug=False, port=5000)
