# =========================================================================
# NEXUS CORE ERP  —  app.py  —  v5.0.0 (consolidated)
# =========================================================================
# Consolidated build merging:
#   - v3.9.0 feature set (Studio Safe Mode, SPRO, SU01, PFCG, SE93, CTST,
#     Standard modules, Excel exports, print views, rollback history).
#   - v5.0.0 backports: TN01 module rewrite (CC / Outlook-style picker,
#     admin delete, paste-lock fix, live clock, new routes & print),
#     Admin_Erp master account, removal of sample users, rich standard
#     module schemas, and login screen that no longer reveals credentials.
# =========================================================================

import io
import json
import sqlite3
import secrets
import datetime
import os
import re
import csv
import math
from pathlib import Path

import bcrypt
from fastapi import FastAPI, Request, Response, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment


# =========================================================================
# STUDIO PROTECTION SET
# =========================================================================
STUDIO_PROTECTED = {
    "TN01", "CTST", "SU01", "SPRO", "SE93", "PFCG", "STUDIO", "MRQ1",
    "PROD", "GCH0", "NCR0", "PLN0", "QCH0", "STD0",
}


# =========================================================================
# DATABASE ENGINE SETUP
# =========================================================================
DATABASE_URL = os.environ.get("DATABASE_URL")
USE_PG = bool(DATABASE_URL and DATABASE_URL.startswith("postgres"))
if USE_PG and DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

DB_PATH = Path(__file__).parent / "nexus_erp.db"


def get_conn():
    if USE_PG:
        import psycopg2, psycopg2.extras
        return psycopg2.connect(DATABASE_URL, sslmode="require", connect_timeout=10)
    con = sqlite3.connect(DB_PATH, check_same_thread=False)
    con.row_factory = sqlite3.Row
    return con


def q(sql: str) -> str:
    return sql.replace("?", "%s") if USE_PG else sql


def fetch_all(sql: str, params=()):
    conn = get_conn()
    try:
        if USE_PG:
            import psycopg2.extras
            with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                cur.execute(q(sql), params)
                return [dict(r) for r in cur.fetchall()]
        return [dict(r) for r in conn.execute(sql, params).fetchall()]
    finally:
        conn.close()


def fetch_one(sql: str, params=()):
    conn = get_conn()
    try:
        if USE_PG:
            import psycopg2.extras
            with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                cur.execute(q(sql), params)
                row = cur.fetchone()
                return dict(row) if row else None
        row = conn.execute(sql, params).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def exec_write(sql: str, params=()):
    conn = get_conn()
    try:
        if USE_PG:
            with conn.cursor() as cur:
                cur.execute(q(sql), params)
        else:
            conn.execute(sql, params)
        conn.commit()
    finally:
        conn.close()


def now() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def esc_py(s) -> str:
    return "" if s is None else str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def json_serial(obj):
    if isinstance(obj, (datetime.datetime, datetime.date)):
        return obj.strftime("%Y-%m-%d %H:%M:%S" if isinstance(obj, datetime.datetime) else "%Y-%m-%d")
    return str(obj)


def infer_field_type(sample_values):
    if not sample_values:
        return "text"
    is_num, is_date, is_long = True, True, False
    date_patterns = [r"^\d{4}-\d{2}-\d{2}", r"^\d{2}/\d{2}/\d{4}", r"^\d{2}-\d{2}-\d{4}"]
    for v in sample_values:
        if v is None or v == "":
            continue
        if isinstance(v, (datetime.datetime, datetime.date)):
            is_num = False
            continue
        if isinstance(v, (int, float)):
            is_date = False
            continue
        s = str(v).strip()
        if len(s) > 70 or "\n" in s:
            is_long, is_num, is_date = True, False, False
            break
        if len(s) > 1 and s[0] == "0" and s.isdigit():
            is_num = False
        else:
            try:
                float(s.replace(",", ""))
            except ValueError:
                is_num = False
        if not any(re.match(p, s) for p in date_patterns):
            is_date = False
    return "date" if is_date else ("number" if is_num else ("textarea" if is_long else "text"))


# =========================================================================
# DATABASE INITIALISATION (v5.0 seed set, v3.9 schema)
# =========================================================================
def init_db():
    conn = get_conn()
    cur = conn.cursor()

    if USE_PG:
        try:
            cur.execute(
                "SELECT column_name FROM information_schema.columns "
                "WHERE table_name='tcodes' AND column_name='color';"
            )
            if cur.fetchone():
                cur.execute(
                    "DROP TABLE IF EXISTS tcodes CASCADE; "
                    "DROP TABLE IF EXISTS tcode_schemas CASCADE; "
                    "DROP TABLE IF EXISTS tcode_layouts CASCADE;"
                )
            cur.execute(
                "SELECT column_name FROM information_schema.columns "
                "WHERE table_name='users' AND column_name='sub_department';"
            )
            if cur.fetchone():
                cur.execute("DROP TABLE IF EXISTS users CASCADE;")
            cur.execute(
                "SELECT column_name FROM information_schema.columns "
                "WHERE table_name='plants' AND column_name='code';"
            )
            if not cur.fetchone():
                cur.execute("DROP TABLE IF EXISTS plants CASCADE;")
            cur.execute(
                "SELECT column_name FROM information_schema.columns "
                "WHERE table_name='org' AND column_name='system_id';"
            )
            if not cur.fetchone():
                cur.execute("DROP TABLE IF EXISTS org CASCADE;")
            cur.execute(
                "SELECT column_name FROM information_schema.columns "
                "WHERE table_name='tcode_records' AND column_name='plant';"
            )
            if not cur.fetchone():
                cur.execute(
                    "ALTER TABLE tcode_records ADD COLUMN IF NOT EXISTS plant TEXT DEFAULT '';"
                )
            cur.execute(
                "SELECT column_name FROM information_schema.columns "
                "WHERE table_name='transfer_notes' AND column_name='cc_users';"
            )
            if not cur.fetchone():
                cur.execute(
                    "ALTER TABLE transfer_notes ADD COLUMN IF NOT EXISTS cc_users TEXT DEFAULT '';"
                )
            conn.commit()
        except Exception as e:
            print("PG auto-migration notice:", e)
            conn.rollback()

    id_type = "SERIAL PRIMARY KEY" if USE_PG else "INTEGER PRIMARY KEY AUTOINCREMENT"

    cur.execute(
        "CREATE TABLE IF NOT EXISTS users ("
        "username TEXT PRIMARY KEY, password_hash TEXT, full_name TEXT, "
        "role TEXT, department TEXT, plants TEXT DEFAULT '', "
        "status TEXT DEFAULT 'approved', created_at TEXT, updated_at TEXT);"
    )
    cur.execute(
        "CREATE TABLE IF NOT EXISTS sessions ("
        "token TEXT PRIMARY KEY, username TEXT, expires TEXT);"
    )
    cur.execute(
        "CREATE TABLE IF NOT EXISTS departments ("
        "name TEXT PRIMARY KEY, active INT DEFAULT 1, created_at TEXT);"
    )
    cur.execute(
        "CREATE TABLE IF NOT EXISTS plants ("
        "code TEXT PRIMARY KEY, name TEXT, active INT DEFAULT 1, created_at TEXT);"
    )
    cur.execute(
        "CREATE TABLE IF NOT EXISTS org ("
        "id INT PRIMARY KEY, name TEXT, system_id TEXT, client TEXT);"
    )
    cur.execute(
        "CREATE TABLE IF NOT EXISTS tcodes ("
        "code TEXT PRIMARY KEY, title TEXT, description TEXT, module TEXT, "
        "allowed_depts TEXT, active INT DEFAULT 1, created_at TEXT, "
        "updated_at TEXT, created_by TEXT);"
    )
    cur.execute(
        "CREATE TABLE IF NOT EXISTS tcode_schemas ("
        "code TEXT PRIMARY KEY, fields TEXT, layout TEXT, "
        "updated_at TEXT, updated_by TEXT);"
    )
    cur.execute(
        f"CREATE TABLE IF NOT EXISTS tcode_records ("
        f"id {id_type}, tcode_code TEXT, plant TEXT DEFAULT '', "
        f"data TEXT, status TEXT DEFAULT 'active', created_by TEXT, "
        f"created_by_name TEXT, created_at TEXT, updated_at TEXT);"
    )
    cur.execute(
        f"CREATE TABLE IF NOT EXISTS tcode_layouts ("
        f"id {id_type}, tcode_code TEXT, layout_code TEXT UNIQUE, name TEXT, "
        f"title TEXT, is_default INT DEFAULT 0, fields TEXT, "
        f"layout_config TEXT, description TEXT, created_by TEXT, "
        f"created_by_name TEXT, created_at TEXT, updated_at TEXT);"
    )
    cur.execute(
        f"CREATE TABLE IF NOT EXISTS access_grants ("
        f"id {id_type}, username TEXT, tcode_code TEXT, granted_by TEXT, "
        f"granted_at TEXT, UNIQUE(username, tcode_code));"
    )
    cur.execute(
        f"CREATE TABLE IF NOT EXISTS transfer_notes ("
        f"id {id_type}, note_no TEXT UNIQUE, layout_key TEXT, route TEXT, "
        f"from_plant TEXT, from_dept TEXT, from_user TEXT, from_user_name TEXT, "
        f"to_plant TEXT, to_dept TEXT, to_user TEXT, to_user_name TEXT, "
        f"entry_count INT DEFAULT 0, sheet_headers TEXT, sheet_data TEXT, "
        f"status TEXT DEFAULT 'sent', remarks TEXT DEFAULT '', created_at TEXT, "
        f"received_at TEXT DEFAULT '', received_by TEXT DEFAULT '', "
        f"received_by_name TEXT DEFAULT '', cc_users TEXT DEFAULT '');"
    )
    cur.execute(
        f"CREATE TABLE IF NOT EXISTS module_specs ("
        f"id {id_type}, tcode_code TEXT, version INT, action TEXT, "
        f"spec TEXT, applied_by TEXT, applied_at TEXT, note TEXT);"
    )
    cur.execute(
        "CREATE TABLE IF NOT EXISTS mrq_materials ("
        "material_code TEXT PRIMARY KEY, description TEXT NOT NULL, "
        "category TEXT DEFAULT 'stores', uom TEXT DEFAULT 'NOS', created_at TEXT);"
    )
    cur.execute(
        f"CREATE TABLE IF NOT EXISTS mrq_requisitions ("
        f"id {id_type}, req_no TEXT UNIQUE, category TEXT NOT NULL, "
        f"from_dept TEXT NOT NULL, from_plant TEXT DEFAULT '', "
        f"to_dept TEXT DEFAULT 'Store', status TEXT DEFAULT 'pending', "
        f"remarks TEXT DEFAULT '', requested_by TEXT, requested_by_name TEXT, "
        f"issued_by TEXT, issued_by_name TEXT, issued_to_name TEXT, "
        f"received_by TEXT, received_by_name TEXT, "
        f"created_at TEXT, updated_at TEXT, cancelled_at TEXT, "
        f"cancelled_by TEXT, cancel_reason TEXT DEFAULT '');"
    )
    cur.execute(
        f"CREATE TABLE IF NOT EXISTS mrq_lines ("
        f"id {id_type}, req_id INT, material_code TEXT, description TEXT, "
        f"qty_required REAL DEFAULT 0, batch_no TEXT DEFAULT '', "
        f"qty_issued REAL DEFAULT 0, uom TEXT DEFAULT 'NOS', "
        f"sort_order INT DEFAULT 0);"
    )
    conn.commit()

    if USE_PG:
        try:
            cur.execute(
                "SELECT setval(pg_get_serial_sequence('tcode_records','id'), "
                "COALESCE((SELECT MAX(id)+1 FROM tcode_records), 1), false);"
            )
            cur.execute(
                "SELECT setval(pg_get_serial_sequence('transfer_notes','id'), "
                "COALESCE((SELECT MAX(id)+1 FROM transfer_notes), 1), false);"
            )
            conn.commit()
        except Exception:
            conn.rollback()

    cur_time = now()

    # --- Master org ---
    cur.execute(q("SELECT 1 FROM org WHERE id=1"))
    if not cur.fetchone():
        cur.execute(
            q("INSERT INTO org (id, name, system_id, client) VALUES (?,?,?,?)"),
            (1, "Thermocables Limited", "NX-01", "GLOBAL"),
        )

    # --- Departments ---
    for dept in ["Production", "Final Testing", "Quality Control", "Commercial",
                 "Store", "Planning", "Engineering"]:
        cur.execute(q("SELECT 1 FROM departments WHERE name=?"), (dept,))
        if not cur.fetchone():
            cur.execute(
                q("INSERT INTO departments (name, active, created_at) VALUES (?,?,?)"),
                (dept, 1, cur_time),
            )

    # --- Plants ---
    for code, name in [("1000", "Gooty Central Plant"),
                       ("1400", "Extrusion Unit 1400"),
                       ("2000", "Hyderabad Works"),
                       ("CORP", "Corporate Headquarters")]:
        cur.execute(q("SELECT 1 FROM plants WHERE code=?"), (code,))
        if not cur.fetchone():
            cur.execute(
                q("INSERT INTO plants (code, name, active, created_at) VALUES (?,?,?,?)"),
                (code, name, 1, cur_time),
            )

    # --- Master Admin (single account, v5.0 credentials) ---
    hashed = bcrypt.hashpw("Swami@265".encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
    cur.execute(q("SELECT username FROM users WHERE username='Admin_Erp'"))
    if cur.fetchone():
        cur.execute(
            q("UPDATE users SET password_hash=?, status='approved', role='master' "
              "WHERE username='Admin_Erp'"),
            (hashed,),
        )
    else:
        cur.execute(
            q("INSERT INTO users (username, password_hash, full_name, role, "
              "department, plants, status, created_at, updated_at) "
              "VALUES (?,?,?,?,?,?,?,?,?)"),
            ("Admin_Erp", hashed, "Master System Administrator", "master",
             "IT & Systems", "1000,1400,2000,CORP", "approved", cur_time, cur_time),
        )

    # --- Purge legacy sample users (v5.0 requirement) ---
    cur.execute(
        q("DELETE FROM users WHERE username IN "
          "('admin', 'prod_user', 'testing_lead', 'comm_lead', "
          "'store_head', 'qc_lead')")
    )

    # --- Standard module definitions (v5.0 rich schemas) ---
    standard_modules = [
        ("TN01", "Transfer Note & Handover Sheet",
         "Inter-department material handover & batch transfer note", "MM",
         "Production,Final Testing,Commercial,Store,Quality Control,Planning", []),
        ("CTST", "Cable Type-Test Calculator",
         "CR, Braid Coverage, Tubular & Dumbbell mechanical tests", "QC",
         "Production,Final Testing,Quality Control,Commercial,Store,Planning,Engineering", []),
        ("MRQ1", "Material Requisition",
         "Department-to-Stores material request (Stores / Inhouse / Conductor)", "MM",
         "Production,Final Testing,Quality Control,Commercial,Store,Planning,Engineering", []),
        ("PROD", "Production Handover Report",
         "Manufacturing batch & lot transfer logs", "PROD",
         "Production,Quality Control", [
             {"key": "job_no", "label": "Job / Order #", "type": "text", "required": True},
             {"key": "batch_no", "label": "Batch / Lot No.", "type": "text", "required": True},
             {"key": "cable_size", "label": "Cable Spec / Size", "type": "text", "required": True},
             {"key": "length_m", "label": "Length (Mtrs)", "type": "number", "required": True},
             {"key": "drum_no", "label": "Drum / Coil #", "type": "text"},
             {"key": "operator", "label": "Line Operator", "type": "text"},
             {"key": "remarks", "label": "Production Notes", "type": "textarea", "width": 2},
         ]),
        ("GCH0", "Quality Green Card Release",
         "Incoming QA & physical cable release inspection", "QC",
         "Quality Control,Production", [
             {"key": "job_no", "label": "Production Order #", "type": "text", "required": True},
             {"key": "drum_no", "label": "Drum No.", "type": "text", "required": True},
             {"key": "cable_desc", "label": "Cable Description", "type": "text", "required": True},
             {"key": "passed_length", "label": "Passed Length (Mtrs)", "type": "number", "required": True},
             {"key": "spark_test", "label": "Spark Test Result", "type": "text"},
             {"key": "inspector", "label": "QC Inspector", "type": "text"},
             {"key": "remarks", "label": "Inspection Remarks", "type": "textarea", "width": 2},
         ]),
        ("NCR0", "Non-Conformance Incidents",
         "QC deviation & defect corrective tracking", "QC",
         "Quality Control,Final Testing,Production", [
             {"key": "ncr_no", "label": "NCR Report #", "type": "text", "required": True},
             {"key": "job_no", "label": "Batch / Lot #", "type": "text", "required": True},
             {"key": "defect_type", "label": "Defect Classification", "type": "text", "required": True},
             {"key": "affected_qty", "label": "Affected Length (Mtrs)", "type": "number"},
             {"key": "root_cause", "label": "Root Cause Analysis", "type": "textarea", "width": 2},
             {"key": "corrective_action", "label": "Corrective Action", "type": "textarea", "width": 2},
         ]),
        ("PLN0", "Production Planning Details",
         "Machine allocation and manufacturing order queue", "PLAN",
         "Planning,Production,Store", [
             {"key": "plan_no", "label": "Schedule Order #", "type": "text", "required": True},
             {"key": "customer", "label": "Customer Name", "type": "text", "required": True},
             {"key": "cable_spec", "label": "Cable Specification", "type": "text", "required": True},
             {"key": "planned_qty", "label": "Target Length (Mtrs)", "type": "number", "required": True},
             {"key": "delivery_date", "label": "Target Delivery Date", "type": "date"},
             {"key": "planner", "label": "Planner In-charge", "type": "text"},
         ]),
        ("QCH0", "Stage QC to Final Clearance",
         "Clearance release slips before shipment", "QC",
         "Quality Control,Final Testing", [
             {"key": "clearance_no", "label": "Clearance Slip #", "type": "text", "required": True},
             {"key": "job_no", "label": "Production Lot #", "type": "text", "required": True},
             {"key": "drum_no", "label": "Drum No.", "type": "text", "required": True},
             {"key": "hv_test", "label": "High Voltage Test (kV)", "type": "text"},
             {"key": "ir_val", "label": "Insulation Resistance (MΩ)", "type": "text"},
             {"key": "passed", "label": "Inspection Verdict", "type": "text", "required": True},
         ]),
        ("STD0", "Inventory Transfer Orders",
         "Inter-plant and store material movements", "INV",
         "Store,Planning,Production", [
             {"key": "transfer_id", "label": "Transfer Order #", "type": "text", "required": True},
             {"key": "material_code", "label": "Material Code", "type": "text", "required": True},
             {"key": "description", "label": "Material Description", "type": "text"},
             {"key": "qty", "label": "Transfer Qty", "type": "number", "required": True},
             {"key": "source_loc", "label": "From Location", "type": "text"},
             {"key": "dest_loc", "label": "To Location", "type": "text"},
         ]),
    ]

    for code, title, desc, mod, depts, default_fields in standard_modules:
        cur.execute(q("SELECT 1 FROM tcodes WHERE code=?"), (code,))
        if not cur.fetchone():
            cur.execute(
                q("INSERT INTO tcodes (code, title, description, module, "
                  "allowed_depts, active, created_at, updated_at, created_by) "
                  "VALUES (?,?,?,?,?,?,?,?,?)"),
                (code, title, desc, mod, depts, 1, cur_time, cur_time, "SYSTEM"),
            )
        else:
            cur.execute(q("UPDATE tcodes SET allowed_depts=? WHERE code=?"), (depts, code))

        if default_fields:
            cur.execute(q("SELECT 1 FROM tcode_schemas WHERE code=?"), (code,))
            if not cur.fetchone():
                cur.execute(
                    q("INSERT INTO tcode_schemas (code, fields, layout, "
                      "updated_at, updated_by) VALUES (?,?,?,?,?)"),
                    (code, json.dumps(default_fields), json.dumps({"columns": 3}),
                     cur_time, "SYSTEM"),
                )
            cur.execute(q("SELECT 1 FROM tcode_layouts WHERE tcode_code=?"), (code,))
            if not cur.fetchone():
                cur.execute(
                    q("INSERT INTO tcode_layouts (tcode_code, layout_code, name, "
                      "title, is_default, fields, layout_config, description, "
                      "created_by, created_by_name, created_at, updated_at) "
                      "VALUES (?,?,?,?,?,?,?,?,?,?,?,?)"),
                    (code, f"GRID_{code}_01", "Default Grid", f"{title} Grid", 1,
                     json.dumps(default_fields), json.dumps({"columns": 3}),
                     "Default System View", "SYSTEM", "System",
                     cur_time, cur_time),
                )

    # CTST/MRQ1 handled by dedicated endpoints — no generic schema
    cur.execute(q("DELETE FROM tcode_schemas WHERE code IN ('CTST','MRQ1','TN01')"))
    cur.execute(q("DELETE FROM tcode_layouts WHERE tcode_code IN ('CTST','MRQ1','TN01')"))

    conn.commit()
    conn.close()


init_db()


# =========================================================================
# AUTHENTICATION HELPERS
# =========================================================================
def get_user(req: Request):
    token = (req.headers.get("Authorization", "").replace("Bearer ", "").strip()
             or req.cookies.get("nexus_session"))
    if not token:
        return None
    session = fetch_one("SELECT * FROM sessions WHERE token=?", (token,))
    if not session:
        return None
    user = fetch_one("SELECT * FROM users WHERE username=?", (session["username"],))
    return user if (user and user["status"] == "approved") else None


def need(req: Request):
    u = get_user(req)
    if not u:
        raise HTTPException(401, "Authentication required.")
    return u


def is_master(user) -> bool:
    return user["role"] == "master" or user["username"] == "Admin_Erp"


def need_admin(req: Request):
    u = need(req)
    if u["role"] not in ["admin", "master"]:
        raise HTTPException(403, "Admin access required.")
    return u


def need_master(req: Request):
    u = need(req)
    if not is_master(u):
        raise HTTPException(403, "Master Administrator clearance required.")
    return u


def check_tcode_access(user, code: str) -> bool:
    code = code.upper().strip()
    if is_master(user) or user["role"] == "admin":
        return True
    tc = fetch_one("SELECT * FROM tcodes WHERE code=? AND active=1", (code,))
    if not tc:
        return False
    allowed = [x.strip() for x in (tc["allowed_depts"] or "").split(",") if x.strip()]
    if user["department"] in allowed:
        return True
    return bool(fetch_one(
        "SELECT 1 FROM access_grants WHERE username=? AND tcode_code=?",
        (user["username"], code),
    ))


# =========================================================================
# FASTAPI APP
# =========================================================================
app = FastAPI(title="Nexus Core ERP", version="5.0.0")
app.add_middleware(
    CORSMiddleware, allow_origins=["*"], allow_credentials=True,
    allow_methods=["*"], allow_headers=["*"],
)


# =========================================================================
# AUTH ENDPOINTS
# =========================================================================
@app.api_route("/api/login", methods=["POST"])
def api_login(res: Response, username: str = Form(...), password: str = Form(...)):
    user = fetch_one("SELECT * FROM users WHERE LOWER(username)=LOWER(?)", (username.strip(),))
    if not user or not bcrypt.checkpw(password.encode("utf-8"), user["password_hash"].encode("utf-8")):
        raise HTTPException(401, "Invalid username or password.")
    if user["status"] != "approved":
        raise HTTPException(403, "Account locked or pending approval.")
    token = secrets.token_hex(32)
    expires = (datetime.datetime.now() + datetime.timedelta(days=7)).strftime("%Y-%m-%d %H:%M:%S")
    exec_write("INSERT INTO sessions (token, username, expires) VALUES (?,?,?)",
               (token, user["username"], expires))
    res.set_cookie("nexus_session", token, httponly=True, samesite="lax", max_age=86400 * 7)
    return {"ok": True, "token": token,
            "user": {"username": user["username"],
                     "full_name": user["full_name"],
                     "role": user["role"]}}


@app.api_route("/api/logout", methods=["POST"])
def api_logout(res: Response, req: Request):
    t = (req.headers.get("Authorization", "").replace("Bearer ", "").strip()
         or req.cookies.get("nexus_session"))
    if t:
        exec_write("DELETE FROM sessions WHERE token=?", (t,))
    res.delete_cookie("nexus_session")
    return {"ok": True}


@app.api_route("/api/me", methods=["GET"])
def api_me(req: Request):
    u = need(req)
    org = fetch_one("SELECT * FROM org WHERE id=1") or {}
    return {"username": u["username"], "full_name": u["full_name"],
            "role": u["role"], "department": u["department"],
            "plants": [p.strip() for p in (u["plants"] or "").split(",") if p.strip()],
            "is_master": is_master(u),
            "system": org.get("system_id", "NX-01"),
            "client": org.get("client", "GLOBAL"),
            "company": org.get("name", "Thermocables Limited")}


# =========================================================================
# SPRO — ENTERPRISE STRUCTURE
# =========================================================================
@app.api_route("/api/spro/overview", methods=["GET"])
def spro_overview(req: Request):
    need(req)
    return {"org": fetch_one("SELECT * FROM org WHERE id=1"),
            "departments": fetch_all("SELECT * FROM departments ORDER BY name"),
            "plants": fetch_all("SELECT * FROM plants ORDER BY code")}


@app.api_route("/api/admin/org/update", methods=["POST"])
def spro_update_org(req: Request, name: str = Form(...), system_id: str = Form(...), client: str = Form(...)):
    need_master(req)
    exec_write("UPDATE org SET name=?, system_id=?, client=? WHERE id=1",
               (name.strip(), system_id.strip(), client.strip()))
    return {"ok": True}


@app.api_route("/api/admin/plants/create", methods=["POST"])
def spro_create_plant(req: Request, code: str = Form(...), name: str = Form(...)):
    need_master(req)
    c = code.upper().strip()
    if fetch_one("SELECT 1 FROM plants WHERE code=?", (c,)):
        raise HTTPException(400, "Plant code exists.")
    exec_write("INSERT INTO plants (code, name, active, created_at) VALUES (?,?,?,?)",
               (c, name.strip(), 1, now()))
    return {"ok": True}


@app.api_route("/api/admin/plants/update", methods=["POST"])
def spro_update_plant(req: Request, code: str = Form(...), name: str = Form(...), active: int = Form(1)):
    need_master(req)
    exec_write("UPDATE plants SET name=?, active=? WHERE code=?",
               (name.strip(), active, code.upper().strip()))
    return {"ok": True}


@app.api_route("/api/admin/plants/delete", methods=["POST"])
def spro_delete_plant(req: Request, code: str = Form(...)):
    need_master(req)
    exec_write("DELETE FROM plants WHERE code=?", (code.strip(),))
    return {"ok": True}


@app.api_route("/api/admin/departments/create", methods=["POST"])
def spro_create_dept(req: Request, name: str = Form(...)):
    need_master(req)
    d = name.strip()
    if fetch_one("SELECT 1 FROM departments WHERE name=?", (d,)):
        raise HTTPException(400, "Department exists.")
    exec_write("INSERT INTO departments (name, active, created_at) VALUES (?,?,?)",
               (d, 1, now()))
    return {"ok": True}


@app.api_route("/api/admin/departments/delete", methods=["POST"])
def spro_delete_dept(req: Request, name: str = Form(...)):
    need_master(req)
    exec_write("DELETE FROM departments WHERE name=?", (name.strip(),))
    return {"ok": True}


# =========================================================================
# SU01 — USER DIRECTORY
# =========================================================================
@app.api_route("/api/su01/users", methods=["GET"])
def su01_list_users(req: Request):
    need_admin(req)
    return fetch_all(
        "SELECT username, full_name, role, department, plants, status, "
        "created_at, updated_at FROM users ORDER BY username"
    )


@app.api_route("/api/su01/create", methods=["POST"])
def su01_create_user(req: Request, username: str = Form(...), password: str = Form(...),
                     full_name: str = Form(...), department: str = Form(...),
                     plants: str = Form(""), role: str = Form("user")):
    caller = need_admin(req)
    u = username.strip()
    if role == "master" and not is_master(caller):
        raise HTTPException(403, "Master Admin clearance required.")
    if fetch_one("SELECT 1 FROM users WHERE LOWER(username)=LOWER(?)", (u,)):
        raise HTTPException(400, "User exists.")
    h = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
    exec_write(
        "INSERT INTO users (username, password_hash, full_name, role, "
        "department, plants, status, created_at, updated_at) "
        "VALUES (?,?,?,?,?,?,?,?,?)",
        (u, h, full_name.strip(), role, department.strip(), plants.strip(),
         "approved", now(), now()),
    )
    return {"ok": True}


@app.api_route("/api/su01/update", methods=["POST"])
def su01_update_user(req: Request, username: str = Form(...), full_name: str = Form(...),
                     department: str = Form(...), plants: str = Form(""),
                     role: str = Form(...), status: str = Form(...)):
    caller = need_admin(req)
    u = username.strip()
    target = fetch_one("SELECT * FROM users WHERE username=?", (u,))
    if not target:
        raise HTTPException(404, "User not found.")
    if is_master(target) and not is_master(caller):
        raise HTTPException(403, "Cannot edit Master Administrator.")
    if role == "master" and not is_master(caller):
        raise HTTPException(403, "Cannot grant Master role.")
    exec_write(
        "UPDATE users SET full_name=?, department=?, plants=?, role=?, "
        "status=?, updated_at=? WHERE username=?",
        (full_name.strip(), department.strip(), plants.strip(), role, status, now(), u),
    )
    return {"ok": True}


@app.api_route("/api/su01/delete", methods=["POST"])
def su01_delete_user(req: Request, username: str = Form(...)):
    caller = need_admin(req)
    u = username.strip()
    target = fetch_one("SELECT * FROM users WHERE username=?", (u,))
    if not target:
        raise HTTPException(404, "User not found.")
    if is_master(target):
        raise HTTPException(400, "Master account is protected.")
    if u == caller["username"]:
        raise HTTPException(400, "Cannot delete your own account.")
    exec_write("DELETE FROM users WHERE username=?", (u,))
    exec_write("DELETE FROM sessions WHERE username=?", (u,))
    exec_write("DELETE FROM access_grants WHERE username=?", (u,))
    return {"ok": True}


@app.api_route("/api/su01/reset-password", methods=["POST"])
def su01_reset_password(req: Request, username: str = Form(...), new_password: str = Form(...)):
    caller = need_admin(req)
    u = username.strip()
    target = fetch_one("SELECT * FROM users WHERE username=?", (u,))
    if not target:
        raise HTTPException(404, "User not found.")
    if is_master(target) and not is_master(caller):
        raise HTTPException(403, "Cannot reset Master password.")
    h = bcrypt.hashpw(new_password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
    exec_write("UPDATE users SET password_hash=?, updated_at=? WHERE username=?",
               (h, now(), u))
    exec_write("DELETE FROM sessions WHERE username=?", (u,))
    return {"ok": True}


# =========================================================================
# PFCG — ACCESS CONTROL
# =========================================================================
@app.api_route("/api/pfcg/matrix", methods=["GET"])
def pfcg_matrix(req: Request):
    need_admin(req)
    return {
        "users": fetch_all("SELECT username, full_name, role, department FROM users ORDER BY username"),
        "tcodes": fetch_all("SELECT code, title, module, allowed_depts FROM tcodes WHERE active=1 ORDER BY code"),
        "grants": fetch_all("SELECT * FROM access_grants"),
    }


@app.api_route("/api/pfcg/grant", methods=["POST"])
def pfcg_grant(req: Request, username: str = Form(...), tcode: str = Form(...)):
    caller = need_admin(req)
    u, tc = username.strip(), tcode.upper().strip()
    if USE_PG:
        exec_write(
            "INSERT INTO access_grants (username, tcode_code, granted_by, granted_at) "
            "VALUES (?,?,?,?) ON CONFLICT (username, tcode_code) DO NOTHING",
            (u, tc, caller["username"], now()),
        )
    else:
        exec_write(
            "INSERT OR IGNORE INTO access_grants (username, tcode_code, granted_by, granted_at) "
            "VALUES (?,?,?,?)", (u, tc, caller["username"], now()),
        )
    return {"ok": True}


@app.api_route("/api/pfcg/revoke", methods=["POST"])
def pfcg_revoke(req: Request, username: str = Form(...), tcode: str = Form(...)):
    need_admin(req)
    exec_write("DELETE FROM access_grants WHERE username=? AND tcode_code=?",
               (username.strip(), tcode.upper().strip()))
    return {"ok": True}


# =========================================================================
# SE93 — MODULE STUDIO (UPLOAD BASED AUTO-CREATE)
# =========================================================================
@app.api_route("/api/se93/list", methods=["GET"])
def se93_list(req: Request):
    user = need(req)
    all_t = fetch_all("SELECT * FROM tcodes ORDER BY code")
    for t in all_t:
        t["layout_count"] = fetch_one(
            "SELECT COUNT(*) AS c FROM tcode_layouts WHERE tcode_code=?", (t["code"],)
        )["c"]
        t["record_count"] = fetch_one(
            "SELECT COUNT(*) AS c FROM tcode_records "
            "WHERE tcode_code=? AND status='active'", (t["code"],)
        )["c"]
        t["has_access"] = check_tcode_access(user, t["code"])
    return all_t


@app.api_route("/api/se93/auto-create", methods=["POST"])
async def se93_auto_create(req: Request, code: str = Form(...), title: str = Form(...),
                           description: str = Form(""), module: str = Form("PROD"),
                           allowed_depts: str = Form("Production,Quality Control"),
                           import_sample_data: int = Form(1),
                           file: UploadFile = File(...)):
    user = need_admin(req)
    c = code.upper().strip()
    if not c:
        raise HTTPException(400, "T-Code identifier required.")
    if fetch_one("SELECT 1 FROM tcodes WHERE code=?", (c,)):
        raise HTTPException(400, f"Module '{c}' already exists.")
    contents = await file.read()
    fn = file.filename.lower()
    headers, data_rows = [], []
    try:
        if fn.endswith(".xlsx") or fn.endswith(".xls"):
            wb = load_workbook(io.BytesIO(contents), data_only=True)
            rows = list(wb.active.iter_rows(values_only=True))
            if rows:
                headers = [str(h).strip() for h in rows[0] if h is not None and str(h).strip()]
                data_rows = rows[1:]
        elif fn.endswith(".csv"):
            rows = list(csv.reader(io.StringIO(contents.decode("utf-8", errors="ignore"))))
            if rows:
                headers = [str(h).strip() for h in rows[0] if str(h).strip()]
                data_rows = rows[1:]
        elif fn.endswith(".json"):
            parsed = json.loads(contents.decode("utf-8", errors="ignore"))
            if isinstance(parsed, list) and parsed and isinstance(parsed[0], dict):
                headers = list(parsed[0].keys())
                for item in parsed:
                    data_rows.append([item.get(k, "") for k in headers])
        else:
            raise HTTPException(400, "Upload .xlsx / .xls / .csv / .json only.")
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(400, f"File parse error: {e}")
    if not headers:
        raise HTTPException(400, "No valid column headers found.")

    col_samples = {idx: [r[idx] for r in data_rows[:25]
                         if idx < len(r) and r[idx] is not None]
                   for idx in range(len(headers))}
    fields = []
    for idx, h in enumerate(headers):
        clean_key = re.sub(r"[^a-zA-Z0-9_]", "_", h.lower().strip()).strip("_") or f"col_{idx+1}"
        ftype = infer_field_type(col_samples[idx])
        fields.append({"key": clean_key, "label": h, "type": ftype,
                       "required": False,
                       "width": 2 if ftype == "textarea" else 1})

    cur_time = now()
    exec_write(
        "INSERT INTO tcodes (code, title, description, module, allowed_depts, "
        "active, created_at, updated_at, created_by) VALUES (?,?,?,?,?,?,?,?,?)",
        (c, title.strip() or f"{c} Module", description.strip(), module.strip(),
         allowed_depts.strip(), 1, cur_time, cur_time, user["username"]),
    )
    exec_write(
        "INSERT INTO tcode_schemas (code, fields, layout, updated_at, updated_by) "
        "VALUES (?,?,?,?,?)",
        (c, json.dumps(fields), json.dumps({"columns": 3}), cur_time, user["username"]),
    )
    exec_write(
        "INSERT INTO tcode_layouts (tcode_code, layout_code, name, title, "
        "is_default, fields, layout_config, description, created_by, "
        "created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
        (c, f"GRID_{c}_01", "Default Grid", title or c, 1, json.dumps(fields),
         json.dumps({"columns": 3}), "Auto-created", user["username"],
         user["full_name"], cur_time, cur_time),
    )

    imported = 0
    if import_sample_data and data_rows:
        def_plant = ((user.get("plants") or "1000").split(",")[0].strip() or "1000")
        for row in data_rows:
            if not any(row):
                continue
            rec, row_plant = {}, def_plant
            for idx in range(len(headers)):
                if idx < len(row):
                    v = row[idx]
                    if isinstance(v, (datetime.datetime, datetime.date)):
                        v = v.strftime("%Y-%m-%d %H:%M:%S" if isinstance(v, datetime.datetime) else "%Y-%m-%d")
                    k = fields[idx]["key"]
                    if k in ["plant", "facility"]:
                        row_plant = str(v or def_plant).strip()
                    else:
                        rec[k] = "" if v is None else v
            exec_write(
                "INSERT INTO tcode_records (tcode_code, plant, data, status, "
                "created_by, created_by_name, created_at, updated_at) "
                "VALUES (?,?,?,?,?,?,?,?)",
                (c, row_plant, json.dumps(rec, default=json_serial), "active",
                 user["username"], user["full_name"], cur_time, cur_time),
            )
            imported += 1
    return {"ok": True, "code": c, "title": title,
            "columns_count": len(fields), "imported_count": imported}


@app.api_route("/api/se93/delete", methods=["POST"])
def se93_delete(req: Request, code: str = Form(...)):
    need_master(req)
    c = code.upper().strip()
    if c in STUDIO_PROTECTED:
        raise HTTPException(403, f"'{c}' is a protected system module.")
    exec_write("DELETE FROM tcodes WHERE code=?", (c,))
    exec_write("DELETE FROM tcode_schemas WHERE code=?", (c,))
    exec_write("DELETE FROM tcode_layouts WHERE tcode_code=?", (c,))
    exec_write("DELETE FROM tcode_records WHERE tcode_code=?", (c,))
    exec_write("DELETE FROM access_grants WHERE tcode_code=?", (c,))
    return {"ok": True}


# =========================================================================
# AI CODE STUDIO — SAFE MODE (unchanged, full v3.9 implementation)
# =========================================================================
VALID_TYPES = {"text", "number", "date", "textarea", "select", "email", "tel"}
RESERVED_CODES = {"TN01", "SU01", "SPRO", "SE93", "PFCG", "STUDIO", "CTST", "MRQ1"}


def _snapshot_module(code: str) -> dict:
    tc = fetch_one("SELECT * FROM tcodes WHERE code=?", (code,))
    schema = fetch_one("SELECT * FROM tcode_schemas WHERE code=?", (code,))
    layouts = fetch_all("SELECT * FROM tcode_layouts WHERE tcode_code=?", (code,))
    records = fetch_all(
        "SELECT id, plant, data, status, created_by, created_by_name, "
        "created_at, updated_at FROM tcode_records WHERE tcode_code=?", (code,),
    )
    grants = fetch_all("SELECT * FROM access_grants WHERE tcode_code=?", (code,))
    if schema:
        try:
            schema["fields"] = json.loads(schema["fields"])
        except Exception:
            pass
    for lay in layouts:
        try:
            lay["fields"] = json.loads(lay["fields"])
        except Exception:
            pass
    for rec in records:
        try:
            rec["data"] = json.loads(rec["data"])
        except Exception:
            pass
    return {"tcode": tc, "schema": schema, "layouts": layouts, "records": records,
            "grants": grants, "snapshot_at": now()}


def _restore_module(code: str, snapshot: dict):
    exec_write("DELETE FROM tcodes WHERE code=?", (code,))
    exec_write("DELETE FROM tcode_schemas WHERE code=?", (code,))
    exec_write("DELETE FROM tcode_layouts WHERE tcode_code=?", (code,))
    exec_write("DELETE FROM tcode_records WHERE tcode_code=?", (code,))
    exec_write("DELETE FROM access_grants WHERE tcode_code=?", (code,))

    tc = snapshot.get("tcode")
    if tc:
        exec_write(
            "INSERT INTO tcodes (code, title, description, module, "
            "allowed_depts, active, created_at, updated_at, created_by) "
            "VALUES (?,?,?,?,?,?,?,?,?)",
            (tc["code"], tc["title"], tc.get("description") or "",
             tc.get("module") or "", tc.get("allowed_depts") or "",
             tc.get("active", 1), tc.get("created_at") or now(),
             tc.get("updated_at") or now(), tc.get("created_by") or ""),
        )
    sch = snapshot.get("schema")
    if sch:
        exec_write(
            "INSERT INTO tcode_schemas (code, fields, layout, updated_at, updated_by) "
            "VALUES (?,?,?,?,?)",
            (sch["code"], json.dumps(sch.get("fields", [])),
             json.dumps(sch.get("layout") or {}),
             sch.get("updated_at") or now(), sch.get("updated_by") or ""),
        )
    for lay in snapshot.get("layouts", []):
        exec_write(
            "INSERT INTO tcode_layouts (tcode_code, layout_code, name, title, "
            "is_default, fields, layout_config, description, created_by, "
            "created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
            (lay["tcode_code"], lay["layout_code"], lay.get("name") or "",
             lay.get("title") or "", lay.get("is_default", 0),
             json.dumps(lay.get("fields", [])),
             lay.get("layout_config") or "{}",
             lay.get("description") or "", lay.get("created_by") or "",
             lay.get("created_by_name") or "",
             lay.get("created_at") or now(), lay.get("updated_at") or now()),
        )
    for rec in snapshot.get("records", []):
        exec_write(
            "INSERT INTO tcode_records (id, tcode_code, plant, data, status, "
            "created_by, created_by_name, created_at, updated_at) "
            "VALUES (?,?,?,?,?,?,?,?,?)",
            (rec["id"], code, rec.get("plant") or "",
             json.dumps(rec.get("data") or {}, default=json_serial),
             rec.get("status") or "active",
             rec.get("created_by") or "", rec.get("created_by_name") or "",
             rec.get("created_at") or now(), rec.get("updated_at") or now()),
        )
    for g in snapshot.get("grants", []):
        exec_write(
            "INSERT INTO access_grants (username, tcode_code, granted_by, granted_at) "
            "VALUES (?,?,?,?)",
            (g["username"], g["tcode_code"], g.get("granted_by") or "",
             g.get("granted_at") or now()),
        )


def _validate_spec(spec: dict):
    errors = []
    if not isinstance(spec, dict):
        return None, ["Spec must be a JSON object."]
    action = (spec.get("action") or "upsert").lower()
    if action not in ("upsert", "delete"):
        errors.append("action must be 'upsert' or 'delete'.")
    code = str(spec.get("code") or "").upper().strip()
    if not re.match(r"^[A-Z][A-Z0-9_]{1,10}$", code):
        errors.append("code must be 2-11 chars, start with a letter, A-Z/0-9/_ only.")
    if code in STUDIO_PROTECTED:
        errors.append(
            f"'{code}' is a protected system module. Remove it from "
            f"STUDIO_PROTECTED in app.py to allow Studio edits."
        )
        return {"action": action, "code": code}, errors
    if code in RESERVED_CODES:
        errors.append(f"'{code}' is reserved.")
    clean = {"action": action, "code": code}
    if action == "delete":
        return clean, errors

    clean["title"] = str(spec.get("title") or code).strip()[:120]
    clean["description"] = str(spec.get("description") or "").strip()[:500]
    clean["module"] = str(spec.get("module") or "CUSTOM").strip()[:16].upper()
    depts = spec.get("allowed_depts") or ""
    if isinstance(depts, list):
        depts = ",".join(str(d).strip() for d in depts)
    clean["allowed_depts"] = str(depts).strip()

    fields = spec.get("fields") or []
    if not isinstance(fields, list) or not fields:
        errors.append("fields must be a non-empty array.")
        return clean, errors

    clean_fields, seen_keys = [], set()
    for i, f in enumerate(fields):
        if not isinstance(f, dict):
            errors.append(f"fields[{i}] must be an object.")
            continue
        key = re.sub(r"[^a-z0-9_]", "_",
                     str(f.get("key") or f.get("label") or f"col_{i+1}").lower().strip())
        key = re.sub(r"_+", "_", key).strip("_") or f"col_{i+1}"
        if key in seen_keys:
            errors.append(f"duplicate field key '{key}'.")
            continue
        seen_keys.add(key)
        ftype = str(f.get("type") or "text").lower()
        if ftype not in VALID_TYPES:
            errors.append(f"field '{key}' has invalid type '{ftype}'.")
            ftype = "text"
        clean_fields.append({
            "key": key,
            "label": str(f.get("label") or key).strip()[:120],
            "type": ftype,
            "required": bool(f.get("required", False)),
            "width": 2 if ftype == "textarea" else int(f.get("width", 1) or 1),
            "options": f.get("options") if ftype == "select" else None,
        })
    clean["fields"] = clean_fields

    seeds = spec.get("seed_rows") or []
    if not isinstance(seeds, list):
        errors.append("seed_rows must be an array.")
        seeds = []
    clean["seed_rows"] = [r for r in seeds if isinstance(r, dict)][:500]
    return clean, errors


def _diff_preview(code: str, clean: dict) -> dict:
    existing = fetch_one("SELECT * FROM tcodes WHERE code=?", (code,))
    if not existing:
        return {"kind": "create", "before_rows": 0, "before_fields": [],
                "after_fields": [f["key"] for f in clean["fields"]],
                "added_fields": [f["key"] for f in clean["fields"]],
                "removed_fields": [], "rows_preserved": True}
    schema = fetch_one("SELECT fields FROM tcode_schemas WHERE code=?", (code,))
    try:
        prev_fields = json.loads(schema["fields"]) if schema else []
    except Exception:
        prev_fields = []
    prev_keys = [f.get("key", "") for f in prev_fields]
    new_keys = [f["key"] for f in clean["fields"]]
    added = [k for k in new_keys if k not in prev_keys]
    removed = [k for k in prev_keys if k not in new_keys]
    before_rows = fetch_one(
        "SELECT COUNT(*) AS c FROM tcode_records WHERE tcode_code=?", (code,)
    )["c"]
    return {"kind": "update", "before_rows": before_rows,
            "before_fields": prev_keys, "after_fields": new_keys,
            "added_fields": added, "removed_fields": removed, "rows_preserved": True}


def _apply_spec(clean: dict, user: dict, note: str = ""):
    code = clean["code"]
    action = clean["action"]
    if code in STUDIO_PROTECTED:
        raise HTTPException(
            403, f"'{code}' is protected. Edit STUDIO_PROTECTED in app.py to unlock."
        )
    cur_time = now()
    existing = fetch_one("SELECT code FROM tcodes WHERE code=?", (code,))
    before_snapshot = _snapshot_module(code) if existing else None

    if action == "delete":
        if not existing:
            raise HTTPException(404, f"Module '{code}' does not exist.")
        exec_write("DELETE FROM tcodes WHERE code=?", (code,))
        exec_write("DELETE FROM tcode_schemas WHERE code=?", (code,))
        exec_write("DELETE FROM tcode_layouts WHERE tcode_code=?", (code,))
        exec_write("DELETE FROM tcode_records WHERE tcode_code=?", (code,))
        exec_write("DELETE FROM access_grants WHERE tcode_code=?", (code,))
        summary = {"action": "delete", "code": code}
    else:
        fields_json = json.dumps(clean["fields"])
        layout_cfg = json.dumps({"columns": 3})
        if existing:
            exec_write(
                "UPDATE tcodes SET title=?, description=?, module=?, "
                "allowed_depts=?, updated_at=?, created_by=? WHERE code=?",
                (clean["title"], clean["description"], clean["module"],
                 clean["allowed_depts"], cur_time, user["username"], code),
            )
            exec_write(
                "UPDATE tcode_schemas SET fields=?, layout=?, updated_at=?, "
                "updated_by=? WHERE code=?",
                (fields_json, layout_cfg, cur_time, user["username"], code),
            )
            default_layout = fetch_one(
                "SELECT id FROM tcode_layouts WHERE tcode_code=? AND is_default=1", (code,)
            )
            if default_layout:
                exec_write(
                    "UPDATE tcode_layouts SET fields=?, layout_config=?, "
                    "updated_at=? WHERE id=?",
                    (fields_json, layout_cfg, cur_time, default_layout["id"]),
                )
            else:
                exec_write(
                    "INSERT INTO tcode_layouts (tcode_code, layout_code, name, "
                    "title, is_default, fields, layout_config, description, "
                    "created_by, created_by_name, created_at, updated_at) "
                    "VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                    (code, f"GRID_{code}_01", "Default Grid", clean["title"], 1,
                     fields_json, layout_cfg, "Managed by AI Studio",
                     user["username"], user["full_name"], cur_time, cur_time),
                )
            summary = {"action": "updated", "code": code}
        else:
            exec_write(
                "INSERT INTO tcodes (code, title, description, module, "
                "allowed_depts, active, created_at, updated_at, created_by) "
                "VALUES (?,?,?,?,?,?,?,?,?)",
                (code, clean["title"], clean["description"], clean["module"],
                 clean["allowed_depts"], 1, cur_time, cur_time, user["username"]),
            )
            exec_write(
                "INSERT INTO tcode_schemas (code, fields, layout, updated_at, "
                "updated_by) VALUES (?,?,?,?,?)",
                (code, fields_json, layout_cfg, cur_time, user["username"]),
            )
            exec_write(
                "INSERT INTO tcode_layouts (tcode_code, layout_code, name, "
                "title, is_default, fields, layout_config, description, "
                "created_by, created_by_name, created_at, updated_at) "
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                (code, f"GRID_{code}_01", "Default Grid", clean["title"], 1,
                 fields_json, layout_cfg, "Created via AI Studio",
                 user["username"], user["full_name"], cur_time, cur_time),
            )
            summary = {"action": "created", "code": code}

        seeded = 0
        field_keys = {f["key"] for f in clean["fields"]}
        for row in clean["seed_rows"]:
            rec = {k: row[k] for k in row if k in field_keys}
            if not rec:
                continue
            plant = str(row.get("plant")
                        or (user.get("plants") or "1000").split(",")[0].strip()
                        or "1000")
            exec_write(
                "INSERT INTO tcode_records (tcode_code, plant, data, status, "
                "created_by, created_by_name, created_at, updated_at) "
                "VALUES (?,?,?,?,?,?,?,?)",
                (code, plant, json.dumps(rec, default=json_serial), "active",
                 user["username"], user["full_name"], cur_time, cur_time),
            )
            seeded += 1
        summary["seeded_rows"] = seeded

    prev = fetch_one(
        "SELECT COALESCE(MAX(version), 0) AS v FROM module_specs WHERE tcode_code=?",
        (code,),
    )
    version = (prev["v"] if prev else 0) + 1
    history_payload = {"clean_spec": clean, "before_snapshot": before_snapshot}
    exec_write(
        "INSERT INTO module_specs (tcode_code, version, action, spec, "
        "applied_by, applied_at, note) VALUES (?,?,?,?,?,?,?)",
        (code, version, action, json.dumps(history_payload, default=json_serial),
         user["username"], cur_time, note or ""),
    )
    summary["version"] = version
    return summary


@app.api_route("/api/studio/validate", methods=["POST"])
def studio_validate(req: Request, payload: dict):
    need_master(req)
    clean, errors = _validate_spec(payload.get("spec", payload))
    result = {"ok": not errors, "errors": errors, "clean": clean, "diff": None,
              "protected": clean and clean.get("code") in STUDIO_PROTECTED}
    if clean and not errors and clean.get("code"):
        result["diff"] = _diff_preview(clean["code"], clean)
    return result


@app.api_route("/api/studio/apply", methods=["POST"])
def studio_apply(req: Request, payload: dict):
    user = need_master(req)
    clean, errors = _validate_spec(payload.get("spec", payload))
    if errors:
        raise HTTPException(400, "Spec invalid: " + "; ".join(errors))
    if clean["action"] == "delete" and not payload.get("confirmed", False):
        raise HTTPException(400, "Delete requires confirmed=true. Review the diff, then re-apply.")
    summary = _apply_spec(clean, user, note=str(payload.get("note", "") or ""))
    return {"ok": True, **summary}


@app.api_route("/api/studio/history", methods=["GET"])
def studio_history(req: Request, code: str = ""):
    need_master(req)
    sql = ("SELECT id, tcode_code, version, action, applied_by, applied_at, note "
           "FROM module_specs")
    params = []
    if code:
        sql += " WHERE tcode_code=?"
        params.append(code.upper().strip())
    sql += " ORDER BY id DESC LIMIT 200"
    return fetch_all(sql, tuple(params))


@app.api_route("/api/studio/spec/{code}", methods=["GET"])
def studio_get_spec(req: Request, code: str):
    need_master(req)
    code = code.upper().strip()
    tc = fetch_one("SELECT * FROM tcodes WHERE code=?", (code,))
    if not tc:
        raise HTTPException(404, "Module not found.")
    schema = fetch_one("SELECT fields FROM tcode_schemas WHERE code=?", (code,))
    try:
        fields = json.loads(schema["fields"]) if schema else []
    except Exception:
        fields = []
    return {"action": "upsert", "code": code, "title": tc["title"],
            "description": tc["description"] or "",
            "module": tc["module"] or "CUSTOM",
            "allowed_depts": tc["allowed_depts"] or "",
            "fields": [{"key": f.get("key"), "label": f.get("label"),
                        "type": f.get("type", "text"),
                        "required": f.get("required", False),
                        "width": f.get("width", 1)} for f in fields],
            "seed_rows": [], "protected": code in STUDIO_PROTECTED}


@app.api_route("/api/studio/rollback/{msid}", methods=["POST"])
def studio_rollback(req: Request, msid: int):
    user = need_master(req)
    row = fetch_one("SELECT * FROM module_specs WHERE id=?", (msid,))
    if not row:
        raise HTTPException(404, "History entry not found.")
    try:
        payload = json.loads(row["spec"])
    except Exception:
        raise HTTPException(400, "Stored spec is corrupt.")
    before_snapshot = payload.get("before_snapshot")
    clean_spec = payload.get("clean_spec") or payload
    code = clean_spec.get("code")
    if not code:
        raise HTTPException(400, "History entry is missing a module code.")
    if code in STUDIO_PROTECTED:
        raise HTTPException(403, f"'{code}' is protected. Rollback blocked.")
    if before_snapshot:
        _restore_module(code, before_snapshot)
    else:
        clean, errors = _validate_spec(clean_spec)
        if errors:
            raise HTTPException(400, "Cannot rollback: " + "; ".join(errors))
        _apply_spec(clean, user, note=f"Rollback v{row['version']}")
    cur_time = now()
    prev = fetch_one(
        "SELECT COALESCE(MAX(version), 0) AS v FROM module_specs WHERE tcode_code=?",
        (code,),
    )
    new_version = (prev["v"] if prev else 0) + 1
    exec_write(
        "INSERT INTO module_specs (tcode_code, version, action, spec, "
        "applied_by, applied_at, note) VALUES (?,?,?,?,?,?,?)",
        (code, new_version, "rollback",
         json.dumps({"clean_spec": clean_spec, "before_snapshot": None}),
         user["username"], cur_time, f"Restored from v{row['version']}"),
    )
    return {"ok": True, "code": code, "version": new_version,
            "restored_from": row["version"]}


# =========================================================================
# CTST — CABLE TYPE-TEST CALCULATOR
# =========================================================================
CTST_TESTS = {
    "cr": {"title": "CR — Conductor Resistance", "icon": "ri-pulse-line",
           "desc": "Temperature-corrected resistance with auto per-km value "
                   "and HIGH / LOW verdict vs spec"},
    "braid": {"title": "Braid / Shield Coverage", "icon": "ri-shield-check-line",
              "desc": "Fill factor and % coverage for wire braid screens"},
    "tubular": {"title": "Tubular Mechanical Test", "icon": "ri-circle-line",
                "desc": "Area, tensile strength, elongation for tube specimen"},
    "dumbbell": {"title": "Dumbbell Mechanical Test", "icon": "ri-shape-line",
                 "desc": "Area, tensile strength, elongation for dumbbell specimen"},
    "hotset": {"title": "Hot Set & Permanent Set", "icon": "ri-temp-hot-line",
               "desc": "Hot-set and permanent-set elongation percentages"},
}


def _n(v):
    try:
        return float(str(v).replace(",", "")) if v is not None and str(v).strip() != "" else None
    except Exception:
        return None


def _compute_by_test(tt: str, d: dict) -> dict:
    out = {}
    if tt == "cr":
        obd, t, L, spec = (_n(d.get("obd_cr_reading")), _n(d.get("temp_c")),
                            _n(d.get("length_m")), _n(d.get("specified_cr")))
        tol = _n(d.get("cr_tolerance_pct")) or 5.0
        observed_cr = None
        if obd is not None and t is not None:
            if -50 <= t <= 100:
                corrected = (250.0 / (230.0 + t)) * obd
                out["cr_1m"] = round(corrected, 6)
                if L and L > 0:
                    observed_cr = corrected / L * 1000.0
                    out["cr_total_per_km"] = round(observed_cr, 6)
                    out["observed_cr"] = round(observed_cr, 6)
            else:
                out["cr_warning"] = "Temp outside −50 to 100 °C"
        if spec and observed_cr is not None and spec != 0:
            dev = (observed_cr - spec) / spec * 100.0
            out["cr_deviation_pct"] = round(dev, 4)
            out["cr_verdict"] = "ON SPEC" if abs(dev) < 0.001 else ("HIGH" if dev > 0 else "LOW")
            out["cr_within_tolerance"] = abs(dev) <= tol
            out["cr_tolerance_pct"] = tol
    elif tt == "braid":
        bd, uD, wpc, carr, lay = (_n(d.get("braid_d")), _n(d.get("under_braid_D")),
                                   _n(d.get("wires_per_carrier")), _n(d.get("carriers")),
                                   _n(d.get("lay_length")))
        if all(v is not None for v in (bd, uD, wpc, carr, lay)) and uD > 0 and lay > 0:
            N = wpc * carr
            F = (N * bd) / (2 * math.pi * uD) * math.sqrt(1 + (math.pi * uD / lay) ** 2)
            out["braid_total_wires"] = int(N)
            out["braid_filling_factor"] = round(F, 6)
            out["braid_coverage_pct"] = round((2 * F - F * F) * 100.0, 3)
    elif tt in ("tubular", "dumbbell"):
        if tt == "tubular":
            D, wall = _n(d.get("tub_outer_d")), _n(d.get("tub_wall"))
            if D and wall and D > wall:
                out["area_mm2"] = round(math.pi * (D - wall) * wall, 4)
            load, final = _n(d.get("break_load_tub")), _n(d.get("elong_tub_final"))
        else:
            w, th = _n(d.get("dumb_width")), _n(d.get("dumb_thk"))
            if w and th:
                out["area_mm2"] = round(w * th, 4)
            load, final = _n(d.get("break_load_dumb")), _n(d.get("elong_dumb_final"))
        init = _n(d.get("initial_gauge"))
        if load is not None and out.get("area_mm2"):
            out["tensile_nmm2"] = round(load / out["area_mm2"], 4)
        if init and final is not None and init != 0:
            out["elongation_pct"] = round((final - init) / init * 100.0, 4)
    elif tt == "hotset":
        init, hs, ps = (_n(d.get("initial_gauge")),
                         _n(d.get("hotset_elong_final")),
                         _n(d.get("permset_cooled")))
        if init and hs is not None and init != 0:
            out["hotset_pct"] = round((hs - init) / init * 100.0, 4)
        if init and ps is not None and init != 0:
            out["permset_pct"] = round((ps - init) / init * 100.0, 4)
    return out


def _filter_by_test(base_sql: str, test_type: str):
    if not test_type:
        return base_sql, []
    if USE_PG:
        return base_sql + " AND data::json->>'_test_type' = %s", [test_type]
    return base_sql + " AND json_extract(data, '$._test_type') = ?", [test_type]


@app.api_route("/api/ctst/tests", methods=["GET"])
def ctst_tests(req: Request):
    need(req)
    return CTST_TESTS


@app.api_route("/api/ctst/records", methods=["GET"])
def ctst_list(req: Request, test_type: str = "", limit: int = 200):
    user = need(req)
    if not check_tcode_access(user, "CTST"):
        raise HTTPException(403, "Access denied.")
    base = "SELECT * FROM tcode_records WHERE tcode_code='CTST' AND status='active'"
    sql, params = _filter_by_test(base, test_type)
    sql += " ORDER BY id DESC LIMIT ?" if not USE_PG else " ORDER BY id DESC LIMIT %s"
    params.append(limit)
    rows = fetch_all(sql, tuple(params))
    for r in rows:
        try:
            r["data"] = json.loads(r["data"])
        except Exception:
            r["data"] = {}
        tt = r["data"].get("_test_type", "")
        r["computed"] = _compute_by_test(tt, r["data"]) if tt else {}
    return rows


@app.api_route("/api/ctst/compute", methods=["POST"])
def ctst_compute(req: Request, payload: dict):
    need(req)
    tt = str(payload.get("_test_type", "")).strip()
    if tt not in CTST_TESTS:
        raise HTTPException(400, "Unknown test type.")
    return {"computed": _compute_by_test(tt, payload)}


@app.api_route("/api/ctst/save", methods=["POST"])
def ctst_save(req: Request, payload: dict):
    user = need(req)
    if not check_tcode_access(user, "CTST"):
        raise HTTPException(403, "Access denied.")
    tt = str(payload.get("_test_type", "")).strip()
    if tt not in CTST_TESTS:
        raise HTTPException(400, "Unknown test type.")
    plant = str(payload.pop("_plant", "")
                or (user.get("plants") or "1000").split(",")[0].strip()
                or "1000").strip()
    exec_write(
        "INSERT INTO tcode_records (tcode_code, plant, data, status, "
        "created_by, created_by_name, created_at, updated_at) "
        "VALUES (?,?,?,?,?,?,?,?)",
        ("CTST", plant, json.dumps(payload, default=json_serial), "active",
         user["username"], user["full_name"], now(), now()),
    )
    return {"ok": True, "computed": _compute_by_test(tt, payload)}


@app.api_route("/api/ctst/records/{rid}", methods=["DELETE"])
def ctst_delete(req: Request, rid: int):
    user = need(req)
    if not check_tcode_access(user, "CTST"):
        raise HTTPException(403, "Access denied.")
    rec = fetch_one("SELECT * FROM tcode_records WHERE id=? AND tcode_code='CTST'", (rid,))
    if not rec:
        raise HTTPException(404, "Record not found.")
    if (rec["created_by"] != user["username"]
            and not is_master(user) and user["role"] != "admin"):
        raise HTTPException(403, "You can only delete your own records.")
    exec_write("UPDATE tcode_records SET status='deleted', updated_at=? WHERE id=?",
               (now(), rid))
    return {"ok": True}


@app.api_route("/api/ctst/export", methods=["GET"])
def ctst_export(req: Request, test_type: str = ""):
    user = need(req)
    if not check_tcode_access(user, "CTST"):
        raise HTTPException(403, "Access denied.")
    base = "SELECT * FROM tcode_records WHERE tcode_code='CTST' AND status='active'"
    sql, params = _filter_by_test(base, test_type)
    sql += " ORDER BY id ASC"
    rows = fetch_all(sql, tuple(params))

    wb = Workbook()
    ws = wb.active
    ws.title = (test_type.upper() if test_type else "CTST") + "_DATA"
    ws.append(["ID", "Plant", "Test Type", "Sample ID", "Cable Type", "Date",
               "Result", "Computed", "Remarks", "Created By", "Created At"])
    for r in rows:
        try:
            d = json.loads(r["data"])
        except Exception:
            d = {}
        tt = d.get("_test_type", "")
        c = _compute_by_test(tt, d)
        computed_str = "; ".join(f"{k}={v}" for k, v in c.items())
        ws.append([r["id"], r["plant"], tt, d.get("sample_id", ""),
                   d.get("cable_type", ""), d.get("test_date", ""),
                   d.get("result", ""), computed_str, d.get("remarks", ""),
                   r["created_by_name"] or r["created_by"], r["created_at"]])
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    tag = test_type.upper() if test_type else "ALL"
    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition":
                 f"attachment; filename=CTST_{tag}_"
                 f"{datetime.date.today().strftime('%Y%m%d')}.xlsx"},
    )


@app.api_route("/api/ctst/print/{rid}", methods=["GET"], response_class=HTMLResponse)
def ctst_print(req: Request, rid: int):
    user = need(req)
    if not check_tcode_access(user, "CTST"):
        return HTMLResponse("Access denied", 403)
    rec = fetch_one("SELECT * FROM tcode_records WHERE id=? AND tcode_code='CTST'", (rid,))
    if not rec:
        return HTMLResponse("Record not found", 404)
    try:
        d = json.loads(rec["data"])
    except Exception:
        d = {}
    tt = d.get("_test_type", "")
    c = _compute_by_test(tt, d)
    org = fetch_one("SELECT * FROM org WHERE id=1") or {}
    title = CTST_TESTS.get(tt, {}).get("title", tt.upper())

    def cell(l, v, u=""):
        return (f"<tr>"
                f"<td style='padding:4px 8px;border:1px solid #cbd5e1;"
                f"font-size:11px;background:#f8fafc'>{l}</td>"
                f"<td style='padding:4px 8px;border:1px solid #cbd5e1;"
                f"font-size:11px;font-weight:600'>{v or ''} {u}</td>"
                f"</tr>")

    labels = {"obd_cr_reading": ("OBD CR Reading", "Ω"), "temp_c": ("Test Temp", "°C"),
              "length_m": ("Length", "m"), "specified_cr": ("Specified CR", "Ω/km"),
              "cr_tolerance_pct": ("Tolerance", "%"), "braid_d": ("Braid Wire Ø", "mm"),
              "under_braid_D": ("Ø Under Braid", "mm"), "wires_per_carrier": ("Wires/Carrier", ""),
              "carriers": ("Carriers", ""), "lay_length": ("Lay Length", "mm"),
              "tub_outer_d": ("Tubular Outer Ø", "mm"), "tub_wall": ("Tubular Wall", "mm"),
              "dumb_width": ("Dumbbell Width", "mm"), "dumb_thk": ("Dumbbell Thickness", "mm"),
              "break_load_tub": ("Break Load (Tub)", "N"),
              "break_load_dumb": ("Break Load (Dumb)", "N"),
              "elong_tub_final": ("Tubular Final Gauge", "mm"),
              "elong_dumb_final": ("Dumbbell Final Gauge", "mm"),
              "initial_gauge": ("Initial Gauge", "mm"),
              "hotset_elong_final": ("Hot Set Final", "mm"),
              "permset_cooled": ("Perm Set Cooled", "mm")}
    comp_labels = {"cr_1m": ("CR 1 m", "Ω"), "cr_total_per_km": ("CR per km", "Ω/km"),
                   "observed_cr": ("Observed CR (auto, Ω/km)", "Ω/km"),
                   "cr_deviation_pct": ("CR Deviation vs Spec", "%"),
                   "cr_verdict": ("vs Spec", ""),
                   "cr_within_tolerance": ("Within Tolerance", ""),
                   "cr_tolerance_pct": ("Tolerance", "%"),
                   "braid_total_wires": ("Total Wires", ""),
                   "braid_filling_factor": ("Filling Factor", ""),
                   "braid_coverage_pct": ("Coverage", "%"),
                   "area_mm2": ("Area", "mm²"),
                   "tensile_nmm2": ("Tensile Strength", "N/mm²"),
                   "elongation_pct": ("Elongation", "%"),
                   "hotset_pct": ("Hot Set", "%"),
                   "permset_pct": ("Permanent Set", "%")}

    irows = "".join(cell(l, d.get(k, ""), u_) for k, (l, u_) in labels.items() if k in d)
    orows = "".join(cell(l, c.get(k, ""), u_) for k, (l, u_) in comp_labels.items() if k in c)
    if c.get("cr_warning"):
        orows += (f"<tr><td colspan='2' style='padding:6px 8px;"
                  f"background:#fef3c7;border:1px solid #fcd34d;"
                  f"font-size:11px;color:#92400e'>⚠ {esc_py(c['cr_warning'])}</td></tr>")

    return HTMLResponse(f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>CTST #{rec['id']}</title>
<style>body {{ font-family: sans-serif; margin: 30px; color: #0f172a; }}
table {{ width: 100%; border-collapse: collapse; margin-bottom: 16px; }}
@media print {{ .no-print {{ display: none; }} body {{ margin: 0; }} }}</style></head><body>
<div class="no-print" style="text-align:right;margin-bottom:12px">
  <button onclick="window.print()" style="background:#4f46e5;color:white;border:none;padding:8px 16px;border-radius:6px;cursor:pointer;font-weight:600">Print / Save PDF</button>
</div>
<h2 style="margin:0">{esc_py(org.get('name', 'Thermocables Limited'))}</h2>
<h4 style="margin:4px 0 12px;color:#4f46e5">{esc_py(title)}</h4>
<div style="font-size:11px;color:#475569;margin-bottom:16px">
Record #{rec['id']} | Plant {esc_py(rec['plant'] or '—')} | Sample {esc_py(d.get('sample_id', '—'))} |
{esc_py(d.get('cable_type', '—'))} | By {esc_py(rec['created_by_name'] or rec['created_by'])} | {esc_py(rec['created_at'])}</div>
<h4 style="font-size:12px;color:#4338ca;text-transform:uppercase;margin-bottom:6px">Inputs</h4>
<table>{irows}</table>
<h4 style="font-size:12px;color:#4338ca;text-transform:uppercase;margin-bottom:6px">Computed Results</h4>
<table>{orows or '<tr><td style="padding:6px;font-size:11px;color:#64748b">None</td></tr>'}</table>
<h4 style="font-size:12px;color:#4338ca;text-transform:uppercase;margin-bottom:6px">Verdict</h4>
<table>{cell('Result', d.get('result', ''))}{cell('Remarks', d.get('remarks', ''))}</table>
</body></html>""")


# =========================================================================
# TN01 — TRANSFER NOTE (v5.0 BACKPORT: CC, admin delete, new routes)
# =========================================================================
TN01_ROUTES = {
    "PROD_TO_QC": {"title": "Production to Final Testing",
                   "route": "Production ➔ Final Testing", "to_dept": "Final Testing",
                   "columns": ["Sr. No.", "Job No.", "Customer Name",
                               "Material Description", "Specified Length",
                               "Entry Length", "Drum Size", "Drum Code", "Batch No."]},
    "QC_TO_COMM": {"title": "Final Testing to Commercial",
                   "route": "Final Testing ➔ Commercial", "to_dept": "Commercial",
                   "columns": ["Sr. No.", "Job / Prod Order #", "Batch / Heat Lot",
                               "Cable Description / Spec", "Drum #",
                               "Tested Length (Mtrs)", "HV Test Result",
                               "Insulation Resistance", "QC Inspector",
                               "Release Status"]},
    "PROD_TO_STORE": {"title": "Production to Store",
                      "route": "Production ➔ Store", "to_dept": "Store",
                      "columns": ["Sr. No.", "Work Order #", "Item Code / Material",
                                  "Batch / Heat Lot", "Quantity (Mtrs/Kgs)",
                                  "Weight (Kgs)", "Drum / Pack #",
                                  "Handover Person", "Remarks"]},
    "QC_TO_PROD": {"title": "Final Testing to Production (Rework)",
                   "route": "Final Testing ➔ Production", "to_dept": "Production",
                   "columns": ["Sr. No.", "Job Order #", "Batch / Heat Lot",
                               "Cable Size", "Defect Observed", "Length (Mtrs)",
                               "QC Engineer", "Rework Instructions"]},
    "COMM_TO_QC": {"title": "Commercial to Final Testing",
                   "route": "Commercial ➔ Final Testing", "to_dept": "Final Testing",
                   "columns": ["Sr. No.", "Customer Name", "Order / PO #",
                               "Item Description", "Ordered Qty",
                               "Inspection Agency", "Target Date",
                               "Special Testing Notes"]},
    "STORE_TO_PROD": {"title": "Store to Production",
                      "route": "Store ➔ Production", "to_dept": "Production",
                      "columns": ["Sr. No.", "Material Code", "Material Description",
                                  "Lot / GRN #", "Issued Qty", "UOM", "Issued By",
                                  "Received For Work Order"]},
}


@app.api_route("/api/tn01/routes", methods=["GET"])
def tn01_get_routes(req: Request):
    need(req)
    return TN01_ROUTES


@app.api_route("/api/tn01/suggest-users", methods=["GET"])
def tn01_suggest_users(req: Request, plant: str = "", dept: str = "", q: str = ""):
    need(req)
    sql = ("SELECT username, full_name, department, plants, role "
           "FROM users WHERE status='approved'")
    params = []
    if dept:
        sql += " AND department = ?"
        params.append(dept)
    if q:
        sql += " AND (LOWER(username) LIKE ? OR LOWER(full_name) LIKE ?)"
        t = f"%{q.strip().lower()}%"
        params.extend([t, t])
    users = fetch_all(sql, tuple(params))
    if plant:
        return [u for u in users
                if not (u.get("plants") or "").strip()
                or plant in [p.strip() for p in (u.get("plants") or "").split(",")]]
    return users


@app.api_route("/api/tn01/layouts", methods=["GET"])
def tn01_get_layouts(req: Request):
    need(req)
    rows = fetch_all("SELECT * FROM tcode_layouts WHERE tcode_code='TN01' "
                     "ORDER BY is_default DESC, id ASC")
    for r in rows:
        try:
            r["fields"] = json.loads(r["fields"])
        except Exception:
            r["fields"] = []
    return rows


@app.api_route("/api/tn01/layouts/save", methods=["POST"])
def tn01_save_layout(req: Request, payload: dict):
    user = need_admin(req)
    name = payload.get("name", "Custom Layout").strip()
    key = payload.get("key", "").strip().upper()
    columns = payload.get("columns", [])
    if not key:
        cnt = fetch_one("SELECT COUNT(*) AS c FROM tcode_layouts WHERE tcode_code='TN01'")["c"]
        key = f"TN01_VAR_{cnt + 1:02d}"
    cur_time = now()
    existing = fetch_one(
        "SELECT id FROM tcode_layouts WHERE tcode_code='TN01' AND layout_code=?", (key,)
    )
    if existing:
        exec_write("UPDATE tcode_layouts SET name=?, title=?, fields=?, updated_at=? WHERE id=?",
                   (name, name, json.dumps(columns), cur_time, existing["id"]))
    else:
        exec_write(
            "INSERT INTO tcode_layouts (tcode_code, layout_code, name, title, "
            "is_default, fields, layout_config, description, created_by, "
            "created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
            ("TN01", key, name, name, 0, json.dumps(columns), "{}",
             "Custom Format", user["username"], user["full_name"], cur_time, cur_time),
        )
    return {"ok": True, "key": key}


@app.api_route("/api/tn01/layouts/{lid}", methods=["DELETE"])
def tn01_delete_layout(req: Request, lid: int):
    need_admin(req)
    exec_write("DELETE FROM tcode_layouts WHERE id=? AND tcode_code='TN01'", (lid,))
    return {"ok": True}


@app.api_route("/api/tn01/send", methods=["POST"])
def tn01_send(req: Request, payload: dict):
    user = need(req)
    today_tag = datetime.date.today().strftime("%Y%m%d")
    cnt = fetch_one(
        "SELECT COUNT(*) AS c FROM transfer_notes WHERE created_at LIKE ?",
        (f"{datetime.date.today().strftime('%Y-%m-%d')}%",),
    )
    note_no = f"TN-{today_tag}-{(cnt['c'] if cnt else 0) + 1:04d}"
    headers = payload.get("headers", [])
    rows = payload.get("rows", [])
    header_count = len(headers) if headers else 0
    normalized_rows = []
    for i, raw in enumerate(rows, 1):
        if not isinstance(raw, (list, tuple)):
            continue
        row = ["" if c is None else str(c) for c in raw]
        while len(row) < header_count:
            row.append("")
        if header_count:
            row = row[:header_count]
        row[0] = str(i)
        normalized_rows.append(row)
    normalized_rows = [r for r in normalized_rows if any(str(x).strip() for x in r[1:])]
    to_user = payload.get("to_user", "").strip()
    target = fetch_one("SELECT username, full_name FROM users WHERE username=?", (to_user,))
    to_user_name = target["full_name"] if target else to_user
    cc_users_raw = payload.get("cc_users") or []
    if isinstance(cc_users_raw, str):
        cc_users_raw = [cc_users_raw]
    cc_users = ",".join(str(x).strip() for x in cc_users_raw if str(x).strip())
    from_plant = (payload.get("from_plant", "").strip()
                  or (user.get("plants") or ["1000"])[0])
    exec_write(
        "INSERT INTO transfer_notes (note_no, layout_key, route, from_plant, "
        "from_dept, from_user, from_user_name, to_plant, to_dept, to_user, "
        "to_user_name, entry_count, sheet_headers, sheet_data, status, "
        "remarks, created_at, cc_users) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
        (note_no, payload.get("layout_key", "STANDARD"),
         payload.get("route", "Transfer"), from_plant, user["department"],
         user["username"], user["full_name"], payload.get("to_plant", "").strip(),
         payload.get("to_dept", "").strip(), to_user, to_user_name,
         len(normalized_rows), json.dumps(headers),
         json.dumps(normalized_rows, default=json_serial), "sent",
         payload.get("remarks", "").strip(), now(), cc_users),
    )
    return {"ok": True, "note_no": note_no, "entries": len(normalized_rows)}


@app.api_route("/api/tn01/inbox", methods=["GET"])
def tn01_inbox(req: Request):
    user = need(req)
    uname = user["username"]
    return fetch_all(
        "SELECT * FROM transfer_notes WHERE to_user = ? "
        "OR (to_dept = ? AND to_user = '') "
        "OR (',' || cc_users || ',') LIKE ? ORDER BY id DESC",
        (uname, user["department"], f"%,{uname},%"),
    )


@app.api_route("/api/tn01/receive/{nid}", methods=["POST"])
def tn01_receive(req: Request, nid: int):
    user = need(req)
    exec_write(
        "UPDATE transfer_notes SET status='received', received_at=?, "
        "received_by=?, received_by_name=? WHERE id=?",
        (now(), user["username"], user["full_name"], nid),
    )
    return {"ok": True}


@app.api_route("/api/tn01/delete/{nid}", methods=["DELETE"])
def tn01_delete(req: Request, nid: int):
    user = need(req)
    if not (is_master(user) or user["role"] == "admin"):
        raise HTTPException(403, "Only Admin can delete history.")
    exec_write("DELETE FROM transfer_notes WHERE id=?", (nid,))
    return {"ok": True}


@app.api_route("/api/tn01/history", methods=["GET"])
def tn01_history(req: Request, plant: str = "", route: str = "",
                 status: str = "", from_dept: str = "", to_dept: str = "",
                 date_from: str = "", date_to: str = ""):
    user = need(req)
    uname = user["username"]
    sql = "SELECT * FROM transfer_notes WHERE 1=1"
    params = []
    if not (is_master(user) or user["role"] == "admin"):
        sql += (" AND (from_user = ? OR to_user = ? OR from_dept = ? "
                "OR to_dept = ? OR (',' || cc_users || ',') LIKE ?)")
        params.extend([user["username"], user["username"],
                       user["department"], user["department"], f"%,{uname},%"])
    if plant:
        sql += " AND (from_plant = ? OR to_plant = ?)"
        params.extend([plant, plant])
    if route:
        sql += " AND route = ?"
        params.append(route)
    if status:
        sql += " AND status = ?"
        params.append(status)
    if is_master(user) or user["role"] == "admin":
        if from_dept:
            sql += " AND from_dept = ?"
            params.append(from_dept)
        if to_dept:
            sql += " AND to_dept = ?"
            params.append(to_dept)
        if date_from:
            sql += " AND created_at >= ?"
            params.append(date_from + " 00:00:00")
        if date_to:
            sql += " AND created_at <= ?"
            params.append(date_to + " 23:59:59")
    sql += " ORDER BY id DESC LIMIT 300"
    return fetch_all(sql, tuple(params))


@app.api_route("/api/tn01/note/{nid}", methods=["GET"])
def tn01_get_note(req: Request, nid: int):
    need(req)
    note = fetch_one("SELECT * FROM transfer_notes WHERE id=?", (nid,))
    if not note:
        raise HTTPException(404, "Note not found.")
    try:
        note["sheet_headers"] = json.loads(note["sheet_headers"])
    except Exception:
        note["sheet_headers"] = []
    try:
        note["sheet_data"] = json.loads(note["sheet_data"])
    except Exception:
        note["sheet_data"] = []
    return note


@app.api_route("/api/tn01/export/{nid}", methods=["GET"])
def tn01_export_excel(req: Request, nid: int):
    user = need(req)
    note = fetch_one("SELECT * FROM transfer_notes WHERE id=?", (nid,))
    if not note:
        raise HTTPException(404, "Note not found.")
    is_sender = (user["username"] == note["from_user"]
                 or user["department"] == note["from_dept"])
    is_receiver = (user["username"] == note["to_user"]
                   or user["department"] == note["to_dept"])
    cc_list = [x.strip() for x in (note.get("cc_users") or "").split(",") if x.strip()]
    is_cc = user["username"] in cc_list
    if not (is_sender or is_receiver or is_cc
            or is_master(user) or user["role"] == "admin"):
        raise HTTPException(403, "Access denied.")
    try:
        headers = json.loads(note["sheet_headers"])
    except Exception:
        headers = []
    try:
        data_rows = json.loads(note["sheet_data"])
    except Exception:
        data_rows = []
    wb = Workbook()
    ws = wb.active
    ws.title = "Transfer_Note"
    org = fetch_one("SELECT * FROM org WHERE id=1") or {}
    ws.append([org.get("name", "Thermocables Limited")])
    ws.append([f"TRANSFER NOTE {note['note_no']} — {note['route']}"])
    ws.append([f"From: {note['from_plant']} - {note['from_dept']}  |  "
               f"To: {note['to_plant']} - {note['to_dept']}"])
    ws.append([f"Sender: {note['from_user_name']}  |  "
               f"Receiver: {note['to_user_name']}  |  "
               f"Status: {note['status'].upper()}"])
    ws.append([f"CC: {note.get('cc_users') or '—'}"])
    ws.append([])
    ws.append(headers)
    for row in data_rows:
        ws.append(row)
    ws.append([])
    ws.append(["Handovered By: ________________", "",
               "Received By: ________________", "", "QC: ________________"])
    header_fill = PatternFill("solid", fgColor="4F46E5")
    header_font = Font(bold=True, color="FFFFFF")
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"].font = Font(bold=True)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=7, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
    for col_idx, h in enumerate(headers, 1):
        ws.column_dimensions[
            ws.cell(row=7, column=col_idx).column_letter
        ].width = max(14, min(40, len(str(h)) + 4))
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition":
                 f"attachment; filename=TransferNote_{note['note_no']}.xlsx"},
    )


@app.api_route("/api/tn01/export-grid", methods=["POST"])
def tn01_export_grid(req: Request, payload: dict):
    user = need(req)
    headers = payload.get("headers", [])
    rows = payload.get("rows", [])
    title = payload.get("title", "Transfer Note (Draft)")
    if not headers:
        raise HTTPException(400, "No columns to export.")
    wb = Workbook()
    ws = wb.active
    ws.title = "Draft_Note"
    org = fetch_one("SELECT * FROM org WHERE id=1") or {}
    ws.append([org.get("name", "Thermocables Limited")])
    ws.append([title])
    ws.append([f"Prepared by: {user['full_name']} ({user['username']}) · {user['department']}"])
    ws.append([f"Generated: {now()}"])
    ws.append([])
    ws.append(headers)
    for r in rows:
        ws.append(r)
    header_fill = PatternFill("solid", fgColor="4F46E5")
    header_font = Font(bold=True, color="FFFFFF")
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"].font = Font(bold=True)
    header_row = 6
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=header_row, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
    for col_idx, h in enumerate(headers, 1):
        ws.column_dimensions[
            ws.cell(row=header_row, column=col_idx).column_letter
        ].width = max(14, min(40, len(str(h)) + 4))
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition":
                 f"attachment; filename=TN01_Draft_"
                 f"{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"},
    )


@app.api_route("/api/tn01/print/{nid}", methods=["GET"], response_class=HTMLResponse)
def tn01_print_preview(req: Request, nid: int):
    need(req)
    note = fetch_one("SELECT * FROM transfer_notes WHERE id=?", (nid,))
    if not note:
        return HTMLResponse("Note not found", 404)
    org = fetch_one("SELECT * FROM org WHERE id=1") or {}
    try:
        headers = json.loads(note["sheet_headers"])
    except Exception:
        headers = []
    try:
        data_rows = json.loads(note["sheet_data"])
    except Exception:
        data_rows = []
    th = "".join(
        f"<th style='border:1px solid #0f172a; padding:6px 8px; "
        f"font-size:11px; background:#e0e7ff'>{h}</th>" for h in headers
    )
    trs = "".join(
        "<tr>" + "".join(
            f'<td style="border:1px solid #cbd5e1; padding:6px 8px; '
            f'font-size:11px">{c}</td>' for c in r
        ) + "</tr>" for r in data_rows
    )
    return HTMLResponse(f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>{note['note_no']}</title>
<style>body {{ font-family: sans-serif; margin: 30px; color: #0f172a; }}
.header {{ border-bottom: 2px solid #4f46e5; padding-bottom: 10px; margin-bottom: 15px; }}
.meta {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px;
font-size: 12px; background: #f8fafc; padding: 12px;
border: 1px solid #e2e8f0; margin-bottom: 20px; }}
table {{ width: 100%; border-collapse: collapse; margin-bottom: 30px; }}
.signs {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px;
margin-top: 40px; text-align: center; font-size: 11px; font-weight: bold; }}
.sign-box {{ border-top: 1px solid #0f172a; padding-top: 8px; }}
@media print {{ .no-print {{ display: none; }} body {{ margin: 0; }} }}</style></head><body>
<div class="no-print" style="margin-bottom:15px;text-align:right">
  <button onclick="window.print()" style="background:#4f46e5;color:white;
  border:none;padding:8px 16px;border-radius:6px;cursor:pointer;font-weight:bold">
  Print / PDF</button>
</div>
<div class="header">
  <h2 style="margin:0">{org.get('name', 'Thermocables')}</h2>
  <h4 style="margin:4px 0 0;color:#4f46e5">TRANSFER NOTE (TN01)</h4>
  <div style="font-family:monospace;font-weight:bold;margin-top:4px">
    {note['note_no']} | {note['created_at']}
  </div>
</div>
<div class="meta">
  <div><b>Route:</b><br>{note['route']}</div>
  <div><b>From:</b><br>{note['from_plant']} - {note['from_dept']}</div>
  <div><b>To:</b><br>{note['to_plant']} - {note['to_dept']}</div>
  <div><b>Entries:</b><br>{note['entry_count']}</div>
  <div><b>Sender:</b><br>{note['from_user_name']}</div>
  <div><b>Recipient:</b><br>{note['to_user_name']}</div>
  <div><b>Status:</b><br>{note['status'].upper()}</div>
  <div><b>Receipt:</b><br>{note['received_at'] or 'Pending'}</div>
</div>
<div style="font-size:11px;margin-bottom:10px"><b>CC:</b> {note.get('cc_users') or '—'}</div>
<table><thead><tr>{th}</tr></thead><tbody>{trs}</tbody></table>
<div class="signs">
  <div class="sign-box">Handovered By ({note['from_user_name']})</div>
  <div class="sign-box">Production | QC</div>
  <div class="sign-box">Received By ({note['received_by_name'] or note['to_user_name']})</div>
</div>
</body></html>""")


# =========================================================================
# MRQ1 — MATERIAL REQUISITION (kept from v5.0)
# =========================================================================
MRQ_CATEGORIES = {
    "stores": {"title": "Stores Issue", "icon": "ri-store-2-line",
               "desc": "General consumables & spares from Stores"},
    "inhouse": {"title": "Inhouse Material Issue", "icon": "ri-home-gear-line",
                "desc": "In-house fabricated / semi-finished material"},
    "conductor": {"title": "Conductor Issue", "icon": "ri-plug-line",
                  "desc": "Copper / aluminium conductor stock"},
}


@app.api_route("/api/mrq/categories", methods=["GET"])
def mrq_categories(req: Request):
    need(req)
    return MRQ_CATEGORIES


@app.api_route("/api/mrq/materials", methods=["GET"])
def mrq_list_materials(req: Request, category: str = ""):
    need(req)
    sql = "SELECT * FROM mrq_materials"
    params = []
    if category:
        sql += " WHERE category=?"
        params.append(category)
    sql += " ORDER BY material_code"
    return fetch_all(sql, tuple(params))


@app.api_route("/api/mrq/materials/import", methods=["POST"])
async def mrq_import_materials(req: Request, file: UploadFile = File(...)):
    need_admin(req)
    contents = await file.read()
    fn = file.filename.lower()
    rows = []
    try:
        if fn.endswith(".xlsx") or contents.startswith(b"PK\x03\x04"):
            wb = load_workbook(io.BytesIO(contents), data_only=True)
            all_r = list(wb.active.iter_rows(values_only=True))
            if all_r:
                h = [str(x).strip().lower() for x in all_r[0]]
                for r in all_r[1:]:
                    rows.append(dict(zip(h, r)))
        elif fn.endswith(".csv"):
            lines = contents.decode("utf-8-sig", errors="ignore").splitlines()
            if lines:
                h = [x.strip().lower().strip('"') for x in lines[0].split(",")]
                for line in lines[1:]:
                    if line.strip():
                        rows.append(dict(zip(h, next(csv.reader([line])))))
    except Exception as e:
        raise HTTPException(400, f"Error: {e}")
    imp, upd = 0, 0
    for r in rows:
        code = str(r.get("material code") or r.get("code") or "").strip().upper()
        desc = str(r.get("description") or r.get("material description") or "").strip()
        cat = str(r.get("category") or "stores").strip().lower()
        uom = str(r.get("uom") or "NOS").strip().upper()
        if not code or not desc:
            continue
        if fetch_one("SELECT 1 FROM mrq_materials WHERE material_code=?", (code,)):
            exec_write("UPDATE mrq_materials SET description=?, category=?, uom=? "
                       "WHERE material_code=?", (desc, cat, uom, code))
            upd += 1
        else:
            exec_write("INSERT INTO mrq_materials (material_code, description, "
                       "category, uom, created_at) VALUES (?,?,?,?,?)",
                       (code, desc, cat, uom, now()))
            imp += 1
    return {"ok": True, "imported": imp, "updated": upd, "total_rows": len(rows)}


@app.api_route("/api/mrq/requisitions", methods=["GET"])
def mrq_list(req: Request, category: str = "", status: str = ""):
    user = need(req)
    sql = "SELECT * FROM mrq_requisitions WHERE 1=1"
    params = []
    if not (is_master(user) or user["role"] == "admin"):
        sql += " AND (from_dept = ? OR to_dept = ?)"
        params.extend([user["department"], user["department"]])
    if category:
        sql += " AND category = ?"
        params.append(category)
    if status:
        sql += " AND status = ?"
        params.append(status)
    sql += " ORDER BY id DESC LIMIT 300"
    rows = fetch_all(sql, tuple(params))
    for r in rows:
        agg = fetch_one(
            "SELECT COUNT(*) AS c, COALESCE(SUM(qty_required),0) AS q "
            "FROM mrq_lines WHERE req_id=?", (r["id"],)
        )
        r["line_count"] = agg["c"]
        r["total_qty"] = agg["q"]
    return rows


@app.api_route("/api/mrq/requisition/{rid}", methods=["GET"])
def mrq_get(req: Request, rid: int):
    need(req)
    r = fetch_one("SELECT * FROM mrq_requisitions WHERE id=?", (rid,))
    if not r:
        raise HTTPException(404, "Not found.")
    r["lines"] = fetch_all("SELECT * FROM mrq_lines WHERE req_id=? ORDER BY sort_order, id", (rid,))
    return r


@app.api_route("/api/mrq/requisition", methods=["POST"])
def mrq_create(req: Request, payload: dict):
    user = need(req)
    category = str(payload.get("category") or "").strip().lower()
    req_no = f"MRQ-{datetime.date.today().strftime('%Y%m%d')}-{secrets.token_hex(2).upper()}"
    exec_write(
        "INSERT INTO mrq_requisitions (req_no, category, from_dept, from_plant, "
        "to_dept, status, remarks, requested_by, requested_by_name, "
        "created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
        (req_no, category, user["department"], "1000", "Store", "pending",
         payload.get("remarks", ""), user["username"], user["full_name"], now(), now()),
    )
    r_id = fetch_one("SELECT id FROM mrq_requisitions WHERE req_no=?", (req_no,))["id"]
    for i, ln in enumerate(payload.get("lines") or []):
        mat = fetch_one("SELECT * FROM mrq_materials WHERE material_code=?",
                        (ln.get("material_code", "").upper(),))
        if mat:
            exec_write(
                "INSERT INTO mrq_lines (req_id, material_code, description, "
                "qty_required, batch_no, qty_issued, uom, sort_order) "
                "VALUES (?,?,?,?,?,?,?,?)",
                (r_id, mat["material_code"], mat["description"],
                 float(ln.get("qty_required") or 0), "", 0, mat["uom"], i),
            )
    return {"ok": True, "id": r_id, "req_no": req_no}


@app.api_route("/api/mrq/requisition/{rid}/issue", methods=["POST"])
def mrq_issue(req: Request, rid: int, payload: dict = None):
    user = need(req)
    payload = payload or {}
    for ln in payload.get("lines") or []:
        exec_write("UPDATE mrq_lines SET batch_no=?, qty_issued=? WHERE id=?",
                   (str(ln.get("batch_no") or "").strip(),
                    float(ln.get("qty_issued") or 0), ln["id"]))
    exec_write(
        "UPDATE mrq_requisitions SET status='issued', issued_by=?, "
        "issued_by_name=?, issued_to_name=?, updated_at=? WHERE id=?",
        (user["username"], user["full_name"],
         str(payload.get("issued_to_name") or "").strip(), now(), rid),
    )
    return {"ok": True}


@app.api_route("/api/mrq/requisition/{rid}/receive", methods=["POST"])
def mrq_receive(req: Request, rid: int):
    user = need(req)
    exec_write(
        "UPDATE mrq_requisitions SET status='received', received_by=?, "
        "received_by_name=?, updated_at=? WHERE id=?",
        (user["username"], user["full_name"], now(), rid),
    )
    return {"ok": True}


@app.api_route("/api/mrq/requisition/{rid}", methods=["DELETE"])
def mrq_delete(req: Request, rid: int):
    need_admin(req)
    exec_write("DELETE FROM mrq_lines WHERE req_id=?", (rid,))
    exec_write("DELETE FROM mrq_requisitions WHERE id=?", (rid,))
    return {"ok": True}


@app.api_route("/api/mrq/print/{rid}", methods=["GET"], response_class=HTMLResponse)
def mrq_print(req: Request, rid: int):
    r = fetch_one("SELECT * FROM mrq_requisitions WHERE id=?", (rid,))
    if not r:
        return HTMLResponse("Not found", 404)
    lines = fetch_all("SELECT * FROM mrq_lines WHERE req_id=? ORDER BY sort_order, id", (rid,))
    trs = "".join(
        f"<tr><td style='border:1px solid #cbd5e1;padding:6px'>{i}</td>"
        f"<td style='border:1px solid #cbd5e1;padding:6px;font-family:monospace'>"
        f"{esc_py(l['material_code'])}</td>"
        f"<td style='border:1px solid #cbd5e1;padding:6px'>{esc_py(l['description'])}</td>"
        f"<td style='border:1px solid #cbd5e1;padding:6px;text-align:right'>"
        f"{l['qty_required']} {l['uom']}</td>"
        f"<td style='border:1px solid #cbd5e1;padding:6px'>{esc_py(l['batch_no'])}</td>"
        f"<td style='border:1px solid #cbd5e1;padding:6px;text-align:right'>"
        f"{l['qty_issued']} {l['uom']}</td></tr>"
        for i, l in enumerate(lines, 1)
    )
    return HTMLResponse(
        f"""<!DOCTYPE html><html><head><meta charset="utf-8">
<title>{r['req_no']}</title>
<style>body {{ font-family: sans-serif; margin: 30px; }}
table {{ width:100%; border-collapse: collapse; }}</style></head><body>
<h2>Material Requisition {r['req_no']}</h2>
<p>Dept: {r['from_dept']} | Status: {r['status'].upper()}</p>
<table border="1"><thead><tr><th>#</th><th>Code</th><th>Description</th>
<th>Qty Req</th><th>Batch</th><th>Qty Issued</th></tr></thead>
<tbody>{trs}</tbody></table></body></html>"""
    )


# =========================================================================
# STANDARD MODULE RECORDS (layouts + records + export)
# =========================================================================
@app.api_route("/api/tcode/{code}/layouts", methods=["GET"])
def get_layouts(code: str, req: Request):
    user = need(req)
    if not check_tcode_access(user, code):
        raise HTTPException(403, "Access denied.")
    rows = fetch_all(
        "SELECT * FROM tcode_layouts WHERE tcode_code=? "
        "ORDER BY is_default DESC, id ASC", (code,)
    )
    for r in rows:
        try:
            r["fields"] = json.loads(r["fields"])
        except Exception:
            r["fields"] = []
    return rows


@app.api_route("/api/tcode/{code}/records", methods=["GET"])
def get_tcode_records(code: str, req: Request, plant: str = "", limit: int = 500):
    user = need(req)
    if not check_tcode_access(user, code):
        raise HTTPException(403, "Access denied.")
    sql = "SELECT * FROM tcode_records WHERE tcode_code=? AND status='active'"
    params = [code]
    if plant:
        sql += " AND plant=?"
        params.append(plant)
    sql += " ORDER BY id DESC LIMIT ?" if not USE_PG else " ORDER BY id DESC LIMIT %s"
    params.append(limit)
    rows = fetch_all(sql, tuple(params))
    for r in rows:
        try:
            r["data"] = json.loads(r["data"])
        except Exception:
            r["data"] = {}
    return rows


@app.api_route("/api/tcode/{code}/records", methods=["POST"])
def create_record(code: str, req: Request, payload: dict):
    user = need(req)
    if not check_tcode_access(user, code):
        raise HTTPException(403, "Access denied.")
    plant = payload.pop("_plant", "").strip()
    exec_write(
        "INSERT INTO tcode_records (tcode_code, plant, data, status, "
        "created_by, created_by_name, created_at, updated_at) "
        "VALUES (?,?,?,?,?,?,?,?)",
        (code, plant, json.dumps(payload, default=json_serial), "active",
         user["username"], user["full_name"], now(), now()),
    )
    return {"ok": True}


@app.api_route("/api/tcode/{code}/records/{rid}", methods=["DELETE"])
def delete_record(code: str, rid: int, req: Request):
    user = need(req)
    if not check_tcode_access(user, code):
        raise HTTPException(403, "Access denied.")
    rec = fetch_one("SELECT * FROM tcode_records WHERE id=? AND tcode_code=?", (rid, code))
    if not rec:
        raise HTTPException(404, "Record not found.")
    if (rec["created_by"] != user["username"]
            and not is_master(user) and user["role"] != "admin"):
        raise HTTPException(403, "You can only delete your own records.")
    exec_write("UPDATE tcode_records SET status='deleted', updated_at=? WHERE id=?",
               (now(), rid))
    return {"ok": True}


@app.api_route("/api/tcode/{code}/export", methods=["GET"])
def export_excel(code: str, req: Request):
    user = need(req)
    if not check_tcode_access(user, code):
        raise HTTPException(403, "Access denied.")
    rows = fetch_all(
        "SELECT * FROM tcode_records WHERE tcode_code=? AND status='active' "
        "ORDER BY id ASC", (code,)
    )
    wb = Workbook()
    ws = wb.active
    ws.title = f"{code}_DATA"
    base_keys = ["id", "plant", "created_by_name", "created_at"]
    dynamic_keys = set()
    parsed_rows = []
    for r in rows:
        try:
            d = json.loads(r["data"]) if isinstance(r["data"], str) else r["data"]
        except Exception:
            d = {}
        dynamic_keys.update(d.keys())
        parsed_rows.append((r, d))
    all_cols = base_keys + sorted(dynamic_keys)
    ws.append(all_cols)
    for r, d in parsed_rows:
        ws.append([r["id"], r["plant"], r["created_by_name"], r["created_at"]]
                  + [d.get(k, "") for k in sorted(dynamic_keys)])
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition":
                 f"attachment; filename=Export_{code}_"
                 f"{datetime.date.today().strftime('%Y%m%d')}.xlsx"},
    )


# =========================================================================
# FRONTEND — single-file HTML UI (v3.9 shell + v5.0 backports)
# =========================================================================
HTML_UI = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Nexus Core — Enterprise Operations</title>
<script src="https://cdn.tailwindcss.com"></script>
<link href="https://cdn.jsdelivr.net/npm/remixicon@4.2.0/fonts/remixicon.css" rel="stylesheet">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">
<style>
body { font-family:'Inter',system-ui,sans-serif; background:#f8fafc; color:#0f172a; font-size:13px; }
.font-mono { font-family:'JetBrains Mono',monospace; }
.btn-primary { background:#4f46e5; color:#fff; border:1px solid #4f46e5; }
.btn-primary:hover { background:#4338ca; }
.btn-secondary { background:#fff; color:#475569; border:1px solid #e2e8f0; }
.btn-secondary:hover { background:#f1f5f9; border-color:#cbd5e1; }
.card { background:#fff; border:1px solid #e2e8f0; border-radius:14px; }
.custom-table { width:100%; border-collapse:collapse; }
.custom-table thead th { background:#f1f5f9; color:#475569; font-size:11px; font-weight:600;
  text-transform:uppercase; padding:10px 14px; border-bottom:1px solid #e2e8f0; text-align:left; }
.custom-table tbody td { padding:11px 14px; border-bottom:1px solid #f1f5f9;
  font-size:13px; color:#334155; }
.excel-grid { border-collapse:separate; border-spacing:0; width:100%; }
.excel-grid th { background:#f1f5f9; color:#475569; font-weight:600; font-size:11px;
  text-transform:uppercase; padding:8px 10px; border-right:1px solid #e2e8f0;
  border-bottom:1px solid #e2e8f0; position:sticky; top:0; z-index:2; text-align:left; }
.excel-grid td { background:#fff; color:#0f172a; font-size:12px; padding:6px 10px;
  border-right:1px solid #f1f5f9; border-bottom:1px solid #f1f5f9; outline:none; min-width:110px; }
.excel-grid td.sr-col { background:#f8fafc; color:#64748b; font-weight:600; text-align:center;
  width:50px; user-select:none; }
.excel-grid td:focus { background:#eef2ff; box-shadow:inset 0 0 0 2px #6366f1; }
.form-input { background:#fff; border:1px solid #cbd5e1; border-radius:8px; padding:8px 12px;
  font-size:13px; width:100%; outline:none; }
.form-input:focus { border-color:#6366f1; }
.form-label { display:block; font-size:11px; font-weight:600; color:#475569; margin-bottom:5px; }
.badge { display:inline-flex; align-items:center; font-size:10px; font-weight:600;
  padding:2px 8px; border-radius:999px; text-transform:uppercase; }
.badge-indigo { background:#eef2ff; color:#4338ca; border:1px solid #c7d2fe; }
.badge-emerald { background:#ecfdf5; color:#047857; border:1px solid #a7f3d0; }
.badge-amber { background:#fffbeb; color:#b45309; border:1px solid #fde68a; }
.badge-red { background:#fef2f2; color:#b91c1c; border:1px solid #fecaca; }
.badge-slate { background:#f1f5f9; color:#475569; border:1px solid #e2e8f0; }
.sidebar-link { color:#475569; border-left:2px solid transparent; transition:all .15s; }
.sidebar-link:hover { background:#f1f5f9; color:#0f172a; }
.sidebar-link.active { background:#eef2ff; color:#4f46e5; border-left-color:#4f46e5; }
</style>
</head>
<body class="h-screen flex flex-col overflow-hidden">
<div id="app_root" class="h-full flex flex-col">
  <header class="h-14 bg-white border-b border-slate-200 px-4 flex items-center justify-between shrink-0 shadow-sm">
    <div class="flex items-center gap-4">
      <div class="flex items-center gap-2.5 cursor-pointer" onclick="goHome()">
        <div class="w-8 h-8 rounded-lg bg-indigo-600 flex items-center justify-center font-bold text-white shadow-sm">
          <i class="ri-pulse-fill text-lg"></i>
        </div>
        <div class="font-bold text-sm tracking-tight text-slate-900">
          Nexus Core <span class="text-[10px] bg-indigo-50 text-indigo-600 border border-indigo-200 font-mono px-1.5 py-0.5 rounded">v5.0</span>
        </div>
      </div>
      <div class="relative ml-4">
        <i class="ri-search-line absolute left-3 top-2.5 text-slate-400 text-xs"></i>
        <input id="cmd_box" placeholder="Jump to module..."
          class="bg-slate-50 border border-slate-200 rounded-lg pl-8 pr-4 py-1.5 text-xs text-slate-900 w-96 font-mono focus:outline-none focus:border-indigo-400 focus:bg-white"
          onkeydown="if(event.key==='Enter') execCmd()">
      </div>
    </div>
    <div class="flex items-center gap-3">
      <div class="flex items-center gap-2 bg-slate-50 border border-slate-200 px-3 py-1 rounded-full text-xs">
        <div class="w-2 h-2 rounded-full bg-emerald-500"></div>
        <span id="user_display" class="font-medium text-slate-700">Connecting...</span>
        <span id="role_badge" class="badge badge-indigo"></span>
      </div>
      <button onclick="logout()" class="text-xs bg-white hover:bg-red-50 text-slate-500 hover:text-red-600 border border-slate-200 p-2 rounded-lg transition">
        <i class="ri-logout-box-r-line"></i>
      </button>
    </div>
  </header>
  <div class="flex-1 flex overflow-hidden">
    <aside class="w-64 bg-white border-r border-slate-200 flex flex-col justify-between shrink-0">
      <div class="p-3 overflow-y-auto flex-1 space-y-4">
        <div>
          <div class="text-[10px] uppercase font-bold text-slate-400 tracking-wider px-2 mb-1.5">Primary Modules</div>
          <div class="space-y-0.5">
            <button onclick="openTcode('CTST')" class="w-full text-left px-2.5 py-2 rounded-lg text-xs font-semibold bg-indigo-50 border border-indigo-200 text-indigo-700 hover:bg-indigo-100 flex items-center justify-between transition">
              <span class="flex items-center gap-2"><i class="ri-flask-line text-indigo-500 text-sm"></i> CTST - Type-Test</span>
              <span class="badge badge-indigo">NEW</span>
            </button>
            <button onclick="openTcode('TN01')" class="w-full text-left px-2.5 py-2 rounded-lg text-xs font-semibold bg-slate-50 border border-slate-200 text-slate-700 hover:bg-slate-100 flex items-center justify-between transition">
              <span class="flex items-center gap-2"><i class="ri-exchange-box-line text-slate-500 text-sm"></i> TN01 - Transfer Note</span>
              <span class="badge badge-slate">EXCEL</span>
            </button>
            <button onclick="openTcode('MRQ1')" class="w-full text-left px-2.5 py-2 rounded-lg text-xs font-semibold bg-emerald-50 border border-emerald-200 text-emerald-700 hover:bg-emerald-100 flex items-center justify-between transition mt-1">
              <span class="flex items-center gap-2"><i class="ri-file-list-3-line text-emerald-500 text-sm"></i> MRQ1 - Requisition</span>
              <span class="badge badge-emerald">NEW</span>
            </button>
          </div>
        </div>
        <div id="admin_nav_section" class="hidden">
          <div class="text-[10px] uppercase font-bold text-indigo-500 tracking-wider px-2 mb-1.5 flex items-center justify-between">
            <span>Administration</span><i class="ri-shield-keyhole-line text-xs"></i>
          </div>
          <div class="space-y-0.5">
            <button onclick="openTcode('STUDIO')" class="sidebar-link w-full text-left px-2.5 py-1.5 rounded-lg text-xs font-medium flex items-center justify-between">
              <span class="flex items-center gap-2"><i class="ri-robot-2-line text-emerald-500"></i> AI Code Studio</span>
              <span class="badge badge-emerald">SAFE</span>
            </button>
            <button onclick="openTcode('SU01')" class="sidebar-link w-full text-left px-2.5 py-1.5 rounded-lg text-xs font-medium flex items-center justify-between">
              <span class="flex items-center gap-2"><i class="ri-user-settings-line text-indigo-500"></i> User Directory (SU01)</span>
              <span class="badge badge-indigo">IAM</span>
            </button>
            <button onclick="openTcode('SPRO')" class="sidebar-link w-full text-left px-2.5 py-1.5 rounded-lg text-xs font-medium flex items-center justify-between">
              <span class="flex items-center gap-2"><i class="ri-building-line text-slate-500"></i> Units & Plants (SPRO)</span>
              <span class="badge badge-slate">MGMT</span>
            </button>
            <button onclick="openTcode('SE93')" class="sidebar-link w-full text-left px-2.5 py-1.5 rounded-lg text-xs font-medium flex items-center justify-between">
              <span class="flex items-center gap-2"><i class="ri-apps-2-line text-slate-500"></i> Module Studio (SE93)</span>
              <span class="badge badge-slate">AUTO</span>
            </button>
            <button onclick="openTcode('PFCG')" class="sidebar-link w-full text-left px-2.5 py-1.5 rounded-lg text-xs font-medium flex items-center gap-2">
              <i class="ri-shield-user-line text-slate-500"></i> Access Policies (PFCG)
            </button>
          </div>
        </div>
        <div>
          <div class="text-[10px] uppercase font-bold text-slate-400 tracking-wider px-2 mb-1.5">Operations Modules</div>
          <div id="modules_list" class="space-y-0.5 text-xs"></div>
        </div>
      </div>
      <div class="p-3 border-t border-slate-200 bg-slate-50/60 text-[11px]">
        <div class="font-bold text-slate-700 truncate" id="side_company">Thermocables Limited</div>
        <div class="text-slate-500 truncate" id="side_plant">Facilities: Authorized</div>
      </div>
    </aside>
    <main id="main_screen" class="flex-1 bg-slate-50 overflow-hidden flex flex-col"></main>
  </div>
  <footer class="h-7 bg-white border-t border-slate-200 px-4 flex items-center justify-between text-[11px] text-slate-500 font-mono shrink-0">
    <div class="flex items-center gap-2 truncate">
      <span id="stat_dot" class="w-2 h-2 rounded-full bg-emerald-500"></span>
      <span id="stat_msg" class="text-slate-600">Ready.</span>
    </div>
    <div class="flex items-center gap-4">
      <span id="stat_tcode" class="text-indigo-600 font-semibold font-mono">DASHBOARD</span>
      <span id="stat_resp">—ms</span>
      <span>READY</span>
    </div>
  </footer>
</div>
<div id="app_dialog" class="hidden fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center p-4">
  <div id="app_dialog_body" class="bg-white text-slate-800 rounded-2xl shadow-2xl border border-slate-200 max-h-[90vh] overflow-y-auto w-full max-w-4xl"></div>
</div>
<script>
// ===== GLOBAL STATE =====
let ME = null, CURRENT_TCODE = null, ALL_TCODES = [], ALL_USERS_LIST = [];
const $ = s => document.querySelector(s);
function esc(str){ if(str===null||str===undefined) return '';
  return String(str).replace(/[&<>"']/g, s => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[s])); }
function setMsg(m, e=false){ const el=$('#stat_msg'), d=$('#stat_dot'); if(!el) return;
  el.textContent=m; d.className = e ? "w-2 h-2 rounded-full bg-red-500" : "w-2 h-2 rounded-full bg-emerald-500";
  el.className = e ? "text-red-600 font-semibold" : "text-slate-600"; }

async function api(p, o={}){
  const t = localStorage.getItem("nexus_token");
  const h = {...(o.headers || {})};
  if(t) h["Authorization"] = `Bearer ${t}`;
  const t0 = performance.now();
  const r = await fetch(p, {...o, headers:h, credentials:'include'});
  if($('#stat_resp')) $('#stat_resp').textContent = Math.round(performance.now()-t0)+'ms';
  if(r.status === 401){ localStorage.removeItem("nexus_token"); renderLogon(); throw 'Auth'; }
  if(!r.ok){ let x = await r.text(); try { x = JSON.parse(x).detail || x; } catch {} throw x; }
  return await r.json();
}

// ===== LOGON (no credentials shown) =====
function renderLogon(err=""){
  $('#app_root').innerHTML = `
  <div class="min-h-screen bg-gradient-to-br from-slate-50 to-indigo-50 flex items-center justify-center p-4">
    <div class="w-full max-w-md bg-white border border-slate-200 rounded-2xl shadow-xl p-8">
      <div class="flex items-center gap-3 mb-6">
        <div class="w-10 h-10 rounded-xl bg-indigo-600 flex items-center justify-center font-bold text-white shadow-sm">
          <i class="ri-pulse-fill text-2xl"></i>
        </div>
        <div>
          <h1 class="text-xl font-bold text-slate-900">Nexus Core ERP</h1>
          <p class="text-xs text-slate-500">Thermocables Industrial Suite</p>
        </div>
      </div>
      ${err ? `<div class="mb-4 p-3 bg-red-50 border border-red-200 rounded-xl text-red-700 text-xs">${esc(err)}</div>` : ''}
      <div class="space-y-4">
        <div>
          <label class="form-label">Username</label>
          <input id="l_u" class="form-input" autofocus>
        </div>
        <div>
          <label class="form-label">Password</label>
          <input id="l_p" type="password" class="form-input" onkeydown="if(event.key==='Enter') handleLogon()">
        </div>
        <button onclick="handleLogon()" class="w-full btn-primary py-2.5 rounded-lg font-semibold text-sm">Sign In</button>
      </div>
    </div>
  </div>`;
}

async function handleLogon(){
  const u = $('#l_u').value.trim(), p = $('#l_p').value;
  if(!u || !p) return renderLogon("Enter credentials.");
  const fd = new FormData();
  fd.append("username", u); fd.append("password", p);
  try {
    const r = await fetch('/api/login', { method:'POST', body:fd });
    if(!r.ok){ let t = await r.text(); try { t = JSON.parse(t).detail || t; } catch {} return renderLogon(t); }
    const d = await r.json();
    if(d.token) localStorage.setItem("nexus_token", d.token);
    location.reload();
  } catch(e) { renderLogon("Error: " + e); }
}

async function logout(){
  await api('/api/logout', { method:'POST' });
  localStorage.removeItem("nexus_token");
  location.reload();
}

async function initSession(){
  try {
    ME = await api('/api/me');
    $('#user_display').textContent = ME.full_name || ME.username;
    $('#role_badge').textContent = ME.role;
    $('#side_company').textContent = ME.company;
    $('#side_plant').textContent = "Facilities: " + (ME.plants.join(", ") || "All");
    if(ME.is_master || ME.role === 'admin') $('#admin_nav_section').classList.remove('hidden');
    setMsg(`Connected: ${ME.username}`);
    await loadModules();
    goHome();
  } catch(e) { renderLogon(); }
}

function execCmd(){
  let v = $('#cmd_box').value.trim(); if(!v) return;
  $('#cmd_box').value = '';
  if(v.startsWith('/n')) v = v.substring(2).trim();
  if(!v) { goHome(); return; }
  openTcode(v.toUpperCase());
}

function goHome(){
  CURRENT_TCODE = null;
  $('#stat_tcode').textContent = "DASHBOARD";
  $('#main_screen').innerHTML = `
    <div class="flex-1 overflow-y-auto p-8 select-text">
      <div class="max-w-5xl mx-auto space-y-8">
        <div>
          <h1 class="text-2xl font-bold text-slate-900">Operations Dashboard</h1>
          <p class="text-sm text-slate-500 mt-1">Thermocables Industrial Platform</p>
        </div>
        <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div class="card p-5">
            <div class="text-[10px] font-semibold text-slate-500 uppercase tracking-wider">Enterprise</div>
            <div class="text-lg font-bold text-slate-900 mt-1.5">${esc(ME.company)}</div>
            <div class="text-xs text-indigo-600 mt-1">Dept: ${esc(ME.department)}</div>
          </div>
          <div class="card p-5">
            <div class="text-[10px] font-semibold text-slate-500 uppercase tracking-wider">Role</div>
            <div class="text-lg font-bold text-emerald-600 mt-1.5 uppercase">${ME.role}</div>
            <div class="text-xs text-slate-500 mt-1">Plants: ${esc(ME.plants.join(', ') || 'All')}</div>
          </div>
          <div class="card p-5">
            <div class="text-[10px] font-semibold text-slate-500 uppercase tracking-wider">Quick</div>
            <div class="text-xs text-slate-600 mt-2 font-mono">
              Launch <b class="text-indigo-600">CTST</b>, <b class="text-indigo-600">TN01</b>, <b class="text-indigo-600">MRQ1</b> or <b class="text-indigo-600">STUDIO</b>.
            </div>
          </div>
        </div>
        <div>
          <h3 class="text-xs font-bold text-slate-500 uppercase tracking-wider mb-4">Modules</h3>
          <div id="quick_grid" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4"></div>
        </div>
      </div>
    </div>`;
  renderQuickGrid();
}

async function loadModules(){
  ALL_TCODES = await api('/api/se93/list');
  $('#modules_list').innerHTML = ALL_TCODES.filter(t => t.has_access).map(t => `
    <button onclick="openTcode('${t.code}')" class="sidebar-link w-full text-left px-2.5 py-1.5 rounded-lg flex items-center justify-between">
      <span class="truncate"><b class="text-indigo-600 font-mono">${t.code}</b> <span class="text-slate-600">- ${esc(t.title)}</span></span>
    </button>`).join('');
}
function renderQuickGrid(){
  const c = $('#quick_grid'); if(!c) return;
  c.innerHTML = ALL_TCODES.filter(t => t.has_access).map(t => `
    <div onclick="openTcode('${t.code}')" class="card hover:border-indigo-300 hover:shadow-md p-5 cursor-pointer transition-all">
      <div class="flex items-center justify-between mb-2">
        <span class="font-mono font-bold text-xs bg-indigo-50 text-indigo-600 border border-indigo-200 px-2 py-0.5 rounded-lg">${t.code}</span>
        <span class="badge badge-slate">${t.module || 'ERP'}</span>
      </div>
      <h4 class="font-semibold text-slate-900 text-sm mt-2">${esc(t.title)}</h4>
      <p class="text-xs text-slate-500 mt-1 line-clamp-2">${esc(t.description || 'Standard module')}</p>
    </div>`).join('');
}

async function openTcode(code){
  code = code.toUpperCase().trim();
  CURRENT_TCODE = code;
  $('#stat_tcode').textContent = code;
  setMsg(`Module ${code}`);
  if(code === 'CTST')   return renderCTST();
  if(code === 'TN01')   return renderTN01();
  if(code === 'MRQ1')   return renderMRQ();
  if(code === 'SU01')   return renderSU01();
  if(code === 'SPRO')   return renderSPRO();
  if(code === 'SE93')   return renderSE93();
  if(code === 'PFCG')   return renderPFCG();
  if(code === 'STUDIO') return renderStudio();
  renderStandardModule(code);
}

// ===== CTST HUB =====
const CTST_TESTS = {
  cr:       {title:"CR — Conductor Resistance", icon:"ri-pulse-line",        desc:"Auto per-km CR + HIGH/LOW verdict vs spec"},
  braid:    {title:"Braid / Shield Coverage",   icon:"ri-shield-check-line", desc:"Fill factor and % coverage"},
  tubular:  {title:"Tubular Mechanical Test",   icon:"ri-circle-line",       desc:"Area, tensile, elongation (tube)"},
  dumbbell: {title:"Dumbbell Mechanical Test",  icon:"ri-shape-line",        desc:"Area, tensile, elongation (dumbbell)"},
  hotset:   {title:"Hot Set & Permanent Set",   icon:"ri-temp-hot-line",     desc:"Hot-set & permanent-set elongation"},
};
const CTST_FIELDS = {
  cr: [{key:"obd_cr_reading",label:"OBD CR Reading (Ω)",step:"any"},{key:"temp_c",label:"Test Temp (°C)",step:"any"},{key:"length_m",label:"Length (m)",step:"any"},{key:"specified_cr",label:"Specified CR (Ω/km)",step:"any"},{key:"cr_tolerance_pct",label:"Tolerance (± %)",step:"any",value:"5"}],
  braid: [{key:"braid_d",label:"Braid Wire Ø (mm)",step:"any"},{key:"under_braid_D",label:"Ø Under Braid (mm)",step:"any"},{key:"wires_per_carrier",label:"Wires/Carrier",step:"1"},{key:"carriers",label:"Carriers",step:"1"},{key:"lay_length",label:"Lay Length (mm)",step:"any"}],
  tubular: [{key:"tub_outer_d",label:"Outer Ø (mm)",step:"any"},{key:"tub_wall",label:"Wall (mm)",step:"any"},{key:"break_load_tub",label:"Break Load (N)",step:"any"},{key:"initial_gauge",label:"Initial Gauge (mm)",step:"any",value:"20"},{key:"elong_tub_final",label:"Final Gauge (mm)",step:"any"}],
  dumbbell: [{key:"dumb_width",label:"Width (mm)",step:"any"},{key:"dumb_thk",label:"Thickness (mm)",step:"any"},{key:"break_load_dumb",label:"Break Load (N)",step:"any"},{key:"initial_gauge",label:"Initial Gauge (mm)",step:"any",value:"20"},{key:"elong_dumb_final",label:"Final Gauge (mm)",step:"any"}],
  hotset: [{key:"initial_gauge",label:"Initial Gauge (mm)",step:"any",value:"20"},{key:"hotset_elong_final",label:"Hot Set Final (mm)",step:"any"},{key:"permset_cooled",label:"Perm. Set Cooled (mm)",step:"any"}],
};

async function renderCTST(testType = null){
  if(testType) return renderCTSTForm(testType);
  const allRows = await api('/api/ctst/records');
  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col p-6 overflow-hidden select-text">
      <div class="flex items-center justify-between pb-4 border-b border-slate-200 shrink-0">
        <div class="flex items-center gap-3">
          <span class="font-mono font-bold text-xs bg-indigo-600 text-white px-2.5 py-1 rounded-lg shadow-sm">CTST</span>
          <div><h2 class="text-lg font-bold text-slate-900">Cable Type-Test Calculator</h2>
          <p class="text-xs text-slate-500">Pick a test — only fill what you need</p></div>
        </div>
        <div class="flex items-center gap-2">
          <a href="/api/ctst/export" class="btn-secondary text-xs font-semibold px-3 py-2 rounded-lg flex items-center gap-1">
            <i class="ri-download-2-line"></i> Export All
          </a>
          <button onclick="renderCTST()" class="btn-secondary text-xs font-semibold p-2 rounded-lg">
            <i class="ri-refresh-line"></i>
          </button>
        </div>
      </div>
      <div class="mt-4 grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 shrink-0">
        ${Object.entries(CTST_TESTS).map(([k,t]) => `
          <div onclick="renderCTST('${k}')" class="card hover:border-indigo-300 hover:shadow-md p-5 cursor-pointer transition-all group">
            <div class="flex items-start justify-between mb-2">
              <div class="w-10 h-10 rounded-xl bg-indigo-50 border border-indigo-200 flex items-center justify-center text-indigo-600 group-hover:bg-indigo-600 group-hover:text-white transition">
                <i class="${t.icon} text-lg"></i>
              </div>
              <span class="badge badge-indigo">${k.toUpperCase()}</span>
            </div>
            <h4 class="font-bold text-slate-900 text-sm mt-2">${t.title}</h4>
            <p class="text-xs text-slate-500 mt-1">${t.desc}</p>
          </div>
        `).join('')}
      </div>
      <div class="flex-1 mt-6 card overflow-hidden flex flex-col">
        <div class="p-3 border-b border-slate-200 bg-slate-50">
          <span class="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
            <i class="ri-history-line text-indigo-500"></i> Recent Records (${allRows.length})
          </span>
        </div>
        <div class="flex-1 overflow-auto">
          <table class="custom-table">
            <thead><tr><th>#</th><th>Type</th><th>Sample</th><th>Result</th><th>Actions</th></tr></thead>
            <tbody>
              ${!allRows.length ? `<tr><td colspan="5" class="text-center py-10 text-slate-400">No tests yet.</td></tr>` : ''}
              ${allRows.map(r => `
                <tr>
                  <td class="font-mono text-slate-400 font-semibold">${r.id}</td>
                  <td><span class="badge badge-indigo">${esc((r.data._test_type||'').toUpperCase())}</span></td>
                  <td class="font-semibold text-slate-900">${esc(r.data.sample_id || '—')}</td>
                  <td>${r.data.result ? `<span class="badge ${r.data.result==='PASS'?'badge-emerald':'badge-red'}">${esc(r.data.result)}</span>` : '—'}</td>
                  <td>
                    <a href="/api/ctst/print/${r.id}" target="_blank" class="text-xs text-indigo-600 font-semibold">Print</a>
                    <button onclick="deleteCTST(${r.id})" class="text-xs text-red-600 font-semibold ml-2">Del</button>
                  </td>
                </tr>`).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>`;
}

function renderCTSTForm(testType){
  const t = CTST_TESTS[testType], fields = CTST_FIELDS[testType] || [];
  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col p-6 overflow-hidden select-text">
      <div class="flex items-center justify-between pb-4 border-b border-slate-200 shrink-0">
        <div class="flex items-center gap-3">
          <button onclick="renderCTST()" class="btn-secondary text-xs p-2 rounded-lg"><i class="ri-arrow-left-line"></i></button>
          <div class="w-10 h-10 rounded-xl bg-indigo-600 flex items-center justify-center text-white shadow-sm">
            <i class="${t.icon} text-lg"></i>
          </div>
          <div><h2 class="text-lg font-bold text-slate-900">${t.title}</h2>
          <p class="text-xs text-slate-500">${t.desc}</p></div>
        </div>
        <a href="/api/ctst/export?test_type=${testType}" class="btn-secondary text-xs font-semibold px-3 py-2 rounded-lg flex items-center gap-1">
          <i class="ri-download-2-line"></i> Export ${testType.toUpperCase()}
        </a>
      </div>
      <div class="flex-1 mt-4 overflow-auto">
        <div class="max-w-3xl mx-auto">
          <div class="card p-6 space-y-4">
            <div><label class="form-label">Sample ID *</label><input id="cf_sample_id" class="form-input font-mono"></div>
            <div class="grid grid-cols-2 gap-4">
              ${fields.map(f => `
                <div><label class="form-label">${f.label}</label>
                  <input id="cf_${f.key}" type="number" step="${f.step||'any'}" value="${f.value||''}" class="form-input">
                </div>`).join('')}
            </div>
            <div class="grid grid-cols-2 gap-4">
              <div><label class="form-label">Result</label>
                <select id="cf_result" class="form-input">
                  <option value="">— Not judged —</option>
                  <option value="PASS">PASS</option><option value="FAIL">FAIL</option>
                  <option value="RETAIN">RETAIN</option>
                </select>
              </div>
              <div><label class="form-label">Facility</label>
                <input id="cf_plant" class="form-input font-mono" value="${esc((ME.plants && ME.plants[0]) || '1000')}">
              </div>
            </div>
            <div class="flex justify-end gap-2 pt-2 border-t border-slate-100">
              <button onclick="renderCTST()" class="btn-secondary px-4 py-2 rounded-lg text-xs font-semibold">Cancel</button>
              <button onclick="saveCTST('${testType}')" class="btn-primary px-5 py-2 rounded-lg text-xs font-bold shadow-sm">Save Test</button>
            </div>
          </div>
        </div>
      </div>
    </div>`;
}

async function saveCTST(testType){
  const fields = CTST_FIELDS[testType] || [], payload = { _test_type: testType };
  payload.sample_id = ($('#cf_sample_id') || {}).value || '';
  if(!payload.sample_id.trim()) return alert("Sample ID required.");
  payload.result = ($('#cf_result') || {}).value || '';
  payload._plant = ($('#cf_plant') || {}).value || '';
  for(const f of fields){ const el = $('#cf_'+f.key); if(el) payload[f.key] = el.value; }
  try {
    await api('/api/ctst/save', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(payload) });
    setMsg("Saved"); renderCTST();
  } catch(e) { alert(e); }
}
async function deleteCTST(id){
  if(!confirm(`Delete test #${id}?`)) return;
  try { await api(`/api/ctst/records/${id}`, {method:'DELETE'}); renderCTST(); setMsg("Deleted"); }
  catch(e) { alert(e); }
}

// ===== TN01 (v5.0 backport, full sheet + CC picker) =====
let TN01_ROUTES = {}, TN01_ACTIVE_ROUTE = "PROD_TO_QC", TN01_HEADERS = [],
    TN01_ROWS = [], TN01_UNDO = [], TN01_REDO = [], TN01_CELL = { r: 0, c: 1 };
let TN01_SELECTED_CC = [];
let isGridUpdating = false;

async function renderTN01(){ TN01_ROUTES = await api('/api/tn01/routes'); renderTN01Main(); }

function renderTN01Main(tab="sheet"){
  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col overflow-hidden select-text">
      <div class="p-3 bg-white border-b border-slate-200 flex items-center justify-between shrink-0">
        <div class="flex items-center gap-3">
          <span class="font-mono font-bold text-xs bg-indigo-600 text-white px-2.5 py-1 rounded-lg shadow-sm">TN01</span>
          <div class="flex items-center bg-slate-100 border border-slate-200 rounded-lg p-1 gap-1 text-xs font-semibold">
            <button onclick="switchTN01Tab('sheet')" class="px-3 py-1.5 rounded-md transition ${tab==='sheet'?'bg-white text-indigo-600 shadow-sm':'text-slate-500'}">Sheet</button>
            <button onclick="switchTN01Tab('receive')" class="px-3 py-1.5 rounded-md transition ${tab==='receive'?'bg-white text-indigo-600 shadow-sm':'text-slate-500'}">Receive</button>
            <button onclick="switchTN01Tab('history')" class="px-3 py-1.5 rounded-md transition ${tab==='history'?'bg-white text-indigo-600 shadow-sm':'text-slate-500'}">History</button>
          </div>
        </div>
        ${tab==='sheet' ? `<button onclick="openSendModal()" class="btn-primary text-xs font-bold px-4 py-2 rounded-lg flex items-center gap-1.5 shadow-sm">
          <i class="ri-send-plane-fill"></i> Send</button>` : ''}
      </div>
      <div id="tn01_content" class="flex-1 overflow-hidden flex flex-col"></div>
    </div>`;
  if(tab==='sheet') renderTN01SheetTab();
  else if(tab==='receive') renderTN01ReceiveTab();
  else if(tab==='history') renderTN01HistoryTab();
}
function switchTN01Tab(t){ renderTN01Main(t); }

function renderTN01SheetTab(){
  if(!TN01_HEADERS.length) applyRouteLayout(TN01_ACTIVE_ROUTE, false);
  $('#tn01_content').innerHTML = `
    <div class="p-2 bg-white border-b border-slate-200 flex items-center justify-between text-xs shrink-0">
      <div class="flex items-center gap-2">
        <label class="font-bold text-slate-600 text-xs">Route:
          <select id="route_selector" onchange="applyRouteLayout(this.value,true)" class="form-input w-auto py-1 font-semibold">
            ${Object.keys(TN01_ROUTES).map(k=>`<option value="${k}" ${k===TN01_ACTIVE_ROUTE?'selected':''}>${TN01_ROUTES[k].route}</option>`).join('')}
          </select>
        </label>
        <button onclick="addGridRow()" class="bg-white text-slate-600 px-2.5 py-1 rounded-md border border-slate-200 hover:bg-slate-50">+ Row</button>
        <button onclick="deleteGridRow()" class="bg-white text-red-500 px-2.5 py-1 rounded-md border border-slate-200 hover:bg-red-50">− Row</button>
      </div>
      <button onclick="exportGridAsExcel()" class="bg-white text-emerald-700 border border-emerald-200 px-3 py-1 rounded-md flex items-center gap-1 font-semibold hover:bg-emerald-50">
        <i class="ri-file-excel-2-line"></i> Export
      </button>
    </div>
    <div id="grid_table_container" class="flex-1 overflow-auto p-2 select-text" tabindex="0"></div>`;
  renderGridTable();
  setupPasteHandler();
}

function applyRouteLayout(k, r=true){
  TN01_ACTIVE_ROUTE = k;
  TN01_HEADERS = [...TN01_ROUTES[k].columns];
  if(r || !TN01_ROWS.length){
    TN01_ROWS = [];
    for(let i=1;i<=8;i++){
      const row = new Array(TN01_HEADERS.length).fill(""); row[0] = i; TN01_ROWS.push(row);
    }
  }
  renderGridTable();
}

function renderGridTable(){
  const c = $('#grid_table_container'); if(!c) return;
  isGridUpdating = true;
  if(document.activeElement && c.contains(document.activeElement)){ document.activeElement.blur(); }
  let h = `<table class="excel-grid"><thead><tr>${TN01_HEADERS.map(x=>`<th>${esc(x)}</th>`).join('')}</tr></thead><tbody>`;
  TN01_ROWS.forEach((row, r) => {
    h += `<tr>${row.map((cell, ci) => `
      <td contenteditable="${ci!==0}" class="${ci===0?'sr-col':''}"
          onfocus="TN01_CELL={r:${r},c:${ci}}"
          onblur="updateGridCell(${r},${ci},this.innerText)">${esc(cell)}</td>`).join('')}</tr>`;
  });
  h += `</tbody></table>`;
  c.innerHTML = h;
  setTimeout(()=>{ isGridUpdating = false; }, 50);
}

function updateGridCell(r, ci, v){
  if(isGridUpdating || ci === 0) return;
  if(!TN01_ROWS[r]) return;
  if(TN01_ROWS[r][ci] !== v.trim()){
    TN01_UNDO.push(JSON.stringify(TN01_ROWS));
    TN01_ROWS[r][ci] = v.trim();
  }
}
function addGridRow(){
  TN01_UNDO.push(JSON.stringify(TN01_ROWS));
  const r = new Array(TN01_HEADERS.length).fill(""); r[0] = TN01_ROWS.length + 1;
  TN01_ROWS.push(r); renderGridTable();
}
function deleteGridRow(){
  if(TN01_ROWS.length <= 1) return;
  TN01_ROWS.pop(); renderGridTable();
}

function setupPasteHandler(){
  const c = $('#grid_table_container'); if(!c) return;
  c.addEventListener('paste', e => {
    e.preventDefault();
    const clip = (e.clipboardData || window.clipboardData).getData('text');
    if(!clip) return;
    isGridUpdating = true;
    if(document.activeElement && c.contains(document.activeElement)){ document.activeElement.blur(); }
    TN01_UNDO.push(JSON.stringify(TN01_ROWS));
    const lines = clip.replace(/\r\n/g,'\n').replace(/\r/g,'\n').split('\n');
    let sr = (TN01_CELL && TN01_CELL.r !== undefined) ? TN01_CELL.r : 0;
    lines.forEach((line, li) => {
      if(!line && li === lines.length - 1) return;
      let cells = line.split('\t');
      if(cells.length > 1 && cells[0].trim() === "" && cells.slice(1).some(x => x.trim() !== "")){ cells.shift(); }
      const tr = sr + li;
      while(TN01_ROWS.length <= tr){
        const nr = new Array(TN01_HEADERS.length).fill(""); nr[0] = TN01_ROWS.length + 1;
        TN01_ROWS.push(nr);
      }
      cells.forEach((v, ci) => {
        const tc = ci + 1;
        if(tc < TN01_HEADERS.length){ TN01_ROWS[tr][tc] = v.trim(); }
      });
    });
    TN01_ROWS.forEach((r, i) => { r[0] = i + 1; });
    renderGridTable();
  });
}

async function exportGridAsExcel(){
  const f = TN01_ROWS.filter(r => r.slice(1).some(x => String(x).trim() !== ''));
  if(!f.length) return alert("No rows to export.");
  const t = TN01_ROUTES[TN01_ACTIVE_ROUTE] ? TN01_ROUTES[TN01_ACTIVE_ROUTE].route : "Draft";
  try {
    const tok = localStorage.getItem("nexus_token");
    const r = await fetch('/api/tn01/export-grid', {
      method:'POST', headers:{'Content-Type':'application/json','Authorization':tok?`Bearer ${tok}`:''},
      body: JSON.stringify({headers: TN01_HEADERS, rows: f, title: `TN01 Draft — ${t}`})
    });
    const b = await r.blob();
    const u = URL.createObjectURL(b);
    const a = document.createElement('a'); a.href = u; a.download = 'TN01_Draft.xlsx';
    document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(u);
  } catch(e) { alert("Failed: " + e); }
}

async function openSendModal(){
  const a = TN01_ROUTES[TN01_ACTIVE_ROUTE];
  const f = TN01_ROWS.filter(r => r.slice(1).some(x => String(x).trim() !== ''));
  if(!f.length) return alert("Enter at least one row before sending.");
  TN01_SELECTED_CC = [];
  const sp = await api('/api/spro/overview');
  openDialog(`
    <div class="px-6 py-4 border-b bg-slate-50 flex justify-between items-center rounded-t-2xl">
      <div><h3 class="font-bold text-sm">Send Transfer Note</h3><p class="text-xs text-slate-500">${a.route}</p></div>
      <span id="live_datetime" class="text-xs font-mono text-slate-500"></span>
    </div>
    <div class="p-6 space-y-4">
      <div class="grid grid-cols-3 gap-3">
        <div><label class="form-label">Plant</label>
          <select id="send_plant" onchange="updateSuggestedUsers()" class="form-input">
            ${sp.plants.map(p=>`<option value="${esc(p.code)}">${esc(p.code)} - ${esc(p.name)}</option>`).join('')}
          </select>
        </div>
        <div><label class="form-label">Dept</label>
          <select id="send_dept" onchange="updateSuggestedUsers()" class="form-input">
            ${sp.departments.map(d=>`<option value="${esc(d.name)}" ${d.name===a.to_dept?'selected':''}>${esc(d.name)}</option>`).join('')}
          </select>
        </div>
        <div><label class="form-label">Receiver *</label><select id="send_user" class="form-input"></select></div>
      </div>
      <div class="relative">
        <label class="form-label">CC (Outlook-style address book)</label>
        <div id="cc_tags" class="flex flex-wrap gap-1.5 p-2 bg-slate-50 border border-slate-200 rounded-lg min-h-[40px] items-center mb-1">
          <input id="cc_search_input" placeholder="Type user name or ID..." oninput="searchCCUsers(this.value)"
            class="bg-transparent text-xs outline-none flex-1 min-w-[140px]">
        </div>
        <div id="cc_suggestions" class="hidden absolute left-0 right-0 top-full bg-white border border-slate-200 shadow-xl rounded-lg z-50 max-h-48 overflow-y-auto"></div>
      </div>
      <div><label class="form-label">Remarks</label><textarea id="send_remarks" class="form-input" rows="2"></textarea></div>
    </div>
    <div class="px-6 py-4 bg-slate-50 border-t flex justify-end gap-2 rounded-b-2xl">
      <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-lg text-xs">Cancel</button>
      <button onclick="submitSendNote()" class="btn-primary px-5 py-2 rounded-lg text-xs font-bold">Send</button>
    </div>`);
  updateSuggestedUsers();
  updateLiveClock();
}

function updateLiveClock(){
  const el = $('#live_datetime'); if(!el) return;
  el.textContent = new Date().toLocaleString('en-GB',{dateStyle:'medium',timeStyle:'medium'});
  setTimeout(updateLiveClock, 1000);
}

async function updateSuggestedUsers(){
  const pl = $('#send_plant').value, dept = $('#send_dept').value;
  const u = await api(`/api/tn01/suggest-users?plant=${pl}&dept=${encodeURIComponent(dept)}`);
  $('#send_user').innerHTML = u.length
    ? u.map(x=>`<option value="${esc(x.username)}">${esc(x.full_name)} (${esc(x.username)})</option>`).join('')
    : '<option value="">(None)</option>';
}

async function searchCCUsers(val){
  const box = $('#cc_suggestions');
  if(!val.trim()){ box.classList.add('hidden'); box.innerHTML=''; return; }
  const list = await api(`/api/tn01/suggest-users?q=${encodeURIComponent(val)}`);
  box.innerHTML = list.length ? list.map(u => `
    <div onclick="addCCTag('${esc(u.username)}','${esc(u.full_name)}')" class="p-2.5 hover:bg-indigo-50 cursor-pointer border-b border-slate-100 flex items-center justify-between text-xs">
      <div><b class="text-slate-900">${esc(u.full_name)}</b>
        <span class="text-slate-400 font-mono">(${esc(u.username)})</span></div>
      <span class="badge badge-slate">${esc(u.department||'')}</span>
    </div>`).join('') : '<div class="p-2.5 text-xs text-slate-400">No match.</div>';
  box.classList.remove('hidden');
}

function addCCTag(username, fullname){
  if(!TN01_SELECTED_CC.includes(username)){ TN01_SELECTED_CC.push(username); renderCCTags(); }
  $('#cc_search_input').value = '';
  $('#cc_suggestions').classList.add('hidden');
}
function removeCCTag(username){
  TN01_SELECTED_CC = TN01_SELECTED_CC.filter(u => u !== username);
  renderCCTags();
}
function renderCCTags(){
  const box = $('#cc_tags'), input = $('#cc_search_input');
  box.querySelectorAll('.cc-pill').forEach(el => el.remove());
  TN01_SELECTED_CC.forEach(u => {
    const span = document.createElement('span');
    span.className = "cc-pill badge badge-indigo text-xs py-1 flex items-center gap-1 font-normal";
    span.innerHTML = `<b>${esc(u)}</b> <button onclick="removeCCTag('${esc(u)}')">&times;</button>`;
    box.insertBefore(span, input);
  });
}

async function submitSendNote(){
  const f = TN01_ROWS.filter(r => r.slice(1).some(x => String(x).trim() !== ''));
  const p = {
    layout_key: TN01_ACTIVE_ROUTE, route: TN01_ROUTES[TN01_ACTIVE_ROUTE].route,
    to_plant: $('#send_plant').value, to_dept: $('#send_dept').value,
    to_user: $('#send_user').value, cc_users: TN01_SELECTED_CC,
    remarks: $('#send_remarks').value, headers: TN01_HEADERS, rows: f
  };
  try {
    const r = await api('/api/tn01/send', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(p) });
    closeDialog();
    alert(`Transfer Sent: ${r.note_no}`);
    switchTN01Tab('history');
  } catch(e) { alert(e); }
}

async function renderTN01ReceiveTab(){
  const ns = await api('/api/tn01/inbox');
  $('#tn01_content').innerHTML = `
    <div class="p-6 flex-1 flex flex-col overflow-hidden">
      <div class="flex items-center justify-between pb-4 border-b">
        <h2 class="text-base font-bold">Incoming Transfer Notes</h2>
        <button onclick="renderTN01ReceiveTab()" class="btn-secondary text-xs px-3 py-1.5 rounded-lg"><i class="ri-refresh-line"></i></button>
      </div>
      <div class="flex-1 mt-4 card overflow-auto">
        <table class="custom-table">
          <thead><tr><th>Note</th><th>Route</th><th>From</th><th>Sender</th><th>Batches</th><th>CC</th><th>Time</th><th>Status</th><th>Actions</th></tr></thead>
          <tbody>
            ${!ns.length ? `<tr><td colspan="9" class="text-center py-12 text-slate-400">No incoming notes.</td></tr>` :
              ns.map(n => `<tr>
                <td class="font-mono font-bold text-indigo-600">${esc(n.note_no)}</td>
                <td>${esc(n.route)}</td>
                <td>${esc(n.from_plant)} - ${esc(n.from_dept)}</td>
                <td>${esc(n.from_user_name)}</td>
                <td class="font-mono">${n.entry_count}</td>
                <td class="font-mono text-xs text-slate-500">${esc(n.cc_users||'—')}</td>
                <td class="font-mono text-xs">${esc(n.created_at)}</td>
                <td><span class="badge ${n.status==='received'?'badge-emerald':'badge-amber'}">${n.status.toUpperCase()}</span></td>
                <td>
                  <button onclick="viewTransferNote(${n.id})" class="btn-primary text-xs px-2.5 py-1 rounded">Open</button>
                  ${n.status==='sent' ? `<button onclick="ackNote(${n.id})" class="bg-emerald-600 text-white text-xs px-2.5 py-1 rounded ml-1">Ack</button>` : ''}
                </td>
              </tr>`).join('')}
          </tbody>
        </table>
      </div>
    </div>`;
}
async function ackNote(id){ await api(`/api/tn01/receive/${id}`, {method:'POST'}); renderTN01ReceiveTab(); }

async function viewTransferNote(id){
  const n = await api(`/api/tn01/note/${id}`);
  openDialog(`
    <div class="px-6 py-4 border-b flex justify-between items-center bg-slate-50">
      <div><span class="font-mono font-bold text-indigo-700">${esc(n.note_no)}</span>
        <h3 class="font-bold text-sm mt-1">${esc(n.route)}</h3></div>
      <button onclick="closeDialog()"><i class="ri-close-line text-lg"></i></button>
    </div>
    <div class="p-6 space-y-4">
      <div class="flex justify-between items-center">
        <span class="text-xs font-bold text-slate-700">${n.entry_count} rows</span>
        <a href="/api/tn01/export/${id}" class="btn-secondary text-xs px-3 py-1.5 rounded flex items-center gap-1 text-emerald-700 font-semibold border-emerald-200 inline-flex">
          <i class="ri-file-excel-2-line"></i> Export as Excel
        </a>
      </div>
      <div class="text-xs"><b>CC:</b> ${esc(n.cc_users||'—')}</div>
      <div class="max-h-72 overflow-auto border border-slate-200 rounded-xl">
        <table class="w-full text-left text-xs">
          <thead><tr class="bg-slate-50 font-bold">${n.sheet_headers.map(h=>`<th class="p-2 border-r border-slate-200">${esc(h)}</th>`).join('')}</tr></thead>
          <tbody>${n.sheet_data.map(r=>`<tr class="border-b border-slate-100">${r.map(c=>`<td class="p-2 border-r border-slate-100">${esc(c)}</td>`).join('')}</tr>`).join('')}</tbody>
        </table>
      </div>
    </div>
    <div class="px-6 py-4 bg-slate-50 border-t flex justify-between">
      <a href="/api/tn01/print/${n.id}" target="_blank" class="btn-secondary text-xs px-3 py-2 rounded flex items-center gap-1"><i class="ri-printer-line"></i> Print</a>
      <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded text-xs">Close</button>
    </div>`);
}

async function renderTN01HistoryTab(){
  const admin = ME.is_master || ME.role === 'admin';
  const h = await api('/api/tn01/history');
  $('#tn01_content').innerHTML = `
    <div class="p-6 flex-1 flex flex-col overflow-hidden">
      <div class="flex items-center justify-between pb-3 border-b">
        <h2 class="text-base font-bold">${admin ? "All Transfer Notes History" : "My Transfer Notes History"}</h2>
        <button onclick="renderTN01HistoryTab()" class="btn-secondary text-xs px-3 py-1.5 rounded-lg"><i class="ri-refresh-line"></i></button>
      </div>
      <div class="flex-1 mt-2 card overflow-auto">
        <table class="custom-table">
          <thead><tr><th>Note</th><th>Route</th><th>From</th><th>To</th><th>Batches</th><th>Sender</th><th>CC</th><th>Time</th><th>Status</th><th>Actions</th></tr></thead>
          <tbody>
            ${!h.length ? `<tr><td colspan="10" class="text-center py-10 text-slate-400">No records.</td></tr>` :
              h.map(x => `<tr>
                <td class="font-mono font-bold text-indigo-600">${esc(x.note_no)}</td>
                <td>${esc(x.route)}</td>
                <td>${esc(x.from_plant)} - ${esc(x.from_dept)}</td>
                <td>${esc(x.to_plant)} - ${esc(x.to_dept)}</td>
                <td class="font-mono">${x.entry_count}</td>
                <td>${esc(x.from_user_name)}</td>
                <td class="font-mono text-xs text-slate-500">${esc(x.cc_users||'—')}</td>
                <td class="font-mono text-xs">${esc(x.created_at)}</td>
                <td><span class="badge ${x.status==='received'?'badge-emerald':'badge-amber'}">${x.status.toUpperCase()}</span></td>
                <td>
                  <button onclick="viewTransferNote(${x.id})" class="text-xs text-indigo-600 font-semibold">View</button>
                  ${admin ? `<button onclick="deleteTN01History(${x.id})" class="text-xs text-red-600 font-bold ml-2">Del</button>` : ''}
                </td>
              </tr>`).join('')}
          </tbody>
        </table>
      </div>
    </div>`;
}

async function deleteTN01History(id){
  if(!confirm("Delete this transfer note history entry?")) return;
  try { await api(`/api/tn01/delete/${id}`, {method:'DELETE'}); setMsg("Deleted"); renderTN01HistoryTab(); }
  catch(e) { alert(e); }
}

// ===== MRQ1 =====
let MRQ_CATEGORIES = {};
async function renderMRQ(category=null){
  if(category) return renderMRQForm(category);
  MRQ_CATEGORIES = await api('/api/mrq/categories');
  const history = await api('/api/mrq/requisitions');
  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col p-6 overflow-hidden select-text">
      <div class="flex items-center justify-between pb-4 border-b shrink-0">
        <div class="flex items-center gap-3">
          <span class="font-mono font-bold text-xs bg-emerald-600 text-white px-2.5 py-1 rounded-lg">MRQ1</span>
          <div><h2 class="text-lg font-bold">Material Requisitions</h2></div>
        </div>
        <button onclick="renderMRQ()" class="btn-secondary text-xs px-3 py-1.5 rounded"><i class="ri-refresh-line"></i></button>
      </div>
      <div class="mt-4 grid grid-cols-3 gap-4 shrink-0">
        ${Object.entries(MRQ_CATEGORIES).map(([k,t]) => `
          <div onclick="renderMRQ('${k}')" class="card hover:border-indigo-300 p-5 cursor-pointer">
            <h4 class="font-bold text-sm">${t.title}</h4>
            <p class="text-xs text-slate-500 mt-1">${t.desc}</p>
          </div>`).join('')}
      </div>
      <div class="flex-1 mt-6 card overflow-auto">
        <table class="custom-table">
          <thead><tr><th>Req #</th><th>Dept</th><th>Status</th><th>Actions</th></tr></thead>
          <tbody>
            ${history.map(r => `<tr>
              <td>${esc(r.req_no)}</td><td>${esc(r.from_dept)}</td>
              <td><span class="badge badge-emerald">${r.status.toUpperCase()}</span></td>
              <td><a href="/api/mrq/print/${r.id}" target="_blank" class="text-xs text-indigo-600 font-semibold">Print</a></td>
            </tr>`).join('')}
          </tbody>
        </table>
      </div>
    </div>`;
}
function renderMRQForm(category){
  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col p-6 overflow-hidden select-text">
      <div class="flex items-center justify-between pb-4 border-b shrink-0">
        <div class="flex items-center gap-2">
          <button onclick="renderMRQ()" class="btn-secondary text-xs p-2 rounded"><i class="ri-arrow-left-line"></i></button>
          <h2 class="text-lg font-bold">Create Requisition (${category.toUpperCase()})</h2>
        </div>
      </div>
      <div class="flex-1 mt-4 overflow-auto">
        <div class="card p-6 space-y-4 max-w-xl mx-auto">
          <div><label class="form-label">Material Code</label><input id="mrq_code" class="form-input"></div>
          <div><label class="form-label">Quantity</label><input id="mrq_qty" type="number" class="form-input"></div>
          <div class="flex justify-end gap-2">
            <button onclick="renderMRQ()" class="btn-secondary text-xs px-4 py-2 rounded">Cancel</button>
            <button onclick="submitMRQSingle('${category}')" class="btn-primary text-xs font-bold px-4 py-2 rounded">Submit</button>
          </div>
        </div>
      </div>
    </div>`;
}
async function submitMRQSingle(cat){
  const code = $('#mrq_code').value.trim();
  const qty = parseFloat($('#mrq_qty').value) || 0;
  if(!code || qty <= 0) return alert("Enter valid code and quantity.");
  try {
    await api('/api/mrq/requisition', { method:'POST', headers:{'Content-Type':'application/json'},
      body:JSON.stringify({category:cat, lines:[{material_code:code, qty_required:qty}]}) });
    alert("Requisition submitted."); renderMRQ();
  } catch(e) { alert(e); }
}

// ===== SU01 =====
async function renderSU01(){
  if(!ME.is_master && ME.role !== 'admin') return setMsg("Denied", true);
  ALL_USERS_LIST = await api('/api/su01/users');
  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col p-6 overflow-hidden select-text">
      <div class="flex items-center justify-between pb-4 border-b">
        <h2 class="text-lg font-bold">User Directory (SU01)</h2>
        <button onclick="showCreateUserModal()" class="btn-primary text-xs px-4 py-2 rounded-lg">+ Add User</button>
      </div>
      <div class="flex-1 mt-4 card overflow-auto">
        <table class="custom-table">
          <thead><tr><th>Username</th><th>Name</th><th>Role</th><th>Dept</th><th>Status</th><th>Actions</th></tr></thead>
          <tbody>
            ${ALL_USERS_LIST.map(u => `<tr>
              <td class="font-mono font-bold text-indigo-600">${esc(u.username)}</td>
              <td>${esc(u.full_name)}</td>
              <td><span class="badge badge-slate">${u.role.toUpperCase()}</span></td>
              <td>${esc(u.department)}</td>
              <td><span class="badge ${u.status==='approved'?'badge-emerald':'badge-red'}">${u.status}</span></td>
              <td>
                <button onclick="showResetPassModal('${esc(u.username)}')" class="text-xs text-amber-600 font-semibold">Reset PW</button>
                ${u.role!=='master' ? `<button onclick="deleteUser('${esc(u.username)}')" class="text-xs text-red-600 font-semibold ml-2">Del</button>` : ''}
              </td>
            </tr>`).join('')}
          </tbody>
        </table>
      </div>
    </div>`;
}
function showCreateUserModal(){
  api('/api/spro/overview').then(d => {
    openDialog(`
      <div class="p-6 space-y-4">
        <h3 class="font-bold text-sm">Create User</h3>
        <div><label class="form-label">Username</label><input id="nu_u" class="form-input"></div>
        <div><label class="form-label">Password</label><input id="nu_p" type="password" class="form-input"></div>
        <div><label class="form-label">Full Name</label><input id="nu_fn" class="form-input"></div>
        <div><label class="form-label">Dept</label>
          <select id="nu_d" class="form-input">${d.departments.map(x=>`<option value="${esc(x.name)}">${esc(x.name)}</option>`).join('')}</select>
        </div>
        <div><label class="form-label">Role</label>
          <select id="nu_r" class="form-input"><option value="user">User</option><option value="admin">Admin</option></select>
        </div>
        <div class="flex justify-end gap-2">
          <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded text-xs">Cancel</button>
          <button onclick="submitCreateUser()" class="btn-primary px-4 py-2 rounded text-xs">Save</button>
        </div>
      </div>`);
  });
}
async function submitCreateUser(){
  const fd = new FormData();
  fd.append("username", $('#nu_u').value); fd.append("password", $('#nu_p').value);
  fd.append("full_name", $('#nu_fn').value); fd.append("department", $('#nu_d').value);
  fd.append("role", $('#nu_r').value); fd.append("plants","1000");
  try { await api('/api/su01/create', {method:'POST', body:fd}); closeDialog(); renderSU01(); }
  catch(e) { alert(e); }
}
function showResetPassModal(u){
  const p = prompt(`New password for ${u}:`); if(!p) return;
  const fd = new FormData(); fd.append("username", u); fd.append("new_password", p);
  api('/api/su01/reset-password', {method:'POST', body:fd})
    .then(()=>setMsg("Password reset.")).catch(e=>alert(e));
}
async function deleteUser(u){
  if(!confirm(`Delete ${u}?`)) return;
  const fd = new FormData(); fd.append("username", u);
  try { await api('/api/su01/delete', {method:'POST', body:fd}); renderSU01(); }
  catch(e) { alert(e); }
}

// ===== SPRO =====
async function renderSPRO(){
  if(!ME.is_master && ME.role !== 'admin') return setMsg("Denied", true);
  const d = await api('/api/spro/overview');
  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col p-6 overflow-hidden">
      <div class="flex items-center justify-between pb-4 border-b"><h2 class="text-lg font-bold">Enterprise Structure (SPRO)</h2></div>
      <div class="grid grid-cols-2 gap-4 flex-1 mt-4 overflow-auto">
        <div class="card overflow-hidden">
          <div class="p-3 border-b bg-slate-50 flex justify-between">
            <span class="text-xs font-bold uppercase">Plants (${d.plants.length})</span>
            <button onclick="createPlant()" class="btn-primary text-xs px-3 py-1.5 rounded">+ Add</button>
          </div>
          <table class="custom-table"><thead><tr><th>Code</th><th>Name</th></tr></thead>
            <tbody>${d.plants.map(p=>`<tr><td class="font-mono font-bold">${esc(p.code)}</td><td>${esc(p.name)}</td></tr>`).join('')}</tbody>
          </table>
        </div>
        <div class="card overflow-hidden">
          <div class="p-3 border-b bg-slate-50 flex justify-between">
            <span class="text-xs font-bold uppercase">Departments (${d.departments.length})</span>
            <button onclick="createDept()" class="btn-primary text-xs px-3 py-1.5 rounded">+ Add</button>
          </div>
          <table class="custom-table"><thead><tr><th>Name</th></tr></thead>
            <tbody>${d.departments.map(x=>`<tr><td>${esc(x.name)}</td></tr>`).join('')}</tbody>
          </table>
        </div>
      </div>
    </div>`;
}
function createPlant(){
  const c = prompt("Plant Code:"); if(!c) return;
  const n = prompt("Plant Name:"); if(!n) return;
  const fd = new FormData(); fd.append("code", c); fd.append("name", n);
  api('/api/admin/plants/create', {method:'POST', body:fd}).then(()=>renderSPRO()).catch(e=>alert(e));
}
function createDept(){
  const d = prompt("Department name:"); if(!d) return;
  const fd = new FormData(); fd.append("name", d.trim());
  api('/api/admin/departments/create', {method:'POST', body:fd}).then(()=>renderSPRO()).catch(e=>alert(e));
}

// ===== SE93 =====
async function renderSE93(){
  if(!ME.is_master && ME.role !== 'admin') return setMsg("Denied", true);
  const list = await api('/api/se93/list');
  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col p-6 overflow-hidden">
      <div class="flex justify-between pb-4 border-b"><h2 class="text-lg font-bold">Module Studio (SE93)</h2></div>
      <div class="grid grid-cols-3 gap-4 flex-1 mt-4 overflow-hidden">
        <div class="col-span-1 card p-4 overflow-auto">
          <h3 class="text-sm font-bold mb-3">Auto-Create Module</h3>
          <form onsubmit="submitSE93(event)" class="space-y-3">
            <div><label class="form-label">Code</label><input name="code" required class="form-input font-mono uppercase"></div>
            <div><label class="form-label">Title</label><input name="title" required class="form-input"></div>
            <div><label class="form-label">Description</label><input name="description" class="form-input"></div>
            <div><label class="form-label">File (.xlsx, .csv)</label><input name="file" type="file" required class="w-full text-xs"></div>
            <button type="submit" class="btn-primary w-full text-xs font-bold py-2 rounded">Create Module</button>
          </form>
        </div>
        <div class="col-span-2 card overflow-auto">
          <table class="custom-table">
            <thead><tr><th>Code</th><th>Title</th><th>Records</th><th>Actions</th></tr></thead>
            <tbody>
              ${list.map(t => `<tr>
                <td class="font-mono font-bold text-indigo-600">${esc(t.code)}</td>
                <td>${esc(t.title)}</td>
                <td class="font-mono">${t.record_count}</td>
                <td><button onclick="openTcode('${esc(t.code)}')" class="text-xs text-indigo-600 font-semibold">Open</button></td>
              </tr>`).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>`;
}
async function submitSE93(ev){
  ev.preventDefault();
  const fd = new FormData(ev.target);
  try {
    const r = await fetch('/api/se93/auto-create', { method:'POST', body:fd,
      headers:{'Authorization':'Bearer '+(localStorage.getItem("nexus_token")||'')} });
    const d = await r.json();
    if(!r.ok) throw d.detail || "Fail";
    alert(`Created ${d.code}`); renderSE93(); loadModules();
  } catch(e) { alert(e); }
}

// ===== PFCG =====
async function renderPFCG(){
  if(!ME.is_master && ME.role !== 'admin') return setMsg("Denied", true);
  const m = await api('/api/pfcg/matrix');
  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col p-6 overflow-hidden">
      <div class="flex justify-between pb-4 border-b">
        <h2 class="text-lg font-bold">Access Control (PFCG)</h2>
        <div class="flex gap-2">
          <select id="pfcg_u" class="form-input w-56">
            ${m.users.map(u=>`<option value="${esc(u.username)}">${esc(u.full_name||u.username)}</option>`).join('')}
          </select>
          <select id="pfcg_t" class="form-input w-56">
            ${m.tcodes.map(t=>`<option value="${esc(t.code)}">${esc(t.code)} - ${esc(t.title)}</option>`).join('')}
          </select>
          <button onclick="grantAccess()" class="btn-primary text-xs px-3 py-2 rounded">+ Grant</button>
        </div>
      </div>
      <div class="flex-1 mt-4 card overflow-auto">
        <table class="custom-table">
          <thead><tr><th>User</th><th>Module</th><th>By</th><th>Action</th></tr></thead>
          <tbody>
            ${m.grants.map(g=>`<tr>
              <td class="font-mono">${esc(g.username)}</td>
              <td class="font-mono">${esc(g.tcode_code)}</td>
              <td>${esc(g.granted_by)}</td>
              <td><button onclick="revokeAccess('${esc(g.username)}','${esc(g.tcode_code)}')" class="text-xs text-red-600 font-semibold">Revoke</button></td>
            </tr>`).join('')}
          </tbody>
        </table>
      </div>
    </div>`;
}
async function grantAccess(){
  const fd = new FormData();
  fd.append("username", $('#pfcg_u').value); fd.append("tcode", $('#pfcg_t').value);
  await api('/api/pfcg/grant', {method:'POST', body:fd}); renderPFCG();
}
async function revokeAccess(u, t){
  if(!confirm("Revoke?")) return;
  const fd = new FormData(); fd.append("username", u); fd.append("tcode", t);
  await api('/api/pfcg/revoke', {method:'POST', body:fd}); renderPFCG();
}

// ===== AI STUDIO (Safe Mode, unchanged behaviour) =====
const STUDIO_EXAMPLE = {
  action: "upsert", code: "DISP0", title: "Dispatch",
  description: "Outward dispatch", module: "SD",
  allowed_depts: "Commercial,Store",
  fields: [
    {key: "challan_no", label: "Challan No.", type: "text", required: true},
    {key: "customer",   label: "Customer",    type: "text"},
    {key: "dispatch_date", label: "Date",     type: "date"},
    {key: "qty",        label: "Qty",         type: "number"},
    {key: "remarks",    label: "Remarks",     type: "textarea", width: 2}
  ],
  seed_rows: []
};

async function renderStudio(){
  if(!ME.is_master) return setMsg("Master only", true);
  const hist = await api('/api/studio/history');
  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col p-6 overflow-hidden select-text">
      <div class="flex items-center justify-between pb-4 border-b border-slate-200 shrink-0">
        <div>
          <h2 class="text-lg font-bold text-slate-900 flex items-center gap-2">
            <i class="ri-robot-2-line text-emerald-500"></i> AI Code Studio
          </h2>
          <p class="text-xs text-slate-500">Paste JSON to create/update/delete modules. Protected codes are locked.</p>
        </div>
        <div class="flex gap-2">
          <button onclick="loadStudioExample()" class="btn-secondary text-xs px-3 py-1.5 rounded-lg"><i class="ri-magic-line"></i> Example</button>
          <button onclick="renderStudio()" class="btn-secondary text-xs px-3 py-1.5 rounded-lg"><i class="ri-refresh-line"></i></button>
        </div>
      </div>
      <div class="grid grid-cols-1 lg:grid-cols-2 gap-4 flex-1 mt-4 overflow-hidden">
        <div class="flex flex-col card overflow-hidden">
          <div class="p-3 border-b border-slate-200 bg-slate-50 flex justify-between items-center">
            <span class="text-xs font-bold text-slate-700 uppercase"><i class="ri-code-s-slash-line text-indigo-500"></i> Module Spec (JSON)</span>
            <div class="flex gap-1.5">
              <button onclick="studioValidate()" class="text-[11px] font-bold px-2.5 py-1 rounded-md bg-amber-50 border border-amber-200 text-amber-700">Validate</button>
              <button onclick="studioCopy()" class="text-[11px] font-semibold px-2.5 py-1 rounded-md bg-white border border-slate-200 text-slate-600">Copy</button>
              <button onclick="studioFormat()" class="text-[11px] font-semibold px-2.5 py-1 rounded-md bg-white border border-slate-200 text-slate-600">Prettify</button>
            </div>
          </div>
          <textarea id="studio_json" spellcheck="false"
            class="flex-1 bg-slate-50 text-emerald-700 font-mono text-[11.5px] leading-relaxed p-3 outline-none resize-none"
            placeholder='Paste JSON here...'></textarea>
          <div class="p-3 border-t border-slate-200 bg-slate-50 flex items-center justify-between">
            <span id="studio_status" class="text-[11px] font-mono text-slate-500">Idle</span>
            <button onclick="studioApply()" class="btn-primary text-xs font-bold px-4 py-1.5 rounded-md flex items-center gap-1">
              <i class="ri-play-fill"></i> Apply
            </button>
          </div>
        </div>
        <div class="flex flex-col gap-4 overflow-hidden">
          <div id="studio_preview" class="card p-4 overflow-auto flex-1 min-h-[200px]">
            <div class="text-xs text-slate-400 text-center pt-10">Validate a spec to see preview.</div>
          </div>
          <div class="card overflow-hidden" style="max-height:44%;">
            <div class="p-2.5 border-b border-slate-200 bg-slate-50 text-xs font-bold text-slate-700 uppercase flex items-center gap-1.5">
              <i class="ri-history-line text-indigo-500"></i> History
            </div>
            <div class="overflow-auto" style="max-height:200px;">
              <table class="custom-table">
                <thead><tr><th>v</th><th>Module</th><th>Action</th><th>By</th><th>When</th><th></th></tr></thead>
                <tbody>
                  ${!hist.length ? `<tr><td colspan="6" class="text-center py-6 text-slate-400 text-xs">None.</td></tr>` :
                    hist.map(h => `<tr>
                      <td class="font-mono text-indigo-600 font-bold">v${h.version}</td>
                      <td class="font-mono font-bold">${esc(h.tcode_code)}</td>
                      <td><span class="badge ${h.action==='delete'?'badge-red':'badge-emerald'}">${h.action}</span></td>
                      <td class="text-xs">${esc(h.applied_by)}</td>
                      <td class="text-xs font-mono text-slate-500">${esc(h.applied_at)}</td>
                      <td>
                        <button onclick="studioRollback(${h.id})" class="text-[11px] text-amber-600 font-semibold">Rollback</button>
                        <button onclick="loadStudioExisting('${esc(h.tcode_code)}')" class="text-[11px] text-indigo-600 font-semibold ml-2">Load</button>
                      </td>
                    </tr>`).join('')}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    </div>`;
}
function loadStudioExample(){
  $('#studio_json').value = JSON.stringify(STUDIO_EXAMPLE, null, 2);
  setStudioStatus("Example loaded.", false);
}
function setStudioStatus(m, e){
  const el = $('#studio_status'); if(!el) return;
  el.textContent = m;
  el.className = "text-[11px] font-mono " + (e ? "text-red-600" : "text-emerald-600");
}
function studioFormat(){
  try {
    const v = JSON.parse($('#studio_json').value);
    $('#studio_json').value = JSON.stringify(v, null, 2);
    setStudioStatus("Formatted.", false);
  } catch(e) { setStudioStatus("Invalid: " + e.message, true); }
}
function studioCopy(){
  navigator.clipboard.writeText($('#studio_json').value);
  setStudioStatus("Copied.", false);
}
async function studioValidate(){
  let s;
  try { s = JSON.parse($('#studio_json').value); }
  catch(e) { setStudioStatus("Invalid: " + e.message, true); return; }
  try {
    const r = await api('/api/studio/validate', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body:JSON.stringify({spec:s})
    });
    if(!r.ok){ setStudioStatus("Errors: " + r.errors.join(" | "), true); renderStudioPreview(r.clean, r.errors, r.diff); return; }
    if(r.protected){ setStudioStatus("PROTECTED — cannot apply via Studio.", true); renderStudioPreview(r.clean, r.errors, r.diff); return; }
    setStudioStatus(r.diff && r.diff.kind === "create" ? "Valid (CREATE)." : "Valid (UPDATE).", false);
    renderStudioPreview(r.clean, [], r.diff);
  } catch(e) { setStudioStatus("Error: " + e, true); }
}
function renderStudioPreview(cl, errs, diff){
  const b = $('#studio_preview'); if(!b) return;
  if(!cl || errs.length){
    b.innerHTML = `<div class="text-xs text-red-600 whitespace-pre-wrap">${esc(errs.join("\n"))}</div>`;
    return;
  }
  if(cl.action === "delete"){
    b.innerHTML = `<div class="text-sm"><b class="text-red-600">DELETE:</b> ${esc(cl.code)}</div>
      <div class="text-xs text-slate-500 mt-2">This will remove the module and all its data. Two-step confirmation required.</div>`;
    return;
  }
  let diffHtml = '';
  if(diff){
    if(diff.kind === "create"){
      diffHtml = `<div class="mt-3 p-2 bg-emerald-50 border border-emerald-200 rounded-lg text-xs text-emerald-800">
        New module — ${diff.after_fields.length} fields will be created.</div>`;
    } else {
      diffHtml = `<div class="mt-3 p-2 bg-slate-50 border border-slate-200 rounded-lg text-xs">
        ${diff.removed_fields.length ? `<div class="text-red-600 font-semibold mb-1">⚠ ${diff.removed_fields.length} field(s) will no longer be visible on old records: ${diff.removed_fields.join(', ')}</div>` : ''}
        ${diff.added_fields.length ? `<div class="text-emerald-600 font-semibold mb-1">+${diff.added_fields.length} new field(s): ${diff.added_fields.join(', ')}</div>` : ''}
        <div class="text-slate-600">${diff.before_rows} existing record(s) will be preserved.</div>
      </div>`;
    }
  }
  b.innerHTML = `
    <div class="mb-3">
      <div class="text-[10px] uppercase font-bold text-slate-400">Preview</div>
      <div class="text-base font-bold text-slate-900">${esc(cl.title)} <span class="font-mono text-indigo-600 text-xs">[${esc(cl.code)}]</span></div>
      <div class="text-xs text-slate-500 mt-0.5">${esc(cl.description || '—')}</div>
    </div>
    <table class="w-full text-left text-[11px]">
      <thead><tr class="bg-slate-50 text-slate-500"><th class="p-1.5">Key</th><th class="p-1.5">Label</th><th class="p-1.5">Type</th></tr></thead>
      <tbody>${cl.fields.map(f => `<tr class="border-b border-slate-100">
        <td class="p-1.5 font-mono text-indigo-600">${esc(f.key)}</td>
        <td class="p-1.5 text-slate-900">${esc(f.label)}</td>
        <td class="p-1.5 text-slate-500">${esc(f.type)}</td>
      </tr>`).join('')}</tbody>
    </table>
    ${diffHtml}`;
}
async function studioApply(){
  let s;
  try { s = JSON.parse($('#studio_json').value); }
  catch(e) { setStudioStatus("Invalid: " + e.message, true); return; }
  const actionLabel = s.action === 'delete' ? 'DELETE' : (s.action === 'upsert' ? 'APPLY' : s.action);
  if(!confirm(`${actionLabel} on ${s.code}?`)) return;
  try {
    const payload = {spec:s};
    if(s.action === 'delete') payload.confirmed = true;
    const r = await api('/api/studio/apply', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body:JSON.stringify(payload)
    });
    setStudioStatus(`OK ${r.action} ${r.code} v${r.version}`, false);
    await loadModules();
    renderStudio();
  } catch(e) { setStudioStatus("Error: " + e, true); alert(e); }
}
async function studioRollback(id){
  if(!confirm("Rollback this version? Schema + data + layouts will be restored from snapshot.")) return;
  try {
    const r = await api(`/api/studio/rollback/${id}`, {method:'POST'});
    setStudioStatus(`Restored from v${r.restored_from} → new v${r.version}`, false);
    await loadModules(); renderStudio();
  } catch(e) { alert(e); }
}
async function loadStudioExisting(code){
  try {
    const s = await api(`/api/studio/spec/${code}`);
    $('#studio_json').value = JSON.stringify(s, null, 2);
    if(s.protected){ setStudioStatus(`${code} is PROTECTED — cannot edit via Studio.`, true); }
    else { setStudioStatus(`Loaded ${code}.`, false); }
    await studioValidate();
  } catch(e) { alert(e); }
}

// ===== STANDARD MODULE =====
async function renderStandardModule(code){
  const ls = await api(`/api/tcode/${code}/layouts`);
  const al = ls.find(l => l.is_default) || ls[0];
  const fs = al ? al.fields : [];
  const recs = await api(`/api/tcode/${code}/records`);
  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col overflow-hidden select-text">
      <div class="p-4 border-b flex items-center justify-between shrink-0 bg-white">
        <div class="flex items-center gap-3">
          <span class="font-mono font-bold text-xs bg-indigo-600 text-white px-2 py-0.5 rounded">${code}</span>
          <h2 class="text-base font-bold">${al ? esc(al.title) : code}</h2>
          <button onclick="showStandardRecordModal('${code}')" class="btn-primary text-xs px-3 py-1.5 rounded flex items-center gap-1">
            <i class="ri-add-line"></i> New
          </button>
        </div>
        <a href="/api/tcode/${code}/export" class="btn-secondary text-xs px-2.5 py-1.5 rounded flex items-center gap-1">
          <i class="ri-download-2-line"></i> Export
        </a>
      </div>
      <div class="flex-1 overflow-auto">
        <table class="custom-table">
          <thead><tr><th>#</th><th>Plant</th>${fs.map(f=>`<th>${esc(f.label)}</th>`).join('')}<th>By</th><th>Action</th></tr></thead>
          <tbody>
            ${!recs.length ? `<tr><td colspan="${fs.length+4}" class="text-center py-10 text-slate-400">No records found. Click '+ New' to create an entry.</td></tr>` : ''}
            ${recs.map(r=>`<tr>
              <td class="font-mono">${r.id}</td>
              <td>${esc(r.plant||'1000')}</td>
              ${fs.map(f=>`<td>${esc(r.data[f.key]||'')}</td>`).join('')}
              <td>${esc(r.created_by_name||r.created_by)}</td>
              <td><button onclick="deleteStandardRecord('${code}',${r.id})" class="text-xs text-red-600 font-semibold">Del</button></td>
            </tr>`).join('')}
          </tbody>
        </table>
      </div>
    </div>`;
}
function showStandardRecordModal(code){
  api(`/api/tcode/${code}/layouts`).then(ls => {
    const al = ls.find(l => l.is_default) || ls[0];
    const fs = al ? al.fields : [];
    openDialog(`
      <div class="p-6 space-y-4">
        <h3 class="font-bold text-sm">New Entry: ${code}</h3>
        <div><label class="form-label">Plant</label><input id="std_f__plant" class="form-input font-mono" value="1000"></div>
        ${fs.map(f => `<div><label class="form-label">${esc(f.label)}</label>
          <input id="std_f_${f.key}" class="form-input"></div>`).join('')}
        <div class="flex justify-end gap-2">
          <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded text-xs">Cancel</button>
          <button onclick="submitStandardRecord('${code}')" class="btn-primary px-4 py-2 rounded text-xs">Save</button>
        </div>
      </div>`);
  });
}
async function submitStandardRecord(code){
  const ls = await api(`/api/tcode/${code}/layouts`);
  const al = ls.find(l => l.is_default) || ls[0];
  const fs = al ? al.fields : [];
  const p = { _plant: $('#std_f__plant').value };
  for(const f of fs){ const el = $('#std_f_'+f.key); p[f.key] = el ? el.value.trim() : ""; }
  try {
    await api(`/api/tcode/${code}/records`, {
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(p)
    });
    closeDialog(); renderStandardModule(code); setMsg("Saved");
  } catch(e) { alert(e); }
}
async function deleteStandardRecord(code, rid){
  if(!confirm(`Delete #${rid}?`)) return;
  try { await api(`/api/tcode/${code}/records/${rid}`, {method:'DELETE'}); renderStandardModule(code); setMsg("Deleted"); }
  catch(e) { alert(e); }
}

// Dialog helpers
function openDialog(c){ $('#app_dialog_body').innerHTML = c; $('#app_dialog').classList.remove('hidden'); }
function closeDialog(){ $('#app_dialog').classList.add('hidden'); $('#app_dialog_body').innerHTML = ''; }

document.addEventListener('click', e => {
  const box = $('#cc_suggestions');
  if(box && !box.contains(e.target) && e.target.id !== 'cc_search_input'){
    box.classList.add('hidden');
  }
});

initSession();
</script>
</body>
</html>
"""


# =========================================================================
# HEALTH + HOME
# =========================================================================
@app.api_route("/", methods=["GET", "HEAD"], response_class=HTMLResponse)
def home():
    return HTMLResponse(HTML_UI)


@app.api_route("/health", methods=["GET", "HEAD"])
def health():
    return {"status": "ACTIVE", "system": "NEXUS-CORE", "version": "5.0.0"}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
