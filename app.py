# =========================================================================
# NEXUS CORE ERP  —  app.py  —  v4.0.1
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

STUDIO_PROTECTED = {
    "TN01", "CTST", "SU01", "SPRO", "SE93", "PFCG", "STUDIO", "MRQ1",
    "PROD", "GCH0", "NCR0", "PLN0", "QCH0", "STD0",
}

DATABASE_URL = os.environ.get("DATABASE_URL")
USE_PG = bool(DATABASE_URL and DATABASE_URL.startswith("postgres"))

if USE_PG:
    import psycopg2
    import psycopg2.extras
    if DATABASE_URL.startswith("postgres://"):
        DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

DB_PATH = Path(__file__).parent / "nexus_erp.db"

def get_conn():
    if USE_PG:
        return psycopg2.connect(DATABASE_URL, sslmode="require")
    con = sqlite3.connect(DB_PATH, check_same_thread=False)
    con.row_factory = sqlite3.Row
    return con

def q(sql: str) -> str:
    return sql.replace("?", "%s") if USE_PG else sql

def fetch_all(sql: str, params=()):
    conn = get_conn()
    try:
        if USE_PG:
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
            conn.commit()
        else:
            conn.execute(sql, params)
            conn.commit()
    finally:
        conn.close()

def now() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def esc_py(s) -> str:
    if s is None: return ""
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def json_serial(obj):
    if isinstance(obj, (datetime.datetime, datetime.date)):
        if isinstance(obj, datetime.datetime):
            return obj.strftime("%Y-%m-%d %H:%M:%S")
        return obj.strftime("%Y-%m-%d")
    return str(obj)

def infer_field_type(sample_values):
    if not sample_values: return "text"
    is_num = True
    is_date = True
    is_long = False
    date_patterns = [r"^\d{4}-\d{2}-\d{2}", r"^\d{2}/\d{2}/\d{4}", r"^\d{2}-\d{2}-\d{4}"]
    for v in sample_values:
        if v is None or v == "": continue
        if isinstance(v, (datetime.datetime, datetime.date)):
            is_num = False
            continue
        if isinstance(v, (int, float)):
            is_date = False
            continue
        s = str(v).strip()
        if len(s) > 70 or "\n" in s:
            is_long = True
            is_num = False
            is_date = False
            break
        if len(s) > 1 and s[0] == "0" and s.isdigit():
            is_num = False
        else:
            try: float(s.replace(",", ""))
            except ValueError: is_num = False
        if not any(re.match(p, s) for p in date_patterns): is_date = False
    if is_date: return "date"
    if is_num: return "number"
    if is_long: return "textarea"
    return "text"

def init_db():
    conn = get_conn()
    cur = conn.cursor()

    if USE_PG:
        try:
            cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name='tcodes' AND column_name='color';")
            if cur.fetchone():
                cur.execute("DROP TABLE IF EXISTS tcodes CASCADE; DROP TABLE IF EXISTS tcode_schemas CASCADE; DROP TABLE IF EXISTS tcode_layouts CASCADE;")
            cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name='users' AND column_name='sub_department';")
            if cur.fetchone(): cur.execute("DROP TABLE IF EXISTS users CASCADE;")
            cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name='plants' AND column_name='code';")
            if not cur.fetchone(): cur.execute("DROP TABLE IF EXISTS plants CASCADE;")
            cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name='org' AND column_name='system_id';")
            if not cur.fetchone(): cur.execute("DROP TABLE IF EXISTS org CASCADE;")
            cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name='tcode_records' AND column_name='plant';")
            if not cur.fetchone(): cur.execute("ALTER TABLE tcode_records ADD COLUMN IF NOT EXISTS plant TEXT DEFAULT '';")
            conn.commit()
        except Exception as e:
            print("Auto-migration notice:", e)
            conn.rollback()

    schema_template = """
    CREATE TABLE IF NOT EXISTS users (username TEXT PRIMARY KEY, password_hash TEXT, full_name TEXT, role TEXT, department TEXT, plants TEXT DEFAULT '', status TEXT DEFAULT 'approved', created_at TEXT, updated_at TEXT);
    CREATE TABLE IF NOT EXISTS sessions (token TEXT PRIMARY KEY, username TEXT, expires TEXT);
    CREATE TABLE IF NOT EXISTS departments (name TEXT PRIMARY KEY, active INT DEFAULT 1, created_at TEXT);
    CREATE TABLE IF NOT EXISTS plants (code TEXT PRIMARY KEY, name TEXT, active INT DEFAULT 1, created_at TEXT);
    CREATE TABLE IF NOT EXISTS org (id INT PRIMARY KEY, name TEXT, system_id TEXT, client TEXT);
    CREATE TABLE IF NOT EXISTS tcodes (code TEXT PRIMARY KEY, title TEXT, description TEXT, module TEXT, allowed_depts TEXT, active INT DEFAULT 1, created_at TEXT, updated_at TEXT, created_by TEXT);
    CREATE TABLE IF NOT EXISTS tcode_schemas (code TEXT PRIMARY KEY, fields TEXT, layout TEXT, updated_at TEXT, updated_by TEXT);
    CREATE TABLE IF NOT EXISTS tcode_records (id {IDTYPE}, tcode_code TEXT, plant TEXT DEFAULT '', data TEXT, status TEXT DEFAULT 'active', created_by TEXT, created_by_name TEXT, created_at TEXT, updated_at TEXT);
    CREATE TABLE IF NOT EXISTS tcode_layouts (id {IDTYPE}, tcode_code TEXT, layout_code TEXT UNIQUE, name TEXT, title TEXT, is_default INT DEFAULT 0, fields TEXT, layout_config TEXT, description TEXT, created_by TEXT, created_by_name TEXT, created_at TEXT, updated_at TEXT);
    CREATE TABLE IF NOT EXISTS access_grants (id {IDTYPE}, username TEXT, tcode_code TEXT, granted_by TEXT, granted_at TEXT, UNIQUE(username, tcode_code));
    CREATE TABLE IF NOT EXISTS transfer_notes (id {IDTYPE}, note_no TEXT UNIQUE, layout_key TEXT, route TEXT, from_plant TEXT, from_dept TEXT, from_user TEXT, from_user_name TEXT, to_plant TEXT, to_dept TEXT, to_user TEXT, to_user_name TEXT, entry_count INT DEFAULT 0, sheet_headers TEXT, sheet_data TEXT, status TEXT DEFAULT 'sent', remarks TEXT DEFAULT '', created_at TEXT, received_at TEXT DEFAULT '', received_by TEXT DEFAULT '', received_by_name TEXT DEFAULT '', cc_users TEXT DEFAULT '');
    CREATE TABLE IF NOT EXISTS module_specs (id {IDTYPE}, tcode_code TEXT, version INT, action TEXT, spec TEXT, applied_by TEXT, applied_at TEXT, note TEXT);
    CREATE TABLE IF NOT EXISTS mrq_materials (material_code TEXT PRIMARY KEY, description TEXT NOT NULL, category TEXT DEFAULT 'stores', uom TEXT DEFAULT 'NOS', created_at TEXT);
    CREATE TABLE IF NOT EXISTS mrq_requisitions (id {IDTYPE}, req_no TEXT UNIQUE, category TEXT NOT NULL, from_dept TEXT NOT NULL, from_plant TEXT DEFAULT '', to_dept TEXT DEFAULT 'Store', status TEXT DEFAULT 'pending', remarks TEXT DEFAULT '', requested_by TEXT, requested_by_name TEXT, issued_by TEXT, issued_by_name TEXT, issued_to_name TEXT, received_by TEXT, received_by_name TEXT, created_at TEXT, updated_at TEXT, cancelled_at TEXT, cancelled_by TEXT, cancel_reason TEXT DEFAULT '');
    CREATE TABLE IF NOT EXISTS mrq_lines (id {IDTYPE}, req_id INT, material_code TEXT, description TEXT, qty_required REAL DEFAULT 0, batch_no TEXT DEFAULT '', qty_issued REAL DEFAULT 0, uom TEXT DEFAULT 'NOS', sort_order INT DEFAULT 0);
    """

    if USE_PG:
        cur.execute(schema_template.replace("{IDTYPE}", "SERIAL PRIMARY KEY"))
        try:
            cur.execute("SELECT setval(pg_get_serial_sequence('tcode_records','id'), COALESCE((SELECT MAX(id)+1 FROM tcode_records), 1), false);")
            cur.execute("SELECT setval(pg_get_serial_sequence('transfer_notes','id'), COALESCE((SELECT MAX(id)+1 FROM transfer_notes), 1), false);")
            conn.commit()
        except Exception: conn.rollback()
    else:
        cur.executescript(schema_template.replace("{IDTYPE}", "INTEGER PRIMARY KEY AUTOINCREMENT"))

    conn.commit()

    for tbl, col, ddl in [
        ("mrq_requisitions", "issued_to_name", "ALTER TABLE mrq_requisitions ADD COLUMN issued_to_name TEXT"),
        ("transfer_notes", "cc_users", "ALTER TABLE transfer_notes ADD COLUMN cc_users TEXT DEFAULT ''"),
    ]:
        try: cur.execute(q(f"SELECT {col} FROM {tbl} LIMIT 1"))
        except Exception:
            conn.rollback()
            try:
                cur.execute(ddl)
                conn.commit()
            except Exception: conn.rollback()

    cur_time = now()
    cur.execute(q("SELECT 1 FROM org WHERE id=1"))
    if not cur.fetchone():
        cur.execute(q("INSERT INTO org (id, name, system_id, client) VALUES (?,?,?,?)"), (1, "Thermocables Limited", "NX-01", "GLOBAL"))

    for dept_name in ["Production", "Final Testing", "Quality Control", "Commercial", "Store", "Planning", "Engineering"]:
        cur.execute(q("SELECT 1 FROM departments WHERE name=?"), (dept_name,))
        if not cur.fetchone():
            cur.execute(q("INSERT INTO departments (name, active, created_at) VALUES (?,?,?)"), (dept_name, 1, cur_time))

    for code, name in [("1000", "Gooty Central Plant"), ("1400", "Extrusion Unit 1400"), ("2000", "Hyderabad Works"), ("CORP", "Corporate Headquarters")]:
        cur.execute(q("SELECT 1 FROM plants WHERE code=?"), (code,))
        if not cur.fetchone():
            cur.execute(q("INSERT INTO plants (code, name, active, created_at) VALUES (?,?,?,?)"), (code, name, 1, cur_time))

    hashed = bcrypt.hashpw("Swami@265".encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
    cur.execute(q("SELECT username FROM users WHERE username='Admin_Erp'"))
    if cur.fetchone():
        cur.execute(q("UPDATE users SET password_hash=?, status='approved', role='master' WHERE username='Admin_Erp'"), (hashed,))
    else:
        cur.execute(q("INSERT INTO users (username, password_hash, full_name, role, department, plants, status, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?)"), ("Admin_Erp", hashed, "Master System Administrator", "master", "IT & Systems", "1000,1400,2000,CORP", "approved", cur_time, cur_time))

    # Clean up default sample accounts
    cur.execute(q("DELETE FROM users WHERE username IN ('prod_user', 'testing_lead', 'comm_lead', 'store_head', 'qc_lead')"))

    for code, title, desc, mod, depts in [
        ("TN01", "Transfer Note & Handover Sheet", "Inter-department material handover & batch transfer note", "MM", "Production,Final Testing,Commercial,Store,Quality Control,Planning"),
        ("CTST", "Cable Type-Test Calculator", "CR, Braid Coverage, Tubular & Dumbbell mechanical tests", "QC", "Production,Final Testing,Quality Control,Commercial,Store,Planning,Engineering"),
        ("MRQ1", "Material Requisition", "Department-to-Stores material request (Stores / Inhouse / Conductor)", "MM", "Production,Final Testing,Quality Control,Commercial,Store,Planning,Engineering"),
        ("PROD", "Production Handover Report", "Manufacturing batch & lot transfer logs", "PROD", "Production,Quality Control"),
        ("GCH0", "Quality Green Card Release", "Incoming QA & physical cable release inspection", "QC", "Quality Control,Production"),
        ("NCR0", "Non-Conformance Incidents", "QC deviation & defect corrective tracking", "QC", "Quality Control,Final Testing,Production"),
        ("PLN0", "Production Planning Details", "Machine allocation and manufacturing order queue", "PLAN", "Planning,Production,Store"),
        ("QCH0", "Stage QC to Final Clearance", "Clearance release slips before shipment", "QC", "Quality Control,Final Testing"),
        ("STD0", "Inventory Transfer Orders", "Inter-plant and store material movements", "INV", "Store,Planning,Production"),
    ]:
        cur.execute(q("SELECT 1 FROM tcodes WHERE code=?"), (code,))
        if not cur.fetchone():
            cur.execute(q("INSERT INTO tcodes (code, title, description, module, allowed_depts, active, created_at, updated_at, created_by) VALUES (?,?,?,?,?,?,?,?,?)"), (code, title, desc, mod, depts, 1, cur_time, cur_time, "SYSTEM"))
        else:
            if code == "CTST":
                cur.execute(q("UPDATE tcodes SET allowed_depts=? WHERE code='CTST'"), (depts,))

    # Fix: delete with correct column names (code for tcode_schemas, tcode_code for tcode_layouts)
    cur.execute(q("DELETE FROM tcode_schemas WHERE code='CTST'"))
    cur.execute(q("DELETE FROM tcode_layouts WHERE tcode_code='CTST'"))
    conn.commit()
    conn.close()

init_db()

def get_user(req: Request):
    token = None
    auth_header = req.headers.get("Authorization")
    if auth_header and auth_header.startswith("Bearer "):
        token = auth_header.split(" ")[1].strip()
    if not token: token = req.cookies.get("nexus_session")
    if not token: return None
    session = fetch_one("SELECT * FROM sessions WHERE token=?", (token,))
    if not session: return None
    user = fetch_one("SELECT * FROM users WHERE username=?", (session["username"],))
    if user and user["status"] == "approved": return user
    return None

def need(req: Request):
    user = get_user(req)
    if not user: raise HTTPException(401, "Authentication required.")
    return user

def is_master(user) -> bool:
    return user["role"] == "master" or user["username"] == "Admin_Erp"

def need_admin(req: Request):
    user = need(req)
    if user["role"] not in ["admin", "master"]: raise HTTPException(403, "Admin access required.")
    return user

def need_master(req: Request):
    user = need(req)
    if not is_master(user): raise HTTPException(403, "Master access required.")
    return user

def check_tcode_access(user, code: str) -> bool:
    code = code.upper().strip()
    if is_master(user) or user["role"] == "admin": return True
    tc = fetch_one("SELECT * FROM tcodes WHERE code=? AND active=1", (code,))
    if not tc: return False
    allowed = [x.strip() for x in (tc["allowed_depts"] or "").split(",") if x.strip()]
    if user["department"] in allowed: return True
    granted = fetch_one("SELECT 1 FROM access_grants WHERE username=? AND tcode_code=?", (user["username"], code))
    return bool(granted)

app = FastAPI(title="Nexus Core ERP", version="4.0.1")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True, allow_methods=["*"], allow_headers=["*"])

@app.post("/api/login")
def api_login(res: Response, username: str = Form(...), password: str = Form(...)):
    username = username.strip()
    user = fetch_one("SELECT * FROM users WHERE LOWER(username)=LOWER(?)", (username,))
    if not user or not bcrypt.checkpw(password.encode("utf-8"), user["password_hash"].encode("utf-8")):
        raise HTTPException(401, "Invalid username or password.")
    if user["status"] != "approved": raise HTTPException(403, "Account is locked or pending approval.")
    token = secrets.token_hex(32)
    expires = (datetime.datetime.now() + datetime.timedelta(days=7)).strftime("%Y-%m-%d %H:%M:%S")
    exec_write("INSERT INTO sessions (token, username, expires) VALUES (?,?,?)", (token, user["username"], expires))
    res.set_cookie("nexus_session", token, httponly=True, samesite="lax", max_age=86400 * 7)
    return {"ok": True, "token": token, "user": {"username": user["username"], "full_name": user["full_name"], "role": user["role"]}}

@app.post("/api/logout")
def api_logout(res: Response, req: Request):
    token = (req.headers.get("Authorization", "").replace("Bearer ", "").strip() or req.cookies.get("nexus_session"))
    if token: exec_write("DELETE FROM sessions WHERE token=?", (token,))
    res.delete_cookie("nexus_session")
    return {"ok": True}

@app.get("/api/me")
def api_me(req: Request):
    user = need(req)
    org_info = fetch_one("SELECT * FROM org WHERE id=1") or {}
    return {"username": user["username"], "full_name": user["full_name"], "role": user["role"], "department": user["department"], "plants": [p for p in (user["plants"] or "").split(",") if p], "is_master": is_master(user), "system": org_info.get("system_id", "NX-01"), "client": org_info.get("client", "GLOBAL"), "company": org_info.get("name", "Thermocables Limited")}

@app.get("/api/spro/overview")
def spro_overview(req: Request):
    need(req)
    return {"org": fetch_one("SELECT * FROM org WHERE id=1"), "departments": fetch_all("SELECT * FROM departments ORDER BY name"), "plants": fetch_all("SELECT * FROM plants ORDER BY code")}

@app.post("/api/admin/org/update")
def spro_update_org(req: Request, name: str = Form(...), system_id: str = Form(...), client: str = Form(...)):
    need_master(req)
    exec_write("UPDATE org SET name=?, system_id=?, client=? WHERE id=1", (name.strip(), system_id.strip(), client.strip()))
    return {"ok": True}

@app.post("/api/admin/plants/create")
def spro_create_plant(req: Request, code: str = Form(...), name: str = Form(...)):
    need_master(req)
    code = code.upper().strip()
    if fetch_one("SELECT 1 FROM plants WHERE code=?", (code,)): raise HTTPException(400, f"Plant '{code}' already exists.")
    exec_write("INSERT INTO plants (code, name, active, created_at) VALUES (?,?,?,?)", (code, name.strip(), 1, now()))
    return {"ok": True}

@app.post("/api/admin/plants/update")
def spro_update_plant(req: Request, code: str = Form(...), name: str = Form(...), active: int = Form(1)):
    need_master(req)
    exec_write("UPDATE plants SET name=?, active=? WHERE code=?", (name.strip(), active, code.upper().strip()))
    return {"ok": True}

@app.post("/api/admin/plants/delete")
def spro_delete_plant(req: Request, code: str = Form(...)):
    need_master(req)
    exec_write("DELETE FROM plants WHERE code=?", (code.strip(),))
    return {"ok": True}

@app.post("/api/admin/departments/create")
def spro_create_dept(req: Request, name: str = Form(...)):
    need_master(req)
    name = name.strip()
    if fetch_one("SELECT 1 FROM departments WHERE name=?", (name,)): raise HTTPException(400, "Department already exists.")
    exec_write("INSERT INTO departments (name, active, created_at) VALUES (?,?,?)", (name, 1, now()))
    return {"ok": True}

@app.post("/api/admin/departments/delete")
def spro_delete_dept(req: Request, name: str = Form(...)):
    need_master(req)
    exec_write("DELETE FROM departments WHERE name=?", (name.strip(),))
    return {"ok": True}

@app.get("/api/su01/users")
def su01_list_users(req: Request):
    need_admin(req)
    return fetch_all("SELECT username, full_name, role, department, plants, status, created_at, updated_at FROM users ORDER BY username")

@app.post("/api/su01/create")
def su01_create_user(req: Request, username: str = Form(...), password: str = Form(...), full_name: str = Form(...), department: str = Form(...), plants: str = Form(""), role: str = Form("user")):
    caller = need_admin(req)
    username = username.strip()
    if role == "master" and not is_master(caller): raise HTTPException(403, "Only Master Admins can create master accounts.")
    if fetch_one("SELECT 1 FROM users WHERE LOWER(username)=LOWER(?)", (username,)): raise HTTPException(400, f"User '{username}' already exists.")
    hashed = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
    cur_time = now()
    exec_write("INSERT INTO users (username, password_hash, full_name, role, department, plants, status, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?)", (username, hashed, full_name.strip(), role, department.strip(), plants.strip(), "approved", cur_time, cur_time))
    return {"ok": True}

@app.post("/api/su01/update")
def su01_update_user(req: Request, username: str = Form(...), full_name: str = Form(...), department: str = Form(...), plants: str = Form(""), role: str = Form(...), status: str = Form(...)):
    caller = need_admin(req)
    target = fetch_one("SELECT * FROM users WHERE username=?", (username,))
    if not target: raise HTTPException(404, "User not found.")
    if is_master(target) and not is_master(caller): raise HTTPException(403, "Cannot modify a Master Administrator.")
    if role == "master" and not is_master(caller): raise HTTPException(403, "Cannot grant Master privileges.")
    exec_write("UPDATE users SET full_name=?, department=?, plants=?, role=?, status=?, updated_at=? WHERE username=?", (full_name.strip(), department.strip(), plants.strip(), role, status, now(), username))
    return {"ok": True}

@app.post("/api/su01/delete")
def su01_delete_user(req: Request, username: str = Form(...)):
    caller = need_admin(req)
    target = fetch_one("SELECT * FROM users WHERE username=?", (username,))
    if not target: raise HTTPException(404, "User not found.")
    if is_master(target): raise HTTPException(400, "Master account is protected from deletion.")
    if target["role"] == "admin" and not is_master(caller): raise HTTPException(403, "Only Master Admins can delete admin accounts.")
    if username == caller["username"]: raise HTTPException(400, "You cannot delete your own account.")
    exec_write("DELETE FROM users WHERE username=?", (username,))
    exec_write("DELETE FROM sessions WHERE username=?", (username,))
    exec_write("DELETE FROM access_grants WHERE username=?", (username,))
    return {"ok": True}

@app.post("/api/su01/reset-password")
def su01_reset_password(req: Request, username: str = Form(...), new_password: str = Form(...)):
    caller = need_admin(req)
    target = fetch_one("SELECT * FROM users WHERE username=?", (username,))
    if not target: raise HTTPException(404, "User not found.")
    if is_master(target) and not is_master(caller): raise HTTPException(403, "Cannot reset master account.")
    hashed = bcrypt.hashpw(new_password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
    exec_write("UPDATE users SET password_hash=?, updated_at=? WHERE username=?", (hashed, now(), username))
    exec_write("DELETE FROM sessions WHERE username=?", (username,))
    return {"ok": True}

@app.get("/api/pfcg/matrix")
def pfcg_matrix(req: Request):
    need_admin(req)
    return {"users": fetch_all("SELECT username, full_name, role, department FROM users ORDER BY username"), "tcodes": fetch_all("SELECT code, title, module, allowed_depts FROM tcodes WHERE active=1 ORDER BY code"), "grants": fetch_all("SELECT * FROM access_grants")}

@app.post("/api/pfcg/grant")
def pfcg_grant(req: Request, username: str = Form(...), tcode: str = Form(...)):
    caller = need_admin(req)
    u_val = username.strip()
    tc_val = tcode.upper().strip()
    if USE_PG:
        exec_write("INSERT INTO access_grants (username, tcode_code, granted_by, granted_at) VALUES (?,?,?,?) ON CONFLICT (username, tcode_code) DO NOTHING", (u_val, tc_val, caller["username"], now()))
    else:
        exec_write("INSERT OR IGNORE INTO access_grants (username, tcode_code, granted_by, granted_at) VALUES (?,?,?,?)", (u_val, tc_val, caller["username"], now()))
    return {"ok": True}

@app.post("/api/pfcg/revoke")
def pfcg_revoke(req: Request, username: str = Form(...), tcode: str = Form(...)):
    need_admin(req)
    exec_write("DELETE FROM access_grants WHERE username=? AND tcode_code=?", (username.strip(), tcode.upper().strip()))
    return {"ok": True}

@app.get("/api/se93/list")
def se93_list(req: Request):
    user = need(req)
    all_tcodes = fetch_all("SELECT * FROM tcodes ORDER BY code")
    for t in all_tcodes:
        t["layout_count"] = fetch_one("SELECT COUNT(*) AS c FROM tcode_layouts WHERE tcode_code=?", (t["code"],))["c"]
        t["record_count"] = fetch_one("SELECT COUNT(*) AS c FROM tcode_records WHERE tcode_code=? AND status='active'", (t["code"],))["c"]
        t["has_access"] = check_tcode_access(user, t["code"])
    return all_tcodes

@app.post("/api/se93/auto-create")
async def se93_auto_create(req: Request, code: str = Form(...), title: str = Form(...), description: str = Form(""), module: str = Form("PROD"), allowed_depts: str = Form("Production,Quality Control"), import_sample_data: int = Form(1), file: UploadFile = File(...)):
    user = need_admin(req)
    code = code.upper().strip()
    if not code: raise HTTPException(400, "T-Code identifier is required.")
    if fetch_one("SELECT 1 FROM tcodes WHERE code=?", (code,)): raise HTTPException(400, f"Module '{code}' already exists.")
    contents = await file.read()
    filename = file.filename.lower()
    headers = []
    data_rows = []
    try:
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
            wb = load_workbook(io.BytesIO(contents), data_only=True)
            rows = list(wb.active.iter_rows(values_only=True))
            if rows:
                headers = [str(h).strip() for h in rows[0] if h is not None and str(h).strip()]
                data_rows = rows[1:]
        elif filename.endswith(".csv"):
            text = contents.decode("utf-8", errors="ignore")
            rows = list(csv.reader(io.StringIO(text)))
            if rows:
                headers = [str(h).strip() for h in rows[0] if str(h).strip()]
                data_rows = rows[1:]
        elif filename.endswith(".json"):
            parsed = json.loads(contents.decode("utf-8", errors="ignore"))
            if isinstance(parsed, list) and parsed and isinstance(parsed[0], dict):
                headers = list(parsed[0].keys())
                for item in parsed: data_rows.append([item.get(k, "") for k in headers])
        else: raise HTTPException(400, "Upload .xlsx / .xls / .csv / .json only.")
    except HTTPException: raise
    except Exception as e: raise HTTPException(400, f"Template error: {str(e)}")
    if not headers: raise HTTPException(400, "No valid column headers found.")
    col_samples = {idx: [] for idx in range(len(headers))}
    for row in data_rows[:25]:
        for ci in range(len(headers)):
            if ci < len(row) and row[ci] is not None: col_samples[ci].append(row[ci])
    fields = []
    for idx, h in enumerate(headers):
        clean_key = re.sub(r"[^a-zA-Z0-9_]", "_", h.lower().strip())
        clean_key = re.sub(r"_+", "_", clean_key).strip("_") or f"col_{idx+1}"
        ftype = infer_field_type(col_samples[idx])
        fields.append({"key": clean_key, "label": h, "type": ftype, "required": False, "width": 2 if ftype == "textarea" else 1})
    cur_time = now()
    exec_write("INSERT INTO tcodes (code, title, description, module, allowed_depts, active, created_at, updated_at, created_by) VALUES (?,?,?,?,?,?,?,?,?)", (code, title.strip() or f"{code} Module", description.strip(), module.strip(), allowed_depts.strip(), 1, cur_time, cur_time, user["username"]))
    fields_json = json.dumps(fields)
    exec_write("INSERT INTO tcode_schemas (code, fields, layout, updated_at, updated_by) VALUES (?,?,?,?,?)", (code, fields_json, json.dumps({"columns": 3}), cur_time, user["username"]))
    exec_write("INSERT INTO tcode_layouts (tcode_code, layout_code, name, title, is_default, fields, layout_config, description, created_by, created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)", (code, f"GRID_{code}_01", "Default Grid", title or code, 1, fields_json, json.dumps({"columns": 3}), "Auto-generated", user["username"], user["full_name"], cur_time, cur_time))
    imported = 0
    if import_sample_data and data_rows:
        conn = get_conn()
        try:
            cur = conn.cursor()
            insert_sql = q("INSERT INTO tcode_records (tcode_code, plant, data, status, created_by, created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?)")
            default_plant = ((user.get("plants") or "1000").split(",")[0].strip() or "1000")
            for row in data_rows:
                if not any(row): continue
                rec = {}
                row_plant = default_plant
                for idx in range(len(headers)):
                    if idx < len(row):
                        v = row[idx]
                        if v is None: v = ""
                        elif isinstance(v, (datetime.datetime, datetime.date)):
                            if isinstance(v, datetime.datetime): v = v.strftime("%Y-%m-%d %H:%M:%S")
                            else: v = v.strftime("%Y-%m-%d")
                        key = fields[idx]["key"]
                        if key in ["plant", "facility"]: row_plant = str(v or default_plant).strip()
                        else: rec[key] = v
                cur.execute(insert_sql, (code, row_plant, json.dumps(rec, default=json_serial), "active", user["username"], user["full_name"], cur_time, cur_time))
                imported += 1
            conn.commit()
        finally: conn.close()
    return {"ok": True, "code": code, "title": title, "columns_count": len(fields), "imported_count": imported}

@app.post("/api/se93/delete")
def se93_delete(req: Request, code: str = Form(...)):
    need_master(req)
    code = code.upper().strip()
    if code in STUDIO_PROTECTED: raise HTTPException(403, f"'{code}' is a protected system module and cannot be deleted.")
    exec_write("DELETE FROM tcodes WHERE code=?", (code,))
    exec_write("DELETE FROM tcode_schemas WHERE code=?", (code,))
    exec_write("DELETE FROM tcode_layouts WHERE tcode_code=?", (code,))
    exec_write("DELETE FROM tcode_records WHERE code=?", (code,))
    exec_write("DELETE FROM access_grants WHERE tcode_code=?", (code,))
    return {"ok": True}

# =========================================================================
# TN01 — TRANSFER NOTE
# =========================================================================
TN01_ROUTES = {
    "PROD_TO_QC": {"title": "Production to Final Testing", "route": "Production ➔ Final Testing", "from_dept": "Production", "to_dept": "Final Testing", "columns": ["Sr. No.", "Job No.", "Customer Name", "Material Description", "Specified Length", "Entry Length", "Drum Size", "Drum Code", "Batch No."]},
    "QC_TO_COMM": {"title": "Final Testing to Commercial", "route": "Final Testing ➔ Commercial", "from_dept": "Final Testing", "to_dept": "Commercial", "columns": ["Sr. No.", "Job / Prod Order #", "Batch / Heat Lot", "Cable Description / Spec", "Drum #", "Tested Length (Mtrs)", "HV Test Result", "Insulation Resistance", "QC Inspector", "Release Status"]},
    "PROD_TO_STORE": {"title": "Production to Store", "route": "Production ➔ Store", "from_dept": "Production", "to_dept": "Store", "columns": ["Sr. No.", "Work Order #", "Item Code / Material", "Batch / Heat Lot", "Quantity (Mtrs/Kgs)", "Weight (Kgs)", "Drum / Pack #", "Handover Person", "Remarks"]},
    "QC_TO_PROD": {"title": "Final Testing to Production (Rework/Reject)", "route": "Final Testing ➔ Production", "from_dept": "Final Testing", "to_dept": "Production", "columns": ["Sr. No.", "Job Order #", "Batch / Heat Lot", "Cable Size", "Defect Observed", "Length (Mtrs)", "QC Engineer", "Rework Instructions"]},
    "COMM_TO_QC": {"title": "Commercial to Final Testing", "route": "Commercial ➔ Final Testing", "from_dept": "Commercial", "to_dept": "Final Testing", "columns": ["Sr. No.", "Customer Name", "Order / PO #", "Item Description", "Ordered Qty", "Inspection Agency", "Target Date", "Special Testing Notes"]},
    "STORE_TO_PROD": {"title": "Store to Production", "route": "Store ➔ Production", "from_dept": "Store", "to_dept": "Production", "columns": ["Sr. No.", "Material Code", "Material Description", "Lot / GRN #", "Issued Qty", "UOM", "Issued By", "Received For Work Order"]},
}

@app.get("/api/tn01/routes")
def tn01_get_routes(req: Request):
    need(req)
    return TN01_ROUTES

@app.get("/api/tn01/suggest-users")
def tn01_suggest_users(req: Request, plant: str = "", dept: str = "", q: str = ""):
    need(req)
    sql = "SELECT username, full_name, department, plants, role FROM users WHERE status='approved'"
    params = []
    if dept:
        sql += " AND department = ?"
        params.append(dept)
    if q:
        sql += " AND (LOWER(username) LIKE ? OR LOWER(full_name) LIKE ?)"
        term = f"%{q.strip().lower()}%"
        params.extend([term, term])
    users = fetch_all(sql, tuple(params))
    if plant:
        return [u for u in users if not (u.get("plants") or "").strip() or plant in [p.strip() for p in (u.get("plants") or "").split(",")]]
    return users

@app.get("/api/tn01/layouts")
def tn01_get_layouts(req: Request):
    need(req)
    rows = fetch_all("SELECT * FROM tcode_layouts WHERE tcode_code='TN01' ORDER BY is_default DESC, id ASC")
    for r in rows:
        try: r["fields"] = json.loads(r["fields"])
        except Exception: r["fields"] = []
    return rows

@app.post("/api/tn01/send")
def tn01_send(req: Request, payload: dict):
    user = need(req)
    cur_time = now()
    today_tag = datetime.date.today().strftime("%Y%m%d")
    cnt = fetch_one("SELECT COUNT(*) AS c FROM transfer_notes WHERE created_at LIKE ?", (f"{datetime.date.today().strftime('%Y-%m-%d')}%",))
    note_no = f"TN-{today_tag}-{(cnt['c'] if cnt else 0) + 1:04d}"
    headers = payload.get("headers", [])
    rows = payload.get("rows", [])
    header_count = len(headers) if headers else 0
    normalized_rows = []
    for i, raw in enumerate(rows, 1):
        if not isinstance(raw, (list, tuple)): continue
        row = ["" if c is None else str(c) for c in raw]
        while len(row) < header_count: row.append("")
        if header_count: row = row[:header_count]
        row[0] = str(i)
        normalized_rows.append(row)
    normalized_rows = [r for r in normalized_rows if any(str(x).strip() for x in r[1:])]
    to_user = payload.get("to_user", "").strip()
    target = fetch_one("SELECT username, full_name FROM users WHERE username=?", (to_user,))
    to_user_name = target["full_name"] if target else to_user
    cc_users_raw = payload.get("cc_users") or []
    if isinstance(cc_users_raw, str): cc_users_raw = [cc_users_raw]
    cc_users = ",".join(str(x).strip() for x in cc_users_raw if str(x).strip())
    from_plant = (payload.get("from_plant", "").strip() or (user.get("plants") or "1000").split(",")[0].strip())
    exec_write("INSERT INTO transfer_notes (note_no, layout_key, route, from_plant, from_dept, from_user, from_user_name, to_plant, to_dept, to_user, to_user_name, entry_count, sheet_headers, sheet_data, status, remarks, created_at, cc_users) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (note_no, payload.get("layout_key", "STANDARD"), payload.get("route", "Transfer"), from_plant, user["department"], user["username"], user["full_name"], payload.get("to_plant", "").strip(), payload.get("to_dept", "").strip(), to_user, to_user_name, len(normalized_rows), json.dumps(headers), json.dumps(normalized_rows, default=json_serial), "sent", payload.get("remarks", "").strip(), cur_time, cc_users))
    return {"ok": True, "note_no": note_no, "entries": len(normalized_rows)}

@app.get("/api/tn01/inbox")
def tn01_inbox(req: Request):
    user = need(req)
    uname = user["username"]
    return fetch_all("SELECT * FROM transfer_notes WHERE to_user = ? OR (to_dept = ? AND to_user = '') OR (',' || cc_users || ',') LIKE ? ORDER BY id DESC", (uname, user["department"], f"%,{uname},%"))

@app.post("/api/tn01/receive/{nid}")
def tn01_receive(req: Request, nid: int):
    user = need(req)
    cur_time = now()
    exec_write("UPDATE transfer_notes SET status='received', received_at=?, received_by=?, received_by_name=? WHERE id=?", (cur_time, user["username"], user["full_name"], nid))
    return {"ok": True, "received_at": cur_time}

@app.delete("/api/tn01/delete/{nid}")
def tn01_delete(req: Request, nid: int):
    user = need(req)
    if not (is_master(user) or user["role"] == "admin"): raise HTTPException(403, "Only Admin can delete history.")
    exec_write("DELETE FROM transfer_notes WHERE id=?", (nid,))
    return {"ok": True}

@app.get("/api/tn01/history")
def tn01_history(req: Request, plant: str = "", route: str = "", status: str = "", from_dept: str = "", to_dept: str = "", date_from: str = "", date_to: str = ""):
    user = need(req)
    uname = user["username"]
    sql = "SELECT * FROM transfer_notes WHERE 1=1"
    params = []
    if not (is_master(user) or user["role"] == "admin"):
        sql += " AND (from_user = ? OR to_user = ? OR from_dept = ? OR to_dept = ? OR (',' || cc_users || ',') LIKE ?)"
        params.extend([user["username"], user["username"], user["department"], user["department"], f"%,{uname},%"])
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

@app.get("/api/tn01/note/{nid}")
def tn01_get_note(req: Request, nid: int):
    need(req)
    note = fetch_one("SELECT * FROM transfer_notes WHERE id=?", (nid,))
    if not note: raise HTTPException(404, "Note not found.")
    try: note["sheet_headers"] = json.loads(note["sheet_headers"])
    except Exception: note["sheet_headers"] = []
    try: note["sheet_data"] = json.loads(note["sheet_data"])
    except Exception: note["sheet_data"] = []
    return note

@app.get("/api/tn01/export/{nid}")
def tn01_export_excel(req: Request, nid: int):
    user = need(req)
    note = fetch_one("SELECT * FROM transfer_notes WHERE id=?", (nid,))
    if not note: raise HTTPException(404, "Note not found.")
    is_sender = (user["username"] == note["from_user"]) or (user["department"] == note["from_dept"])
    is_receiver = (user["username"] == note["to_user"]) or (user["department"] == note["to_dept"])
    cc_hit = f",{user['username']}," in f",{note.get('cc_users') or ''},"
    if not (is_sender or is_receiver or cc_hit or is_master(user) or user["role"] == "admin"): raise HTTPException(403, "Access denied.")
    try: headers = json.loads(note["sheet_headers"])
    except Exception: headers = []
    try: data_rows = json.loads(note["sheet_data"])
    except Exception: data_rows = []
    wb = Workbook()
    ws = wb.active
    ws.title = "Transfer_Note"
    org = fetch_one("SELECT * FROM org WHERE id=1") or {}
    ws.append([org.get("name", "Thermocables Limited")])
    ws.append([f"TRANSFER NOTE {note['note_no']} — {note['route']}"])
    ws.append([f"From: {note['from_plant']} - {note['from_dept']}  |  To: {note['to_plant']} - {note['to_dept']}"])
    ws.append([f"Sender: {note['from_user_name']}  |  Receiver: {note['to_user_name']}  |  Status: {note['status'].upper()}"])
    cc_line = f"  |  CC: {note.get('cc_users')}" if note.get("cc_users") else ""
    ws.append([f"Created: {note['created_at']}  |  Received: {note['received_at'] or 'Pending'}{cc_line}"])
    ws.append([])
    ws.append(headers)
    for row in data_rows: ws.append(row)
    ws.append([])
    ws.append(["Handovered By: ________________", "", "Received By: ________________", "", "QC: ________________"])
    from openpyxl.styles import Font, Alignment, PatternFill
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
        ws.column_dimensions[ws.cell(row=7, column=col_idx).column_letter].width = max(14, min(40, len(str(h)) + 4))
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return StreamingResponse(bio, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f"attachment; filename=TransferNote_{note['note_no']}.xlsx"})

@app.post("/api/tn01/export-grid")
def tn01_export_grid(req: Request, payload: dict):
    user = need(req)
    headers = payload.get("headers", [])
    rows = payload.get("rows", [])
    title = payload.get("title", "Transfer Note (Draft)")
    if not headers: raise HTTPException(400, "No columns to export.")
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
    for r in rows: ws.append(r)
    from openpyxl.styles import Font, Alignment, PatternFill
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
        ws.column_dimensions[ws.cell(row=header_row, column=col_idx).column_letter].width = max(14, min(40, len(str(h)) + 4))
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return StreamingResponse(bio, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f"attachment; filename=TN01_Draft_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"})

@app.get("/api/tn01/print/{nid}", response_class=HTMLResponse)
def tn01_print_preview(req: Request, nid: int):
    need(req)
    note = fetch_one("SELECT * FROM transfer_notes WHERE id=?", (nid,))
    if not note: return HTMLResponse("Note not found", 404)
    org = fetch_one("SELECT * FROM org WHERE id=1") or {}
    try: headers = json.loads(note["sheet_headers"])
    except Exception: headers = []
    try: data_rows = json.loads(note["sheet_data"])
    except Exception: data_rows = []
    th = "".join(f"<th style='border:1px solid #0f172a; padding:6px 8px; font-size:11px; background:#e0e7ff'>{h}</th>" for h in headers)
    trs = "".join(f"<tr>{''.join(f'<td style=\"border:1px solid #cbd5e1; padding:6px 8px; font-size:11px\">{c}</td>' for c in r)}</tr>" for r in data_rows)
    return HTMLResponse(f"""<!DOCTYPE html><html><head><meta charset="utf-8"><title>{note['note_no']}</title><style>body {{ font-family: sans-serif; margin: 30px; color: #0f172a; }} .header {{ border-bottom: 2px solid #4f46e5; padding-bottom: 10px; margin-bottom: 15px; }} .meta {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; font-size: 12px; background: #f8fafc; padding: 12px; border: 1px solid #e2e8f0; margin-bottom: 20px; }} table {{ width: 100%; border-collapse: collapse; margin-bottom: 30px; }} .signs {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-top: 40px; text-align: center; font-size: 11px; font-weight: bold; }} .sign-box {{ border-top: 1px solid #0f172a; padding-top: 8px; }} @media print {{ .no-print {{ display: none; }} body {{ margin: 0; }} }}</style></head><body><div class="no-print" style="margin-bottom:15px;text-align:right"><button onclick="window.print()" style="background:#4f46e5;color:white;border:none;padding:8px 16px;border-radius:6px;cursor:pointer;font-weight:bold">Print / PDF</button></div><div class="header"><h2 style="margin:0">{org.get('name', 'Thermocables')}</h2><h4 style="margin:4px 0 0;color:#4f46e5">TRANSFER NOTE (TN01)</h4><div style="font-family:monospace;font-weight:bold;margin-top:4px">{note['note_no']} | {note['created_at']}</div></div><div class="meta"><div><b>Route:</b><br>{note['route']}</div><div><b>From:</b><br>{note['from_plant']} - {note['from_dept']}</div><div><b>To:</b><br>{note['to_plant']} - {note['to_dept']}</div><div><b>Entries:</b><br>{note['entry_count']}</div><div><b>Sender:</b><br>{note['from_user_name']}</div><div><b>Recipient:</b><br>{note['to_user_name']}</div><div><b>Status:</b><br>{note['status'].upper()}</div><div><b>Receipt:</b><br>{note['received_at'] or 'Pending'}</div></div><div style="font-size:11px;margin-bottom:10px"><b>CC:</b> {note.get('cc_users') or '—'}</div><table><thead><tr>{th}</tr></thead><tbody>{trs}</tbody></table><div class="signs"><div class="sign-box">Handovered By ({note['from_user_name']})</div><div class="sign-box">Production | QC</div><div class="sign-box">Received By ({note['received_by_name'] or note['to_user_name']})</div></div></body></html>""")

@app.get("/api/tcode/{code}/layouts")
def get_layouts(code: str, req: Request):
    user = need(req)
    if not check_tcode_access(user, code): raise HTTPException(403, "Access denied.")
    rows = fetch_all("SELECT * FROM tcode_layouts WHERE tcode_code=? ORDER BY is_default DESC, id ASC", (code,))
    for r in rows:
        try: r["fields"] = json.loads(r["fields"])
        except Exception: r["fields"] = []
    return rows

@app.get("/api/tcode/{code}/records")
def get_tcode_records(code: str, req: Request, plant: str = "", limit: int = 500):
    user = need(req)
    if not check_tcode_access(user, code): raise HTTPException(403, "Access denied.")
    sql = "SELECT * FROM tcode_records WHERE tcode_code=? AND status='active'"
    params = [code]
    if plant:
        sql += " AND plant=?"
        params.append(plant)
    sql += " ORDER BY id DESC LIMIT ?"
    params.append(limit)
    rows = fetch_all(sql, tuple(params))
    for r in rows:
        try: r["data"] = json.loads(r["data"])
        except Exception: r["data"] = {}
    return rows

@app.post("/api/tcode/{code}/records")
def create_record(code: str, req: Request, payload: dict):
    user = need(req)
    if not check_tcode_access(user, code): raise HTTPException(403, "Access denied.")
    plant = payload.pop("_plant", "").strip()
    cur_time = now()
    exec_write("INSERT INTO tcode_records (tcode_code, plant, data, status, created_by, created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?)", (code, plant, json.dumps(payload, default=json_serial), "active", user["username"], user["full_name"], cur_time, cur_time))
    return {"ok": True}

@app.delete("/api/tcode/{code}/records/{rid}")
def delete_record(code: str, rid: int, req: Request):
    user = need(req)
    if not check_tcode_access(user, code): raise HTTPException(403, "Access denied.")
    rec = fetch_one("SELECT * FROM tcode_records WHERE id=? AND tcode_code=?", (rid, code))
    if not rec: raise HTTPException(404, "Record not found.")
    if (rec["created_by"] != user["username"] and not is_master(user) and user["role"] != "admin"): raise HTTPException(403, "You can only delete your own records.")
    exec_write("UPDATE tcode_records SET status='deleted', updated_at=? WHERE id=?", (now(), rid))
    return {"ok": True}

@app.get("/api/tcode/{code}/export")
def export_excel(code: str, req: Request):
    user = need(req)
    if not check_tcode_access(user, code): raise HTTPException(403, "Access denied.")
    rows = fetch_all("SELECT * FROM tcode_records WHERE tcode_code=? AND status='active' ORDER BY id ASC", (code,))
    wb = Workbook()
    ws = wb.active
    ws.title = f"{code}_DATA"
    base_keys = ["id", "plant", "created_by_name", "created_at"]
    dynamic_keys = set()
    parsed_rows = []
    for r in rows:
        try: d = json.loads(r["data"]) if isinstance(r["data"], str) else r["data"]
        except Exception: d = {}
        dynamic_keys.update(d.keys())
        parsed_rows.append((r, d))
    all_cols = base_keys + sorted(dynamic_keys)
    ws.append(all_cols)
    for r, d in parsed_rows:
        ws.append([r["id"], r["plant"], r["created_by_name"], r["created_at"]] + [d.get(k, "") for k in sorted(dynamic_keys)])
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return StreamingResponse(bio, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f"attachment; filename=Export_{code}_{datetime.date.today().strftime('%Y%m%d')}.xlsx"})

HTML_UI = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Nexus Core — Enterprise Operations</title>
<script src="https://cdn.tailwindcss.com"></script>
<link href="https://cdn.jsdelivr.net/npm/remixicon@4.2.0/fonts/remixicon.css" rel="stylesheet">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">
<style>
body{font-family:'Inter',system-ui,sans-serif;background:#f8fafc;color:#0f172a;font-size:13px}
.font-mono{font-family:'JetBrains Mono',monospace}
.btn-primary{background:#4f46e5;color:#fff;border:1px solid #4f46e5;transition:background .15s,box-shadow .15s}
.btn-primary:hover{background:#4338ca;box-shadow:0 4px 12px rgba(79,70,229,.20)}
.btn-secondary{background:#fff;color:#475569;border:1px solid #e2e8f0;transition:background .15s,border-color .15s}
.btn-secondary:hover{background:#f1f5f9;border-color:#cbd5e1}
.card{background:#fff;border:1px solid #e2e8f0;border-radius:14px;box-shadow:0 1px 2px rgba(15,23,42,.03)}
.custom-table{width:100%;border-collapse:collapse}
.custom-table thead th{background:#f1f5f9;color:#475569;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.03em;padding:10px 14px;border-bottom:1px solid #e2e8f0;text-align:left}
.custom-table tbody td{padding:11px 14px;border-bottom:1px solid #f1f5f9;font-size:13px;color:#334155}
.custom-table tbody tr:hover{background:#f8fafc}
.excel-grid{border-collapse:separate;border-spacing:0;width:100%}
.excel-grid th{background:#f1f5f9;color:#475569;font-weight:600;font-size:11px;text-transform:uppercase;padding:8px 10px;border-right:1px solid #e2e8f0;border-bottom:1px solid #e2e8f0;position:sticky;top:0;z-index:2;text-align:left}
.excel-grid td{background:#fff;color:#0f172a;font-size:12px;padding:6px 10px;border-right:1px solid #f1f5f9;border-bottom:1px solid #f1f5f9;outline:none;min-width:110px}
.excel-grid td.sr-col{background:#f8fafc;color:#64748b;font-weight:600;text-align:center;width:50px;user-select:none}
.excel-grid td:focus{background:#eef2ff;box-shadow:inset 0 0 0 2px #6366f1}
.form-input{background:#fff;border:1px solid #cbd5e1;border-radius:8px;padding:8px 12px;font-size:13px;color:#0f172a;width:100%;outline:none;transition:border-color .15s,box-shadow .15s;font-family:inherit}
.form-input:focus{border-color:#6366f1;box-shadow:0 0 0 3px rgba(99,102,241,.12)}
.form-label{display:block;font-size:11px;font-weight:600;color:#475569;margin-bottom:5px;letter-spacing:.01em}
.badge{display:inline-flex;align-items:center;font-size:10px;font-weight:600;padding:2px 8px;border-radius:999px;letter-spacing:.02em;text-transform:uppercase}
.badge-indigo{background:#eef2ff;color:#4338ca;border:1px solid #c7d2fe}
.badge-emerald{background:#ecfdf5;color:#047857;border:1px solid #a7f3d0}
.badge-amber{background:#fffbeb;color:#b45309;border:1px solid #fde68a}
.badge-red{background:#fef2f2;color:#b91c1c;border:1px solid #fecaca}
.badge-slate{background:#f1f5f9;color:#475569;border:1px solid #e2e8f0}
.sidebar-link{color:#475569;border-left:2px solid transparent;transition:all .15s ease}
.sidebar-link:hover{background:#f1f5f9;color:#0f172a}
.sidebar-link.active{background:#eef2ff;color:#4f46e5;border-left-color:#4f46e5}
</style>
</head>
<body class="h-screen flex flex-col overflow-hidden">
<div id="app_root" class="h-full flex flex-col">
  <header class="h-14 bg-white border-b border-slate-200 px-4 flex items-center justify-between shrink-0 shadow-sm">
    <div class="flex items-center gap-4">
      <div class="flex items-center gap-2.5 cursor-pointer" onclick="goHome()">
        <div class="w-8 h-8 rounded-lg bg-indigo-600 flex items-center justify-center font-bold text-white shadow-sm"><i class="ri-pulse-fill text-lg"></i></div>
        <div class="font-bold text-sm tracking-tight text-slate-900 flex items-center gap-1.5">Nexus Core <span class="text-[10px] bg-indigo-50 text-indigo-600 border border-indigo-200 font-mono px-1.5 py-0.5 rounded">v4.0.1</span></div>
      </div>
      <div class="relative ml-4">
        <i class="ri-search-line absolute left-3 top-2.5 text-slate-400 text-xs"></i>
        <input id="cmd_box" placeholder="Jump to module..." class="bg-slate-50 border border-slate-200 rounded-lg pl-8 pr-4 py-1.5 text-xs text-slate-900 w-96 font-mono focus:outline-none focus:border-indigo-400 focus:bg-white" onkeydown="if(event.key==='Enter') execCmd()">
      </div>
    </div>
    <div class="flex items-center gap-3">
      <div class="flex items-center gap-2 bg-slate-50 border border-slate-200 px-3 py-1 rounded-full text-xs">
        <div class="w-2 h-2 rounded-full bg-emerald-500"></div>
        <span id="user_display" class="font-medium text-slate-700">Connecting...</span>
        <span id="role_badge" class="badge badge-indigo"></span>
      </div>
      <button onclick="logout()" class="text-xs bg-white hover:bg-red-50 text-slate-500 hover:text-red-600 border border-slate-200 p-2 rounded-lg transition"><i class="ri-logout-box-r-line"></i></button>
    </div>
  </header>
  <div class="flex-1 flex overflow-hidden">
    <aside class="w-64 bg-white border-r border-slate-200 flex flex-col justify-between shrink-0">
      <div class="p-3 overflow-y-auto flex-1 space-y-4">
        <div>
          <div class="text-[10px] uppercase font-bold text-slate-400 tracking-wider px-2 mb-1.5">Special Modules</div>
          <div class="space-y-0.5">
            <button onclick="openTcode('TN01')" class="w-full text-left px-2.5 py-2 rounded-lg text-xs font-semibold bg-slate-50 border border-slate-200 text-slate-700 hover:bg-slate-100 flex items-center justify-between transition">
              <span class="flex items-center gap-2"><i class="ri-exchange-box-line text-slate-500 text-sm"></i> TN01 - Transfer Note</span>
              <span class="badge badge-slate">EXCEL</span>
            </button>
            <button onclick="openTcode('MRQ1')" class="w-full text-left px-2.5 py-2 rounded-lg text-xs font-semibold bg-emerald-50 border border-emerald-200 text-emerald-700 hover:bg-emerald-100 flex items-center justify-between transition mt-1">
              <span class="flex items-center gap-2"><i class="ri-file-list-3-line text-emerald-500 text-sm"></i> MRQ1 - Requisition</span>
            </button>
            <button onclick="openTcode('CTST')" class="w-full text-left px-2.5 py-2 rounded-lg text-xs font-semibold bg-indigo-50 border border-indigo-200 text-indigo-700 hover:bg-indigo-100 flex items-center justify-between transition mt-1">
              <span class="flex items-center gap-2"><i class="ri-flask-line text-indigo-500 text-sm"></i> CTST - Type-Test</span>
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
      <span id="stat_resp">Ready</span>
    </div>
  </footer>
</div>
<div id="app_dialog" class="hidden fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center p-4">
  <div id="app_dialog_body" class="bg-white text-slate-800 rounded-2xl shadow-2xl border border-slate-200 max-h-[90vh] overflow-y-auto w-full max-w-4xl"></div>
</div>
<script>
let ME = null;
let CURRENT_TCODE = null;
let ALL_TCODES = [];
const $ = s => document.querySelector(s);

function esc(str){if(str===null||str===undefined)return '';return String(str).replace(/[&<>"']/g,s=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[s]));}
function setMsg(m,e=false){const el=$('#stat_msg'),d=$('#stat_dot');if(!el)return;el.textContent=m;d.className=e?"w-2 h-2 rounded-full bg-red-500":"w-2 h-2 rounded-full bg-emerald-500";el.className=e?"text-red-600 font-semibold":"text-slate-600";}

async function api(p,o={}){
  const t=localStorage.getItem("nexus_token");
  const h={...(o.headers||{})};
  if(t)h["Authorization"]=`Bearer ${t}`;
  const t0=performance.now();
  const r=await fetch(p,{...o,headers:h,credentials:'include'});
  if($('#stat_resp'))$('#stat_resp').textContent=Math.round(performance.now()-t0)+'ms';
  if(r.status===401){localStorage.removeItem("nexus_token");renderLogon();throw 'Auth';}
  if(!r.ok){let x=await r.text();try{x=JSON.parse(x).detail||x;}catch{}throw x;}
  return await r.json();
}

function renderLogon(err=""){
  $('#app_root').innerHTML=`
  <div class="min-h-screen bg-gradient-to-br from-slate-50 to-indigo-50 flex items-center justify-center p-4">
    <div class="w-full max-w-md bg-white border border-slate-200 rounded-2xl shadow-xl p-8">
      <div class="flex items-center gap-3 mb-6">
        <div class="w-10 h-10 rounded-xl bg-indigo-600 flex items-center justify-center font-bold text-white shadow-sm"><i class="ri-pulse-fill text-2xl"></i></div>
        <div><h1 class="text-xl font-bold text-slate-900">Nexus Core ERP</h1><p class="text-xs text-slate-500">Thermocables Industrial Suite</p></div>
      </div>
      ${err?`<div class="mb-4 p-3 bg-red-50 border border-red-200 rounded-xl text-red-700 text-xs">${esc(err)}</div>`:''}
      <div class="space-y-4">
        <div><label class="form-label">Username</label><input id="l_u" class="form-input" placeholder="Admin_Erp" autofocus></div>
        <div><label class="form-label">Password</label><input id="l_p" type="password" class="form-input" placeholder="••••••••" onkeydown="if(event.key==='Enter')handleLogon()"></div>
        <button onclick="handleLogon()" class="w-full btn-primary py-2.5 rounded-lg font-semibold text-sm">Sign In</button>
      </div>
    </div>
  </div>`;
}
async function handleLogon(){
  const u=$('#l_u').value.trim(),p=$('#l_p').value;
  if(!u||!p)return renderLogon("Enter credentials.");
  const fd=new FormData();fd.append("username",u);fd.append("password",p);
  try{const r=await fetch('/api/login',{method:'POST',body:fd});if(!r.ok){let t=await r.text();try{t=JSON.parse(t).detail||t;}catch{}return renderLogon(t);}
    const d=await r.json();if(d.token)localStorage.setItem("nexus_token",d.token);location.reload();}
  catch(e){renderLogon("Error: "+e);}
}
async function logout(){await api('/api/logout',{method:'POST'});localStorage.removeItem("nexus_token");location.reload();}

async function initSession(){
  try{
    ME=await api('/api/me');
    $('#user_display').textContent=ME.full_name||ME.username;
    $('#role_badge').textContent=ME.role;
    $('#side_company').textContent=ME.company;
    $('#side_plant').textContent="Facilities: "+(ME.plants.join(", ")||"All");
    if(ME.is_master||ME.role==='admin')$('#admin_nav_section').classList.remove('hidden');
    setMsg(`Connected: ${ME.username}`);
    await loadModules();goHome();
  }catch(e){renderLogon();}
}

function execCmd(){let v=$('#cmd_box').value.trim();if(!v)return;$('#cmd_box').value='';if(v.startsWith('/n'))v=v.substring(2).trim();if(!v){goHome();return;}openTcode(v.toUpperCase());}

function goHome(){
  CURRENT_TCODE=null;$('#stat_tcode').textContent="DASHBOARD";
  $('#main_screen').innerHTML=`
    <div class="flex-1 overflow-y-auto p-8 select-text">
      <div class="max-w-5xl mx-auto space-y-8">
        <div><h1 class="text-2xl font-bold text-slate-900">Operations Dashboard</h1><p class="text-sm text-slate-500 mt-1">Thermocables Industrial Platform</p></div>
        <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div class="card p-5"><div class="text-[10px] font-semibold text-slate-500 uppercase tracking-wider">Enterprise</div><div class="text-lg font-bold text-slate-900 mt-1.5">${esc(ME.company)}</div><div class="text-xs text-indigo-600 mt-1">Dept: ${esc(ME.department)}</div></div>
          <div class="card p-5"><div class="text-[10px] font-semibold text-slate-500 uppercase tracking-wider">Role</div><div class="text-lg font-bold text-emerald-600 mt-1.5 uppercase">${ME.role}</div><div class="text-xs text-slate-500 mt-1">Plants: ${esc(ME.plants.join(', ')||'All')}</div></div>
          <div class="card p-5"><div class="text-[10px] font-semibold text-slate-500 uppercase tracking-wider">Quick</div><div class="text-xs text-slate-600 mt-2 font-mono">Launch <b class="text-indigo-600">TN01</b> to transfer, <b class="text-indigo-600">MRQ1</b> for materials.</div></div>
        </div>
        <div><h3 class="text-xs font-bold text-slate-500 uppercase tracking-wider mb-4">Modules</h3><div id="quick_grid" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4"></div></div>
      </div>
    </div>`;
  renderQuickGrid();
}

async function loadModules(){
  ALL_TCODES=await api('/api/se93/list');
  $('#modules_list').innerHTML=ALL_TCODES.filter(t=>t.has_access).map(t=>`
    <button onclick="openTcode('${t.code}')" class="sidebar-link w-full text-left px-2.5 py-1.5 rounded-lg flex items-center justify-between">
      <span class="truncate"><b class="text-indigo-600 font-mono">${t.code}</b> <span class="text-slate-600">- ${esc(t.title)}</span></span>
    </button>`).join('');
}
function renderQuickGrid(){
  const c=$('#quick_grid');if(!c)return;
  c.innerHTML=ALL_TCODES.filter(t=>t.has_access).map(t=>`
    <div onclick="openTcode('${t.code}')" class="card hover:border-indigo-300 hover:shadow-md p-5 cursor-pointer transition-all">
      <div class="flex items-center justify-between mb-2">
        <span class="font-mono font-bold text-xs bg-indigo-50 text-indigo-600 border border-indigo-200 px-2 py-0.5 rounded-lg">${t.code}</span>
        <span class="badge badge-slate">${t.module||'ERP'}</span>
      </div>
      <h4 class="font-semibold text-slate-900 text-sm mt-2">${esc(t.title)}</h4>
      <p class="text-xs text-slate-500 mt-1 line-clamp-2">${esc(t.description||'Standard module')}</p>
    </div>`).join('');
}
async function openTcode(code){
  code=code.toUpperCase().trim();CURRENT_TCODE=code;$('#stat_tcode').textContent=code;setMsg(`Module ${code}`);
  if(code==='TN01')return renderTN01();
  if(code==='MRQ1')return renderMRQ();
  if(code==='CTST')return renderCTST();
  if(code==='SU01')return renderSU01();
  if(code==='SPRO')return renderSPRO();
  if(code==='SE93')return renderSE93();
  if(code==='PFCG')return renderPFCG();
  renderStandardModule(code);
}

// ===== TN01 =====
let TN01_ROUTES={},TN01_ACTIVE_ROUTE="PROD_TO_QC",TN01_HEADERS=[],TN01_ROWS=[],TN01_UNDO=[],TN01_REDO=[],TN01_CELL={r:0,c:1};
let TN01_SELECTED_CC=[];

async function renderTN01(){TN01_ROUTES=await api('/api/tn01/routes');renderTN01Main();}
function renderTN01Main(tab="sheet"){
  $('#main_screen').innerHTML=`
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
        ${tab==='sheet'?`<button onclick="openSendModal()" class="btn-primary text-xs font-bold px-4 py-2 rounded-lg flex items-center gap-1.5 shadow-sm"><i class="ri-send-plane-fill"></i> Send</button>`:''}
      </div>
      <div id="tn01_content" class="flex-1 overflow-hidden flex flex-col"></div>
    </div>`;
  if(tab==='sheet')renderTN01SheetTab();
  else if(tab==='receive')renderTN01ReceiveTab();
  else if(tab==='history')renderTN01HistoryTab();
}
function switchTN01Tab(t){renderTN01Main(t);}

function renderTN01SheetTab(){
  if(!TN01_HEADERS.length)applyRouteLayout(TN01_ACTIVE_ROUTE,false);
  $('#tn01_content').innerHTML=`
    <div class="p-2 bg-white border-b border-slate-200 flex items-center justify-between text-xs shrink-0">
      <div class="flex items-center gap-2">
        <label class="font-bold text-slate-600 text-xs">Route: <select id="route_selector" onchange="applyRouteLayout(this.value,true)" class="form-input w-auto py-1 font-semibold">${Object.keys(TN01_ROUTES).map(k=>`<option value="${k}" ${k===TN01_ACTIVE_ROUTE?'selected':''}>${TN01_ROUTES[k].route}</option>`).join('')}</select></label>
        <button onclick="addGridRow()" class="bg-white text-slate-600 px-2.5 py-1 rounded-md border border-slate-200 hover:bg-slate-50">+ Row</button>
        <button onclick="deleteGridRow()" class="bg-white text-red-500 px-2.5 py-1 rounded-md border border-slate-200 hover:bg-red-50">− Row</button>
      </div>
      <button onclick="exportGridAsExcel()" class="bg-white text-emerald-700 border border-emerald-200 px-3 py-1 rounded-md flex items-center gap-1 font-semibold hover:bg-emerald-50"><i class="ri-file-excel-2-line"></i> Export</button>
    </div>
    <div id="grid_table_container" class="flex-1 overflow-auto p-2 select-text" tabindex="0"></div>`;
  renderGridTable();
  setupPasteHandler();
}

function applyRouteLayout(k,r=true){
  TN01_ACTIVE_ROUTE=k;
  TN01_HEADERS=[...TN01_ROUTES[k].columns];
  if(r||!TN01_ROWS.length){
    TN01_ROWS=[];
    for(let i=1;i<=8;i++){const row=new Array(TN01_HEADERS.length).fill("");row[0]=i;TN01_ROWS.push(row);}
  }
  renderGridTable();
}

function renderGridTable(){
  const c=$('#grid_table_container');if(!c)return;
  let h=`<table class="excel-grid"><thead><tr>${TN01_HEADERS.map(x=>`<th>${esc(x)}</th>`).join('')}</tr></thead><tbody>`;
  TN01_ROWS.forEach((row,r)=>{
    h+=`<tr>${row.map((cell,ci)=>`<td contenteditable="${ci!==0}" class="${ci===0?'sr-col':''}" onfocus="TN01_CELL={r:${r},c:${ci}}" onblur="updateGridCell(${r},${ci},this.innerText)">${esc(cell)}</td>`).join('')}</tr>`;
  });
  h+=`</tbody></table>`;
  c.innerHTML=h;
}

function updateGridCell(r,ci,v){if(ci===0)return;if(TN01_ROWS[r][ci]!==v.trim()){TN01_UNDO.push(JSON.stringify(TN01_ROWS));TN01_ROWS[r][ci]=v.trim();}}
function addGridRow(){TN01_UNDO.push(JSON.stringify(TN01_ROWS));const r=new Array(TN01_HEADERS.length).fill("");r[0]=TN01_ROWS.length+1;TN01_ROWS.push(r);renderGridTable();}
function deleteGridRow(){if(TN01_ROWS.length<=1)return;TN01_ROWS.pop();renderGridTable();}

function setupPasteHandler(){
  const c=$('#grid_table_container');if(!c)return;
  c.addEventListener('paste',e=>{
    e.preventDefault();
    const clip=(e.clipboardData||window.clipboardData).getData('text');
    if(!clip)return;
    TN01_UNDO.push(JSON.stringify(TN01_ROWS));
    const lines=clip.replace(/\r\n/g,'\n').replace(/\r/g,'\n').split('\n');
    let sr=TN01_CELL.r||0;
    
    lines.forEach((line,li)=>{
      if(!line&&li===lines.length-1)return;
      let cells=line.split('\t');
      
      const tr=sr+li;
      while(TN01_ROWS.length<=tr){
        const nr=new Array(TN01_HEADERS.length).fill("");
        nr[0]=TN01_ROWS.length+1;
        TN01_ROWS.push(nr);
      }
      
      // Auto-assign: col 0 is ALWAYS Sr. No. Paste strictly fills starting at col 1 onwards.
      cells.forEach((v,ci)=>{
        const tc=ci+1;
        if(tc<TN01_HEADERS.length){
          TN01_ROWS[tr][tc]=v.trim();
        }
      });
    });
    TN01_ROWS.forEach((r,i)=>{r[0]=i+1;});
    renderGridTable();
  });
}

async function exportGridAsExcel(){
  const f=TN01_ROWS.filter(r=>r.slice(1).some(x=>String(x).trim()!==''));
  if(!f.length)return alert("No rows.");
  const t=TN01_ROUTES[TN01_ACTIVE_ROUTE]?TN01_ROUTES[TN01_ACTIVE_ROUTE].route:"Draft";
  try{
    const tok=localStorage.getItem("nexus_token");
    const r=await fetch('/api/tn01/export-grid',{method:'POST',headers:{'Content-Type':'application/json','Authorization':tok?`Bearer ${tok}`:''},body:JSON.stringify({headers:TN01_HEADERS,rows:f,title:`TN01 Draft — ${t}`})});
    const b=await r.blob();
    const u=URL.createObjectURL(b);
    const a=document.createElement('a');a.href=u;a.download='TN01_Draft.xlsx';document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(u);
  }catch(e){alert("Failed: "+e);}
}

async function openSendModal(){
  const a=TN01_ROUTES[TN01_ACTIVE_ROUTE];
  const f=TN01_ROWS.filter(r=>r.slice(1).some(x=>String(x).trim()!==''));
  if(!f.length)return alert("Enter rows.");
  TN01_SELECTED_CC = [];
  const sp=await api('/api/spro/overview');
  openDialog(`
    <div class="px-6 py-4 border-b bg-slate-50 flex justify-between items-center rounded-t-2xl">
      <div><h3 class="font-bold text-sm">Send Transfer Note</h3><p class="text-xs text-slate-500">${a.route}</p></div>
      <span id="live_datetime" class="text-xs font-mono text-slate-500"></span>
    </div>
    <div class="p-6 space-y-4">
      <div class="grid grid-cols-3 gap-3">
        <div><label class="form-label">Plant</label><select id="send_plant" onchange="updateSuggestedUsers()" class="form-input">${sp.plants.map(p=>`<option value="${esc(p.code)}">${esc(p.code)} - ${esc(p.name)}</option>`).join('')}</select></div>
        <div><label class="form-label">Dept</label><select id="send_dept" onchange="updateSuggestedUsers()" class="form-input">${sp.departments.map(d=>`<option value="${esc(d.name)}" ${d.name===a.to_dept?'selected':''}>${esc(d.name)}</option>`).join('')}</select></div>
        <div><label class="form-label">Receiver *</label><select id="send_user" class="form-input"></select></div>
      </div>

      <!-- OUTLOOK-STYLE CC AUTO-SUGGESTION -->
      <div class="relative">
        <label class="form-label">CC (Outlook style address book search)</label>
        <div id="cc_tags" class="flex flex-wrap gap-1.5 p-2 bg-slate-50 border border-slate-200 rounded-lg min-h-[40px] items-center mb-1">
          <input id="cc_search_input" placeholder="Type user name or ID..." oninput="searchCCUsers(this.value)" class="bg-transparent text-xs outline-none flex-1 min-w-[140px]">
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
  const el=$('#live_datetime');if(!el)return;
  const d=new Date();
  el.textContent=d.toLocaleString('en-GB',{dateStyle:'medium',timeStyle:'medium'});
  setTimeout(updateLiveClock,1000);
}

async function updateSuggestedUsers(){
  const pl=$('#send_plant').value,dept=$('#send_dept').value;
  const u=await api(`/api/tn01/suggest-users?plant=${pl}&dept=${encodeURIComponent(dept)}`);
  $('#send_user').innerHTML=u.length?u.map(x=>`<option value="${esc(x.username)}">${esc(x.full_name)} (${esc(x.username)})</option>`).join(''):'<option value="">(None)</option>';
}

async function searchCCUsers(val){
  const box=$('#cc_suggestions');
  if(!val.trim()){box.classList.add('hidden');box.innerHTML='';return;}
  const list=await api(`/api/tn01/suggest-users?q=${encodeURIComponent(val)}`);
  if(!list.length){
    box.innerHTML='<div class="p-2.5 text-xs text-slate-400">No matching user found.</div>';
  }else{
    box.innerHTML=list.map(u=>`
      <div onclick="addCCTag('${esc(u.username)}', '${esc(u.full_name)}')" class="p-2.5 hover:bg-indigo-50 cursor-pointer border-b border-slate-100 flex items-center justify-between text-xs">
        <div><b class="text-slate-900">${esc(u.full_name)}</b> <span class="text-slate-400 font-mono">(${esc(u.username)})</span></div>
        <span class="badge badge-slate">${esc(u.department||'')}</span>
      </div>`).join('');
  }
  box.classList.remove('hidden');
}

function addCCTag(username, fullname){
  if(!TN01_SELECTED_CC.includes(username)){
    TN01_SELECTED_CC.push(username);
    renderCCTags();
  }
  $('#cc_search_input').value='';
  $('#cc_suggestions').classList.add('hidden');
}

function removeCCTag(username){
  TN01_SELECTED_CC=TN01_SELECTED_CC.filter(u=>u!==username);
  renderCCTags();
}

function renderCCTags(){
  const box=$('#cc_tags');
  const input=$('#cc_search_input');
  box.querySelectorAll('.cc-pill').forEach(el=>el.remove());
  TN01_SELECTED_CC.forEach(u=>{
    const span=document.createElement('span');
    span.className="cc-pill badge badge-indigo text-xs py-1 flex items-center gap-1 font-normal";
    span.innerHTML=`<b>${esc(u)}</b> <button onclick="removeCCTag('${esc(u)}')">&times;</button>`;
    box.insertBefore(span,input);
  });
}

async function submitSendNote(){
  const f=TN01_ROWS.filter(r=>r.slice(1).some(x=>String(x).trim()!==''));
  const p={
    layout_key:TN01_ACTIVE_ROUTE,
    route:TN01_ROUTES[TN01_ACTIVE_ROUTE].route,
    to_plant:$('#send_plant').value,
    to_dept:$('#send_dept').value,
    to_user:$('#send_user').value,
    cc_users:TN01_SELECTED_CC,
    remarks:$('#send_remarks').value,
    headers:TN01_HEADERS,
    rows:f
  };
  try{
    const r=await api('/api/tn01/send',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(p)});
    closeDialog();
    alert(`Transfer Sent: ${r.note_no}`);
    switchTN01Tab('history');
  }catch(e){alert(e);}
}

async function renderTN01ReceiveTab(){
  const ns=await api('/api/tn01/inbox');
  $('#tn01_content').innerHTML=`
    <div class="p-6 flex-1 flex flex-col overflow-hidden">
      <div class="flex items-center justify-between pb-4 border-b"><h2 class="text-base font-bold">Incoming Transfer Notes</h2><button onclick="renderTN01ReceiveTab()" class="btn-secondary text-xs px-3 py-1.5 rounded-lg"><i class="ri-refresh-line"></i></button></div>
      <div class="flex-1 mt-4 card overflow-auto"><table class="custom-table"><thead><tr><th>Note</th><th>Route</th><th>From</th><th>Sender</th><th>Batches</th><th>CC</th><th>Time</th><th>Status</th><th>Actions</th></tr></thead><tbody>
        ${!ns.length?`<tr><td colspan="9" class="text-center py-12 text-slate-400">No incoming notes.</td></tr>`:ns.map(n=>`<tr><td class="font-mono font-bold text-indigo-600">${esc(n.note_no)}</td><td>${esc(n.route)}</td><td>${esc(n.from_plant)} - ${esc(n.from_dept)}</td><td>${esc(n.from_user_name)}</td><td class="font-mono">${n.entry_count}</td><td class="font-mono text-xs text-slate-500">${esc(n.cc_users||'—')}</td><td class="font-mono text-xs">${esc(n.created_at)}</td><td><span class="badge ${n.status==='received'?'badge-emerald':'badge-amber'}">${n.status.toUpperCase()}</span></td><td><button onclick="viewTransferNote(${n.id})" class="btn-primary text-xs px-2.5 py-1 rounded">Open</button>${n.status==='sent'?`<button onclick="ackNote(${n.id})" class="bg-emerald-600 text-white text-xs px-2.5 py-1 rounded ml-1">Ack</button>`:''}</td></tr>`).join('')}
      </tbody></table></div>
    </div>`;
}
async function ackNote(id){await api(`/api/tn01/receive/${id}`,{method:'POST'});renderTN01ReceiveTab();}

async function viewTransferNote(id){
  const n=await api(`/api/tn01/note/${id}`);
  openDialog(`
    <div class="px-6 py-4 border-b flex justify-between items-center bg-slate-50">
      <div><span class="font-mono font-bold text-indigo-700">${esc(n.note_no)}</span><h3 class="font-bold text-sm mt-1">${esc(n.route)}</h3></div>
      <button onclick="closeDialog()"><i class="ri-close-line text-lg"></i></button>
    </div>
    <div class="p-6 space-y-4">
      <div class="flex justify-between items-center">
        <span class="text-xs font-bold text-slate-700">${n.entry_count} rows</span>
        <a href="/api/tn01/export/${id}" class="btn-secondary text-xs px-3 py-1.5 rounded flex items-center gap-1 text-emerald-700 font-semibold border-emerald-200 inline-flex"><i class="ri-file-excel-2-line"></i> Export as Excel</a>
      </div>
      <div class="text-xs"><b>CC:</b> ${esc(n.cc_users||'—')}</div>
      <div class="max-h-72 overflow-auto border border-slate-200 rounded-xl">
        <table class="w-full text-left text-xs">
          <thead><tr class="bg-slate-50 font-bold">${n.sheet_headers.map(h=>`<th class="p-2 border-r border-slate-200">${esc(h)}</th>`).join('')}</tr></thead>
          <tbody>${n.sheet_data.map(r=>`<tr class="border-b">${r.map(c=>`<td class="p-2 border-r">${esc(c)}</td>`).join('')}</tr>`).join('')}</tbody>
        </table>
      </div>
    </div>
    <div class="px-6 py-4 bg-slate-50 border-t flex justify-between">
      <a href="/api/tn01/print/${n.id}" target="_blank" class="btn-secondary text-xs px-3 py-2 rounded flex items-center gap-1"><i class="ri-printer-line"></i> Print</a>
      <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded text-xs">Close</button>
    </div>`);
}

async function renderTN01HistoryTab(){
  const admin=ME.is_master||ME.role==='admin';
  const h=await api('/api/tn01/history');
  $('#tn01_content').innerHTML=`
    <div class="p-6 flex-1 flex flex-col overflow-hidden">
      <div class="flex items-center justify-between pb-3 border-b">
        <h2 class="text-base font-bold">${admin?"All Transfer Notes History":"My Transfer Notes History"}</h2>
        <button onclick="renderTN01HistoryTab()" class="btn-secondary text-xs px-3 py-1.5 rounded-lg"><i class="ri-refresh-line"></i></button>
      </div>
      <div class="flex-1 mt-2 card overflow-auto"><table class="custom-table"><thead><tr><th>Note</th><th>Route</th><th>From</th><th>To</th><th>Batches</th><th>Sender</th><th>CC</th><th>Time</th><th>Status</th><th>Actions</th></tr></thead><tbody>
        ${!h.length?`<tr><td colspan="10" class="text-center py-10 text-slate-400">No records.</td></tr>`:h.map(x=>`
          <tr>
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
              ${admin?`<button onclick="deleteTN01History(${x.id})" class="text-xs text-red-600 font-bold ml-2">Del</button>`:''}
            </td>
          </tr>`).join('')}
      </tbody></table></div>
    </div>`;
}

async function deleteTN01History(id){
  if(!confirm("Delete this transfer note history entry?"))return;
  try{await api(`/api/tn01/delete/${id}`,{method:'DELETE'});setMsg("Deleted");renderTN01HistoryTab();}
  catch(e){alert(e);}
}

// ===== SU01 =====
async function renderSU01(){
  if(!ME.is_master&&ME.role!=='admin')return setMsg("Denied",true);
  const users=await api('/api/su01/users');
  $('#main_screen').innerHTML=`
    <div class="flex-1 flex flex-col p-6 overflow-hidden select-text">
      <div class="flex items-center justify-between pb-4 border-b">
        <h2 class="text-lg font-bold">User Directory (SU01)</h2>
        <button onclick="showCreateUserModal()" class="btn-primary text-xs px-4 py-2 rounded-lg">+ Add User</button>
      </div>
      <div class="flex-1 mt-4 card overflow-auto"><table class="custom-table"><thead><tr><th>Username</th><th>Name</th><th>Role</th><th>Dept</th><th>Plants</th><th>Status</th><th>Actions</th></tr></thead><tbody>
        ${users.map(u=>`
          <tr>
            <td class="font-mono font-bold text-indigo-600">${esc(u.username)}</td>
            <td>${esc(u.full_name)}</td>
            <td><span class="badge badge-slate">${u.role.toUpperCase()}</span></td>
            <td>${esc(u.department)}</td>
            <td class="font-mono text-xs">${esc(u.plants||'ALL')}</td>
            <td><span class="badge ${u.status==='approved'?'badge-emerald':'badge-red'}">${u.status}</span></td>
            <td>
              <button onclick="showEditUserModal('${esc(u.username)}')" class="text-xs text-indigo-600 font-semibold">Edit</button>
              <button onclick="showResetPassModal('${esc(u.username)}')" class="text-xs text-amber-600 font-semibold ml-2">Reset</button>
              ${u.role!=='master'?`<button onclick="deleteUser('${esc(u.username)}')" class="text-xs text-red-600 font-semibold ml-2">Del</button>`:''}
            </td>
          </tr>`).join('')}
      </tbody></table></div>
    </div>`;
}

async function showCreateUserModal(){
  const d=await api('/api/spro/overview');
  openDialog(`
    <div class="px-6 py-4 border-b bg-slate-50 flex justify-between"><h3 class="font-bold text-sm">Create User</h3><button onclick="closeDialog()"><i class="ri-close-line"></i></button></div>
    <div class="p-6 space-y-4">
      <div class="grid grid-cols-2 gap-4"><div><label class="form-label">Username</label><input id="nu_u" class="form-input font-mono"></div><div><label class="form-label">Password</label><input id="nu_p" type="password" class="form-input"></div></div>
      <div><label class="form-label">Full Name</label><input id="nu_fn" class="form-input"></div>
      <div class="grid grid-cols-2 gap-4"><div><label class="form-label">Dept</label><select id="nu_d" class="form-input">${d.departments.map(x=>`<option value="${esc(x.name)}">${esc(x.name)}</option>`).join('')}</select></div><div><label class="form-label">Role</label><select id="nu_r" class="form-input"><option value="user">User</option><option value="admin">Admin</option></select></div></div>
      <div><label class="form-label">Plants</label><input id="nu_pl" class="form-input font-mono"></div>
    </div>
    <div class="px-6 py-4 bg-slate-50 border-t flex justify-end gap-2"><button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded text-xs">Cancel</button><button onclick="submitCreateUser()" class="btn-primary px-4 py-2 rounded text-xs">Save</button></div>`);
}

async function submitCreateUser(){
  const fd=new FormData();fd.append("username",$('#nu_u').value);fd.append("password",$('#nu_p').value);fd.append("full_name",$('#nu_fn').value);fd.append("department",$('#nu_d').value);fd.append("role",$('#nu_r').value);fd.append("plants",$('#nu_pl').value);
  try{await api('/api/su01/create',{method:'POST',body:fd});closeDialog();renderSU01();}catch(e){alert(e);}
}

async function showEditUserModal(u){
  const users=await api('/api/su01/users');
  const user=users.find(x=>x.username===u);
  const d=await api('/api/spro/overview');
  openDialog(`
    <div class="px-6 py-4 border-b bg-slate-50 flex justify-between"><h3 class="font-bold text-sm">Edit: ${esc(u)}</h3><button onclick="closeDialog()"><i class="ri-close-line"></i></button></div>
    <div class="p-6 space-y-4">
      <div><label class="form-label">Name</label><input id="eu_fn" value="${esc(user.full_name)}" class="form-input"></div>
      <div class="grid grid-cols-2 gap-4"><div><label class="form-label">Dept</label><select id="eu_d" class="form-input">${d.departments.map(x=>`<option value="${esc(x.name)}" ${x.name===user.department?'selected':''}>${esc(x.name)}</option>`).join('')}</select></div><div><label class="form-label">Role</label><select id="eu_r" class="form-input"><option value="user" ${user.role==='user'?'selected':''}>User</option><option value="admin" ${user.role==='admin'?'selected':''}>Admin</option></select></div></div>
      <div class="grid grid-cols-2 gap-4"><div><label class="form-label">Plants</label><input id="eu_pl" value="${esc(user.plants||'')}" class="form-input font-mono"></div><div><label class="form-label">Status</label><select id="eu_st" class="form-input"><option value="approved" ${user.status==='approved'?'selected':''}>Approved</option><option value="locked" ${user.status==='locked'?'selected':''}>Locked</option></select></div></div>
    </div>
    <div class="px-6 py-4 bg-slate-50 border-t flex justify-end gap-2"><button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded text-xs">Cancel</button><button onclick="submitEditUser('${esc(u)}')" class="btn-primary px-4 py-2 rounded text-xs">Save</button></div>`);
}

async function submitEditUser(u){
  const fd=new FormData();
  fd.append("username",u);fd.append("full_name",$('#eu_fn').value);fd.append("department",$('#eu_d').value);fd.append("role",$('#eu_r').value);fd.append("plants",$('#eu_pl').value);fd.append("status",$('#eu_st').value);
  try{await api('/api/su01/update',{method:'POST',body:fd});closeDialog();renderSU01();}catch(e){alert(e);}
}

function showResetPassModal(u){
  const p=prompt(`New password for ${u}:`);if(!p)return;
  const fd=new FormData();fd.append("username",u);fd.append("new_password",p);
  api('/api/su01/reset-password',{method:'POST',body:fd}).then(()=>setMsg("Password reset.")).catch(e=>alert(e));
}

async function deleteUser(u){
  if(!confirm(`Delete ${u}?`))return;
  const fd=new FormData();fd.append("username",u);
  try{await api('/api/su01/delete',{method:'POST',body:fd});renderSU01();}catch(e){alert(e);}
}

async function renderSPRO(){
  if(!ME.is_master&&ME.role!=='admin')return setMsg("Denied",true);
  const d=await api('/api/spro/overview');
  $('#main_screen').innerHTML=`
    <div class="flex-1 flex flex-col p-6 overflow-hidden">
      <div class="flex items-center justify-between pb-4 border-b"><h2 class="text-lg font-bold">Enterprise Structure (SPRO)</h2></div>
      <div class="grid grid-cols-2 gap-4 flex-1 mt-4 overflow-auto">
        <div class="card overflow-hidden">
          <div class="p-3 border-b bg-slate-50 flex justify-between"><span class="text-xs font-bold uppercase">Plants (${d.plants.length})</span><button onclick="createPlant()" class="btn-primary text-xs px-3 py-1.5 rounded">+ Add</button></div>
          <table class="custom-table"><thead><tr><th>Code</th><th>Name</th><th>Status</th></tr></thead><tbody>
            ${d.plants.map(p=>`<tr><td class="font-mono font-bold text-indigo-600">${esc(p.code)}</td><td>${esc(p.name)}</td><td><span class="badge ${p.active?'badge-emerald':'badge-red'}">${p.active?'ACTIVE':'OFF'}</span></td></tr>`).join('')}
          </tbody></table>
        </div>
        <div class="card overflow-hidden">
          <div class="p-3 border-b bg-slate-50 flex justify-between"><span class="text-xs font-bold uppercase">Departments (${d.departments.length})</span><button onclick="createDept()" class="btn-primary text-xs px-3 py-1.5 rounded">+ Add</button></div>
          <table class="custom-table"><thead><tr><th>Name</th><th>Status</th></tr></thead><tbody>
            ${d.departments.map(x=>`<tr><td>${esc(x.name)}</td><td><span class="badge ${x.active?'badge-emerald':'badge-red'}">${x.active?'ACTIVE':'OFF'}</span></td></tr>`).join('')}
          </tbody></table>
        </div>
      </div>
    </div>`;
}

function createPlant(){
  const c=prompt("Plant Code (e.g. 1500):");if(!c)return;
  const n=prompt("Plant Name:");if(!n)return;
  const fd=new FormData();fd.append("code",c);fd.append("name",n);
  api('/api/admin/plants/create',{method:'POST',body:fd}).then(()=>renderSPRO()).catch(e=>alert(e));
}

function createDept(){
  const d=prompt("Department name:");if(!d)return;
  const fd=new FormData();fd.append("name",d.trim());
  api('/api/admin/departments/create',{method:'POST',body:fd}).then(()=>renderSPRO()).catch(e=>alert(e));
}

async function renderPFCG(){
  if(!ME.is_master&&ME.role!=='admin')return setMsg("Denied",true);
  const m=await api('/api/pfcg/matrix');
  $('#main_screen').innerHTML=`
    <div class="flex-1 flex flex-col p-6 overflow-hidden">
      <div class="flex justify-between pb-4 border-b">
        <h2 class="text-lg font-bold">Access Control (PFCG)</h2>
        <div class="flex gap-2">
          <select id="pfcg_u" class="form-input w-56">${m.users.map(u=>`<option value="${esc(u.username)}">${esc(u.full_name||u.username)}</option>`).join('')}</select>
          <select id="pfcg_t" class="form-input w-56">${m.tcodes.map(t=>`<option value="${esc(t.code)}">${esc(t.code)} - ${esc(t.title)}</option>`).join('')}</select>
          <button onclick="grantAccess()" class="btn-primary text-xs px-3 py-2 rounded">+ Grant</button>
        </div>
      </div>
      <div class="flex-1 mt-4 card overflow-auto"><table class="custom-table"><thead><tr><th>User</th><th>Module</th><th>By</th><th>At</th><th>Action</th></tr></thead><tbody>
        ${m.grants.map(g=>`<tr><td class="font-mono">${esc(g.username)}</td><td class="font-mono">${esc(g.tcode_code)}</td><td>${esc(g.granted_by)}</td><td class="text-xs">${esc(g.granted_at)}</td><td><button onclick="revokeAccess('${esc(g.username)}','${esc(g.tcode_code)}')" class="text-xs text-red-600 font-semibold">Revoke</button></td></tr>`).join('')}
      </tbody></table></div>
    </div>`;
}

async function grantAccess(){
  const fd=new FormData();fd.append("username",$('#pfcg_u').value);fd.append("tcode",$('#pfcg_t').value);
  await api('/api/pfcg/grant',{method:'POST',body:fd});renderPFCG();
}
async function revokeAccess(u,t){
  if(!confirm("Revoke?"))return;
  const fd=new FormData();fd.append("username",u);fd.append("tcode",t);
  await api('/api/pfcg/revoke',{method:'POST',body:fd});renderPFCG();
}

async function renderSE93(){
  if(!ME.is_master&&ME.role!=='admin')return setMsg("Denied",true);
  const list=await api('/api/se93/list');
  $('#main_screen').innerHTML=`
    <div class="flex-1 flex flex-col p-6 overflow-hidden">
      <div class="flex justify-between pb-4 border-b"><h2 class="text-lg font-bold">Module Studio (SE93)</h2></div>
      <div class="grid grid-cols-3 gap-4 flex-1 mt-4 overflow-hidden">
        <div class="col-span-1 card p-4 overflow-auto">
          <h3 class="text-sm font-bold mb-3">Auto-Create Module</h3>
          <form onsubmit="submitSE93(event)" class="space-y-3">
            <div><label class="form-label">Code</label><input name="code" required class="form-input font-mono uppercase"></div>
            <div><label class="form-label">Title</label><input name="title" required class="form-input"></div>
            <div><label class="form-label">Description</label><input name="description" class="form-input"></div>
            <div><label class="form-label">Module</label><input name="module" value="PROD" class="form-input font-mono uppercase"></div>
            <div><label class="form-label">Allowed Depts</label><input name="allowed_depts" value="Production,Quality Control" class="form-input"></div>
            <div><label class="form-label">Import File</label><input name="file" type="file" required class="w-full text-xs file:bg-indigo-600 file:text-white file:border-0 file:rounded file:px-2 file:py-1"></div>
            <button type="submit" class="btn-primary w-full text-xs font-bold py-2 rounded">Create Module</button>
          </form>
        </div>
        <div class="col-span-2 card overflow-auto"><table class="custom-table"><thead><tr><th>Code</th><th>Title</th><th>Records</th><th>Actions</th></tr></thead><tbody>
          ${list.map(t=>`<tr><td class="font-mono font-bold text-indigo-600">${esc(t.code)}</td><td>${esc(t.title)}</td><td class="font-mono">${t.record_count}</td><td><button onclick="openTcode('${esc(t.code)}')" class="text-xs text-indigo-600 font-semibold">Open</button></td></tr>`).join('')}
        </tbody></table></div>
      </div>
    </div>`;
}

async function submitSE93(ev){
  ev.preventDefault();
  const fd=new FormData(ev.target);
  try{
    const r=await fetch('/api/se93/auto-create',{method:'POST',body:fd,headers:{'Authorization':'Bearer '+(localStorage.getItem("nexus_token")||'')}});
    const d=await r.json();if(!r.ok)throw d.detail||"Fail";
    alert(`Created ${d.code}`);renderSE93();loadModules();
  }catch(e){alert(e);}
}

// Dialog Helpers
function openDialog(c){$('#app_dialog_body').innerHTML=c;$('#app_dialog').classList.remove('hidden');}
function closeDialog(){$('#app_dialog').classList.add('hidden');$('#app_dialog_body').innerHTML='';}

// Close suggestion box when clicking outside
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

# Support both GET and HEAD for Render health check
@app.api_route("/", methods=["GET", "HEAD"], response_class=HTMLResponse)
def home():
    return HTMLResponse(HTML_UI)

@app.api_route("/health", methods=["GET", "HEAD"])
def health():
    return {"status": "ACTIVE", "system": "NEXUS-CORE", "version": "4.0.1"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
