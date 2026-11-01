import io, json, sqlite3, secrets, datetime, os
from pathlib import Path
import bcrypt
from fastapi import FastAPI, Request, Response, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook, Workbook

# --- DATABASE ENGINE SETUP ---
DATABASE_URL = os.environ.get("DATABASE_URL")
USE_PG = bool(DATABASE_URL and DATABASE_URL.startswith("postgres"))
if USE_PG:
    import psycopg2
    import psycopg2.extras
    if DATABASE_URL.startswith("postgres://"):
        DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

DB_PATH = Path(__file__).parent / "sap_protransfer.db"

def get_conn():
    if USE_PG:
        return psycopg2.connect(DATABASE_URL, sslmode='require')
    else:
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
                return cur.fetchall()
        else:
            return [dict(r) for r in conn.execute(sql, params).fetchall()]
    finally:
        conn.close()

def fetch_one(sql: str, params=()):
    conn = get_conn()
    try:
        if USE_PG:
            with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                cur.execute(q(sql), params)
                return cur.fetchone()
        else:
            r = conn.execute(sql, params).fetchone()
            return dict(r) if r else None
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

def now():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# --- INITIALIZATION & AUTO-MIGRATION ---
def init_db():
    conn = get_conn()
    cur = conn.cursor()

    # 1. AUTO-MIGRATION: Cleanly upgrade outdated PostgreSQL tables from older app versions
    if USE_PG:
        try:
            cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name = 'plants' AND column_name = 'code';")
            if not cur.fetchone():
                cur.execute("DROP TABLE IF EXISTS plants CASCADE;")

            cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name = 'org' AND column_name = 'system_id';")
            if not cur.fetchone():
                cur.execute("DROP TABLE IF EXISTS org CASCADE;")

            cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name = 'tcode_records' AND column_name = 'plant';")
            if not cur.fetchone():
                cur.execute("ALTER TABLE tcode_records ADD COLUMN IF NOT EXISTS plant TEXT DEFAULT '';")

            cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name = 'users' AND column_name = 'plants';")
            if not cur.fetchone():
                cur.execute("ALTER TABLE users ADD COLUMN IF NOT EXISTS plants TEXT DEFAULT '';")

            cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name = 'tcodes' AND column_name = 'module';")
            if not cur.fetchone():
                cur.execute("ALTER TABLE tcodes ADD COLUMN IF NOT EXISTS module TEXT DEFAULT 'PP';")

            conn.commit()
        except Exception as e:
            print("Auto-migration notice:", e)
            conn.rollback()

    # 2. CREATE STANDARD SAP SCHEMAS
    if USE_PG:
        cur.execute('''
        CREATE TABLE IF NOT EXISTS users (username TEXT PRIMARY KEY, password_hash TEXT, full_name TEXT, role TEXT, department TEXT, plants TEXT DEFAULT '', status TEXT DEFAULT 'approved', created_at TEXT, updated_at TEXT);
        CREATE TABLE IF NOT EXISTS sessions (token TEXT PRIMARY KEY, username TEXT, expires TEXT);
        CREATE TABLE IF NOT EXISTS departments (name TEXT PRIMARY KEY, active INT DEFAULT 1, created_at TEXT);
        CREATE TABLE IF NOT EXISTS plants (code TEXT PRIMARY KEY, name TEXT, active INT DEFAULT 1, created_at TEXT);
        CREATE TABLE IF NOT EXISTS org (id INT PRIMARY KEY, name TEXT, system_id TEXT, client TEXT);
        CREATE TABLE IF NOT EXISTS tcodes (code TEXT PRIMARY KEY, title TEXT, description TEXT, module TEXT, allowed_depts TEXT, active INT DEFAULT 1, created_at TEXT, updated_at TEXT, created_by TEXT);
        CREATE TABLE IF NOT EXISTS tcode_schemas (code TEXT PRIMARY KEY, fields TEXT, layout TEXT, updated_at TEXT, updated_by TEXT);
        CREATE TABLE IF NOT EXISTS tcode_records (id SERIAL PRIMARY KEY, tcode_code TEXT, plant TEXT DEFAULT '', data TEXT, status TEXT DEFAULT 'active', created_by TEXT, created_by_name TEXT, created_at TEXT, updated_at TEXT);
        CREATE TABLE IF NOT EXISTS tcode_layouts (id SERIAL PRIMARY KEY, tcode_code TEXT, layout_code TEXT UNIQUE, name TEXT, title TEXT, is_default INT DEFAULT 0, fields TEXT, layout_config TEXT, description TEXT, created_by TEXT, created_by_name TEXT, created_at TEXT, updated_at TEXT);
        CREATE TABLE IF NOT EXISTS access_grants (id SERIAL PRIMARY KEY, username TEXT, tcode_code TEXT, granted_by TEXT, granted_at TEXT, UNIQUE(username, tcode_code));
        ''')
    else:
        cur.executescript('''
        CREATE TABLE IF NOT EXISTS users (username TEXT PRIMARY KEY, password_hash TEXT, full_name TEXT, role TEXT, department TEXT, plants TEXT DEFAULT '', status TEXT DEFAULT 'approved', created_at TEXT, updated_at TEXT);
        CREATE TABLE IF NOT EXISTS sessions (token TEXT PRIMARY KEY, username TEXT, expires TEXT);
        CREATE TABLE IF NOT EXISTS departments (name TEXT PRIMARY KEY, active INT DEFAULT 1, created_at TEXT);
        CREATE TABLE IF NOT EXISTS plants (code TEXT PRIMARY KEY, name TEXT, active INT DEFAULT 1, created_at TEXT);
        CREATE TABLE IF NOT EXISTS org (id INT PRIMARY KEY, name TEXT, system_id TEXT, client TEXT);
        CREATE TABLE IF NOT EXISTS tcodes (code TEXT PRIMARY KEY, title TEXT, description TEXT, module TEXT, allowed_depts TEXT, active INT DEFAULT 1, created_at TEXT, updated_at TEXT, created_by TEXT);
        CREATE TABLE IF NOT EXISTS tcode_schemas (code TEXT PRIMARY KEY, fields TEXT, layout TEXT, updated_at TEXT, updated_by TEXT);
        CREATE TABLE IF NOT EXISTS tcode_records (id INTEGER PRIMARY KEY AUTOINCREMENT, tcode_code TEXT, plant TEXT DEFAULT '', data TEXT, status TEXT DEFAULT 'active', created_by TEXT, created_by_name TEXT, created_at TEXT, updated_at TEXT);
        CREATE TABLE IF NOT EXISTS tcode_layouts (id INTEGER PRIMARY KEY AUTOINCREMENT, tcode_code TEXT, layout_code TEXT UNIQUE, name TEXT, title TEXT, is_default INT DEFAULT 0, fields TEXT, layout_config TEXT, description TEXT, created_by TEXT, created_by_name TEXT, created_at TEXT, updated_at TEXT);
        CREATE TABLE IF NOT EXISTS access_grants (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT, tcode_code TEXT, granted_by TEXT, granted_at TEXT, UNIQUE(username, tcode_code));
        ''')
    conn.commit()

    cur_time = now()

    # Org Record
    cur.execute(q("SELECT 1 FROM org WHERE id=1"))
    if not cur.fetchone():
        cur.execute(q("INSERT INTO org VALUES (?,?,?,?)"), (1, "Thermocables Enterprise Group", "PRD", "100"))

    # Default Departments
    for d in ["Production", "Quality Control", "Final QC", "Commercial", "Planning", "Store", "Plant Maintenance"]:
        cur.execute(q("SELECT 1 FROM departments WHERE name=?"), (d,))
        if not cur.fetchone():
            cur.execute(q("INSERT INTO departments VALUES (?,?,?)"), (d, 1, cur_time))

    # Default Plants
    for p_code, p_name in [("1000", "Gooty Central Plant"), ("1400", "Plant 1400 Extrusion"), ("2000", "Hyderabad Cable Works"), ("CORP", "TCL Corporate HQ")]:
        cur.execute(q("SELECT 1 FROM plants WHERE code=?"), (p_code,))
        if not cur.fetchone():
            cur.execute(q("INSERT INTO plants VALUES (?,?,?,?)"), (p_code, p_name, 1, cur_time))

    # Super Admin Account (Password: Admin@123)
    cur.execute(q("SELECT 1 FROM users WHERE username='admin'"))
    if not cur.fetchone():
        ph = bcrypt.hashpw("Admin@123".encode(), bcrypt.gensalt()).decode()
        cur.execute(q("INSERT INTO users VALUES (?,?,?,?,?,?,?,?,?)"),
                    ("admin", ph, "SAP Super Administrator", "master", "BASIS / IT", "1000,1400,2000,CORP", "approved", cur_time, cur_time))

    # Standard Business T-Codes
    default_tcodes = [
        ("PROD", "Production Handover Report", "Outer Sheathing Production tracking", "PP", "Production,Quality Control"),
        ("GCH0", "Green Card Inspection Check", "Incoming QA & Cable Release Cards", "QM", "Quality Control,Production"),
        ("NCR0", "Non-Conformance Records", "Internal & Customer QC Deviation Logging", "QM", "Quality Control,Final QC,Production"),
        ("PLN0", "Master Planning & Scheduling", "Plant Order scheduling and lot runs", "PP", "Planning,Production,Store"),
        ("QCH0", "Quality Control to Final QC", "QC Stage clearance handover slips", "QM", "Quality Control,Final QC"),
        ("STD0", "Store Material Transfer", "Inter-plant and store stock issue", "MM", "Store,Planning,Production")
    ]

    for code, title, desc, mod, depts in default_tcodes:
        cur.execute(q("SELECT 1 FROM tcodes WHERE code=?"), (code,))
        if not cur.fetchone():
            cur.execute(q("INSERT INTO tcodes VALUES (?,?,?,?,?,?,?,?,?)"),
                        (code, title, desc, mod, depts, 1, cur_time, cur_time, "SAP_SYSTEM"))
            fields = json.dumps([
                {"key": "job_no", "label": "Job / Production Order", "type": "text", "required": True, "width": 1},
                {"key": "batch_no", "label": "Batch / Lot No", "type": "text", "required": True, "width": 1},
                {"key": "cable_size", "label": "Cable Spec / Size", "type": "text", "required": True, "width": 1},
                {"key": "quantity", "label": "Quantity (Mtrs)", "type": "number", "required": True, "width": 1},
                {"key": "target_date", "label": "Completion Date", "type": "date", "required": False, "width": 1},
                {"key": "operator", "label": "Plant Operator", "type": "text", "required": False, "width": 1},
                {"key": "remarks", "label": "Engineering Remarks", "type": "textarea", "required": False, "width": 3}
            ])
            cur.execute(q("INSERT INTO tcode_schemas VALUES (?,?,?,?,?)"), (code, fields, json.dumps({"columns": 3}), cur_time, "SAP_SYSTEM"))
            cur.execute(q("INSERT INTO tcode_layouts (tcode_code, layout_code, name, title, is_default, fields, layout_config, description, created_by, created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)"),
                        (code, f"/1SAP/{code}_DEFAULT", "Standard SAP ALV", f"Standard Grid: {title}", 1, fields, json.dumps({"columns": 3}), "Default SAP Table View", "admin", "SAP Super Administrator", cur_time, cur_time))

    conn.commit()
    conn.close()

init_db()

# --- AUTH & ACCESS HELPERS ---
def get_user(req: Request):
    t = req.cookies.get("sap_session")
    if not t:
        return None
    s = fetch_one("SELECT * FROM sessions WHERE token=?", (t,))
    if not s:
        return None
    u = fetch_one("SELECT * FROM users WHERE username=?", (s["username"],))
    if not u or u["status"] != "approved":
        return None
    return u

def need(req: Request):
    u = get_user(req)
    if not u:
        raise HTTPException(401, "Authorization required. Please log into SAP System.")
    return u

def is_master(u):
    return u["role"] == "master" or u["username"] == "admin"

def need_admin(req: Request):
    u = need(req)
    if u["role"] not in ["admin", "master"]:
        raise HTTPException(403, "Access Denied: Authorization object S_TCODE / S_TABU_DIS required.")
    return u

def need_master(req: Request):
    u = need(req)
    if not is_master(u):
        raise HTTPException(403, "Access Denied: Superuser Authorization (SAP_ALL / MASTER) required.")
    return u

def check_tcode_access(u, code: str):
    code = code.upper().strip()
    if is_master(u) or u["role"] == "admin":
        return True
    tc = fetch_one("SELECT * FROM tcodes WHERE code=? AND active=1", (code,))
    if not tc:
        return False
    allowed = [x.strip() for x in (tc["allowed_depts"] or "").split(",") if x.strip()]
    if u["department"] in allowed:
        return True
    grant = fetch_one("SELECT 1 FROM access_grants WHERE username=? AND tcode_code=?", (u["username"], code))
    return bool(grant)

# --- FASTAPI APP ---
app = FastAPI(title="SAP Enterprise GUI - ProTransfer ERP", version="7.70-REV2")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True, allow_methods=["*"], allow_headers=["*"])

# --- AUTH APIS ---
@app.post("/api/login")
def api_login(res: Response, username: str = Form(...), password: str = Form(...)):
    username = username.lower().strip()
    u = fetch_one("SELECT * FROM users WHERE username=?", (username,))
    if not u or not bcrypt.checkpw(password.encode(), u["password_hash"].encode()):
        raise HTTPException(401, "RFC Error: User or Password Incorrect")
    if u["status"] != "approved":
        raise HTTPException(403, "RFC Error: User Account Locked or Not Approved by BASIS")

    token = secrets.token_hex(32)
    exp = (datetime.datetime.now() + datetime.timedelta(days=7)).strftime("%Y-%m-%d %H:%M:%S")
    exec_write("INSERT INTO sessions VALUES (?,?,?)", (token, username, exp))
    res.set_cookie("sap_session", token, httponly=True, samesite="lax", max_age=86400 * 7)
    return {"ok": True, "user": {"username": u["username"], "full_name": u["full_name"], "role": u["role"]}}

@app.post("/api/logout")
def api_logout(res: Response, req: Request):
    t = req.cookies.get("sap_session")
    if t:
        exec_write("DELETE FROM sessions WHERE token=?", (t,))
    res.delete_cookie("sap_session")
    return {"ok": True}

@app.get("/api/me")
def api_me(req: Request):
    u = need(req)
    org_info = fetch_one("SELECT * FROM org WHERE id=1") or {}
    return {
        "username": u["username"],
        "full_name": u["full_name"],
        "role": u["role"],
        "department": u["department"],
        "plants": [p for p in (u["plants"] or "").split(",") if p],
        "is_master": is_master(u),
        "system": org_info.get("system_id", "PRD"),
        "client": org_info.get("client", "100"),
        "company": org_info.get("name", "Thermocables")
    }

# --- SAP SPRO: SYSTEM ORG & ENTERPRISE STRUCTURE (PLANTS / DEPTS) ---
@app.get("/api/spro/overview")
def spro_overview(req: Request):
    need(req)
    return {
        "org": fetch_one("SELECT * FROM org WHERE id=1"),
        "departments": fetch_all("SELECT * FROM departments ORDER BY name"),
        "plants": fetch_all("SELECT * FROM plants ORDER BY code")
    }

@app.post("/api/admin/plants/create")
def spro_create_plant(req: Request, code: str = Form(...), name: str = Form(...)):
    need_master(req)
    code = code.upper().strip()
    if fetch_one("SELECT 1 FROM plants WHERE code=?", (code,)):
        raise HTTPException(400, f"Plant code {code} already exists.")
    exec_write("INSERT INTO plants VALUES (?,?,?,?)", (code, name.strip(), 1, now()))
    return {"ok": True}

@app.post("/api/admin/plants/update")
def spro_update_plant(req: Request, code: str = Form(...), name: str = Form(...), active: int = Form(...)):
    need_master(req)
    exec_write("UPDATE plants SET name=?, active=? WHERE code=?", (name.strip(), active, code.strip()))
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
    if fetch_one("SELECT 1 FROM departments WHERE name=?", (name,)):
        raise HTTPException(400, "Department already exists.")
    exec_write("INSERT INTO departments VALUES (?,?,?)", (name, 1, now()))
    return {"ok": True}

@app.post("/api/admin/departments/delete")
def spro_delete_dept(req: Request, name: str = Form(...)):
    need_master(req)
    exec_write("DELETE FROM departments WHERE name=?", (name.strip(),))
    return {"ok": True}

# --- SAP SU01: USER MAINTENANCE ---
@app.get("/api/su01/users")
def su01_list_users(req: Request):
    need_admin(req)
    return fetch_all("SELECT username, full_name, role, department, plants, status, created_at, updated_at FROM users ORDER BY username")

@app.post("/api/su01/create")
def su01_create_user(req: Request, username: str = Form(...), password: str = Form(...), full_name: str = Form(...),
                     department: str = Form(...), plants: str = Form(""), role: str = Form("user")):
    caller = need_admin(req)
    username = username.lower().strip()
    if role == "master" and not is_master(caller):
        raise HTTPException(403, "Authorization violation: Only MASTER superuser can grant SAP_ALL/Master role.")
    if fetch_one("SELECT 1 FROM users WHERE username=?", (username,)):
        raise HTTPException(400, f"User {username} already exists in client.")

    ph = bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()
    cur_time = now()
    exec_write("INSERT INTO users VALUES (?,?,?,?,?,?,?,?,?)",
               (username, ph, full_name.strip(), role, department.strip(), plants.strip(), "approved", cur_time, cur_time))
    return {"ok": True}

@app.post("/api/su01/update")
def su01_update_user(req: Request, username: str = Form(...), full_name: str = Form(...),
                     department: str = Form(...), plants: str = Form(""), role: str = Form(...), status: str = Form(...)):
    caller = need_admin(req)
    username = username.lower().strip()
    target = fetch_one("SELECT * FROM users WHERE username=?", (username,))
    if not target:
        raise HTTPException(404, "User not found")
    if is_master(target) and not is_master(caller):
        raise HTTPException(403, "Security Error: Cannot modify MASTER user")
    if role == "master" and not is_master(caller):
        raise HTTPException(403, "Cannot elevate role to MASTER")

    exec_write("UPDATE users SET full_name=?, department=?, plants=?, role=?, status=?, updated_at=? WHERE username=?",
               (full_name.strip(), department.strip(), plants.strip(), role, status, now(), username))
    return {"ok": True}

@app.post("/api/su01/delete")
def su01_delete_user(req: Request, username: str = Form(...)):
    caller = need_admin(req)
    username = username.lower().strip()
    target = fetch_one("SELECT * FROM users WHERE username=?", (username,))
    if not target:
        raise HTTPException(404, "User not found")
    if is_master(target):
        raise HTTPException(400, "Security Violation: Master Administrator cannot be deleted!")
    if target["role"] == "admin" and not is_master(caller):
        raise HTTPException(403, "Security Error: Only Master Admin can delete other Administrators")
    if username == caller["username"]:
        raise HTTPException(400, "Cannot delete your own active logon account.")

    exec_write("DELETE FROM users WHERE username=?", (username,))
    exec_write("DELETE FROM sessions WHERE username=?", (username,))
    exec_write("DELETE FROM access_grants WHERE username=?", (username,))
    return {"ok": True}

@app.post("/api/su01/reset-password")
def su01_reset_password(req: Request, username: str = Form(...), new_password: str = Form(...)):
    caller = need_admin(req)
    username = username.lower().strip()
    target = fetch_one("SELECT * FROM users WHERE username=?", (username,))
    if not target:
        raise HTTPException(404, "User not found")
    if is_master(target) and not is_master(caller):
        raise HTTPException(403, "Cannot change password of Master Admin")

    ph = bcrypt.hashpw(new_password.encode(), bcrypt.gensalt()).decode()
    exec_write("UPDATE users SET password_hash=?, updated_at=? WHERE username=?", (ph, now(), username))
    exec_write("DELETE FROM sessions WHERE username=?", (username,))
    return {"ok": True}

# --- SAP PFCG: AUTHORIZATIONS & TCODE GRANTS ---
@app.get("/api/pfcg/matrix")
def pfcg_matrix(req: Request):
    need_admin(req)
    users = fetch_all("SELECT username, full_name, role, department FROM users ORDER BY username")
    tcodes = fetch_all("SELECT code, title, module, allowed_depts FROM tcodes WHERE active=1 ORDER BY code")
    grants = fetch_all("SELECT * FROM access_grants")
    return {"users": users, "tcodes": tcodes, "grants": grants}

@app.post("/api/pfcg/grant")
def pfcg_grant(req: Request, username: str = Form(...), tcode: str = Form(...)):
    caller = need_admin(req)
    username = username.lower().strip()
    tcode = tcode.upper().strip()
    if USE_PG:
        exec_write("INSERT INTO access_grants (username, tcode_code, granted_by, granted_at) VALUES (?,?,?,?) ON CONFLICT (username, tcode_code) DO NOTHING",
                   (username, tcode, caller["username"], now()))
    else:
        exec_write("INSERT OR IGNORE INTO access_grants (username, tcode_code, granted_by, granted_at) VALUES (?,?,?,?)",
                   (username, tcode, caller["username"], now()))
    return {"ok": True}

@app.post("/api/pfcg/revoke")
def pfcg_revoke(req: Request, username: str = Form(...), tcode: str = Form(...)):
    need_admin(req)
    exec_write("DELETE FROM access_grants WHERE username=? AND tcode_code=?", (username.lower().strip(), tcode.upper().strip()))
    return {"ok": True}

# --- SAP SE93: TRANSACTION CODES & REPOSITORY ---
@app.get("/api/se93/list")
def se93_list(req: Request):
    u = need(req)
    all_t = fetch_all("SELECT * FROM tcodes ORDER BY code")
    for t in all_t:
        t["layout_count"] = fetch_one("SELECT COUNT(*) as c FROM tcode_layouts WHERE tcode_code=?", (t["code"],))["c"]
        t["record_count"] = fetch_one("SELECT COUNT(*) as c FROM tcode_records WHERE tcode_code=? AND status='active'", (t["code"],))["c"]
        t["has_access"] = check_tcode_access(u, t["code"])
    return all_t

@app.post("/api/se93/create")
def se93_create(req: Request, code: str = Form(...), title: str = Form(...), description: str = Form(...),
                module: str = Form("PP"), allowed_depts: str = Form(...)):
    u = need_admin(req)
    code = code.upper().strip()
    if fetch_one("SELECT 1 FROM tcodes WHERE code=?", (code,)):
        raise HTTPException(400, f"Transaction code {code} already exists in SAP Repository.")

    cur_time = now()
    exec_write("INSERT INTO tcodes VALUES (?,?,?,?,?,?,?,?,?)",
               (code, title.strip(), description.strip(), module.strip(), allowed_depts.strip(), 1, cur_time, cur_time, u["username"]))

    starter_fields = json.dumps([
        {"key": "job_no", "label": "Job / Document No", "type": "text", "required": True, "width": 1},
        {"key": "batch_no", "label": "Batch / Heat Lot", "type": "text", "required": True, "width": 1},
        {"key": "description", "label": "Item Description", "type": "text", "required": False, "width": 2},
        {"key": "quantity", "label": "Target Quantity", "type": "number", "required": False, "width": 1},
        {"key": "remarks", "label": "Operational Notes", "type": "textarea", "required": False, "width": 3}
    ])
    exec_write("INSERT INTO tcode_schemas VALUES (?,?,?,?,?)", (code, starter_fields, json.dumps({"columns": 3}), cur_time, u["username"]))
    exec_write("INSERT INTO tcode_layouts (tcode_code, layout_code, name, title, is_default, fields, layout_config, description, created_by, created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
               (code, f"/1SAP/{code}_DEF", "Default View", f"{title} Grid", 1, starter_fields, json.dumps({"columns": 3}), "Generated Default", u["username"], u["full_name"], cur_time, cur_time))
    return {"ok": True}

@app.post("/api/se93/update")
def se93_update(req: Request, code: str = Form(...), title: str = Form(...), description: str = Form(...),
                module: str = Form(...), allowed_depts: str = Form(...), active: int = Form(...)):
    need_admin(req)
    exec_write("UPDATE tcodes SET title=?, description=?, module=?, allowed_depts=?, active=?, updated_at=? WHERE code=?",
               (title.strip(), description.strip(), module.strip(), allowed_depts.strip(), active, now(), code.upper().strip()))
    return {"ok": True}

@app.post("/api/se93/delete")
def se93_delete(req: Request, code: str = Form(...)):
    need_master(req)
    code = code.upper().strip()
    exec_write("DELETE FROM tcodes WHERE code=?", (code,))
    exec_write("DELETE FROM tcode_schemas WHERE code=?", (code,))
    exec_write("DELETE FROM tcode_layouts WHERE code=?", (code,))
    exec_write("DELETE FROM tcode_records WHERE code=?", (code,))
    exec_write("DELETE FROM access_grants WHERE tcode_code=?", (code,))
    return {"ok": True}

# --- SAP ALV LAYOUT ENGINE ---
@app.get("/api/tcode/{code}/layouts")
def get_layouts(code: str, req: Request):
    u = need(req)
    code = code.upper().strip()
    if not check_tcode_access(u, code):
        raise HTTPException(403, f"No authorization for T-Code {code}")
    rows = fetch_all("SELECT * FROM tcode_layouts WHERE tcode_code=? ORDER BY is_default DESC, id ASC", (code,))
    for r in rows:
        try:
            r["fields"] = json.loads(r["fields"])
        except:
            r["fields"] = []
    return rows

@app.post("/api/tcode/{code}/layouts/save")
def save_layout(code: str, req: Request, payload: dict):
    u = need(req)
    code = code.upper().strip()
    if not check_tcode_access(u, code):
        raise HTTPException(403, f"No authorization for {code}")

    l_name = payload.get("name", "Custom Layout").strip()
    l_title = payload.get("title", l_name).strip()
    fields_json = json.dumps(payload.get("fields", []))
    config_json = json.dumps(payload.get("layout_config", {"columns": 3}))

    l_code = payload.get("layout_code", "").strip()
    if not l_code:
        cnt = fetch_one("SELECT COUNT(*) as c FROM tcode_layouts WHERE tcode_code=?", (code,))["c"]
        l_code = f"/VAR/{code}_{cnt+1:02d}"

    cur_time = now()
    existing = fetch_one("SELECT id FROM tcode_layouts WHERE tcode_code=? AND layout_code=?", (code, l_code))
    if existing:
        exec_write("UPDATE tcode_layouts SET name=?, title=?, fields=?, layout_config=?, updated_at=? WHERE id=?",
                   (l_name, l_title, fields_json, config_json, cur_time, existing["id"]))
    else:
        exec_write("INSERT INTO tcode_layouts (tcode_code, layout_code, name, title, is_default, fields, layout_config, description, created_by, created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                   (code, l_code, l_name, l_title, 0, fields_json, config_json, "User Custom Variant", u["username"], u["full_name"], cur_time, cur_time))
    return {"ok": True, "layout_code": l_code}

@app.post("/api/tcode/{code}/layouts/{lid}/set-default")
def set_default_layout(code: str, lid: int, req: Request):
    need_admin(req)
    code = code.upper().strip()
    exec_write("UPDATE tcode_layouts SET is_default=0 WHERE tcode_code=?", (code,))
    exec_write("UPDATE tcode_layouts SET is_default=1, updated_at=? WHERE id=? AND tcode_code=?", (now(), lid, code))
    return {"ok": True}

@app.delete("/api/tcode/{code}/layouts/{lid}")
def delete_layout(code: str, lid: int, req: Request):
    u = need(req)
    code = code.upper().strip()
    lay = fetch_one("SELECT * FROM tcode_layouts WHERE id=? AND tcode_code=?", (lid, code))
    if not lay:
        raise HTTPException(404, "Layout variant not found")
    if lay["is_default"] == 1:
        raise HTTPException(400, "Cannot delete default ALV layout")
    if lay["created_by"] != u["username"] and not is_master(u):
        raise HTTPException(403, "You can only delete your own layout variants")

    exec_write("DELETE FROM tcode_layouts WHERE id=?", (lid,))
    return {"ok": True}

# --- SAP DATA RECORDS & ALV GRID CONTROLLER ---
@app.get("/api/tcode/{code}/records")
def get_tcode_records(code: str, req: Request, plant: str = "", limit: int = 500, offset: int = 0):
    u = need(req)
    code = code.upper().strip()
    if not check_tcode_access(u, code):
        raise HTTPException(403, f"Security Violation: User {u['username']} has no access to {code}")

    if plant:
        rows = fetch_all("SELECT * FROM tcode_records WHERE tcode_code=? AND plant=? AND status='active' ORDER BY id DESC LIMIT ? OFFSET ?", (code, plant, limit, offset))
    else:
        rows = fetch_all("SELECT * FROM tcode_records WHERE tcode_code=? AND status='active' ORDER BY id DESC LIMIT ? OFFSET ?", (code, limit, offset))

    for r in rows:
        try:
            r["data"] = json.loads(r["data"])
        except:
            r["data"] = {}
    return rows

@app.post("/api/tcode/{code}/records")
def create_record(code: str, req: Request, payload: dict):
    u = need(req)
    code = code.upper().strip()
    if not check_tcode_access(u, code):
        raise HTTPException(403, f"No access to {code}")

    plant = payload.pop("_plant", "").strip()
    cur_time = now()
    exec_write("INSERT INTO tcode_records (tcode_code, plant, data, status, created_by, created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?)",
               (code, plant, json.dumps(payload), "active", u["username"], u["full_name"], cur_time, cur_time))
    return {"ok": True}

@app.put("/api/tcode/{code}/records/{rid}")
def update_record(code: str, rid: int, req: Request, payload: dict):
    u = need(req)
    code = code.upper().strip()
    if not check_tcode_access(u, code):
        raise HTTPException(403, f"No access to {code}")

    rec = fetch_one("SELECT * FROM tcode_records WHERE id=? AND tcode_code=?", (rid, code))
    if not rec:
        raise HTTPException(404, "Record not found")
    if rec["created_by"] != u["username"] and not is_master(u) and u["role"] != "admin":
        raise HTTPException(403, "Access Denied: Only record owner or BASIS Admin can modify this entry.")

    plant = payload.pop("_plant", rec["plant"]).strip()
    exec_write("UPDATE tcode_records SET plant=?, data=?, updated_at=? WHERE id=?",
               (plant, json.dumps(payload), now(), rid))
    return {"ok": True}

@app.delete("/api/tcode/{code}/records/{rid}")
def delete_record(code: str, rid: int, req: Request):
    u = need(req)
    code = code.upper().strip()
    if not check_tcode_access(u, code):
        raise HTTPException(403, f"No access to {code}")

    rec = fetch_one("SELECT * FROM tcode_records WHERE id=? AND tcode_code=?", (rid, code))
    if not rec:
        raise HTTPException(404, "Record not found")
    if rec["created_by"] != u["username"] and not is_master(u) and u["role"] != "admin":
        raise HTTPException(403, "Cannot delete records of other plant operators")

    exec_write("UPDATE tcode_records SET status='deleted', updated_at=? WHERE id=?", (now(), rid))
    return {"ok": True}

# --- EXCEL IMPORT / EXPORT (SAP SPREADSHEET INTERFACE) ---
@app.get("/api/tcode/{code}/export")
def export_excel(code: str, req: Request):
    u = need(req)
    code = code.upper().strip()
    if not check_tcode_access(u, code):
        raise HTTPException(403, f"Export denied for {code}")

    rows = fetch_all("SELECT * FROM tcode_records WHERE tcode_code=? AND status='active' ORDER BY id ASC", (code,))
    wb = Workbook()
    ws = wb.active
    ws.title = f"{code}_EXPORT"

    keys = ["id", "plant", "created_by_name", "created_at"]
    dynamic_keys = set()
    parsed_rows = []
    for r in rows:
        try:
            d = json.loads(r["data"]) if isinstance(r["data"], str) else r["data"]
        except:
            d = {}
        dynamic_keys.update(d.keys())
        parsed_rows.append((r, d))

    all_cols = keys + sorted(list(dynamic_keys))
    ws.append(all_cols)

    for r, d in parsed_rows:
        row_vals = [r["id"], r["plant"], r["created_by_name"], r["created_at"]]
        for k in sorted(list(dynamic_keys)):
            row_vals.append(d.get(k, ""))
        ws.append(row_vals)

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=SAP_{code}_{datetime.date.today().strftime('%Y%m%d')}.xlsx"}
    )

@app.post("/api/tcode/{code}/import")
def import_excel(code: str, req: Request, file: UploadFile = File(...)):
    u = need(req)
    code = code.upper().strip()
    if not check_tcode_access(u, code):
        raise HTTPException(403, f"Import denied for {code}")

    contents = file.file.read()
    wb = load_workbook(io.BytesIO(contents))
    ws = wb.active
    data_rows = list(ws.iter_rows(values_only=True))
    if not data_rows:
        return {"ok": True, "count": 0}

    headers = [str(h).strip() for h in data_rows[0] if h is not None]
    count = 0
    cur_time = now()

    for row in data_rows[1:]:
        if not any(row):
            continue
        rec = {}
        plant_val = ""
        for idx, col in enumerate(headers):
            if idx < len(row):
                val = row[idx]
                if col.lower() == "plant":
                    plant_val = str(val or "")
                elif col not in ["id", "created_by_name", "created_at"]:
                    rec[col] = val if val is not None else ""
        exec_write("INSERT INTO tcode_records (tcode_code, plant, data, status, created_by, created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?)",
                   (code, plant_val, json.dumps(rec), "active", u["username"], u["full_name"], cur_time, cur_time))
        count += 1
    return {"ok": True, "count": count}

# --- SAP EASY ACCESS & GUI THEMED FRONTEND ---
HTML_SAP_UI = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>SAP Easy Access - Thermocables Enterprise ERP</title>
<script src="https://cdn.tailwindcss.com"></script>
<link href="https://cdn.jsdelivr.net/npm/remixicon@4.2.0/fonts/remixicon.css" rel="stylesheet">
<style>
  body { font-family: 'Segoe UI', Tahoma, Arial, sans-serif; background-color: #dbe4ee; font-size: 13px; color: #1e293b; }
  .sap-btn { background: linear-gradient(180deg, #f8fafc 0%, #e2e8f0 100%); border: 1px solid #94a3b8; border-radius: 3px; font-size: 11px; font-weight: 600; padding: 3px 10px; box-shadow: inset 0 1px 0 rgba(255,255,255,0.8); }
  .sap-btn:hover { background: #f1f5f9; border-color: #64748b; }
  .sap-btn-primary { background: linear-gradient(180deg, #0284c7 0%, #0369a1 100%); color: white; border: 1px solid #075985; }
  .sap-btn-primary:hover { background: #0284c7; }
  .sap-input { border: 1px solid #94a3b8; border-radius: 2px; padding: 3px 6px; font-size: 12px; background: white; }
  .sap-input:focus { outline: 1px solid #0284c7; border-color: #0284c7; }
  .sap-table th { background: linear-gradient(180deg, #e2e8f0 0%, #cbd5e1 100%); border-right: 1px solid #94a3b8; border-bottom: 2px solid #64748b; padding: 4px 8px; font-size: 11px; text-transform: uppercase; }
  .sap-table td { border-bottom: 1px solid #e2e8f0; border-right: 1px solid #e2e8f0; padding: 4px 8px; font-size: 12px; }
  .sap-table tr:hover { background-color: #eff6ff; }
</style>
</head>
<body class="h-screen flex flex-col select-none overflow-hidden">

<div class="bg-[#1e293b] text-slate-200 text-xs px-3 py-1 flex items-center justify-between border-b border-slate-700">
  <div class="flex items-center gap-4">
    <span class="font-bold flex items-center gap-1.5 text-sky-400 tracking-wider"><i class="ri-instance-line"></i> SAP GUI 7.70</span>
    <span class="hover:bg-slate-700 px-2 py-0.5 rounded cursor-pointer" onclick="goHome()">Menu</span>
    <span class="hover:bg-slate-700 px-2 py-0.5 rounded cursor-pointer" onclick="goHome()">Edit</span>
    <span class="hover:bg-slate-700 px-2 py-0.5 rounded cursor-pointer" onclick="goHome()">Favorites</span>
    <span class="hover:bg-slate-700 px-2 py-0.5 rounded cursor-pointer" onclick="showHelp()">System</span>
  </div>
  <div id="sys_pill" class="text-[11px] text-slate-300 flex items-center gap-3">
    <span>Client: <b id="lbl_client">100</b></span>
    <span>Sys: <b id="lbl_sys">PRD</b></span>
    <span id="user_display" class="bg-slate-800 text-sky-300 px-2 py-0.5 rounded border border-slate-600 font-mono">Not logged in</span>
    <button onclick="logout()" class="hover:text-red-400 font-bold ml-2">Log off</button>
  </div>
</div>

<div class="bg-[#e9edf2] border-b border-slate-400 px-3 py-1.5 flex items-center justify-between">
  <div class="flex items-center gap-2">
    <div class="flex items-center bg-white border border-slate-500 rounded px-1.5 py-0.5 shadow-inner">
      <input id="cmd_box" placeholder="/n..." class="outline-none w-28 uppercase font-mono text-xs font-bold text-blue-900" onkeydown="if(event.key==='Enter') execCmd()">
      <button onclick="execCmd()" title="Enter (OK-Code)" class="text-green-700 hover:text-green-800 font-bold text-sm px-1"><i class="ri-checkbox-circle-fill"></i></button>
      <button onclick="$('#cmd_box').value=''" title="Clear" class="text-slate-400 hover:text-slate-600 text-xs px-0.5"><i class="ri-close-line"></i></button>
    </div>
    <div class="h-5 w-[1px] bg-slate-400 mx-1"></div>
    <button onclick="goHome()" title="Back (F3)" class="sap-btn flex items-center gap-1 text-emerald-800"><i class="ri-arrow-go-back-line"></i> Back</button>
    <button onclick="goHome()" title="Exit (Shift+F3)" class="sap-btn flex items-center gap-1 text-amber-800"><i class="ri-logout-box-r-line"></i> Exit</button>
    <button onclick="goHome()" title="Cancel (F12)" class="sap-btn flex items-center gap-1 text-red-800"><i class="ri-close-circle-line"></i> Cancel</button>
    <div class="h-5 w-[1px] bg-slate-400 mx-1"></div>
    <button onclick="window.print()" title="Print" class="sap-btn"><i class="ri-printer-line"></i></button>
  </div>
  <div class="flex items-center gap-2">
    <span id="title_crumb" class="font-bold text-slate-700 text-xs uppercase tracking-wide">SAP Easy Access</span>
  </div>
</div>

<div class="flex-1 flex overflow-hidden">
  <div id="tree_panel" class="w-72 bg-[#f1f5f9] border-r border-slate-400 flex flex-col justify-between">
    <div class="p-2 border-b border-slate-300 bg-slate-200 font-bold text-xs flex justify-between items-center text-slate-700">
      <span><i class="ri-folder-shared-line"></i> SAP Easy Access Menu</span>
      <button onclick="loadUserTree()" title="Refresh" class="hover:text-blue-700"><i class="ri-refresh-line"></i></button>
    </div>
    <div id="tree_view" class="flex-1 overflow-y-auto p-2 text-xs space-y-1"></div>
    <div class="p-2 bg-slate-200 border-t border-slate-300 text-[11px] text-slate-600">
      <div><b>Thermocables Limited</b></div>
      <div id="plant_info_tree" class="truncate text-slate-500">Plant: All Authorized</div>
    </div>
  </div>

  <div id="sap_screen" class="flex-1 bg-white flex flex-col overflow-hidden"></div>
</div>

<div class="bg-[#1e293b] text-slate-300 text-xs px-3 py-1 flex items-center justify-between border-t border-slate-700">
  <div class="flex items-center gap-2 truncate">
    <span id="status_icon" class="text-emerald-400"><i class="ri-checkbox-circle-fill"></i></span>
    <span id="status_msg" class="font-mono text-[11px] text-slate-200">System Ready.</span>
  </div>
  <div class="flex items-center gap-4 text-[10px] text-slate-400 font-mono">
    <span>PRD</span>
    <span>100</span>
    <span id="stat_tcode">SESSION_01</span>
    <span id="stat_resp">12ms</span>
    <span>INS</span>
  </div>
</div>

<div id="sap_dialog" class="hidden fixed inset-0 bg-slate-900/60 z-50 flex items-center justify-center p-4">
  <div id="sap_dialog_body" class="bg-white border-2 border-slate-600 rounded-sm shadow-2xl max-h-[90vh] overflow-y-auto w-full max-w-3xl"></div>
</div>

<script>
let ME = null;
let CURRENT_TCODE = null;
let ACTIVE_LAYOUTS = [];
let CURRENT_LAYOUT = null;
let ALL_TCODES = [];
const $ = s => document.querySelector(s);

function esc(str) {
  if (str === null || str === undefined) return '';
  return String(str).replace(/[&<>"']/g, function (s) {
    return {'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[s];
  });
}

function setMsg(msg, isErr = false, isWarn = false) {
  const el = $('#status_msg'), ico = $('#status_icon');
  el.textContent = msg;
  if (isErr) {
    ico.innerHTML = '<i class="ri-error-warning-fill text-red-500"></i>';
    el.className = 'font-mono text-[11px] text-red-400 font-bold';
  } else if (isWarn) {
    ico.innerHTML = '<i class="ri-alert-fill text-amber-400"></i>';
    el.className = 'font-mono text-[11px] text-amber-300';
  } else {
    ico.innerHTML = '<i class="ri-checkbox-circle-fill text-emerald-400"></i>';
    el.className = 'font-mono text-[11px] text-emerald-300';
  }
}

async function api(path, opts = {}) {
  const t0 = performance.now();
  const res = await fetch(path, {...opts, credentials: 'include'});
  const t1 = performance.now();
  $('#stat_resp').textContent = Math.round(t1 - t0) + 'ms';
  if (res.status === 401) {
    renderLogon();
    throw 'Auth';
  }
  if (!res.ok) {
    let t = await res.text();
    try { t = JSON.parse(t).detail || t; } catch {}
    throw t;
  }
  return await res.json();
}

function renderLogon() {
  $('#sap_screen').innerHTML = `
  <div class="h-full flex items-center justify-center bg-[#dbe4ee]">
    <div class="w-96 bg-white border-2 border-slate-500 shadow-xl p-6 rounded-sm">
      <div class="flex items-center gap-3 border-b-2 border-sky-800 pb-3 mb-4">
        <div class="w-10 h-10 bg-sky-900 text-white flex items-center justify-center font-bold text-xl rounded">SAP</div>
        <div>
          <h2 class="font-bold text-base text-slate-800">SAP GUI 7.70 Logon</h2>
          <div class="text-[11px] text-slate-500">Thermocables System PRD (Client 100)</div>
        </div>
      </div>
      <div class="space-y-3">
        <div>
          <label class="block text-xs font-bold text-slate-600 mb-1">User ID</label>
          <input id="l_u" class="sap-input w-full uppercase" placeholder="e.g. ADMIN" autofocus>
        </div>
        <div>
          <label class="block text-xs font-bold text-slate-600 mb-1">Password</label>
          <input id="l_p" type="password" class="sap-input w-full" placeholder="••••••••">
        </div>
        <button onclick="handleLogon()" class="sap-btn sap-btn-primary w-full py-2 mt-2">Log On</button>
      </div>
    </div>
  </div>`;
}

async function handleLogon() {
  const u = $('#l_u').value.trim(), p = $('#l_p').value;
  if (!u || !p) return setMsg("Enter User and Password", true);
  const fd = new FormData();
  fd.append("username", u); fd.append("password", p);
  try {
    await api('/api/login', {method: 'POST', body: fd});
    setMsg("RFC Connection Established.");
    initSession();
  } catch(e) {
    setMsg(e, true);
  }
}

async function logout() {
  await api('/api/logout', {method: 'POST'});
  location.reload();
}

async function initSession() {
  try {
    ME = await api('/api/me');
    $('#user_display').textContent = `${ME.username.toUpperCase()} (${ME.role.toUpperCase()})`;
    $('#lbl_client').textContent = ME.client;
    $('#lbl_sys').textContent = ME.system;
    $('#plant_info_tree').textContent = "Authorized Plants: " + (ME.plants.join(", ") || "ALL");
    setMsg(`User ${ME.username} authenticated successfully.`);
    loadUserTree();
    goHome();
  } catch(e) {
    renderLogon();
  }
}

function execCmd() {
  let val = $('#cmd_box').value.trim();
  if (!val) return;
  $('#cmd_box').value = '';
  if (val.startsWith('/n') || val.startsWith('/N')) {
    val = val.substring(2).trim();
  }
  if (!val) {
    goHome();
    return;
  }
  openTcode(val.toUpperCase());
}

function goHome() {
  CURRENT_TCODE = null;
  $('#stat_tcode').textContent = "SESSION_01";
  $('#title_crumb').textContent = "SAP Easy Access - User Menu";
  $('#sap_screen').innerHTML = `
    <div class="p-8 flex-1 overflow-y-auto bg-slate-50">
      <div class="max-w-4xl mx-auto">
        <div class="border-b-2 border-sky-800 pb-3 mb-6 flex justify-between items-end">
          <div>
            <h1 class="text-xl font-bold text-slate-800">SAP Easy Access</h1>
            <p class="text-xs text-slate-500">Enterprise Core - Thermocables ProTransfer System</p>
          </div>
          <div class="text-right text-xs">
            <span class="bg-sky-100 text-sky-800 border border-sky-300 px-2 py-1 rounded">System PRD - Client 100</span>
          </div>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-3 gap-4 mb-8">
          <div class="border border-slate-300 p-4 bg-white rounded shadow-sm">
            <div class="text-xs font-bold text-slate-500 uppercase">Organization</div>
            <div class="text-lg font-bold text-slate-800 mt-1">${esc(ME.company)}</div>
            <div class="text-xs text-slate-500 mt-1">Dept: <b>${esc(ME.department)}</b></div>
          </div>
          <div class="border border-slate-300 p-4 bg-white rounded shadow-sm">
            <div class="text-xs font-bold text-slate-500 uppercase">Profile Role</div>
            <div class="text-lg font-bold text-slate-800 mt-1">${ME.is_master ? 'SAP_ALL (MASTER)' : (ME.role === 'admin' ? 'ADMIN (S_ALL)' : 'OPERATOR')}</div>
            <div class="text-xs text-slate-500 mt-1">Plants: ${esc(ME.plants.join(', ') || 'ALL')}</div>
          </div>
          <div class="border border-slate-300 p-4 bg-white rounded shadow-sm">
            <div class="text-xs font-bold text-slate-500 uppercase">Command Shortcut</div>
            <div class="text-xs text-slate-600 mt-2">Enter <code class="bg-slate-100 px-1 border rounded text-sky-700 font-bold">/nPROD</code> or <code class="bg-slate-100 px-1 border rounded text-sky-700 font-bold">/nSU01</code> in the command box.</div>
          </div>
        </div>

        <h3 class="font-bold text-xs uppercase text-slate-600 mb-3 tracking-wider">Quick Transaction Launcher</h3>
        <div id="quick_tiles" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3"></div>
      </div>
    </div>`;

  renderQuickTiles();
}

async function loadUserTree() {
  ALL_TCODES = await api('/api/se93/list');
  let treeHtml = `
    <div class="font-bold text-slate-700 py-1 flex items-center gap-1.5"><i class="ri-folder-star-line text-amber-600"></i> Favorites</div>
    <div class="pl-4 space-y-0.5 border-l border-slate-300 ml-2 mb-2">
      <div onclick="openTcode('PROD')" class="cursor-pointer hover:bg-slate-300 px-1.5 py-0.5 rounded flex items-center gap-1"><i class="ri-file-list-2-line text-blue-700"></i> PROD - Handover</div>
      <div onclick="openTcode('GCH0')" class="cursor-pointer hover:bg-slate-300 px-1.5 py-0.5 rounded flex items-center gap-1"><i class="ri-shield-check-line text-green-700"></i> GCH0 - Green Card</div>
    </div>
    <div class="font-bold text-slate-700 py-1 flex items-center gap-1.5"><i class="ri-folder-line text-blue-600"></i> SAP Menu</div>
    <div class="pl-4 space-y-0.5 border-l border-slate-300 ml-2">
  `;

  if (ME.is_master || ME.role === 'admin') {
    treeHtml += `
      <div class="font-bold text-slate-600 pt-1 text-[11px]"><i class="ri-shield-keyhole-line text-red-600"></i> Basis Administration</div>
      <div class="pl-3 space-y-0.5 border-l border-slate-300 ml-1">
        <div onclick="openTcode('SU01')" class="cursor-pointer hover:bg-slate-300 px-1.5 py-0.5 rounded flex items-center gap-1"><i class="ri-user-settings-line text-sky-700"></i> SU01 - User Master</div>
        <div onclick="openTcode('PFCG')" class="cursor-pointer hover:bg-slate-300 px-1.5 py-0.5 rounded flex items-center gap-1"><i class="ri-lock-password-line text-sky-700"></i> PFCG - Authorizations</div>
        <div onclick="openTcode('SE93')" class="cursor-pointer hover:bg-slate-300 px-1.5 py-0.5 rounded flex items-center gap-1"><i class="ri-code-box-line text-sky-700"></i> SE93 - Maintain T-Codes</div>
        <div onclick="openTcode('SPRO')" class="cursor-pointer hover:bg-slate-300 px-1.5 py-0.5 rounded flex items-center gap-1"><i class="ri-building-line text-sky-700"></i> SPRO - Enterprise Config</div>
      </div>
    `;
  }

  treeHtml += `
    <div class="font-bold text-slate-600 pt-1 text-[11px]"><i class="ri-apps-line text-indigo-600"></i> Logistics & Operations</div>
    <div class="pl-3 space-y-0.5 border-l border-slate-300 ml-1">
  `;

  ALL_TCODES.filter(t => t.has_access).forEach(t => {
    treeHtml += `<div onclick="openTcode('${t.code}')" class="cursor-pointer hover:bg-slate-300 px-1.5 py-0.5 rounded flex items-center justify-between">
      <span><b class="text-sky-800">${t.code}</b> - ${esc(t.title)}</span>
    </div>`;
  });

  treeHtml += `</div></div>`;
  $('#tree_view').innerHTML = treeHtml;
}

function renderQuickTiles() {
  const container = $('#quick_tiles');
  if (!container) return;
  const list = ALL_TCODES.filter(t => t.has_access);
  container.innerHTML = list.map(t => `
    <div onclick="openTcode('${t.code}')" class="bg-white border border-slate-300 hover:border-sky-600 p-3 rounded cursor-pointer shadow-sm hover:shadow transition-all">
      <div class="flex justify-between items-start">
        <span class="font-bold font-mono text-sky-800 text-sm bg-sky-50 px-1.5 py-0.5 border border-sky-200 rounded">${t.code}</span>
        <span class="text-[10px] bg-slate-100 text-slate-600 px-1 rounded uppercase font-semibold">${t.module || 'ERP'}</span>
      </div>
      <div class="font-semibold text-slate-800 text-xs mt-2">${esc(t.title)}</div>
      <div class="text-[11px] text-slate-500 mt-0.5 line-clamp-1">${esc(t.description || 'SAP Transaction Standard')}</div>
      <div class="mt-3 flex justify-between items-center text-[10px] text-slate-400 border-t pt-1.5">
        <span>${t.record_count} Records</span>
        <span>${t.layout_count} ALV Layouts</span>
      </div>
    </div>
  `).join('');
}

async function openTcode(code) {
  code = code.toUpperCase().trim();
  CURRENT_TCODE = code;
  $('#stat_tcode').textContent = code;
  setMsg(`Transaction ${code} executing...`);

  if (code === 'SU01') return renderSU01();
  if (code === 'SE93') return renderSE93();
  if (code === 'SPRO') return renderSPRO();
  if (code === 'PFCG') return renderPFCG();

  const tc = ALL_TCODES.find(x => x.code === code);
  if (!tc) {
    setMsg(`Transaction ${code} does not exist in repository`, true);
    return;
  }
  if (!tc.has_access) {
    setMsg(`No authorization to start transaction ${code}`, true);
    return;
  }
  renderBusinessTcode(tc);
}

// SU01: USER MAINTENANCE
async function renderSU01() {
  if (!ME.is_master && ME.role !== 'admin') {
    setMsg("Authorization S_USER_GRP missing", true);
    return;
  }
  $('#title_crumb').textContent = "SU01: User Maintenance Master Record";
  const users = await api('/api/su01/users');

  $('#sap_screen').innerHTML = `
    <div class="p-4 flex-1 flex flex-col overflow-hidden bg-slate-100">
      <div class="bg-white border border-slate-400 p-2.5 mb-3 flex items-center justify-between rounded-sm">
        <div class="flex items-center gap-2">
          <button onclick="showCreateUserModal()" class="sap-btn sap-btn-primary"><i class="ri-user-add-line"></i> Create User (F5)</button>
          <button onclick="renderSU01()" class="sap-btn"><i class="ri-refresh-line"></i> Refresh</button>
        </div>
        <span class="text-xs text-slate-600 font-semibold">User Master Records: <b>${users.length}</b></span>
      </div>
      <div class="flex-1 bg-white border border-slate-400 overflow-auto">
        <table class="w-full sap-table border-collapse">
          <thead><tr><th>User ID</th><th>Full Name</th><th>Role</th><th>Department</th><th>Assigned Plants</th><th>Status</th><th>Actions</th></tr></thead>
          <tbody>
            ${users.map(u => `
              <tr>
                <td class="font-mono font-bold text-sky-800">${esc(u.username)}</td>
                <td>${esc(u.full_name)}</td>
                <td><span class="px-1.5 py-0.5 rounded text-[10px] uppercase font-bold ${u.role==='master'?'bg-red-100 text-red-800 border border-red-300':(u.role==='admin'?'bg-amber-100 text-amber-800 border border-amber-300':'bg-slate-100 text-slate-700')}">${u.role}</span></td>
                <td>${esc(u.department)}</td>
                <td>${esc(u.plants || 'ALL')}</td>
                <td><span class="px-1.5 py-0.2 rounded text-[10px] ${u.status==='approved'?'text-emerald-700 bg-emerald-50':'text-red-700 bg-red-50'}">${u.status}</span></td>
                <td class="space-x-1">
                  <button onclick="showEditUserModal('${esc(u.username)}')" class="sap-btn text-blue-700">Change</button>
                  <button onclick="showResetPassModal('${esc(u.username)}')" class="sap-btn text-amber-700">Password</button>
                  <button onclick="deleteUser('${esc(u.username)}')" class="sap-btn text-red-700 ${u.role==='master'?'opacity-30 cursor-not-allowed':''}">Delete</button>
                </td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    </div>`;
  setMsg("SU01: Ready.");
}

function showCreateUserModal() {
  api('/api/spro/overview').then(data => {
    openDialog(`
      <div class="border-b border-slate-400 bg-slate-200 px-4 py-2 font-bold text-slate-800 flex justify-between">
        <span>SU01 - Create User Master Record</span>
        <button onclick="closeDialog()"><i class="ri-close-line"></i></button>
      </div>
      <div class="p-4 space-y-3">
        <div class="grid grid-cols-2 gap-3">
          <div><label class="block text-xs font-bold text-slate-600">User ID</label><input id="nu_u" class="sap-input w-full uppercase"></div>
          <div><label class="block text-xs font-bold text-slate-600">Initial Password</label><input id="nu_p" type="password" class="sap-input w-full"></div>
        </div>
        <div><label class="block text-xs font-bold text-slate-600">Full Name</label><input id="nu_fn" class="sap-input w-full"></div>
        <div class="grid grid-cols-2 gap-3">
          <div>
            <label class="block text-xs font-bold text-slate-600">Department</label>
            <select id="nu_d" class="sap-input w-full">${data.departments.map(d=>`<option value="${esc(d.name)}">${esc(d.name)}</option>`).join('')}</select>
          </div>
          <div>
            <label class="block text-xs font-bold text-slate-600">Role Profile</label>
            <select id="nu_r" class="sap-input w-full">
              <option value="user">User (Standard)</option>
              <option value="admin">Admin (S_ALL)</option>
              ${ME.is_master ? '<option value="master">Master Superuser (SAP_ALL)</option>' : ''}
            </select>
          </div>
        </div>
        <div>
          <label class="block text-xs font-bold text-slate-600">Plant Authorization (comma-separated)</label>
          <input id="nu_pl" class="sap-input w-full" placeholder="e.g. 1000, 1400 (leave blank for ALL)">
        </div>
      </div>
      <div class="border-t border-slate-300 p-3 bg-slate-100 flex justify-end gap-2">
        <button onclick="submitCreateUser()" class="sap-btn sap-btn-primary">Save (Ctrl+S)</button>
        <button onclick="closeDialog()" class="sap-btn">Cancel</button>
      </div>`);
  });
}

async function submitCreateUser() {
  const fd = new FormData();
  fd.append("username", $('#nu_u').value);
  fd.append("password", $('#nu_p').value);
  fd.append("full_name", $('#nu_fn').value);
  fd.append("department", $('#nu_d').value);
  fd.append("role", $('#nu_r').value);
  fd.append("plants", $('#nu_pl').value);
  try {
    await api('/api/su01/create', {method: 'POST', body: fd});
    closeDialog();
    renderSU01();
    setMsg("User created successfully.");
  } catch(e) {
    alert("Error: " + e);
  }
}

async function showEditUserModal(u) {
  const users = await api('/api/su01/users');
  const user = users.find(x => x.username === u);
  const data = await api('/api/spro/overview');
  openDialog(`
    <div class="border-b border-slate-400 bg-slate-200 px-4 py-2 font-bold text-slate-800 flex justify-between">
      <span>SU01 - Change User Master: ${esc(u)}</span>
      <button onclick="closeDialog()"><i class="ri-close-line"></i></button>
    </div>
    <div class="p-4 space-y-3">
      <div><label class="block text-xs font-bold text-slate-600">Full Name</label><input id="eu_fn" value="${esc(user.full_name)}" class="sap-input w-full"></div>
      <div class="grid grid-cols-2 gap-3">
        <div>
          <label class="block text-xs font-bold text-slate-600">Department</label>
          <select id="eu_d" class="sap-input w-full">
            ${data.departments.map(d=>`<option value="${esc(d.name)}" ${d.name===user.department?'selected':''}>${esc(d.name)}</option>`).join('')}
          </select>
        </div>
        <div>
          <label class="block text-xs font-bold text-slate-600">Role</label>
          <select id="eu_r" class="sap-input w-full">
            <option value="user" ${user.role==='user'?'selected':''}>User</option>
            <option value="admin" ${user.role==='admin'?'selected':''}>Admin</option>
            ${ME.is_master ? `<option value="master" ${user.role==='master'?'selected':''}>Master</option>` : ''}
          </select>
        </div>
      </div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs font-bold text-slate-600">Plants</label><input id="eu_pl" value="${esc(user.plants||'')}" class="sap-input w-full"></div>
        <div>
          <label class="block text-xs font-bold text-slate-600">Status</label>
          <select id="eu_st" class="sap-input w-full">
            <option value="approved" ${user.status==='approved'?'selected':''}>Approved / Unlocked</option>
            <option value="locked" ${user.status==='locked'?'selected':''}>Locked</option>
          </select>
        </div>
      </div>
    </div>
    <div class="border-t border-slate-300 p-3 bg-slate-100 flex justify-end gap-2">
      <button onclick="submitEditUser('${esc(u)}')" class="sap-btn sap-btn-primary">Update</button>
      <button onclick="closeDialog()" class="sap-btn">Cancel</button>
    </div>`);
}

async function submitEditUser(u) {
  const fd = new FormData();
  fd.append("username", u);
  fd.append("full_name", $('#eu_fn').value);
  fd.append("department", $('#eu_d').value);
  fd.append("role", $('#eu_r').value);
  fd.append("plants", $('#eu_pl').value);
  fd.append("status", $('#eu_st').value);
  try {
    await api('/api/su01/update', {method: 'POST', body: fd});
    closeDialog();
    renderSU01();
    setMsg("User master record updated.");
  } catch(e) {
    alert("Error: " + e);
  }
}

function showResetPassModal(u) {
  const p = prompt(`Enter new password for ${u}:`);
  if (!p) return;
  const fd = new FormData();
  fd.append("username", u);
  fd.append("new_password", p);
  api('/api/su01/reset-password', {method: 'POST', body: fd}).then(() => {
    setMsg(`Password reset for ${u}`);
  }).catch(e => alert(e));
}

async function deleteUser(u) {
  if (!confirm(`Are you sure you want to delete user ${u}?`)) return;
  const fd = new FormData(); fd.append("username", u);
  try {
    await api('/api/su01/delete', {method: 'POST', body: fd});
    renderSU01();
    setMsg(`User ${u} deleted.`);
  } catch(e) {
    alert(e);
  }
}

// SE93: TRANSACTION CODES & REPOSITORY
async function renderSE93() {
  if (!ME.is_master && ME.role !== 'admin') {
    setMsg("Authorization S_DEVELOP missing", true);
    return;
  }
  $('#title_crumb').textContent = "SE93: Maintain Transaction Repository";
  const list = await api('/api/se93/list');

  $('#sap_screen').innerHTML = `
    <div class="p-4 flex-1 flex flex-col overflow-hidden bg-slate-100">
      <div class="bg-white border border-slate-400 p-2.5 mb-3 flex items-center justify-between rounded-sm">
        <div class="flex items-center gap-2">
          <button onclick="showCreateTcodeModal()" class="sap-btn sap-btn-primary"><i class="ri-add-line"></i> Create T-Code (F5)</button>
          <button onclick="renderSE93()" class="sap-btn"><i class="ri-refresh-line"></i> Refresh</button>
        </div>
        <span class="text-xs text-slate-600 font-semibold">Repository T-Codes: <b>${list.length}</b></span>
      </div>
      <div class="flex-1 bg-white border border-slate-400 overflow-auto">
        <table class="w-full sap-table border-collapse">
          <thead><tr><th>T-Code</th><th>Description / Title</th><th>Module</th><th>Authorized Depts</th><th>Status</th><th>Actions</th></tr></thead>
          <tbody>
            ${list.map(t => `
              <tr>
                <td class="font-mono font-bold text-sky-800">${esc(t.code)}</td>
                <td>${esc(t.title)}</td>
                <td><span class="font-mono bg-slate-100 px-1 border">${esc(t.module || 'ERP')}</span></td>
                <td>${esc(t.allowed_depts)}</td>
                <td><span class="text-[10px] ${t.active?'text-emerald-700':'text-red-700'}">${t.active?'ACTIVE':'INACTIVE'}</span></td>
                <td class="space-x-1">
                  <button onclick="openTcode('${esc(t.code)}')" class="sap-btn text-emerald-800">Execute</button>
                  <button onclick="showEditTcodeModal('${esc(t.code)}')" class="sap-btn text-blue-800">Edit</button>
                  <button onclick="deleteTcode('${esc(t.code)}')" class="sap-btn text-red-800">Delete</button>
                </td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    </div>`;
  setMsg("SE93: Ready.");
}

function showCreateTcodeModal() {
  openDialog(`
    <div class="border-b border-slate-400 bg-slate-200 px-4 py-2 font-bold text-slate-800 flex justify-between">
      <span>SE93 - Create Transaction Code</span>
      <button onclick="closeDialog()"><i class="ri-close-line"></i></button>
    </div>
    <div class="p-4 space-y-3">
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs font-bold text-slate-600">T-Code</label><input id="nt_c" class="sap-input w-full uppercase" placeholder="e.g. ZPR0"></div>
        <div><label class="block text-xs font-bold text-slate-600">Module</label><input id="nt_m" class="sap-input w-full uppercase" value="PP"></div>
      </div>
      <div><label class="block text-xs font-bold text-slate-600">Transaction Title</label><input id="nt_t" class="sap-input w-full" placeholder="e.g. Cable Core Processing"></div>
      <div><label class="block text-xs font-bold text-slate-600">Description</label><textarea id="nt_d" class="sap-input w-full" rows="2"></textarea></div>
      <div><label class="block text-xs font-bold text-slate-600">Allowed Departments (comma separated)</label><input id="nt_ad" class="sap-input w-full" value="Production,Quality Control"></div>
    </div>
    <div class="border-t border-slate-300 p-3 bg-slate-100 flex justify-end gap-2">
      <button onclick="submitCreateTcode()" class="sap-btn sap-btn-primary">Generate</button>
      <button onclick="closeDialog()" class="sap-btn">Cancel</button>
    </div>`);
}

async function submitCreateTcode() {
  const fd = new FormData();
  fd.append("code", $('#nt_c').value);
  fd.append("module", $('#nt_m').value);
  fd.append("title", $('#nt_t').value);
  fd.append("description", $('#nt_d').value);
  fd.append("allowed_depts", $('#nt_ad').value);
  try {
    await api('/api/se93/create', {method: 'POST', body: fd});
    closeDialog();
    renderSE93();
    loadUserTree();
    setMsg("T-Code generated in SAP Repository.");
  } catch(e) {
    alert("Error: " + e);
  }
}

async function showEditTcodeModal(code) {
  const list = await api('/api/se93/list');
  const t = list.find(x => x.code === code);
  openDialog(`
    <div class="border-b border-slate-400 bg-slate-200 px-4 py-2 font-bold text-slate-800 flex justify-between">
      <span>SE93 - Maintain Transaction: ${esc(code)}</span>
      <button onclick="closeDialog()"><i class="ri-close-line"></i></button>
    </div>
    <div class="p-4 space-y-3">
      <div><label class="block text-xs font-bold text-slate-600">Title</label><input id="et_t" value="${esc(t.title)}" class="sap-input w-full"></div>
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs font-bold text-slate-600">Module</label><input id="et_m" value="${esc(t.module||'ERP')}" class="sap-input w-full uppercase"></div>
        <div>
          <label class="block text-xs font-bold text-slate-600">Status</label>
          <select id="et_st" class="sap-input w-full">
            <option value="1" ${t.active? 'selected':''}>Active</option>
            <option value="0" ${!t.active? 'selected':''}>Inactive</option>
          </select>
        </div>
      </div>
      <div><label class="block text-xs font-bold text-slate-600">Allowed Departments</label><input id="et_ad" value="${esc(t.allowed_depts)}" class="sap-input w-full"></div>
      <div><label class="block text-xs font-bold text-slate-600">Description</label><textarea id="et_d" class="sap-input w-full">${esc(t.description||'')}</textarea></div>
    </div>
    <div class="border-t border-slate-300 p-3 bg-slate-100 flex justify-end gap-2">
      <button onclick="submitEditTcode('${esc(code)}')" class="sap-btn sap-btn-primary">Save Changes</button>
      <button onclick="closeDialog()" class="sap-btn">Cancel</button>
    </div>`);
}

async function submitEditTcode(code) {
  const fd = new FormData();
  fd.append("code", code);
  fd.append("title", $('#et_t').value);
  fd.append("module", $('#et_m').value);
  fd.append("active", $('#et_st').value);
  fd.append("allowed_depts", $('#et_ad').value);
  fd.append("description", $('#et_d').value);
  try {
    await api('/api/se93/update', {method: 'POST', body: fd});
    closeDialog();
    renderSE93();
    setMsg(`T-Code ${code} updated.`);
  } catch(e) {
    alert(e);
  }
}

async function deleteTcode(code) {
  if (!confirm(`Warning: Deleting ${code} permanently destroys all layouts and records linked to it. Continue?`)) return;
  const fd = new FormData(); fd.append("code", code);
  try {
    await api('/api/se93/delete', {method: 'POST', body: fd});
    renderSE93();
    loadUserTree();
    setMsg(`T-Code ${code} deleted.`);
  } catch(e) {
    alert(e);
  }
}

// SPRO: ENTERPRISE STRUCTURE (PLANTS / DEPTS)
async function renderSPRO() {
  if (!ME.is_master) {
    setMsg("SPRO Configuration restricted to MASTER Administrator", true);
    return;
  }
  $('#title_crumb').textContent = "SPRO: Customizing: Enterprise Structure";
  const data = await api('/api/spro/overview');

  $('#sap_screen').innerHTML = `
    <div class="p-4 flex-1 flex flex-col overflow-hidden bg-slate-100">
      <div class="grid grid-cols-1 md:grid-cols-2 gap-4 flex-1 overflow-auto">
        <div class="bg-white border border-slate-400 flex flex-col rounded-sm">
          <div class="bg-slate-200 border-b border-slate-300 p-2 font-bold text-xs flex justify-between items-center">
            <span><i class="ri-building-line text-sky-800"></i> Define Plants</span>
            <button onclick="showCreatePlantModal()" class="sap-btn sap-btn-primary">+ Add Plant</button>
          </div>
          <div class="flex-1 overflow-auto">
            <table class="w-full sap-table border-collapse">
              <thead><tr><th>Code</th><th>Plant Name</th><th>Status</th><th>Actions</th></tr></thead>
              <tbody>
                ${data.plants.map(p => `
                  <tr>
                    <td class="font-mono font-bold">${esc(p.code)}</td>
                    <td>${esc(p.name)}</td>
                    <td><span class="text-[10px] ${p.active?'text-emerald-700':'text-red-700'}">${p.active?'Active':'Disabled'}</span></td>
                    <td><button onclick="deletePlant('${esc(p.code)}')" class="sap-btn text-red-700">Delete</button></td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        </div>
        <div class="bg-white border border-slate-400 flex flex-col rounded-sm">
          <div class="bg-slate-200 border-b border-slate-300 p-2 font-bold text-xs flex justify-between items-center">
            <span><i class="ri-team-line text-sky-800"></i> Define Departments</span>
            <button onclick="showCreateDeptModal()" class="sap-btn sap-btn-primary">+ Add Dept</button>
          </div>
          <div class="flex-1 overflow-auto">
            <table class="w-full sap-table border-collapse">
              <thead><tr><th>Department Name</th><th>Status</th><th>Actions</th></tr></thead>
              <tbody>
                ${data.departments.map(d => `
                  <tr>
                    <td class="font-bold">${esc(d.name)}</td>
                    <td><span class="text-[10px] ${d.active?'text-emerald-700':'text-red-700'}">${d.active?'Active':'Disabled'}</span></td>
                    <td><button onclick="deleteDept('${esc(d.name)}')" class="sap-btn text-red-700">Delete</button></td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>`;
  setMsg("SPRO: Ready.");
}

function showCreatePlantModal() {
  const c = prompt("Enter 4-digit Plant Code (e.g. 1500):");
  if (!c) return;
  const n = prompt("Enter Plant Description:");
  if (!n) return;
  const fd = new FormData(); fd.append("code", c); fd.append("name", n);
  api('/api/admin/plants/create', {method: 'POST', body: fd}).then(() => { renderSPRO(); setMsg("Plant configured."); }).catch(e => alert(e));
}

function deletePlant(code) {
  if (!confirm(`Delete Plant ${code}?`)) return;
  const fd = new FormData(); fd.append("code", code);
  api('/api/admin/plants/delete', {method: 'POST', body: fd}).then(() => { renderSPRO(); setMsg("Plant removed."); }).catch(e => alert(e));
}

function showCreateDeptModal() {
  const d = prompt("Enter Department Name:");
  if (!d) return;
  const fd = new FormData(); fd.append("name", d);
  api('/api/admin/departments/create', {method: 'POST', body: fd}).then(() => { renderSPRO(); setMsg("Department created."); }).catch(e => alert(e));
}

function deleteDept(name) {
  if (!confirm(`Delete Department ${name}?`)) return;
  const fd = new FormData(); fd.append("name", name);
  api('/api/admin/departments/delete', {method: 'POST', body: fd}).then(() => { renderSPRO(); setMsg("Department removed."); }).catch(e => alert(e));
}

// PFCG: AUTHORIZATIONS MATRIX
async function renderPFCG() {
  if (!ME.is_master && ME.role !== 'admin') {
    setMsg("Authorization S_USER_AGR missing", true);
    return;
  }
  $('#title_crumb').textContent = "PFCG: Role & Authorization Assignment Matrix";
  const m = await api('/api/pfcg/matrix');

  $('#sap_screen').innerHTML = `
    <div class="p-4 flex-1 flex flex-col overflow-hidden bg-slate-100">
      <div class="bg-white border border-slate-400 p-2.5 mb-3 flex items-center justify-between rounded-sm">
        <span class="text-xs text-slate-700 font-bold">Explicit T-Code Grants to Users</span>
        <div class="flex items-center gap-2">
          <select id="pfcg_u" class="sap-input text-xs">${m.users.map(u=>`<option value="${esc(u.username)}">${esc(u.username)} (${esc(u.full_name)})</option>`).join('')}</select>
          <select id="pfcg_t" class="sap-input text-xs">${m.tcodes.map(t=>`<option value="${esc(t.code)}">${esc(t.code)} - ${esc(t.title)}</option>`).join('')}</select>
          <button onclick="grantAccess()" class="sap-btn sap-btn-primary">+ Grant T-Code</button>
        </div>
      </div>
      <div class="flex-1 bg-white border border-slate-400 overflow-auto">
        <table class="w-full sap-table border-collapse">
          <thead><tr><th>User ID</th><th>Granted T-Code</th><th>Granted By</th><th>Timestamp</th><th>Action</th></tr></thead>
          <tbody>
            ${m.grants.map(g => `
              <tr>
                <td class="font-mono font-bold text-sky-800">${esc(g.username)}</td>
                <td class="font-mono font-bold">${esc(g.tcode_code)}</td>
                <td>${esc(g.granted_by)}</td>
                <td>${esc(g.granted_at)}</td>
                <td><button onclick="revokeAccess('${esc(g.username)}','${esc(g.tcode_code)}')" class="sap-btn text-red-700">Revoke</button></td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    </div>`;
  setMsg("PFCG: Ready.");
}

async function grantAccess() {
  const u = $('#pfcg_u').value, t = $('#pfcg_t').value;
  const fd = new FormData(); fd.append("username", u); fd.append("tcode", t);
  await api('/api/pfcg/grant', {method: 'POST', body: fd});
  renderPFCG();
  setMsg("Authorization granted.");
}

async function revokeAccess(u, t) {
  if (!confirm(`Revoke ${t} from ${u}?`)) return;
  const fd = new FormData(); fd.append("username", u); fd.append("tcode", t);
  await api('/api/pfcg/revoke', {method: 'POST', body: fd});
  renderPFCG();
  setMsg("Authorization revoked.");
}

// BUSINESS T-CODES & ALV GRID
async function renderBusinessTcode(tc) {
  $('#title_crumb').textContent = `${tc.code}: ${tc.title}`;
  ACTIVE_LAYOUTS = await api(`/api/tcode/${tc.code}/layouts`);
  CURRENT_LAYOUT = ACTIVE_LAYOUTS.find(l => l.is_default) || ACTIVE_LAYOUTS[0];

  const sproData = await api('/api/spro/overview');
  const allowedPlants = sproData.plants.filter(p => !ME.plants.length || ME.plants.includes(p.code));

  $('#sap_screen').innerHTML = `
    <div class="flex-1 flex flex-col overflow-hidden bg-slate-100">
      <div class="bg-[#f8fafc] border-b border-slate-300 p-2 flex items-center justify-between">
        <div class="flex items-center gap-2">
          <button onclick="showRecordModal()" class="sap-btn sap-btn-primary flex items-center gap-1"><i class="ri-add-line"></i> New Entry</button>
          <div class="h-4 w-[1px] bg-slate-300"></div>
          <label class="text-xs font-bold text-slate-600 flex items-center gap-1">Plant:
            <select id="rec_plant_filter" onchange="reloadRecords()" class="sap-input text-xs font-normal">
              <option value="">(All Plants)</option>
              ${allowedPlants.map(p=>`<option value="${esc(p.code)}">${esc(p.code)} - ${esc(p.name)}</option>`).join('')}
            </select>
          </label>
          <div class="h-4 w-[1px] bg-slate-300"></div>
          <label class="text-xs font-bold text-slate-600 flex items-center gap-1">ALV Variant:
            <select id="alv_variant_sel" onchange="switchLayout(this.value)" class="sap-input text-xs font-normal">
              ${ACTIVE_LAYOUTS.map(l => `<option value="${esc(l.layout_code)}" ${CURRENT_LAYOUT && l.layout_code===CURRENT_LAYOUT.layout_code?'selected':''}>${esc(l.name)}</option>`).join('')}
            </select>
          </label>
          <button onclick="showLayoutDesigner()" title="Change Layout Variant" class="sap-btn"><i class="ri-layout-masonry-line"></i> Layout Painter</button>
        </div>
        <div class="flex items-center gap-2">
          <a href="/api/tcode/${tc.code}/export" class="sap-btn flex items-center gap-1 text-emerald-800"><i class="ri-file-excel-2-line"></i> Export ALV</a>
          <label class="sap-btn cursor-pointer flex items-center gap-1 text-sky-800">
            <i class="ri-upload-2-line"></i> Import Spreadsheet
            <input type="file" class="hidden" onchange="importSpreadsheet(event)">
          </label>
          <button onclick="reloadRecords()" class="sap-btn"><i class="ri-refresh-line"></i></button>
        </div>
      </div>
      <div id="alv_container" class="flex-1 bg-white border-t border-slate-300 overflow-auto"></div>
    </div>`;
  reloadRecords();
}

async function switchLayout(lcode) {
  CURRENT_LAYOUT = ACTIVE_LAYOUTS.find(l => l.layout_code === lcode);
  reloadRecords();
}

async function reloadRecords() {
  const p = $('#rec_plant_filter') ? $('#rec_plant_filter').value : '';
  const recs = await api(`/api/tcode/${CURRENT_TCODE}/records?plant=${p}`);
  const fields = CURRENT_LAYOUT ? CURRENT_LAYOUT.fields : [];

  let html = `
    <table class="w-full sap-table border-collapse">
      <thead>
        <tr>
          <th style="width:50px;">Doc #</th>
          <th style="width:80px;">Plant</th>
          ${fields.map(f => `<th>${esc(f.label)}</th>`).join('')}
          <th style="width:120px;">Created By</th>
          <th style="width:130px;">Timestamp</th>
          <th style="width:90px;">Actions</th>
        </tr>
      </thead>
      <tbody>`;

  if (!recs.length) {
    html += `<tr><td colspan="${fields.length + 5}" class="text-center py-8 text-slate-400">No records found matching criteria in database.</td></tr>`;
  } else {
    recs.forEach(r => {
      html += `<tr>
        <td class="font-mono text-slate-500 font-bold">${r.id}</td>
        <td><span class="bg-slate-100 font-mono text-[11px] px-1 border rounded">${esc(r.plant || 'CORP')}</span></td>
        ${fields.map(f => `<td>${esc(r.data[f.key] || '')}</td>`).join('')}
        <td>${esc(r.created_by_name || r.created_by)}</td>
        <td class="text-[11px] text-slate-500">${esc(r.created_at)}</td>
        <td class="space-x-1">
          <button onclick="editRecord(${r.id})" class="sap-btn text-blue-700">Edit</button>
          <button onclick="deleteRecord(${r.id})" class="sap-btn text-red-700">Del</button>
        </td>
      </tr>`;
    });
  }

  html += `</tbody></table>`;
  $('#alv_container').innerHTML = html;
  setMsg(`T-Code ${CURRENT_TCODE}: ${recs.length} records retrieved.`);
}

function showRecordModal(editId = null, existingData = {}, existingPlant = '') {
  const fields = CURRENT_LAYOUT ? CURRENT_LAYOUT.fields : [];
  openDialog(`
    <div class="border-b border-slate-400 bg-slate-200 px-4 py-2 font-bold text-slate-800 flex justify-between">
      <span>${editId ? 'Change Document Entry #' + editId : 'New Entry: ' + CURRENT_TCODE}</span>
      <button onclick="closeDialog()"><i class="ri-close-line"></i></button>
    </div>
    <div class="p-4 space-y-3">
      <div>
        <label class="block text-xs font-bold text-slate-600 mb-1">Production Plant</label>
        <input id="rec_f__plant" class="sap-input w-full font-mono" value="${esc(existingPlant || (ME.plants[0]||'1000'))}">
      </div>
      <div class="grid grid-cols-2 gap-3">
        ${fields.map(f => `
          <div class="col-span-${f.width||1}">
            <label class="block text-xs font-bold text-slate-600 mb-1">${esc(f.label)} ${f.required ? '<span class="text-red-500">*</span>':''}</label>
            ${f.type==='textarea' ?
              `<textarea id="rec_f_${f.key}" class="sap-input w-full" rows="2">${esc(existingData[f.key]||'')}</textarea>` :
              `<input type="${f.type==='number'?'number':(f.type==='date'?'date':'text')}" id="rec_f_${f.key}" value="${esc(existingData[f.key]||'')}" class="sap-input w-full">`
            }
          </div>
        `).join('')}
      </div>
    </div>
    <div class="border-t border-slate-300 p-3 bg-slate-100 flex justify-end gap-2">
      <button onclick="submitRecord(${editId})" class="sap-btn sap-btn-primary">Post Document</button>
      <button onclick="closeDialog()" class="sap-btn">Cancel</button>
    </div>`);
}

async function editRecord(rid) {
  const recs = await api(`/api/tcode/${CURRENT_TCODE}/records`);
  const r = recs.find(x => x.id === rid);
  if (!r) return alert("Record not found");
  showRecordModal(rid, r.data, r.plant);
}

async function submitRecord(rid) {
  const fields = CURRENT_LAYOUT ? CURRENT_LAYOUT.fields : [];
  const payload = { _plant: $('#rec_f__plant').value };
  for (let f of fields) {
    const el = $('#rec_f_' + f.key);
    if (f.required && !el.value.trim()) {
      alert(`Required field: ${f.label}`);
      return;
    }
    payload[f.key] = el.value.trim();
  }

  const url = rid ? `/api/tcode/${CURRENT_TCODE}/records/${rid}` : `/api/tcode/${CURRENT_TCODE}/records`;
  const method = rid ? 'PUT' : 'POST';

  try {
    await api(url, {
      method,
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify(payload)
    });
    closeDialog();
    reloadRecords();
    setMsg("Document posted successfully.");
  } catch(e) {
    alert(e);
  }
}

async function deleteRecord(rid) {
  if (!confirm(`Delete record entry #${rid}?`)) return;
  try {
    await api(`/api/tcode/${CURRENT_TCODE}/records/${rid}`, {method: 'DELETE'});
    reloadRecords();
    setMsg(`Record #${rid} deleted.`);
  } catch(e) {
    alert(e);
  }
}

async function importSpreadsheet(e) {
  const file = e.target.files[0];
  if (!file) return;
  const fd = new FormData(); fd.append("file", file);
  setMsg("Uploading and processing spreadsheet...");
  try {
    const res = await api(`/api/tcode/${CURRENT_TCODE}/import`, {method: 'POST', body: fd});
    reloadRecords();
    setMsg(`Successfully imported ${res.count} records into ${CURRENT_TCODE}.`);
  } catch(err) {
    alert(err);
  }
}

// SCREEN PAINTER & ALV LAYOUT DESIGNER
function showLayoutDesigner() {
  const fields = CURRENT_LAYOUT ? CURRENT_LAYOUT.fields : [];
  openDialog(`
    <div class="border-b border-slate-400 bg-slate-200 px-4 py-2 font-bold text-slate-800 flex justify-between">
      <span>Screen Painter: ALV Layout Designer for ${CURRENT_TCODE}</span>
      <button onclick="closeDialog()"><i class="ri-close-line"></i></button>
    </div>
    <div class="p-4 space-y-4">
      <div class="grid grid-cols-2 gap-3">
        <div><label class="block text-xs font-bold text-slate-600">Layout Name</label><input id="lay_name" value="${esc(CURRENT_LAYOUT.name)}" class="sap-input w-full"></div>
        <div><label class="block text-xs font-bold text-slate-600">Variant Code</label><input id="lay_code" value="${esc(CURRENT_LAYOUT.layout_code)}" class="sap-input w-full font-mono uppercase"></div>
      </div>
      <div class="border border-slate-300 p-3 bg-slate-50 rounded">
        <div class="flex justify-between items-center mb-2">
          <span class="font-bold text-xs text-slate-700">Column Schema Fields</span>
          <button onclick="addSchemaRow()" class="sap-btn sap-btn-primary">+ Add Column</button>
        </div>
        <div id="schema_rows" class="space-y-2 max-h-56 overflow-y-auto">
          ${fields.map(f => `
            <div class="flex items-center gap-2 schema-row bg-white p-1.5 border border-slate-300 rounded">
              <input placeholder="field_key" value="${esc(f.key)}" class="sap-input w-28 font-mono text-xs skey">
              <input placeholder="Column Label" value="${esc(f.label)}" class="sap-input flex-1 text-xs slbl">
              <select class="sap-input text-xs stype">
                <option value="text" ${f.type==='text'?'selected':''}>Text</option>
                <option value="number" ${f.type==='number'?'selected':''}>Number</option>
                <option value="date" ${f.type==='date'?'selected':''}>Date</option>
                <option value="textarea" ${f.type==='textarea'?'selected':''}>Text Area</option>
              </select>
              <button onclick="this.parentElement.remove()" class="text-red-600 px-1"><i class="ri-delete-bin-line"></i></button>
            </div>
          `).join('')}
        </div>
      </div>
    </div>
    <div class="border-t border-slate-300 p-3 bg-slate-100 flex justify-between">
      <div>
        ${ME.is_master ? `<button onclick="makeLayoutDefault(${CURRENT_LAYOUT.id})" class="sap-btn">Set As Global Default</button>` : ''}
      </div>
      <div class="flex gap-2">
        <button onclick="submitLayoutSave()" class="sap-btn sap-btn-primary">Save Variant</button>
        <button onclick="closeDialog()" class="sap-btn">Cancel</button>
      </div>
    </div>`);
}

function addSchemaRow() {
  const div = document.createElement('div');
  div.className = 'flex items-center gap-2 schema-row bg-white p-1.5 border border-slate-300 rounded';
  div.innerHTML = `
    <input placeholder="field_key" class="sap-input w-28 font-mono text-xs skey">
    <input placeholder="Column Label" class="sap-input flex-1 text-xs slbl">
    <select class="sap-input text-xs stype">
      <option value="text">Text</option>
      <option value="number">Number</option>
      <option value="date">Date</option>
      <option value="textarea">Text Area</option>
    </select>
    <button onclick="this.parentElement.remove()" class="text-red-600 px-1"><i class="ri-delete-bin-line"></i></button>
  `;
  $('#schema_rows').appendChild(div);
}

async function submitLayoutSave() {
  const rows = document.querySelectorAll('.schema-row');
  const fields = [];
  rows.forEach(r => {
    const key = r.querySelector('.skey').value.trim().toLowerCase().replace(/\\s+/g, '_');
    const label = r.querySelector('.slbl').value.trim();
    const type = r.querySelector('.stype').value;
    if (key && label) {
      fields.push({key, label, type, required: false, width: 1});
    }
  });

  const payload = {
    name: $('#lay_name').value.trim(),
    layout_code: $('#lay_code').value.trim(),
    fields
  };

  try {
    await api(`/api/tcode/${CURRENT_TCODE}/layouts/save`, {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify(payload)
    });
    closeDialog();
    renderBusinessTcode(ALL_TCODES.find(x => x.code === CURRENT_TCODE));
    setMsg("Layout variant saved.");
  } catch(e) {
    alert(e);
  }
}

async function makeLayoutDefault(lid) {
  await api(`/api/tcode/${CURRENT_TCODE}/layouts/${lid}/set-default`, {method: 'POST'});
  closeDialog();
  renderBusinessTcode(ALL_TCODES.find(x => x.code === CURRENT_TCODE));
  setMsg("Layout set as client default.");
}

function openDialog(content) {
  $('#sap_dialog_body').innerHTML = content;
  $('#sap_dialog').classList.remove('hidden');
}

function closeDialog() {
  $('#sap_dialog').classList.add('hidden');
  $('#sap_dialog_body').innerHTML = '';
}

function showHelp() {
  alert("Thermocables SAP GUI\\nSystem: PRD (Client 100)\\nCommand examples: /nSU01, /nSE93, /nSPRO, /nPFCG, /nPROD.");
}

initSession();
</script>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
def home():
    return HTMLResponse(HTML_SAP_UI)

@app.get("/health")
def health():
    return {"status": "UP", "rfc_engine": "ACTIVE", "system": "PRD", "client": "100"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
