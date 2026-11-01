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

DB_PATH = Path(__file__).parent / "nexus_erp.db"

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

    if USE_PG:
        try:
            # Upgrade outdated tables from previous builds
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
                cur.execute("ALTER TABLE tcodes ADD COLUMN IF NOT EXISTS module TEXT DEFAULT 'PROD';")

            conn.commit()
        except Exception as e:
            print("Auto-migration notice:", e)
            conn.rollback()

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
        cur.execute(q("INSERT INTO org VALUES (?,?,?,?)"), (1, "Thermocables Group", "NX-01", "GLOBAL"))

    # Default Departments
    for d in ["Production", "Quality Control", "Final Inspection", "Commercial", "Planning", "Logistics", "Engineering"]:
        cur.execute(q("SELECT 1 FROM departments WHERE name=?"), (d,))
        if not cur.fetchone():
            cur.execute(q("INSERT INTO departments VALUES (?,?,?)"), (d, 1, cur_time))

    # Default Plants
    for p_code, p_name in [("1000", "Gooty Central Plant"), ("1400", "Extrusion Unit 1400"), ("2000", "Hyderabad Works"), ("CORP", "Corporate Headquarters")]:
        cur.execute(q("SELECT 1 FROM plants WHERE code=?"), (p_code,))
        if not cur.fetchone():
            cur.execute(q("INSERT INTO plants VALUES (?,?,?,?)"), (p_code, p_name, 1, cur_time))

    # Master Administrator Guarantee (Always sync Admin@123 as active)
    ph = bcrypt.hashpw("Admin@123".encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
    cur.execute(q("SELECT username FROM users WHERE username='admin'"))
    if cur.fetchone():
        cur.execute(q("UPDATE users SET password_hash=?, status='approved', role='master' WHERE username='admin'"), (ph,))
    else:
        cur.execute(q("INSERT INTO users VALUES (?,?,?,?,?,?,?,?,?)"),
                    ("admin", ph, "Master System Administrator", "master", "IT & Systems", "1000,1400,2000,CORP", "approved", cur_time, cur_time))

    # Default Functional Modules / T-Codes
    default_modules = [
        ("PROD", "Production Handover Report", "Manufacturing batch & lot transfer logs", "PROD", "Production,Quality Control"),
        ("GCH0", "Quality Green Card Release", "Incoming QA & physical cable release inspection", "QC", "Quality Control,Production"),
        ("NCR0", "Non-Conformance Incidents", "QC deviation & defect corrective tracking", "QC", "Quality Control,Final Inspection,Production"),
        ("PLN0", "Production Planning & Schedule", "Machine allocation and manufacturing order queue", "PLAN", "Planning,Production,Logistics"),
        ("QCH0", "Stage QC to Final Clearance", "Clearance release slips before shipment", "QC", "Quality Control,Final Inspection"),
        ("STD0", "Inventory Transfer Orders", "Inter-plant and store material movements", "INV", "Logistics,Planning,Production")
    ]

    for code, title, desc, mod, depts in default_modules:
        cur.execute(q("SELECT 1 FROM tcodes WHERE code=?"), (code,))
        if not cur.fetchone():
            cur.execute(q("INSERT INTO tcodes VALUES (?,?,?,?,?,?,?,?,?)"),
                        (code, title, desc, mod, depts, 1, cur_time, cur_time, "SYSTEM"))
            fields = json.dumps([
                {"key": "order_no", "label": "Production Order #", "type": "text", "required": True, "width": 1},
                {"key": "batch_no", "label": "Batch / Heat Lot", "type": "text", "required": True, "width": 1},
                {"key": "cable_spec", "label": "Cable Spec / Core Details", "type": "text", "required": True, "width": 1},
                {"key": "quantity_meters", "label": "Length (Mtrs)", "type": "number", "required": True, "width": 1},
                {"key": "target_date", "label": "Schedule Date", "type": "date", "required": False, "width": 1},
                {"key": "lead_tech", "label": "Supervisor / Tech", "type": "text", "required": False, "width": 1},
                {"key": "notes", "label": "Operational Remarks", "type": "textarea", "required": False, "width": 3}
            ])
            cur.execute(q("INSERT INTO tcode_schemas VALUES (?,?,?,?,?)"), (code, fields, json.dumps({"columns": 3}), cur_time, "SYSTEM"))
            cur.execute(q("INSERT INTO tcode_layouts (tcode_code, layout_code, name, title, is_default, fields, layout_config, description, created_by, created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)"),
                        (code, f"GRID_{code}_DEF", "Standard Data View", f"{title} Grid", 1, fields, json.dumps({"columns": 3}), "System Default Table Layout", "admin", "Master Administrator", cur_time, cur_time))

    conn.commit()
    conn.close()

init_db()

# --- AUTH & ACCESS HELPERS ---
def get_user(req: Request):
    token = None
    auth_header = req.headers.get("Authorization")
    if auth_header and auth_header.startswith("Bearer "):
        token = auth_header.split(" ")[1].strip()
    if not token:
        token = req.cookies.get("nexus_session") or req.cookies.get("sap_session")
    if not token:
        return None

    s = fetch_one("SELECT * FROM sessions WHERE token=?", (token,))
    if not s:
        return None
    u = fetch_one("SELECT * FROM users WHERE username=?", (s["username"],))
    if not u or u["status"] != "approved":
        return None
    return u

def need(req: Request):
    u = get_user(req)
    if not u:
        raise HTTPException(401, "Authentication required. Please sign in.")
    return u

def is_master(u):
    return u["role"] == "master" or u["username"] == "admin"

def need_admin(req: Request):
    u = need(req)
    if u["role"] not in ["admin", "master"]:
        raise HTTPException(403, "Access restricted to Administrators.")
    return u

def need_master(req: Request):
    u = need(req)
    if not is_master(u):
        raise HTTPException(403, "Security Access Denied: Master Administrator clearance required.")
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
app = FastAPI(title="Nexus Core ERP - Industrial Operations Suite", version="3.2.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True, allow_methods=["*"], allow_headers=["*"])

# --- AUTH APIS ---
@app.post("/api/login")
def api_login(res: Response, username: str = Form(...), password: str = Form(...)):
    username = username.lower().strip()
    u = fetch_one("SELECT * FROM users WHERE username=?", (username,))
    if not u or not bcrypt.checkpw(password.encode('utf-8'), u["password_hash"].encode('utf-8')):
        raise HTTPException(401, "Invalid username or password.")
    if u["status"] != "approved":
        raise HTTPException(403, "Account is locked or pending administrative approval.")

    token = secrets.token_hex(32)
    exp = (datetime.datetime.now() + datetime.timedelta(days=7)).strftime("%Y-%m-%d %H:%M:%S")
    exec_write("INSERT INTO sessions VALUES (?,?,?)", (token, username, exp))

    # Dual session: Set cookie AND return token for Bearer header storage
    res.set_cookie("nexus_session", token, httponly=True, samesite="lax", max_age=86400 * 7)
    return {
        "ok": True,
        "token": token,
        "user": {
            "username": u["username"],
            "full_name": u["full_name"],
            "role": u["role"]
        }
    }

@app.post("/api/logout")
def api_logout(res: Response, req: Request):
    token = req.headers.get("Authorization", "").replace("Bearer ", "").strip() or req.cookies.get("nexus_session")
    if token:
        exec_write("DELETE FROM sessions WHERE token=?", (token,))
    res.delete_cookie("nexus_session")
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
        "system": org_info.get("system_id", "NX-01"),
        "client": org_info.get("client", "GLOBAL"),
        "company": org_info.get("name", "Thermocables Group")
    }

# --- SYSTEM SETUP & UNITS (SPRO) ---
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
        raise HTTPException(400, f"Facility code '{code}' already exists.")
    exec_write("INSERT INTO plants VALUES (?,?,?,?)", (code, name.strip(), 1, now()))
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

# --- USER IDENTITY (SU01) ---
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
        raise HTTPException(403, "Only Master Admins can create superuser accounts.")
    if fetch_one("SELECT 1 FROM users WHERE username=?", (username,)):
        raise HTTPException(400, f"User '{username}' already exists.")

    ph = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
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
        raise HTTPException(403, "Cannot modify a Master Administrator.")
    if role == "master" and not is_master(caller):
        raise HTTPException(403, "Cannot grant Master privileges.")

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
        raise HTTPException(400, "Master Administrator account is protected from deletion.")
    if target["role"] == "admin" and not is_master(caller):
        raise HTTPException(403, "Only Master Admins can delete administrative accounts.")
    if username == caller["username"]:
        raise HTTPException(400, "You cannot delete your own session.")

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
        raise HTTPException(403, "Cannot reset password for Master Administrator.")

    ph = bcrypt.hashpw(new_password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
    exec_write("UPDATE users SET password_hash=?, updated_at=? WHERE username=?", (ph, now(), username))
    exec_write("DELETE FROM sessions WHERE username=?", (username,))
    return {"ok": True}

# --- ACCESS CONTROL MATRIX (PFCG) ---
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

# --- MODULES & REGISTRY STUDIO (SE93) ---
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
                module: str = Form("PROD"), allowed_depts: str = Form(...)):
    u = need_admin(req)
    code = code.upper().strip()
    if fetch_one("SELECT 1 FROM tcodes WHERE code=?", (code,)):
        raise HTTPException(400, f"Module code '{code}' already exists.")

    cur_time = now()
    exec_write("INSERT INTO tcodes VALUES (?,?,?,?,?,?,?,?,?)",
               (code, title.strip(), description.strip(), module.strip(), allowed_depts.strip(), 1, cur_time, cur_time, u["username"]))

    starter_fields = json.dumps([
        {"key": "order_no", "label": "Document / Order #", "type": "text", "required": True, "width": 1},
        {"key": "batch_lot", "label": "Batch / Heat Lot", "type": "text", "required": True, "width": 1},
        {"key": "item_desc", "label": "Item Description", "type": "text", "required": False, "width": 2},
        {"key": "quantity", "label": "Volume / Qty", "type": "number", "required": False, "width": 1},
        {"key": "remarks", "label": "Workstation Notes", "type": "textarea", "required": False, "width": 3}
    ])
    exec_write("INSERT INTO tcode_schemas VALUES (?,?,?,?,?)", (code, starter_fields, json.dumps({"columns": 3}), cur_time, u["username"]))
    exec_write("INSERT INTO tcode_layouts (tcode_code, layout_code, name, title, is_default, fields, layout_config, description, created_by, created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
               (code, f"GRID_{code}_01", "Standard Grid", f"{title} Data View", 1, starter_fields, json.dumps({"columns": 3}), "Generated layout", u["username"], u["full_name"], cur_time, cur_time))
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

# --- VIEW & LAYOUT DESIGNER ---
@app.get("/api/tcode/{code}/layouts")
def get_layouts(code: str, req: Request):
    u = need(req)
    code = code.upper().strip()
    if not check_tcode_access(u, code):
        raise HTTPException(403, f"No authorization for code {code}")
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
        raise HTTPException(403, "Access denied")

    l_name = payload.get("name", "Custom View").strip()
    l_title = payload.get("title", l_name).strip()
    fields_json = json.dumps(payload.get("fields", []))
    config_json = json.dumps(payload.get("layout_config", {"columns": 3}))

    l_code = payload.get("layout_code", "").strip()
    if not l_code:
        cnt = fetch_one("SELECT COUNT(*) as c FROM tcode_layouts WHERE tcode_code=?", (code,))["c"]
        l_code = f"VIEW_{code}_{cnt+1:02d}"

    cur_time = now()
    existing = fetch_one("SELECT id FROM tcode_layouts WHERE tcode_code=? AND layout_code=?", (code, l_code))
    if existing:
        exec_write("UPDATE tcode_layouts SET name=?, title=?, fields=?, layout_config=?, updated_at=? WHERE id=?",
                   (l_name, l_title, fields_json, config_json, cur_time, existing["id"]))
    else:
        exec_write("INSERT INTO tcode_layouts (tcode_code, layout_code, name, title, is_default, fields, layout_config, description, created_by, created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                   (code, l_code, l_name, l_title, 0, fields_json, config_json, "Custom view", u["username"], u["full_name"], cur_time, cur_time))
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
        raise HTTPException(404, "Layout view not found")
    if lay["is_default"] == 1:
        raise HTTPException(400, "Cannot delete the default view")
    if lay["created_by"] != u["username"] and not is_master(u):
        raise HTTPException(403, "You can only delete your own custom views")

    exec_write("DELETE FROM tcode_layouts WHERE id=?", (lid,))
    return {"ok": True}

# --- DATA RECORDS & GRID ---
@app.get("/api/tcode/{code}/records")
def get_tcode_records(code: str, req: Request, plant: str = "", limit: int = 500, offset: int = 0):
    u = need(req)
    code = code.upper().strip()
    if not check_tcode_access(u, code):
        raise HTTPException(403, f"Unauthorized: User {u['username']} lacks permission for {code}")

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
        raise HTTPException(403, "Access denied")

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
        raise HTTPException(403, "Access denied")

    rec = fetch_one("SELECT * FROM tcode_records WHERE id=? AND tcode_code=?", (rid, code))
    if not rec:
        raise HTTPException(404, "Record not found")
    if rec["created_by"] != u["username"] and not is_master(u) and u["role"] != "admin":
        raise HTTPException(403, "You can only edit records you submitted or require administrative clearance.")

    plant = payload.pop("_plant", rec["plant"]).strip()
    exec_write("UPDATE tcode_records SET plant=?, data=?, updated_at=? WHERE id=?",
               (plant, json.dumps(payload), now(), rid))
    return {"ok": True}

@app.delete("/api/tcode/{code}/records/{rid}")
def delete_record(code: str, rid: int, req: Request):
    u = need(req)
    code = code.upper().strip()
    if not check_tcode_access(u, code):
        raise HTTPException(403, "Access denied")

    rec = fetch_one("SELECT * FROM tcode_records WHERE id=? AND tcode_code=?", (rid, code))
    if not rec:
        raise HTTPException(404, "Record not found")
    if rec["created_by"] != u["username"] and not is_master(u) and u["role"] != "admin":
        raise HTTPException(403, "Permission denied")

    exec_write("UPDATE tcode_records SET status='deleted', updated_at=? WHERE id=?", (now(), rid))
    return {"ok": True}

# --- EXCEL SPREADSHEET ENGINE ---
@app.get("/api/tcode/{code}/export")
def export_excel(code: str, req: Request):
    u = need(req)
    code = code.upper().strip()
    if not check_tcode_access(u, code):
        raise HTTPException(403, "Export denied")

    rows = fetch_all("SELECT * FROM tcode_records WHERE tcode_code=? AND status='active' ORDER BY id ASC", (code,))
    wb = Workbook()
    ws = wb.active
    ws.title = f"{code}_DATA"

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
        headers={"Content-Disposition": f"attachment; filename=Export_{code}_{datetime.date.today().strftime('%Y%m%d')}.xlsx"}
    )

@app.post("/api/tcode/{code}/import")
def import_excel(code: str, req: Request, file: UploadFile = File(...)):
    u = need(req)
    code = code.upper().strip()
    if not check_tcode_access(u, code):
        raise HTTPException(403, "Import denied")

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

# --- MODERN SLATE & INDIGO UI (NON-SAP LOOK) ---
HTML_MODERN_UI = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Nexus Core - Enterprise Operations Platform</title>
<script src="https://cdn.tailwindcss.com"></script>
<link href="https://cdn.jsdelivr.net/npm/remixicon@4.2.0/fonts/remixicon.css" rel="stylesheet">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">
<style>
  body { font-family: 'Plus Jakarta Sans', sans-serif; background-color: #0b0f17; color: #1e293b; }
  .font-mono { font-family: 'JetBrains Mono', monospace; }
  .btn-primary { background: #4f46e5; color: white; transition: all 0.15s ease; }
  .btn-primary:hover { background: #4338ca; }
  .btn-secondary { background: white; border: 1px solid #e2e8f0; color: #334155; transition: all 0.15s ease; }
  .btn-secondary:hover { background: #f8fafc; border-color: #cbd5e1; }
  .custom-table th { background: #f8fafc; color: #64748b; font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em; padding: 10px 14px; border-bottom: 1px solid #e2e8f0; }
  .custom-table td { padding: 12px 14px; border-bottom: 1px solid #f1f5f9; font-size: 13px; }
  .custom-table tr:hover { background-color: #f8fafc; }
</style>
</head>
<body class="h-screen flex flex-col select-none overflow-hidden bg-slate-900 text-slate-100">

<!-- APP SHELL -->
<div id="app_root" class="h-full flex flex-col">

  <!-- TOP HEADER -->
  <header class="h-14 bg-slate-900/90 backdrop-blur border-b border-slate-800 px-4 flex items-center justify-between z-10 shrink-0">
    <div class="flex items-center gap-4">
      <div class="flex items-center gap-2.5 cursor-pointer" onclick="goHome()">
        <div class="w-8 h-8 rounded-lg bg-indigo-600 flex items-center justify-center font-bold text-white shadow-md shadow-indigo-600/30">
          <i class="ri-pulse-fill text-lg"></i>
        </div>
        <div>
          <div class="font-bold text-sm tracking-tight text-white flex items-center gap-1.5">
            Nexus Core <span class="text-[10px] bg-indigo-500/20 text-indigo-300 border border-indigo-500/30 font-mono px-1.5 py-0.5 rounded">v3.2</span>
          </div>
        </div>
      </div>

      <!-- GLOBAL COMMAND JUMP BAR -->
      <div class="relative ml-4">
        <i class="ri-search-line absolute left-3 top-2.5 text-slate-400 text-xs"></i>
        <input id="cmd_box" placeholder="Jump to module (e.g. PROD, GCH0, SU01, SPRO)..." 
               class="bg-slate-800/80 border border-slate-700/80 rounded-lg pl-8 pr-16 py-1.5 text-xs text-slate-100 placeholder:text-slate-500 w-80 focus:outline-none focus:border-indigo-500 transition-all font-mono"
               onkeydown="if(event.key==='Enter') execCmd()">
        <span class="absolute right-2 top-2 text-[10px] bg-slate-700/60 text-slate-400 px-1.5 py-0.5 rounded font-mono">Enter</span>
      </div>
    </div>

    <!-- USER & ORG PILL -->
    <div id="nav_right" class="flex items-center gap-3">
      <div id="user_chip" class="flex items-center gap-2 bg-slate-800 border border-slate-700/60 px-3 py-1 rounded-full text-xs">
        <div class="w-2 h-2 rounded-full bg-emerald-400 animate-pulse"></div>
        <span id="user_display" class="font-medium text-slate-200">Connecting...</span>
        <span id="role_badge" class="bg-indigo-950 text-indigo-300 border border-indigo-700/40 text-[10px] px-1.5 py-0.5 rounded-full font-semibold uppercase"></span>
      </div>
      <button onclick="logout()" title="Sign Out" class="text-xs bg-slate-800 hover:bg-red-950/40 text-slate-300 hover:text-red-400 border border-slate-700 hover:border-red-900/40 p-2 rounded-lg transition-colors">
        <i class="ri-logout-box-r-line"></i>
      </button>
    </div>
  </header>

  <!-- BODY: SIDEBAR + WORKSPACE -->
  <div class="flex-1 flex overflow-hidden">
    
    <!-- LEFT SIDEBAR -->
    <aside id="sidebar" class="w-64 bg-slate-950 border-r border-slate-800 flex flex-col justify-between shrink-0">
      <div class="p-3 overflow-y-auto flex-1 space-y-4">
        <div>
          <div class="text-[11px] uppercase font-bold text-slate-500 tracking-wider px-2 mb-1.5">Quick Launch</div>
          <div id="quick_links" class="space-y-0.5">
            <button onclick="openTcode('PROD')" class="w-full text-left px-2.5 py-1.5 rounded-lg text-xs font-medium text-slate-300 hover:bg-slate-800/80 hover:text-white flex items-center justify-between group">
              <span class="flex items-center gap-2"><i class="ri-git-commit-line text-indigo-400"></i> PROD - Handover</span>
              <span class="text-[10px] text-slate-500 font-mono opacity-0 group-hover:opacity-100">Open</span>
            </button>
            <button onclick="openTcode('GCH0')" class="w-full text-left px-2.5 py-1.5 rounded-lg text-xs font-medium text-slate-300 hover:bg-slate-800/80 hover:text-white flex items-center justify-between group">
              <span class="flex items-center gap-2"><i class="ri-shield-check-line text-emerald-400"></i> GCH0 - Green Card</span>
              <span class="text-[10px] text-slate-500 font-mono opacity-0 group-hover:opacity-100">Open</span>
            </button>
          </div>
        </div>

        <div id="admin_nav_section" class="hidden">
          <div class="text-[11px] uppercase font-bold text-indigo-400 tracking-wider px-2 mb-1.5 flex items-center justify-between">
            <span>Administration</span>
            <i class="ri-shield-keyhole-line text-xs"></i>
          </div>
          <div class="space-y-0.5">
            <button onclick="openTcode('SU01')" class="w-full text-left px-2.5 py-1.5 rounded-lg text-xs font-medium text-slate-300 hover:bg-slate-800 hover:text-white flex items-center gap-2">
              <i class="ri-user-settings-line text-slate-400"></i> User Directory (SU01)
            </button>
            <button onclick="openTcode('PFCG')" class="w-full text-left px-2.5 py-1.5 rounded-lg text-xs font-medium text-slate-300 hover:bg-slate-800 hover:text-white flex items-center gap-2">
              <i class="ri-shield-user-line text-slate-400"></i> Access Policies (PFCG)
            </button>
            <button onclick="openTcode('SE93')" class="w-full text-left px-2.5 py-1.5 rounded-lg text-xs font-medium text-slate-300 hover:bg-slate-800 hover:text-white flex items-center gap-2">
              <i class="ri-apps-2-line text-slate-400"></i> Module Studio (SE93)
            </button>
            <button onclick="openTcode('SPRO')" class="w-full text-left px-2.5 py-1.5 rounded-lg text-xs font-medium text-slate-300 hover:bg-slate-800 hover:text-white flex items-center gap-2">
              <i class="ri-building-line text-slate-400"></i> Units & Plants (SPRO)
            </button>
          </div>
        </div>

        <div>
          <div class="text-[11px] uppercase font-bold text-slate-500 tracking-wider px-2 mb-1.5">Operations Modules</div>
          <div id="modules_list" class="space-y-0.5 text-xs"></div>
        </div>
      </div>

      <div class="p-3 border-t border-slate-800/80 bg-slate-950/60 text-[11px]">
        <div class="font-bold text-slate-300 truncate" id="side_company">Thermocables Group</div>
        <div class="text-slate-500 truncate" id="side_plant">Facility: Authorized</div>
      </div>
    </aside>

    <!-- RIGHT MAIN WORKSPACE -->
    <main id="main_screen" class="flex-1 bg-slate-900 overflow-hidden flex flex-col"></main>
  </div>

  <!-- STATUS FOOTER -->
  <footer class="h-7 bg-slate-950 border-t border-slate-800 px-4 flex items-center justify-between text-[11px] text-slate-400 font-mono shrink-0">
    <div class="flex items-center gap-2 truncate">
      <span id="stat_dot" class="w-2 h-2 rounded-full bg-emerald-400"></span>
      <span id="stat_msg" class="text-slate-300">Operational readiness OK.</span>
    </div>
    <div class="flex items-center gap-4 text-slate-500">
      <span id="stat_tcode" class="text-indigo-400 font-semibold font-mono">DASHBOARD</span>
      <span id="stat_resp">8ms</span>
      <span>READY</span>
    </div>
  </footer>

</div>

<!-- DIALOG MODAL -->
<div id="app_dialog" class="hidden fixed inset-0 bg-slate-950/70 backdrop-blur-sm z-50 flex items-center justify-center p-4">
  <div id="app_dialog_body" class="bg-white text-slate-800 rounded-2xl shadow-2xl border border-slate-200 max-h-[90vh] overflow-y-auto w-full max-w-2xl"></div>
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
  return String(str).replace(/[&<>"']/g, s => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[s]));
}

function setMsg(msg, isErr = false) {
  const el = $('#stat_msg'), dot = $('#stat_dot');
  el.textContent = msg;
  if (isErr) {
    dot.className = "w-2 h-2 rounded-full bg-red-400";
    el.className = "text-red-400 font-semibold";
  } else {
    dot.className = "w-2 h-2 rounded-full bg-emerald-400";
    el.className = "text-slate-300";
  }
}

async function api(path, opts = {}) {
  const token = localStorage.getItem("nexus_token");
  const headers = {...(opts.headers || {})};
  if (token) {
    headers["Authorization"] = `Bearer ${token}`;
  }
  const t0 = performance.now();
  const res = await fetch(path, {...opts, headers, credentials: 'include'});
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

// MODERN SIGN-IN SCREEN
function renderLogon(errorMsg = "") {
  $('#app_root').innerHTML = `
  <div class="min-h-screen bg-slate-950 flex items-center justify-center p-4 relative overflow-hidden select-text">
    <div class="absolute -top-32 -left-32 w-96 h-96 bg-indigo-600/15 rounded-full blur-3xl pointer-events-none"></div>
    <div class="absolute -bottom-32 -right-32 w-96 h-96 bg-emerald-600/15 rounded-full blur-3xl pointer-events-none"></div>

    <div class="w-full max-w-md bg-slate-900 border border-slate-800 rounded-2xl shadow-2xl p-8 z-10">
      <div class="flex items-center gap-3 mb-6">
        <div class="w-10 h-10 rounded-xl bg-indigo-600 flex items-center justify-center font-bold text-white shadow-lg shadow-indigo-600/40">
          <i class="ri-pulse-fill text-2xl"></i>
        </div>
        <div>
          <h1 class="text-xl font-bold text-white tracking-tight">Nexus Core ERP</h1>
          <p class="text-xs text-slate-400">Enterprise Industrial System - Thermocables</p>
        </div>
      </div>

      ${errorMsg ? `<div class="mb-4 p-3 bg-red-950/40 border border-red-800/60 rounded-xl text-red-300 text-xs flex items-center gap-2"><i class="ri-error-warning-line text-base"></i> ${esc(errorMsg)}</div>` : ''}

      <div class="space-y-4">
        <div>
          <label class="block text-xs font-semibold text-slate-300 mb-1.5">Username</label>
          <input id="l_u" class="w-full bg-slate-800 border border-slate-700 rounded-xl px-3.5 py-2.5 text-sm text-white placeholder-slate-500 focus:outline-none focus:border-indigo-500 transition-colors" placeholder="e.g. admin" autofocus>
        </div>
        <div>
          <label class="block text-xs font-semibold text-slate-300 mb-1.5">Password</label>
          <input id="l_p" type="password" class="w-full bg-slate-800 border border-slate-700 rounded-xl px-3.5 py-2.5 text-sm text-white placeholder-slate-500 focus:outline-none focus:border-indigo-500 transition-colors" placeholder="••••••••" onkeydown="if(event.key==='Enter') handleLogon()">
        </div>
        <button onclick="handleLogon()" class="w-full btn-primary py-2.5 rounded-xl font-semibold text-sm shadow-md shadow-indigo-600/30">Sign In to Workspace</button>
      </div>

      <div class="mt-6 pt-5 border-t border-slate-800/80 flex items-center justify-between text-[11px] text-slate-500">
        <span>Default Admin: <b class="text-slate-400">admin</b></span>
        <span>Password: <b class="text-slate-400">Admin@123</b></span>
      </div>
    </div>
  </div>`;
}

async function handleLogon() {
  const u = $('#l_u').value.trim(), p = $('#l_p').value;
  if (!u || !p) return renderLogon("Please enter both username and password.");
  const fd = new FormData();
  fd.append("username", u); fd.append("password", p);
  try {
    const res = await fetch('/api/login', {method: 'POST', body: fd});
    if (!res.ok) {
      let t = await res.text();
      try { t = JSON.parse(t).detail || t; } catch {}
      return renderLogon(t || "Login failed");
    }
    const data = await res.json();
    if (data.token) {
      localStorage.setItem("nexus_token", data.token);
    }
    location.reload();
  } catch(e) {
    renderLogon("Connection error: " + e);
  }
}

async function logout() {
  await api('/api/logout', {method: 'POST'});
  localStorage.removeItem("nexus_token");
  location.reload();
}

async function initSession() {
  try {
    ME = await api('/api/me');
    $('#user_display').textContent = ME.full_name || ME.username;
    $('#role_badge').textContent = ME.role;
    $('#side_company').textContent = ME.company;
    $('#side_plant').textContent = "Facilities: " + (ME.plants.join(", ") || "All Authorized");

    if (ME.is_master || ME.role === 'admin') {
      $('#admin_nav_section').classList.remove('hidden');
    }
    setMsg(`Connected as ${ME.username}`);
    await loadModules();
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
  if (!val) { goHome(); return; }
  openTcode(val.toUpperCase());
}

function goHome() {
  CURRENT_TCODE = null;
  $('#stat_tcode').textContent = "DASHBOARD";
  $('#main_screen').innerHTML = `
    <div class="flex-1 overflow-y-auto p-8 bg-slate-900 select-text">
      <div class="max-w-5xl mx-auto space-y-8">
        <div>
          <h1 class="text-2xl font-bold text-white tracking-tight">Operations Dashboard</h1>
          <p class="text-sm text-slate-400 mt-1">Thermocables Industrial Management Platform</p>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div class="bg-slate-800/60 border border-slate-700/60 rounded-2xl p-5">
            <div class="text-xs font-semibold text-slate-400 uppercase tracking-wider">Enterprise Org</div>
            <div class="text-xl font-bold text-white mt-1.5">${esc(ME.company)}</div>
            <div class="text-xs text-indigo-400 mt-1 font-medium">Department: ${esc(ME.department)}</div>
          </div>
          <div class="bg-slate-800/60 border border-slate-700/60 rounded-2xl p-5">
            <div class="text-xs font-semibold text-slate-400 uppercase tracking-wider">Security Profile</div>
            <div class="text-xl font-bold text-emerald-400 mt-1.5 uppercase">${ME.role}</div>
            <div class="text-xs text-slate-400 mt-1 truncate">Plants: ${esc(ME.plants.join(', ') || 'Global')}</div>
          </div>
          <div class="bg-slate-800/60 border border-slate-700/60 rounded-2xl p-5">
            <div class="text-xs font-semibold text-slate-400 uppercase tracking-wider">Command Shortcut</div>
            <div class="text-xs text-slate-300 mt-2 font-mono">Type module code in the search bar (e.g. <b class="text-indigo-400">PROD</b>, <b class="text-indigo-400">GCH0</b>)</div>
          </div>
        </div>

        <div>
          <h3 class="text-sm font-bold text-slate-300 uppercase tracking-wider mb-4">Operations Modules</h3>
          <div id="quick_grid" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4"></div>
        </div>
      </div>
    </div>`;

  renderQuickGrid();
}

async function loadModules() {
  ALL_TCODES = await api('/api/se93/list');
  const allowed = ALL_TCODES.filter(t => t.has_access);
  $('#modules_list').innerHTML = allowed.map(t => `
    <button onclick="openTcode('${t.code}')" class="w-full text-left px-2.5 py-1.5 rounded-lg text-slate-300 hover:bg-slate-800 hover:text-white flex items-center justify-between group">
      <span class="truncate"><b class="text-indigo-400 font-mono">${t.code}</b> <span class="text-slate-400">- ${esc(t.title)}</span></span>
    </button>
  `).join('');
}

function renderQuickGrid() {
  const container = $('#quick_grid');
  if (!container) return;
  const list = ALL_TCODES.filter(t => t.has_access);
  container.innerHTML = list.map(t => `
    <div onclick="openTcode('${t.code}')" class="bg-slate-800/40 hover:bg-slate-800/80 border border-slate-700/60 hover:border-indigo-500/50 p-5 rounded-2xl cursor-pointer transition-all">
      <div class="flex items-center justify-between mb-2">
        <span class="font-mono font-bold text-sm bg-indigo-500/10 text-indigo-400 border border-indigo-500/30 px-2 py-0.5 rounded-lg">${t.code}</span>
        <span class="text-[10px] bg-slate-700 text-slate-300 px-1.5 py-0.5 rounded uppercase font-semibold">${t.module || 'ERP'}</span>
      </div>
      <h4 class="font-semibold text-white text-sm mt-2">${esc(t.title)}</h4>
      <p class="text-xs text-slate-400 mt-1 line-clamp-2">${esc(t.description || 'Standard module flow')}</p>
      <div class="mt-4 pt-3 border-t border-slate-700/50 flex justify-between text-[11px] text-slate-500 font-mono">
        <span>${t.record_count} Records</span>
        <span>${t.layout_count} Views</span>
      </div>
    </div>
  `).join('');
}

async function openTcode(code) {
  code = code.toUpperCase().trim();
  CURRENT_TCODE = code;
  $('#stat_tcode').textContent = code;
  setMsg(`Module ${code} loaded`);

  if (code === 'SU01') return renderSU01();
  if (code === 'SE93') return renderSE93();
  if (code === 'SPRO') return renderSPRO();
  if (code === 'PFCG') return renderPFCG();

  const tc = ALL_TCODES.find(x => x.code === code);
  if (!tc) return setMsg(`Module ${code} not found`, true);
  if (!tc.has_access) return setMsg(`Permission denied for ${code}`, true);
  renderBusinessModule(tc);
}

// SU01: USER DIRECTORY
async function renderSU01() {
  if (!ME.is_master && ME.role !== 'admin') return setMsg("Permission denied", true);
  const users = await api('/api/su01/users');

  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col bg-slate-900 p-6 overflow-hidden select-text">
      <div class="flex items-center justify-between pb-5 border-b border-slate-800 shrink-0">
        <div>
          <h2 class="text-lg font-bold text-white tracking-tight">User Directory & IAM</h2>
          <p class="text-xs text-slate-400">Manage user credentials, unit access, and roles</p>
        </div>
        <button onclick="showCreateUserModal()" class="btn-primary text-xs font-semibold px-4 py-2 rounded-xl flex items-center gap-1.5">
          <i class="ri-user-add-line"></i> Add New User
        </button>
      </div>

      <div class="flex-1 mt-4 bg-slate-950 border border-slate-800 rounded-2xl overflow-auto">
        <table class="w-full text-left custom-table">
          <thead>
            <tr>
              <th>Username</th><th>Full Name</th><th>Role</th><th>Department</th><th>Facilities</th><th>Status</th><th>Actions</th>
            </tr>
          </thead>
          <tbody class="text-slate-300">
            ${users.map(u => `
              <tr class="border-b border-slate-800/60 hover:bg-slate-900/60">
                <td class="font-mono font-bold text-indigo-400">${esc(u.username)}</td>
                <td class="font-medium text-white">${esc(u.full_name)}</td>
                <td><span class="text-[10px] font-semibold px-2 py-0.5 rounded-full ${u.role==='master'?'bg-red-950 text-red-300 border border-red-800':(u.role==='admin'?'bg-amber-950 text-amber-300 border border-amber-800':'bg-slate-800 text-slate-300')}">${u.role.toUpperCase()}</span></td>
                <td>${esc(u.department)}</td>
                <td class="font-mono text-xs">${esc(u.plants || 'ALL')}</td>
                <td><span class="text-[10px] font-semibold px-2 py-0.5 rounded-full ${u.status==='approved'?'bg-emerald-950 text-emerald-300 border border-emerald-800':'bg-red-950 text-red-300 border border-red-800'}">${u.status}</span></td>
                <td class="space-x-1">
                  <button onclick="showEditUserModal('${esc(u.username)}')" class="text-xs text-indigo-400 hover:text-indigo-300 px-2 py-1">Edit</button>
                  <button onclick="showResetPassModal('${esc(u.username)}')" class="text-xs text-amber-400 hover:text-amber-300 px-2 py-1">Reset PW</button>
                  <button onclick="deleteUser('${esc(u.username)}')" class="text-xs text-red-400 hover:text-red-300 px-2 py-1 ${u.role==='master'?'opacity-20 pointer-events-none':''}">Delete</button>
                </td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    </div>`;
}

function showCreateUserModal() {
  api('/api/spro/overview').then(data => {
    openDialog(`
      <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl">
        <h3 class="font-bold text-slate-800 text-sm">Create New Account</h3>
        <button onclick="closeDialog()" class="text-slate-400 hover:text-slate-600"><i class="ri-close-line text-lg"></i></button>
      </div>
      <div class="p-6 space-y-4">
        <div class="grid grid-cols-2 gap-4">
          <div><label class="block text-xs font-semibold text-slate-600 mb-1">Username</label><input id="nu_u" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs"></div>
          <div><label class="block text-xs font-semibold text-slate-600 mb-1">Initial Password</label><input id="nu_p" type="password" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs"></div>
        </div>
        <div><label class="block text-xs font-semibold text-slate-600 mb-1">Full Name</label><input id="nu_fn" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs"></div>
        <div class="grid grid-cols-2 gap-4">
          <div>
            <label class="block text-xs font-semibold text-slate-600 mb-1">Department</label>
            <select id="nu_d" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs bg-white">
              ${data.departments.map(d=>`<option value="${esc(d.name)}">${esc(d.name)}</option>`).join('')}
            </select>
          </div>
          <div>
            <label class="block text-xs font-semibold text-slate-600 mb-1">Role Profile</label>
            <select id="nu_r" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs bg-white">
              <option value="user">Standard User</option>
              <option value="admin">Administrator</option>
              ${ME.is_master ? '<option value="master">Master Superuser</option>' : ''}
            </select>
          </div>
        </div>
        <div>
          <label class="block text-xs font-semibold text-slate-600 mb-1">Plant Access (comma-separated, leave blank for All)</label>
          <input id="nu_pl" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs font-mono" placeholder="1000, 1400">
        </div>
      </div>
      <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl">
        <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-xl text-xs font-semibold">Cancel</button>
        <button onclick="submitCreateUser()" class="btn-primary px-4 py-2 rounded-xl text-xs font-semibold">Create Account</button>
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
    setMsg("User created successfully");
  } catch(e) {
    alert(e);
  }
}

async function showEditUserModal(u) {
  const users = await api('/api/su01/users');
  const user = users.find(x => x.username === u);
  const data = await api('/api/spro/overview');
  openDialog(`
    <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl">
      <h3 class="font-bold text-slate-800 text-sm">Update User: ${esc(u)}</h3>
      <button onclick="closeDialog()"><i class="ri-close-line text-lg"></i></button>
    </div>
    <div class="p-6 space-y-4">
      <div><label class="block text-xs font-semibold text-slate-600 mb-1">Full Name</label><input id="eu_fn" value="${esc(user.full_name)}" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs"></div>
      <div class="grid grid-cols-2 gap-4">
        <div>
          <label class="block text-xs font-semibold text-slate-600 mb-1">Department</label>
          <select id="eu_d" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs bg-white">
            ${data.departments.map(d=>`<option value="${esc(d.name)}" ${d.name===user.department?'selected':''}>${esc(d.name)}</option>`).join('')}
          </select>
        </div>
        <div>
          <label class="block text-xs font-semibold text-slate-600 mb-1">Role</label>
          <select id="eu_r" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs bg-white">
            <option value="user" ${user.role==='user'?'selected':''}>Standard User</option>
            <option value="admin" ${user.role==='admin'?'selected':''}>Administrator</option>
            ${ME.is_master ? `<option value="master" ${user.role==='master'?'selected':''}>Master</option>` : ''}
          </select>
        </div>
      </div>
      <div class="grid grid-cols-2 gap-4">
        <div><label class="block text-xs font-semibold text-slate-600 mb-1">Facility Access</label><input id="eu_pl" value="${esc(user.plants||'')}" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs font-mono"></div>
        <div>
          <label class="block text-xs font-semibold text-slate-600 mb-1">Status</label>
          <select id="eu_st" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs bg-white">
            <option value="approved" ${user.status==='approved'?'selected':''}>Active</option>
            <option value="locked" ${user.status==='locked'?'selected':''}>Locked</option>
          </select>
        </div>
      </div>
    </div>
    <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl">
      <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-xl text-xs font-semibold">Cancel</button>
      <button onclick="submitEditUser('${esc(u)}')" class="btn-primary px-4 py-2 rounded-xl text-xs font-semibold">Save Changes</button>
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
    setMsg("User updated");
  } catch(e) {
    alert(e);
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
  if (!confirm(`Permanently remove account ${u}?`)) return;
  const fd = new FormData(); fd.append("username", u);
  try {
    await api('/api/su01/delete', {method: 'POST', body: fd});
    renderSU01();
    setMsg(`User ${u} deleted`);
  } catch(e) {
    alert(e);
  }
}

// SE93: MODULE REGISTRY STUDIO
async function renderSE93() {
  if (!ME.is_master && ME.role !== 'admin') return setMsg("Permission denied", true);
  const list = await api('/api/se93/list');

  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col bg-slate-900 p-6 overflow-hidden select-text">
      <div class="flex items-center justify-between pb-5 border-b border-slate-800 shrink-0">
        <div>
          <h2 class="text-lg font-bold text-white tracking-tight">Module Studio (SE93)</h2>
          <p class="text-xs text-slate-400">Design custom transaction codes, data schemas, and access tags</p>
        </div>
        <button onclick="showCreateTcodeModal()" class="btn-primary text-xs font-semibold px-4 py-2 rounded-xl flex items-center gap-1.5">
          <i class="ri-add-line"></i> Create New Module
        </button>
      </div>

      <div class="flex-1 mt-4 bg-slate-950 border border-slate-800 rounded-2xl overflow-auto">
        <table class="w-full text-left custom-table">
          <thead><tr><th>Code</th><th>Module Name</th><th>Category</th><th>Authorized Depts</th><th>Status</th><th>Actions</th></tr></thead>
          <tbody class="text-slate-300">
            ${list.map(t => `
              <tr class="border-b border-slate-800/60 hover:bg-slate-900/60">
                <td class="font-mono font-bold text-indigo-400">${esc(t.code)}</td>
                <td class="font-medium text-white">${esc(t.title)}</td>
                <td><span class="font-mono text-xs bg-slate-800 px-2 py-0.5 rounded text-slate-300">${esc(t.module||'ERP')}</span></td>
                <td>${esc(t.allowed_depts)}</td>
                <td><span class="text-[10px] font-semibold px-2 py-0.5 rounded-full ${t.active?'bg-emerald-950 text-emerald-300':'bg-red-950 text-red-300'}">${t.active?'ACTIVE':'DISABLED'}</span></td>
                <td class="space-x-1">
                  <button onclick="openTcode('${esc(t.code)}')" class="text-xs text-emerald-400 hover:text-emerald-300 px-2 py-1">Launch</button>
                  <button onclick="showEditTcodeModal('${esc(t.code)}')" class="text-xs text-indigo-400 hover:text-indigo-300 px-2 py-1">Edit</button>
                  <button onclick="deleteTcode('${esc(t.code)}')" class="text-xs text-red-400 hover:text-red-300 px-2 py-1">Delete</button>
                </td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    </div>`;
}

function showCreateTcodeModal() {
  openDialog(`
    <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl">
      <h3 class="font-bold text-slate-800 text-sm">Register Operations Module</h3>
      <button onclick="closeDialog()"><i class="ri-close-line text-lg"></i></button>
    </div>
    <div class="p-6 space-y-4">
      <div class="grid grid-cols-2 gap-4">
        <div><label class="block text-xs font-semibold text-slate-600 mb-1">Code</label><input id="nt_c" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs font-mono uppercase" placeholder="e.g. ZCORE"></div>
        <div><label class="block text-xs font-semibold text-slate-600 mb-1">Category Code</label><input id="nt_m" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs uppercase" value="PROD"></div>
      </div>
      <div><label class="block text-xs font-semibold text-slate-600 mb-1">Module Title</label><input id="nt_t" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs" placeholder="e.g. Cable Core Processing"></div>
      <div><label class="block text-xs font-semibold text-slate-600 mb-1">Operational Description</label><textarea id="nt_d" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs" rows="2"></textarea></div>
      <div><label class="block text-xs font-semibold text-slate-600 mb-1">Permitted Departments (comma-separated)</label><input id="nt_ad" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs" value="Production,Quality Control"></div>
    </div>
    <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl">
      <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-xl text-xs font-semibold">Cancel</button>
      <button onclick="submitCreateTcode()" class="btn-primary px-4 py-2 rounded-xl text-xs font-semibold">Register Module</button>
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
    loadModules();
    setMsg("Module registered");
  } catch(e) {
    alert(e);
  }
}

async function showEditTcodeModal(code) {
  const list = await api('/api/se93/list');
  const t = list.find(x => x.code === code);
  openDialog(`
    <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl">
      <h3 class="font-bold text-slate-800 text-sm">Edit Module: ${esc(code)}</h3>
      <button onclick="closeDialog()"><i class="ri-close-line text-lg"></i></button>
    </div>
    <div class="p-6 space-y-4">
      <div><label class="block text-xs font-semibold text-slate-600 mb-1">Title</label><input id="et_t" value="${esc(t.title)}" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs"></div>
      <div class="grid grid-cols-2 gap-4">
        <div><label class="block text-xs font-semibold text-slate-600 mb-1">Category</label><input id="et_m" value="${esc(t.module||'ERP')}" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs"></div>
        <div>
          <label class="block text-xs font-semibold text-slate-600 mb-1">Status</label>
          <select id="et_st" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs bg-white">
            <option value="1" ${t.active?'selected':''}>Active</option>
            <option value="0" ${!t.active?'selected':''}>Disabled</option>
          </select>
        </div>
      </div>
      <div><label class="block text-xs font-semibold text-slate-600 mb-1">Permitted Departments</label><input id="et_ad" value="${esc(t.allowed_depts)}" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs"></div>
      <div><label class="block text-xs font-semibold text-slate-600 mb-1">Description</label><textarea id="et_d" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs">${esc(t.description||'')}</textarea></div>
    </div>
    <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl">
      <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-xl text-xs font-semibold">Cancel</button>
      <button onclick="submitEditTcode('${esc(code)}')" class="btn-primary px-4 py-2 rounded-xl text-xs font-semibold">Save Changes</button>
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
    setMsg("Module updated");
  } catch(e) {
    alert(e);
  }
}

async function deleteTcode(code) {
  if (!confirm(`Warning: Deleting ${code} permanently wipes all data entries and layouts. Continue?`)) return;
  const fd = new FormData(); fd.append("code", code);
  try {
    await api('/api/se93/delete', {method: 'POST', body: fd});
    renderSE93();
    loadModules();
    setMsg(`Module ${code} removed`);
  } catch(e) {
    alert(e);
  }
}

// SPRO: ENTERPRISE SETUP
async function renderSPRO() {
  if (!ME.is_master) return setMsg("Restricted to Master Administrators", true);
  const data = await api('/api/spro/overview');

  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col bg-slate-900 p-6 overflow-hidden select-text">
      <div class="pb-5 border-b border-slate-800 shrink-0">
        <h2 class="text-lg font-bold text-white tracking-tight">Enterprise Organization Units (SPRO)</h2>
        <p class="text-xs text-slate-400">Configure production plants, facilities, and departments</p>
      </div>

      <div class="grid grid-cols-1 md:grid-cols-2 gap-6 flex-1 mt-4 overflow-auto">
        <!-- PLANTS -->
        <div class="bg-slate-950 border border-slate-800 rounded-2xl flex flex-col overflow-hidden">
          <div class="p-4 border-b border-slate-800 flex justify-between items-center bg-slate-900/60">
            <span class="text-xs font-bold text-white uppercase tracking-wider">Manufacturing Plants</span>
            <button onclick="createPlantModal()" class="btn-primary text-xs font-semibold px-3 py-1.5 rounded-lg">+ Add Facility</button>
          </div>
          <div class="flex-1 overflow-auto">
            <table class="w-full text-left custom-table">
              <thead><tr><th>Code</th><th>Facility Name</th><th>Status</th><th>Action</th></tr></thead>
              <tbody class="text-slate-300">
                ${data.plants.map(p => `
                  <tr class="border-b border-slate-800/60">
                    <td class="font-mono font-bold text-indigo-400">${esc(p.code)}</td>
                    <td class="font-medium text-white">${esc(p.name)}</td>
                    <td><span class="text-[10px] font-semibold px-2 py-0.5 rounded-full ${p.active?'bg-emerald-950 text-emerald-300':'bg-red-950 text-red-300'}">${p.active?'ACTIVE':'DISABLED'}</span></td>
                    <td><button onclick="deletePlant('${esc(p.code)}')" class="text-xs text-red-400 hover:text-red-300">Remove</button></td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        </div>

        <!-- DEPARTMENTS -->
        <div class="bg-slate-950 border border-slate-800 rounded-2xl flex flex-col overflow-hidden">
          <div class="p-4 border-b border-slate-800 flex justify-between items-center bg-slate-900/60">
            <span class="text-xs font-bold text-white uppercase tracking-wider">Operational Departments</span>
            <button onclick="createDeptModal()" class="btn-primary text-xs font-semibold px-3 py-1.5 rounded-lg">+ Add Dept</button>
          </div>
          <div class="flex-1 overflow-auto">
            <table class="w-full text-left custom-table">
              <thead><tr><th>Department</th><th>Status</th><th>Action</th></tr></thead>
              <tbody class="text-slate-300">
                ${data.departments.map(d => `
                  <tr class="border-b border-slate-800/60">
                    <td class="font-medium text-white">${esc(d.name)}</td>
                    <td><span class="text-[10px] font-semibold px-2 py-0.5 rounded-full ${d.active?'bg-emerald-950 text-emerald-300':'bg-red-950 text-red-300'}">${d.active?'ACTIVE':'DISABLED'}</span></td>
                    <td><button onclick="deleteDept('${esc(d.name)}')" class="text-xs text-red-400 hover:text-red-300">Remove</button></td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>`;
}

function createPlantModal() {
  const c = prompt("Enter Facility Code (e.g. 1500):");
  if (!c) return;
  const n = prompt("Enter Facility Description:");
  if (!n) return;
  const fd = new FormData(); fd.append("code", c); fd.append("name", n);
  api('/api/admin/plants/create', {method: 'POST', body: fd}).then(() => { renderSPRO(); setMsg("Facility added"); }).catch(e => alert(e));
}

function deletePlant(code) {
  if (!confirm(`Delete facility ${code}?`)) return;
  const fd = new FormData(); fd.append("code", code);
  api('/api/admin/plants/delete', {method: 'POST', body: fd}).then(() => { renderSPRO(); setMsg("Facility removed"); }).catch(e => alert(e));
}

function createDeptModal() {
  const d = prompt("Enter Department Name:");
  if (!d) return;
  const fd = new FormData(); fd.append("name", d);
  api('/api/admin/departments/create', {method: 'POST', body: fd}).then(() => { renderSPRO(); setMsg("Department added"); }).catch(e => alert(e));
}

function deleteDept(name) {
  if (!confirm(`Delete department ${name}?`)) return;
  const fd = new FormData(); fd.append("name", name);
  api('/api/admin/departments/delete', {method: 'POST', body: fd}).then(() => { renderSPRO(); setMsg("Department removed"); }).catch(e => alert(e));
}

// PFCG: ACCESS POLICIES
async function renderPFCG() {
  if (!ME.is_master && ME.role !== 'admin') return setMsg("Permission denied", true);
  const m = await api('/api/pfcg/matrix');

  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col bg-slate-900 p-6 overflow-hidden select-text">
      <div class="flex items-center justify-between pb-5 border-b border-slate-800 shrink-0">
        <div>
          <h2 class="text-lg font-bold text-white tracking-tight">Access Control & Permissions (PFCG)</h2>
          <p class="text-xs text-slate-400">Explicit module clearances granted to specific operators</p>
        </div>
        <div class="flex items-center gap-2">
          <select id="pfcg_u" class="bg-slate-800 border border-slate-700 text-xs rounded-xl px-3 py-2 text-white">
            ${m.users.map(u=>`<option value="${esc(u.username)}">${esc(u.username)} (${esc(u.full_name)})</option>`).join('')}
          </select>
          <select id="pfcg_t" class="bg-slate-800 border border-slate-700 text-xs rounded-xl px-3 py-2 text-white">
            ${m.tcodes.map(t=>`<option value="${esc(t.code)}">${esc(t.code)} - ${esc(t.title)}</option>`).join('')}
          </select>
          <button onclick="grantAccess()" class="btn-primary text-xs font-semibold px-3 py-2 rounded-xl">+ Grant</button>
        </div>
      </div>

      <div class="flex-1 mt-4 bg-slate-950 border border-slate-800 rounded-2xl overflow-auto">
        <table class="w-full text-left custom-table">
          <thead><tr><th>User</th><th>Permitted Module</th><th>Approved By</th><th>Granted At</th><th>Action</th></tr></thead>
          <tbody class="text-slate-300">
            ${m.grants.map(g => `
              <tr class="border-b border-slate-800/60">
                <td class="font-mono font-bold text-indigo-400">${esc(g.username)}</td>
                <td class="font-mono font-bold text-white">${esc(g.tcode_code)}</td>
                <td>${esc(g.granted_by)}</td>
                <td class="text-slate-400 text-xs">${esc(g.granted_at)}</td>
                <td><button onclick="revokeAccess('${esc(g.username)}','${esc(g.tcode_code)}')" class="text-xs text-red-400 hover:text-red-300">Revoke</button></td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    </div>`;
}

async function grantAccess() {
  const u = $('#pfcg_u').value, t = $('#pfcg_t').value;
  const fd = new FormData(); fd.append("username", u); fd.append("tcode", t);
  await api('/api/pfcg/grant', {method: 'POST', body: fd});
  renderPFCG();
  setMsg("Permission granted");
}

async function revokeAccess(u, t) {
  if (!confirm(`Revoke ${t} clearance from ${u}?`)) return;
  const fd = new FormData(); fd.append("username", u); fd.append("tcode", t);
  await api('/api/pfcg/revoke', {method: 'POST', body: fd});
  renderPFCG();
  setMsg("Permission revoked");
}

// BUSINESS MODULE VIEW & DATA GRID
async function renderBusinessModule(tc) {
  ACTIVE_LAYOUTS = await api(`/api/tcode/${tc.code}/layouts`);
  CURRENT_LAYOUT = ACTIVE_LAYOUTS.find(l => l.is_default) || ACTIVE_LAYOUTS[0];

  const sproData = await api('/api/spro/overview');
  const allowedPlants = sproData.plants.filter(p => !ME.plants.length || ME.plants.includes(p.code));

  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col bg-slate-900 overflow-hidden select-text">
      <!-- WORKSPACE TOOLBAR -->
      <div class="p-4 border-b border-slate-800 flex items-center justify-between shrink-0">
        <div class="flex items-center gap-3">
          <div>
            <div class="flex items-center gap-2">
              <span class="font-mono font-bold text-xs bg-indigo-500/10 text-indigo-400 border border-indigo-500/30 px-2 py-0.5 rounded-lg">${tc.code}</span>
              <h2 class="text-base font-bold text-white">${esc(tc.title)}</h2>
            </div>
          </div>
          <div class="h-5 w-[1px] bg-slate-800 mx-2"></div>
          <button onclick="showRecordModal()" class="btn-primary text-xs font-semibold px-3 py-1.5 rounded-xl flex items-center gap-1.5 shadow-md shadow-indigo-600/30">
            <i class="ri-add-line"></i> New Entry
          </button>
        </div>

        <div class="flex items-center gap-2">
          <!-- PLANT SELECTOR -->
          <select id="rec_plant_filter" onchange="reloadRecords()" class="bg-slate-800 border border-slate-700 text-xs rounded-xl px-2.5 py-1.5 text-slate-200">
            <option value="">(All Facilities)</option>
            ${allowedPlants.map(p=>`<option value="${esc(p.code)}">${esc(p.code)} - ${esc(p.name)}</option>`).join('')}
          </select>

          <!-- VIEW SELECTOR -->
          <select id="view_sel" onchange="switchLayout(this.value)" class="bg-slate-800 border border-slate-700 text-xs rounded-xl px-2.5 py-1.5 text-slate-200">
            ${ACTIVE_LAYOUTS.map(l => `<option value="${esc(l.layout_code)}" ${CURRENT_LAYOUT && l.layout_code===CURRENT_LAYOUT.layout_code?'selected':''}>${esc(l.name)}</option>`).join('')}
          </select>
          <button onclick="showLayoutDesigner()" title="Layout Studio" class="bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs font-semibold px-2.5 py-1.5 border border-slate-700 rounded-xl">
            <i class="ri-layout-masonry-line"></i>
          </button>

          <a href="/api/tcode/${tc.code}/export" class="bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs font-semibold px-2.5 py-1.5 border border-slate-700 rounded-xl flex items-center gap-1">
            <i class="ri-download-2-line"></i> Export
          </a>
          <label class="bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs font-semibold px-2.5 py-1.5 border border-slate-700 rounded-xl flex items-center gap-1 cursor-pointer">
            <i class="ri-upload-2-line"></i> Import
            <input type="file" class="hidden" onchange="importSpreadsheet(event)">
          </label>
          <button onclick="reloadRecords()" class="bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs font-semibold p-2 border border-slate-700 rounded-xl">
            <i class="ri-refresh-line"></i>
          </button>
        </div>
      </div>

      <!-- DATA TABLE -->
      <div id="data_grid_container" class="flex-1 bg-slate-950 border-t border-slate-800 overflow-auto"></div>
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
    <table class="w-full text-left custom-table">
      <thead>
        <tr>
          <th style="width:60px;">Doc #</th>
          <th style="width:90px;">Facility</th>
          ${fields.map(f => `<th>${esc(f.label)}</th>`).join('')}
          <th style="width:130px;">Logged By</th>
          <th style="width:140px;">Timestamp</th>
          <th style="width:100px;">Actions</th>
        </tr>
      </thead>
      <tbody class="text-slate-300">`;

  if (!recs.length) {
    html += `<tr><td colspan="${fields.length + 5}" class="text-center py-12 text-slate-500">No active records found for this unit.</td></tr>`;
  } else {
    recs.forEach(r => {
      html += `
        <tr class="border-b border-slate-800/60 hover:bg-slate-900/60">
          <td class="font-mono text-slate-500 font-semibold">${r.id}</td>
          <td><span class="font-mono text-xs bg-slate-800 border border-slate-700 px-2 py-0.5 rounded text-indigo-300">${esc(r.plant || 'CORP')}</span></td>
          ${fields.map(f => `<td>${esc(r.data[f.key] || '')}</td>`).join('')}
          <td class="text-xs font-medium text-slate-400">${esc(r.created_by_name || r.created_by)}</td>
          <td class="text-xs text-slate-500 font-mono">${esc(r.created_at)}</td>
          <td class="space-x-1">
            <button onclick="editRecord(${r.id})" class="text-xs text-indigo-400 hover:text-indigo-300">Edit</button>
            <button onclick="deleteRecord(${r.id})" class="text-xs text-red-400 hover:text-red-300">Del</button>
          </td>
        </tr>`;
    });
  }

  html += `</tbody></table>`;
  $('#data_grid_container').innerHTML = html;
  setMsg(`${recs.length} records retrieved`);
}

function showRecordModal(editId = null, existingData = {}, existingPlant = '') {
  const fields = CURRENT_LAYOUT ? CURRENT_LAYOUT.fields : [];
  openDialog(`
    <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl">
      <h3 class="font-bold text-slate-800 text-sm">${editId ? 'Edit Entry #' + editId : 'Post Entry: ' + CURRENT_TCODE}</h3>
      <button onclick="closeDialog()"><i class="ri-close-line text-lg"></i></button>
    </div>
    <div class="p-6 space-y-4">
      <div>
        <label class="block text-xs font-semibold text-slate-600 mb-1">Facility Code</label>
        <input id="rec_f__plant" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs font-mono" value="${esc(existingPlant || (ME.plants[0]||'1000'))}">
      </div>
      <div class="grid grid-cols-2 gap-4">
        ${fields.map(f => `
          <div class="col-span-${f.width||1}">
            <label class="block text-xs font-semibold text-slate-600 mb-1">${esc(f.label)} ${f.required ? '<span class="text-red-500">*</span>':''}</label>
            ${f.type==='textarea' ?
              `<textarea id="rec_f_${f.key}" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs" rows="2">${esc(existingData[f.key]||'')}</textarea>` :
              `<input type="${f.type==='number'?'number':(f.type==='date'?'date':'text')}" id="rec_f_${f.key}" value="${esc(existingData[f.key]||'')}" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs">`
            }
          </div>
        `).join('')}
      </div>
    </div>
    <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl">
      <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-xl text-xs font-semibold">Cancel</button>
      <button onclick="submitRecord(${editId})" class="btn-primary px-4 py-2 rounded-xl text-xs font-semibold">Save Entry</button>
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
      alert(`Required: ${f.label}`);
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
    setMsg("Entry saved");
  } catch(e) {
    alert(e);
  }
}

async function deleteRecord(rid) {
  if (!confirm(`Delete entry #${rid}?`)) return;
  try {
    await api(`/api/tcode/${CURRENT_TCODE}/records/${rid}`, {method: 'DELETE'});
    reloadRecords();
    setMsg(`Record #${rid} deleted`);
  } catch(e) {
    alert(e);
  }
}

async function importSpreadsheet(e) {
  const file = e.target.files[0];
  if (!file) return;
  const fd = new FormData(); fd.append("file", file);
  setMsg("Processing spreadsheet...");
  try {
    const res = await api(`/api/tcode/${CURRENT_TCODE}/import`, {method: 'POST', body: fd});
    reloadRecords();
    setMsg(`Imported ${res.count} records successfully`);
  } catch(err) {
    alert(err);
  }
}

// LAYOUT STUDIO
function showLayoutDesigner() {
  const fields = CURRENT_LAYOUT ? CURRENT_LAYOUT.fields : [];
  openDialog(`
    <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl">
      <h3 class="font-bold text-slate-800 text-sm">View Studio: ${CURRENT_TCODE}</h3>
      <button onclick="closeDialog()"><i class="ri-close-line text-lg"></i></button>
    </div>
    <div class="p-6 space-y-4">
      <div class="grid grid-cols-2 gap-4">
        <div><label class="block text-xs font-semibold text-slate-600 mb-1">View Name</label><input id="lay_name" value="${esc(CURRENT_LAYOUT.name)}" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs"></div>
        <div><label class="block text-xs font-semibold text-slate-600 mb-1">Identifier</label><input id="lay_code" value="${esc(CURRENT_LAYOUT.layout_code)}" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs font-mono uppercase"></div>
      </div>
      <div class="border border-slate-200 p-4 rounded-xl bg-slate-50/50">
        <div class="flex justify-between items-center mb-3">
          <span class="font-bold text-xs text-slate-700">Columns Schema</span>
          <button onclick="addSchemaRow()" class="btn-primary text-xs font-semibold px-3 py-1.5 rounded-lg">+ Add Column</button>
        </div>
        <div id="schema_rows" class="space-y-2 max-h-56 overflow-y-auto">
          ${fields.map(f => `
            <div class="flex items-center gap-2 schema-row bg-white p-2 border border-slate-200 rounded-xl">
              <input placeholder="key" value="${esc(f.key)}" class="w-28 border border-slate-200 rounded-lg p-1.5 text-xs font-mono skey">
              <input placeholder="Column Label" value="${esc(f.label)}" class="flex-1 border border-slate-200 rounded-lg p-1.5 text-xs slbl">
              <select class="border border-slate-200 rounded-lg p-1.5 text-xs bg-white stype">
                <option value="text" ${f.type==='text'?'selected':''}>Text</option>
                <option value="number" ${f.type==='number'?'selected':''}>Number</option>
                <option value="date" ${f.type==='date'?'selected':''}>Date</option>
                <option value="textarea" ${f.type==='textarea'?'selected':''}>Textarea</option>
              </select>
              <button onclick="this.parentElement.remove()" class="text-red-500 hover:text-red-700 px-1"><i class="ri-delete-bin-line"></i></button>
            </div>
          `).join('')}
        </div>
      </div>
    </div>
    <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-between items-center rounded-b-2xl">
      <div>
        ${ME.is_master ? `<button onclick="makeLayoutDefault(${CURRENT_LAYOUT.id})" class="btn-secondary text-xs px-3 py-2 rounded-xl">Set As Global Default</button>` : ''}
      </div>
      <div class="flex gap-2">
        <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-xl text-xs font-semibold">Cancel</button>
        <button onclick="submitLayoutSave()" class="btn-primary px-4 py-2 rounded-xl text-xs font-semibold">Save Layout</button>
      </div>
    </div>`);
}

function addSchemaRow() {
  const div = document.createElement('div');
  div.className = 'flex items-center gap-2 schema-row bg-white p-2 border border-slate-200 rounded-xl';
  div.innerHTML = `
    <input placeholder="key" class="w-28 border border-slate-200 rounded-lg p-1.5 text-xs font-mono skey">
    <input placeholder="Column Label" class="flex-1 border border-slate-200 rounded-lg p-1.5 text-xs slbl">
    <select class="border border-slate-200 rounded-lg p-1.5 text-xs bg-white stype">
      <option value="text">Text</option>
      <option value="number">Number</option>
      <option value="date">Date</option>
      <option value="textarea">Textarea</option>
    </select>
    <button onclick="this.parentElement.remove()" class="text-red-500 hover:text-red-700 px-1"><i class="ri-delete-bin-line"></i></button>
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
    renderBusinessModule(ALL_TCODES.find(x => x.code === CURRENT_TCODE));
    setMsg("View updated");
  } catch(e) {
    alert(e);
  }
}

async function makeLayoutDefault(lid) {
  await api(`/api/tcode/${CURRENT_TCODE}/layouts/${lid}/set-default`, {method: 'POST'});
  closeDialog();
  renderBusinessModule(ALL_TCODES.find(x => x.code === CURRENT_TCODE));
  setMsg("Global default view updated");
}

function openDialog(content) {
  $('#app_dialog_body').innerHTML = content;
  $('#app_dialog').classList.remove('hidden');
}

function closeDialog() {
  $('#app_dialog').classList.add('hidden');
  $('#app_dialog_body').innerHTML = '';
}

// STARTUP
initSession();
</script>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
def home():
    return HTMLResponse(HTML_MODERN_UI)

@app.get("/health")
def health():
    return {"status": "ACTIVE", "system": "NEXUS-CORE", "version": "3.2.0"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
