import io, json, sqlite3, secrets, datetime, os, re, csv, math
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

def json_serial(obj):
    if isinstance(obj, (datetime.datetime, datetime.date)):
        return obj.strftime("%Y-%m-%d %H:%M:%S") if isinstance(obj, datetime.datetime) else obj.strftime("%Y-%m-%d")
    return str(obj)

def infer_field_type(sample_values):
    if not sample_values:
        return "text"
    is_num, is_date, is_long = True, True, False
    date_patterns = [r'^\d{4}-\d{2}-\d{2}', r'^\d{2}/\d{2}/\d{4}', r'^\d{2}-\d{2}-\d{4}']
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
        if len(s) > 70 or '\n' in s:
            is_long, is_num, is_date = True, False, False
            break
        if len(s) > 1 and s[0] == '0' and s.isdigit():
            is_num = False
        else:
            try:
                float(s.replace(',', ''))
            except ValueError:
                is_num = False
        if not any(re.match(p, s) for p in date_patterns):
            is_date = False
    return "date" if is_date else ("number" if is_num else ("textarea" if is_long else "text"))

# --- INITIALIZATION & AUTO-MIGRATION ---
def init_db():
    conn = get_conn()
    cur = conn.cursor()

    if USE_PG:
        try:
            cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name = 'tcodes' AND column_name = 'color';")
            if cur.fetchone():
                cur.execute("DROP TABLE IF EXISTS tcodes CASCADE; DROP TABLE IF EXISTS tcode_schemas CASCADE; DROP TABLE IF EXISTS tcode_layouts CASCADE;")

            cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name = 'users' AND column_name = 'sub_department';")
            if cur.fetchone():
                cur.execute("DROP TABLE IF EXISTS users CASCADE;")

            cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name = 'plants' AND column_name = 'code';")
            if not cur.fetchone():
                cur.execute("DROP TABLE IF EXISTS plants CASCADE;")

            cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name = 'org' AND column_name = 'system_id';")
            if not cur.fetchone():
                cur.execute("DROP TABLE IF EXISTS org CASCADE;")

            cur.execute("SELECT column_name FROM information_schema.columns WHERE table_name = 'tcode_records' AND column_name = 'plant';")
            if not cur.fetchone():
                cur.execute("ALTER TABLE tcode_records ADD COLUMN IF NOT EXISTS plant TEXT DEFAULT '';")

            conn.commit()
        except Exception as e:
            print("Auto-migration notice:", e)
            conn.rollback()

    schema_sql = '''
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
    CREATE TABLE IF NOT EXISTS transfer_notes (
        id SERIAL PRIMARY KEY, note_no TEXT UNIQUE, layout_key TEXT, route TEXT,
        from_plant TEXT, from_dept TEXT, from_user TEXT, from_user_name TEXT,
        to_plant TEXT, to_dept TEXT, to_user TEXT, to_user_name TEXT,
        entry_count INT DEFAULT 0, sheet_headers TEXT, sheet_data TEXT,
        status TEXT DEFAULT 'sent', remarks TEXT DEFAULT '', created_at TEXT,
        received_at TEXT DEFAULT '', received_by TEXT DEFAULT '', received_by_name TEXT DEFAULT ''
    );
    CREATE TABLE IF NOT EXISTS module_specs (
        id SERIAL PRIMARY KEY, tcode_code TEXT, version INT,
        action TEXT, spec TEXT, applied_by TEXT, applied_at TEXT, note TEXT
    );
    ''' if USE_PG else '''
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
    CREATE TABLE IF NOT EXISTS transfer_notes (
        id INTEGER PRIMARY KEY AUTOINCREMENT, note_no TEXT UNIQUE, layout_key TEXT, route TEXT,
        from_plant TEXT, from_dept TEXT, from_user TEXT, from_user_name TEXT,
        to_plant TEXT, to_dept TEXT, to_user TEXT, to_user_name TEXT,
        entry_count INT DEFAULT 0, sheet_headers TEXT, sheet_data TEXT,
        status TEXT DEFAULT 'sent', remarks TEXT DEFAULT '', created_at TEXT,
        received_at TEXT DEFAULT '', received_by TEXT DEFAULT '', received_by_name TEXT DEFAULT ''
    );
    CREATE TABLE IF NOT EXISTS module_specs (
        id INTEGER PRIMARY KEY AUTOINCREMENT, tcode_code TEXT, version INT,
        action TEXT, spec TEXT, applied_by TEXT, applied_at TEXT, note TEXT
    );
    '''

    if USE_PG:
        cur.execute(schema_sql)
    else:
        cur.executescript(schema_sql)
    conn.commit()

    if USE_PG:
        try:
            cur.execute("SELECT setval(pg_get_serial_sequence('tcode_records', 'id'), COALESCE((SELECT MAX(id)+1 FROM tcode_records), 1), false);")
            cur.execute("SELECT setval(pg_get_serial_sequence('transfer_notes', 'id'), COALESCE((SELECT MAX(id)+1 FROM transfer_notes), 1), false);")
            conn.commit()
        except Exception:
            conn.rollback()

    cur_time = now()

    cur.execute(q("SELECT 1 FROM org WHERE id=1"))
    if not cur.fetchone():
        cur.execute(q("INSERT INTO org (id, name, system_id, client) VALUES (?,?,?,?)"), (1, "Thermocables Limited", "NX-01", "GLOBAL"))

    for d in ["Production", "Final Testing", "Quality Control", "Commercial", "Store", "Planning", "Engineering"]:
        cur.execute(q("SELECT 1 FROM departments WHERE name=?"), (d,))
        if not cur.fetchone():
            cur.execute(q("INSERT INTO departments (name, active, created_at) VALUES (?,?,?)"), (d, 1, cur_time))

    for p_code, p_name in [("1000", "Gooty Central Plant"), ("1400", "Extrusion Unit 1400"), ("2000", "Hyderabad Works"), ("CORP", "Corporate Headquarters")]:
        cur.execute(q("SELECT 1 FROM plants WHERE code=?"), (p_code,))
        if not cur.fetchone():
            cur.execute(q("INSERT INTO plants (code, name, active, created_at) VALUES (?,?,?,?)"), (p_code, p_name, 1, cur_time))

    ph = bcrypt.hashpw("Admin@123".encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
    cur.execute(q("SELECT username FROM users WHERE username='admin'"))
    if cur.fetchone():
        cur.execute(q("UPDATE users SET password_hash=?, status='approved', role='master' WHERE username='admin'"), (ph,))
    else:
        cur.execute(q("INSERT INTO users (username, password_hash, full_name, role, department, plants, status, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?)"),
                    ("admin", ph, "Master System Administrator", "master", "IT & Systems", "1000,1400,2000,CORP", "approved", cur_time, cur_time))

    sample_users = [
        ("prod_user", "Production Supervisor", "Production", "1000"),
        ("testing_lead", "Final Testing Incharge", "Final Testing", "1000"),
        ("comm_lead", "Commercial Manager", "Commercial", "1000"),
        ("store_head", "Store Incharge", "Store", "1000"),
        ("qc_lead", "Quality Control Incharge", "Quality Control", "1000")
    ]
    for su_name, su_full, su_dept, su_pl in sample_users:
        cur.execute(q("SELECT 1 FROM users WHERE username=?"), (su_name,))
        if not cur.fetchone():
            cur.execute(q("INSERT INTO users (username, password_hash, full_name, role, department, plants, status, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?)"),
                        (su_name, ph, su_full, "user", su_dept, su_pl, "approved", cur_time, cur_time))

    default_modules = [
        ("TN01", "Transfer Note & Handover Sheet", "Inter-department material handover & batch transfer note", "MM", "Production,Final Testing,Commercial,Store,Quality Control,Planning"),
        ("CTST", "Cable Type-Test Calculator", "CR, HV, Braid Coverage, Tubular & Dumbbell mechanical tests", "QC", "Quality Control,Final Testing,Engineering"),
        ("PROD", "Production Handover Report", "Manufacturing batch & lot transfer logs", "PROD", "Production,Quality Control"),
        ("GCH0", "Quality Green Card Release", "Incoming QA & physical cable release inspection", "QC", "Quality Control,Production"),
        ("NCR0", "Non-Conformance Incidents", "QC deviation & defect corrective tracking", "QC", "Quality Control,Final Testing,Production"),
        ("PLN0", "Production Planning Details", "Machine allocation and manufacturing order queue", "PLAN", "Planning,Production,Store"),
        ("QCH0", "Stage QC to Final Clearance", "Clearance release slips before shipment", "QC", "Quality Control,Final Testing"),
        ("STD0", "Inventory Transfer Orders", "Inter-plant and store material movements", "INV", "Store,Planning,Production")
    ]

    for code, title, desc, mod, depts in default_modules:
        cur.execute(q("SELECT 1 FROM tcodes WHERE code=?"), (code,))
        if not cur.fetchone():
            cur.execute(q("INSERT INTO tcodes (code, title, description, module, allowed_depts, active, created_at, updated_at, created_by) VALUES (?,?,?,?,?,?,?,?,?)"),
                        (code, title, desc, mod, depts, 1, cur_time, cur_time, "SYSTEM"))

    # Seed CTST schema + default layout if missing
    cur.execute(q("SELECT 1 FROM tcode_schemas WHERE code='CTST'"))
    if not cur.fetchone():
        ctst_fields = [
            {"key": "sample_id", "label": "Sample / Batch ID", "type": "text", "required": True, "width": 1},
            {"key": "cable_type", "label": "Cable Type / Size", "type": "text", "width": 1},
            {"key": "test_date", "label": "Test Date", "type": "date", "width": 1},
            {"key": "obd_cr_reading", "label": "OBD CR Reading (Ω)", "type": "number", "width": 1},
            {"key": "temp_c", "label": "Temp (°C)", "type": "number", "width": 1},
            {"key": "length_m", "label": "Length (m)", "type": "number", "width": 1},
            {"key": "specified_cr", "label": "Specified CR", "type": "number", "width": 1},
            {"key": "observed_cr", "label": "Observed CR", "type": "number", "width": 1},
            {"key": "cr_max", "label": "Obs. Max CR", "type": "number", "width": 1},
            {"key": "cr_min", "label": "Obs. Min CR", "type": "number", "width": 1},
            {"key": "hv_length", "label": "HV Test Length (m)", "type": "number", "width": 1},
            {"key": "hv_top_pos", "label": "HV Fault from Top (m)", "type": "number", "width": 1},
            {"key": "hv_bottom_pos", "label": "HV Fault from Bottom (m)", "type": "number", "width": 1},
            {"key": "braid_d", "label": "Braid Wire Ø (mm)", "type": "number", "width": 1},
            {"key": "under_braid_D", "label": "Ø Under Braid (mm)", "type": "number", "width": 1},
            {"key": "wires_per_carrier", "label": "Wires per Carrier", "type": "number", "width": 1},
            {"key": "carriers", "label": "Carriers", "type": "number", "width": 1},
            {"key": "lay_length", "label": "Lay Length (mm)", "type": "number", "width": 1},
            {"key": "tub_outer_d", "label": "Tubular Outer Ø (mm)", "type": "number", "width": 1},
            {"key": "tub_wall", "label": "Tubular Wall (mm)", "type": "number", "width": 1},
            {"key": "dumb_width", "label": "Dumbbell Width (mm)", "type": "number", "width": 1},
            {"key": "dumb_thk", "label": "Dumbbell Thickness (mm)", "type": "number", "width": 1},
            {"key": "break_load_tub", "label": "Break Load – Tubular (N)", "type": "number", "width": 1},
            {"key": "break_load_dumb", "label": "Break Load – Dumbbell (N)", "type": "number", "width": 1},
            {"key": "elong_tub_final", "label": "Tubular Gauge Final (mm)", "type": "number", "width": 1},
            {"key": "elong_dumb_final", "label": "Dumbbell Gauge Final (mm)", "type": "number", "width": 1},
            {"key": "initial_gauge", "label": "Initial Gauge (mm)", "type": "number", "width": 1},
            {"key": "hotset_elong_final", "label": "Hot Set Gauge Final (mm)", "type": "number", "width": 1},
            {"key": "permset_cooled", "label": "Perm Set Cooled Gauge (mm)", "type": "number", "width": 1},
            {"key": "result", "label": "Overall Result", "type": "select", "width": 1, "options": ["PASS", "FAIL", "RETAIN"]},
            {"key": "remarks", "label": "Remarks", "type": "textarea", "width": 2}
        ]
        cur.execute(q("INSERT INTO tcode_schemas (code, fields, layout, updated_at, updated_by) VALUES (?,?,?,?,?)"),
                    ("CTST", json.dumps(ctst_fields), json.dumps({"columns": 3}), cur_time, "SYSTEM"))
        cur.execute(q("INSERT INTO tcode_layouts (tcode_code, layout_code, name, title, is_default, fields, layout_config, description, created_by, created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)"),
                    ("CTST", "GRID_CTST_01", "Cable Type-Test Grid", "Cable Type-Test Calculator", 1, json.dumps(ctst_fields), json.dumps({"columns": 3}), "Seeded from Calculations Formulas.xlsx", "SYSTEM", "System", cur_time, cur_time))

    conn.commit()
    conn.close()

init_db()

# --- AUTH HELPERS ---
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
    return u if (u and u["status"] == "approved") else None

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
        raise HTTPException(403, "Master Administrator clearance required.")
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
    return bool(fetch_one("SELECT 1 FROM access_grants WHERE username=? AND tcode_code=?", (u["username"], code)))

# --- FASTAPI APP ---
app = FastAPI(title="Nexus Core ERP", version="3.7.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True, allow_methods=["*"], allow_headers=["*"])

# --- AUTH APIS ---
@app.post("/api/login")
def api_login(res: Response, username: str = Form(...), password: str = Form(...)):
    username = username.lower().strip()
    u = fetch_one("SELECT * FROM users WHERE username=?", (username,))
    if not u or not bcrypt.checkpw(password.encode('utf-8'), u["password_hash"].encode('utf-8')):
        raise HTTPException(401, "Invalid username or password.")
    if u["status"] != "approved":
        raise HTTPException(403, "Account is locked or pending approval.")

    token = secrets.token_hex(32)
    exp = (datetime.datetime.now() + datetime.timedelta(days=7)).strftime("%Y-%m-%d %H:%M:%S")
    exec_write("INSERT INTO sessions (token, username, expires) VALUES (?,?,?)", (token, username, exp))

    res.set_cookie("nexus_session", token, httponly=True, samesite="lax", max_age=86400 * 7)
    return {"ok": True, "token": token, "user": {"username": u["username"], "full_name": u["full_name"], "role": u["role"]}}

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
        "username": u["username"], "full_name": u["full_name"], "role": u["role"],
        "department": u["department"], "plants": [p for p in (u["plants"] or "").split(",") if p],
        "is_master": is_master(u), "system": org_info.get("system_id", "NX-01"),
        "client": org_info.get("client", "GLOBAL"), "company": org_info.get("name", "Thermocables Limited")
    }

# --- SYSTEM UNITS, PLANTS & DEPTS (SPRO) ---
@app.get("/api/spro/overview")
def spro_overview(req: Request):
    need(req)
    return {
        "org": fetch_one("SELECT * FROM org WHERE id=1"),
        "departments": fetch_all("SELECT * FROM departments ORDER BY name"),
        "plants": fetch_all("SELECT * FROM plants ORDER BY code")
    }

@app.post("/api/admin/org/update")
def spro_update_org(req: Request, name: str = Form(...), system_id: str = Form(...), client: str = Form(...)):
    need_master(req)
    exec_write("UPDATE org SET name=?, system_id=?, client=? WHERE id=1", (name.strip(), system_id.strip(), client.strip()))
    return {"ok": True}

@app.post("/api/admin/plants/create")
def spro_create_plant(req: Request, code: str = Form(...), name: str = Form(...)):
    need_master(req)
    code = code.upper().strip()
    if fetch_one("SELECT 1 FROM plants WHERE code=?", (code,)):
        raise HTTPException(400, f"Plant code '{code}' already exists.")
    exec_write("INSERT INTO plants (code, name, active, created_at) VALUES (?,?,?,?)", (code, name.strip(), 1, now()))
    return {"ok": True}

@app.post("/api/admin/plants/update")
def spro_update_plant(req: Request, code: str = Form(...), name: str = Form(...), active: int = Form(1)):
    need_master(req)
    code = code.upper().strip()
    exec_write("UPDATE plants SET name=?, active=? WHERE code=?", (name.strip(), active, code))
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
    exec_write("INSERT INTO departments (name, active, created_at) VALUES (?,?,?)", (name, 1, now()))
    return {"ok": True}

@app.post("/api/admin/departments/delete")
def spro_delete_dept(req: Request, name: str = Form(...)):
    need_master(req)
    exec_write("DELETE FROM departments WHERE name=?", (name.strip(),))
    return {"ok": True}

# --- USER IDENTITY & IAM (SU01) ---
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
    exec_write("INSERT INTO users (username, password_hash, full_name, role, department, plants, status, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?)",
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

# --- ACCESS CONTROL (PFCG) ---
@app.get("/api/pfcg/matrix")
def pfcg_matrix(req: Request):
    need_admin(req)
    return {
        "users": fetch_all("SELECT username, full_name, role, department FROM users ORDER BY username"),
        "tcodes": fetch_all("SELECT code, title, module, allowed_depts FROM tcodes WHERE active=1 ORDER BY code"),
        "grants": fetch_all("SELECT * FROM access_grants")
    }

@app.post("/api/pfcg/grant")
def pfcg_grant(req: Request, username: str = Form(...), tcode: str = Form(...)):
    caller = need_admin(req)
    u_val, tc_val = username.lower().strip(), tcode.upper().strip()
    if USE_PG:
        exec_write("INSERT INTO access_grants (username, tcode_code, granted_by, granted_at) VALUES (?,?,?,?) ON CONFLICT (username, tcode_code) DO NOTHING",
                   (u_val, tc_val, caller["username"], now()))
    else:
        exec_write("INSERT OR IGNORE INTO access_grants (username, tcode_code, granted_by, granted_at) VALUES (?,?,?,?)",
                   (u_val, tc_val, caller["username"], now()))
    return {"ok": True}

@app.post("/api/pfcg/revoke")
def pfcg_revoke(req: Request, username: str = Form(...), tcode: str = Form(...)):
    need_admin(req)
    exec_write("DELETE FROM access_grants WHERE username=? AND tcode_code=?", (username.lower().strip(), tcode.upper().strip()))
    return {"ok": True}

# --- SE93 AUTO T-CODE CREATION ---
@app.get("/api/se93/list")
def se93_list(req: Request):
    u = need(req)
    all_t = fetch_all("SELECT * FROM tcodes ORDER BY code")
    for t in all_t:
        t["layout_count"] = fetch_one("SELECT COUNT(*) as c FROM tcode_layouts WHERE tcode_code=?", (t["code"],))["c"]
        t["record_count"] = fetch_one("SELECT COUNT(*) as c FROM tcode_records WHERE tcode_code=? AND status='active'", (t["code"],))["c"]
        t["has_access"] = check_tcode_access(u, t["code"])
    return all_t

@app.post("/api/se93/auto-create")
async def se93_auto_create(
    req: Request, code: str = Form(...), title: str = Form(...), description: str = Form(""),
    module: str = Form("PROD"), allowed_depts: str = Form("Production,Quality Control"),
    import_sample_data: int = Form(1), file: UploadFile = File(...)
):
    u = need_admin(req)
    code = code.upper().strip()
    if not code:
        raise HTTPException(400, "T-Code identifier is required.")
    if fetch_one("SELECT 1 FROM tcodes WHERE code=?", (code,)):
        raise HTTPException(400, f"Module '{code}' already exists.")

    contents = await file.read()
    filename = file.filename.lower()
    headers, data_rows = [], []

    try:
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
            wb = load_workbook(io.BytesIO(contents), data_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
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
            if isinstance(parsed, list) and len(parsed) > 0 and isinstance(parsed[0], dict):
                headers = list(parsed[0].keys())
                for item in parsed:
                    data_rows.append([item.get(k, "") for k in headers])
        else:
            raise HTTPException(400, "Please upload an Excel (.xlsx/.xls), CSV, or JSON template.")
    except Exception as e:
        raise HTTPException(400, f"Template error: {str(e)}")

    if not headers:
        raise HTTPException(400, "No valid column headers found.")

    col_samples = {idx: [] for idx in range(len(headers))}
    for r in data_rows[:25]:
        for c_idx in range(len(headers)):
            if c_idx < len(r) and r[c_idx] is not None:
                col_samples[c_idx].append(r[c_idx])

    fields = []
    for idx, h in enumerate(headers):
        clean_key = re.sub(r'[^a-zA-Z0-9_]', '_', h.lower().strip())
        clean_key = re.sub(r'_+', '_', clean_key).strip('_') or f"col_{idx+1}"
        f_type = infer_field_type(col_samples[idx])
        fields.append({"key": clean_key, "label": h, "type": f_type, "required": False, "width": 2 if f_type == "textarea" else 1})

    cur_time = now()
    exec_write("INSERT INTO tcodes (code, title, description, module, allowed_depts, active, created_at, updated_at, created_by) VALUES (?,?,?,?,?,?,?,?,?)",
               (code, title.strip() or f"{code} Module", description.strip(), module.strip(), allowed_depts.strip(), 1, cur_time, cur_time, u["username"]))

    fields_json = json.dumps(fields)
    exec_write("INSERT INTO tcode_schemas (code, fields, layout, updated_at, updated_by) VALUES (?,?,?,?,?)",
               (code, fields_json, json.dumps({"columns": 3}), cur_time, u["username"]))

    exec_write("INSERT INTO tcode_layouts (tcode_code, layout_code, name, title, is_default, fields, layout_config, description, created_by, created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
               (code, f"GRID_{code}_01", "Default Grid", f"{title or code} Data View", 1, fields_json, json.dumps({"columns": 3}), "Auto-generated from file", u["username"], u["full_name"], cur_time, cur_time))

    imported_count = 0
    if import_sample_data and data_rows:
        conn = get_conn()
        try:
            cur = conn.cursor()
            insert_sql = q("INSERT INTO tcode_records (tcode_code, plant, data, status, created_by, created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?)")
            def_plant = (u.get("plants") or "1000").split(",")[0].strip() or "1000"

            for r in data_rows:
                if not any(r):
                    continue
                rec = {}
                row_plant = def_plant
                for idx, h in enumerate(headers):
                    if idx < len(r):
                        val = r[idx]
                        if val is None: val = ""
                        elif isinstance(val, (datetime.datetime, datetime.date)): val = val.strftime("%Y-%m-%d %H:%M:%S") if isinstance(val, datetime.datetime) else val.strftime("%Y-%m-%d")
                        key = fields[idx]["key"]
                        if key in ["plant", "facility"]: row_plant = str(val or def_plant).strip()
                        else: rec[key] = val

                cur.execute(insert_sql, (code, row_plant, json.dumps(rec, default=json_serial), "active", u["username"], u["full_name"], cur_time, cur_time))
                imported_count += 1
            conn.commit()
        finally:
            conn.close()

    return {"ok": True, "code": code, "title": title, "columns_count": len(fields), "imported_count": imported_count}

@app.post("/api/se93/delete")
def se93_delete(req: Request, code: str = Form(...)):
    need_master(req)
    code = code.upper().strip()
    exec_write("DELETE FROM tcodes WHERE code=?", (code,))
    exec_write("DELETE FROM tcode_schemas WHERE code=?", (code,))
    exec_write("DELETE FROM tcode_layouts WHERE tcode_code=?", (code,))
    exec_write("DELETE FROM tcode_records WHERE tcode_code=?", (code,))
    exec_write("DELETE FROM access_grants WHERE tcode_code=?", (code,))
    return {"ok": True}

# =========================================================================
# --- AI CODE STUDIO : DECLARATIVE MODULE INSTALLER (PASTE-TO-BUILD) ---
# =========================================================================
VALID_TYPES = {"text", "number", "date", "textarea", "select", "email", "tel"}
RESERVED_CODES = {"TN01", "SU01", "SPRO", "SE93", "PFCG", "STUDIO"}

def _validate_spec(spec: dict):
    errors = []
    if not isinstance(spec, dict):
        return None, ["Spec must be a JSON object."]

    action = (spec.get("action") or "upsert").lower()
    if action not in ("upsert", "delete"):
        errors.append("action must be 'upsert' or 'delete'.")

    code = str(spec.get("code") or "").upper().strip()
    if not re.match(r'^[A-Z][A-Z0-9_]{1,10}$', code):
        errors.append("code must be 2-11 chars, start with a letter, A-Z/0-9/_ only.")
    if code in RESERVED_CODES:
        errors.append(f"'{code}' is a reserved system module.")

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

    clean_fields, seen = [], set()
    for i, f in enumerate(fields):
        if not isinstance(f, dict):
            errors.append(f"fields[{i}] must be an object."); continue
        key = re.sub(r'[^a-z0-9_]', '_', str(f.get("key") or f.get("label") or f"col_{i+1}").lower().strip())
        key = re.sub(r'_+', '_', key).strip('_') or f"col_{i+1}"
        if key in seen:
            errors.append(f"duplicate field key '{key}'."); continue
        seen.add(key)
        ftype = str(f.get("type") or "text").lower()
        if ftype not in VALID_TYPES:
            errors.append(f"field '{key}' has invalid type '{ftype}'."); ftype = "text"
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

    presets = spec.get("presets") or []
    clean["presets"] = [p for p in presets if isinstance(p, dict)][:50]
    return clean, errors

def _apply_spec(clean: dict, user: dict, note: str = ""):
    code, action = clean["code"], clean["action"]
    cur_time = now()
    existing = fetch_one("SELECT code FROM tcodes WHERE code=?", (code,))

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
                "UPDATE tcodes SET title=?, description=?, module=?, allowed_depts=?, "
                "updated_at=?, created_by=? WHERE code=?",
                (clean["title"], clean["description"], clean["module"],
                 clean["allowed_depts"], cur_time, user["username"], code))
            exec_write(
                "UPDATE tcode_schemas SET fields=?, layout=?, updated_at=?, updated_by=? WHERE code=?",
                (fields_json, layout_cfg, cur_time, user["username"], code))
            dl = fetch_one("SELECT id FROM tcode_layouts WHERE tcode_code=? AND is_default=1", (code,))
            if dl:
                exec_write(
                    "UPDATE tcode_layouts SET fields=?, layout_config=?, updated_at=? WHERE id=?",
                    (fields_json, layout_cfg, cur_time, dl["id"]))
            else:
                exec_write(
                    "INSERT INTO tcode_layouts (tcode_code, layout_code, name, title, is_default, "
                    "fields, layout_config, description, created_by, created_by_name, created_at, updated_at) "
                    "VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                    (code, f"GRID_{code}_01", "Default Grid", clean["title"], 1, fields_json,
                     layout_cfg, "Managed by AI Studio", user["username"], user["full_name"],
                     cur_time, cur_time))
            summary = {"action": "updated", "code": code}
        else:
            exec_write(
                "INSERT INTO tcodes (code, title, description, module, allowed_depts, active, "
                "created_at, updated_at, created_by) VALUES (?,?,?,?,?,?,?,?,?)",
                (code, clean["title"], clean["description"], clean["module"],
                 clean["allowed_depts"], 1, cur_time, cur_time, user["username"]))
            exec_write(
                "INSERT INTO tcode_schemas (code, fields, layout, updated_at, updated_by) VALUES (?,?,?,?,?)",
                (code, fields_json, layout_cfg, cur_time, user["username"]))
            exec_write(
                "INSERT INTO tcode_layouts (tcode_code, layout_code, name, title, is_default, "
                "fields, layout_config, description, created_by, created_by_name, created_at, updated_at) "
                "VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                (code, f"GRID_{code}_01", "Default Grid", clean["title"], 1, fields_json,
                 layout_cfg, "Created via AI Studio", user["username"], user["full_name"],
                 cur_time, cur_time))
            summary = {"action": "created", "code": code}

        seeded = 0
        field_keys = {f["key"] for f in clean["fields"]}
        for row in clean["seed_rows"]:
            rec = {k: row[k] for k in row if k in field_keys}
            if not rec:
                continue
            plant = str(row.get("plant") or (user.get("plants") or "1000").split(",")[0].strip() or "1000")
            exec_write(
                "INSERT INTO tcode_records (tcode_code, plant, data, status, created_by, "
                "created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?)",
                (code, plant, json.dumps(rec, default=json_serial), "active",
                 user["username"], user["full_name"], cur_time, cur_time))
            seeded += 1
        summary["seeded_rows"] = seeded

    prev = fetch_one("SELECT COALESCE(MAX(version),0) as v FROM module_specs WHERE tcode_code=?", (code,))
    version = (prev["v"] if prev else 0) + 1
    exec_write(
        "INSERT INTO module_specs (tcode_code, version, action, spec, applied_by, applied_at, note) "
        "VALUES (?,?,?,?,?,?,?)",
        (code, version, action, json.dumps(clean), user["username"], cur_time, note or ""))
    summary["version"] = version
    return summary


@app.post("/api/studio/validate")
def studio_validate(req: Request, payload: dict):
    need_master(req)
    clean, errors = _validate_spec(payload.get("spec", payload))
    existing = bool(clean and clean.get("code") and fetch_one("SELECT 1 FROM tcodes WHERE code=?", (clean["code"],)))
    return {"ok": not errors, "errors": errors, "clean": clean, "will_overwrite": existing}


@app.post("/api/studio/apply")
def studio_apply(req: Request, payload: dict):
    user = need_master(req)
    clean, errors = _validate_spec(payload.get("spec", payload))
    if errors:
        raise HTTPException(400, "Spec invalid: " + "; ".join(errors))
    summary = _apply_spec(clean, user, note=str(payload.get("note", "") or ""))
    return {"ok": True, **summary}


@app.get("/api/studio/history")
def studio_history(req: Request, code: str = ""):
    need_master(req)
    sql = "SELECT id, tcode_code, version, action, applied_by, applied_at, note FROM module_specs"
    params = []
    if code:
        sql += " WHERE tcode_code=?"
        params.append(code.upper().strip())
    sql += " ORDER BY id DESC LIMIT 200"
    return fetch_all(sql, tuple(params))


@app.get("/api/studio/spec/{code}")
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
    return {
        "action": "upsert",
        "code": code,
        "title": tc["title"],
        "description": tc["description"] or "",
        "module": tc["module"] or "CUSTOM",
        "allowed_depts": tc["allowed_depts"] or "",
        "fields": [{"key": f.get("key"), "label": f.get("label"),
                    "type": f.get("type", "text"), "required": f.get("required", False),
                    "width": f.get("width", 1)} for f in fields],
        "seed_rows": [], "presets": [],
    }


@app.post("/api/studio/rollback/{msid}")
def studio_rollback(req: Request, msid: int):
    user = need_master(req)
    row = fetch_one("SELECT * FROM module_specs WHERE id=?", (msid,))
    if not row:
        raise HTTPException(404, "History entry not found.")
    try:
        prev_spec = json.loads(row["spec"])
    except Exception:
        raise HTTPException(400, "Stored spec is corrupt.")
    prev_spec["action"] = "upsert" if row["action"] != "delete" else "delete"
    clean, errors = _validate_spec(prev_spec)
    if errors:
        raise HTTPException(400, "Cannot rollback: " + "; ".join(errors))
    summary = _apply_spec(clean, user, note=f"Rollback to v{row['version']}")
    return {"ok": True, **summary}


# =========================================================================
# --- CTST : CABLE TYPE-TEST CALCULATOR (Based on Calculations Formulas.xlsx) ---
# =========================================================================
# All formulas match the source spreadsheet and are computed server-side so the
# user only enters raw measurements. Fixes applied vs. the source sheet:
#  * elongation values multiplied by 100 (source sheet relied on cell formatting)
#  * temp sanity check for CR correction (source sheet had 262.2 °C entry)
#  * concentricity uses measured values, not spec limits
#  * PI() instead of hardcoded 3.143

def _compute_ctst(d: dict) -> dict:
    """Compute all derived values. Returns dict of computed fields (None-safe)."""
    out = {}
    def num(v):
        try:
            if v is None or str(v).strip() == "": return None
            return float(str(v).replace(",", ""))
        except Exception:
            return None

    # ---- CR (Conductor Resistance) ----
    obd_cr = num(d.get("obd_cr_reading"))
    temp_c = num(d.get("temp_c"))
    length_m = num(d.get("length_m"))
    if obd_cr is not None and temp_c is not None:
        if -50 <= temp_c <= 100:  # sanity: real CR test happens at ambient
            factor = 250.0 / (230.0 + temp_c)
            out["cr_1m"] = round(factor * obd_cr, 6)
            if length_m and length_m > 0:
                out["cr_total_per_km"] = round(factor * obd_cr / length_m * 1000.0, 6)
        else:
            out["cr_warning"] = "Temp outside ambient range — check entry"

    # ---- CR Percentage ----
    spec_cr = num(d.get("specified_cr"))
    obs_cr = num(d.get("observed_cr"))
    if spec_cr and obs_cr is not None and spec_cr != 0:
        out["cr_deviation_pct"] = round((obs_cr - spec_cr) / spec_cr * 100.0, 4)

    cr_max = num(d.get("cr_max"))
    cr_min = num(d.get("cr_min"))
    if cr_max and cr_min and (cr_max + cr_min) != 0:
        out["cr_variation_pct"] = round((cr_max - cr_min) / (cr_max + cr_min) * 100.0, 4)

    # ---- HV fault position interpolation ----
    hv_len = num(d.get("hv_length"))
    hv_top = num(d.get("hv_top_pos"))
    hv_bot = num(d.get("hv_bottom_pos"))
    if hv_len and hv_top is not None and hv_bot is not None:
        total = hv_top + hv_bot
        if total > 0:
            out["hv_from_top_m"] = round((hv_top * hv_len) / total, 3)
            out["hv_from_bottom_m"] = round((hv_bot * hv_len) / total, 3)

    # ---- Braid / Shield Coverage ----
    bd = num(d.get("braid_d"))
    uD = num(d.get("under_braid_D"))
    wpc = num(d.get("wires_per_carrier"))
    carr = num(d.get("carriers"))
    lay = num(d.get("lay_length"))
    if all(v is not None for v in (bd, uD, wpc, carr, lay)) and uD > 0 and lay > 0:
        N_total = wpc * carr
        F = (N_total * bd) / (2 * math.pi * uD) * math.sqrt(1 + (math.pi * uD / lay) ** 2)
        out["braid_total_wires"] = int(N_total)
        out["braid_filling_factor"] = round(F, 6)
        out["braid_coverage_pct"] = round((2 * F - F * F) * 100.0, 3)

    # ---- Tubular & Dumbbell Areas ----
    tub_D = num(d.get("tub_outer_d"))
    tub_t = num(d.get("tub_wall"))
    if tub_D and tub_t and tub_D > tub_t:
        out["tub_area_mm2"] = round(math.pi * (tub_D - tub_t) * tub_t, 4)

    dumb_w = num(d.get("dumb_width"))
    dumb_t = num(d.get("dumb_thk"))
    if dumb_w and dumb_t:
        out["dumb_area_mm2"] = round(dumb_w * dumb_t, 4)

    # ---- Tensile Strength ----
    bl_tub = num(d.get("break_load_tub"))
    bl_dumb = num(d.get("break_load_dumb"))
    if bl_tub is not None and out.get("tub_area_mm2"):
        out["tensile_tub_nmm2"] = round(bl_tub / out["tub_area_mm2"], 4)
    if bl_dumb is not None and out.get("dumb_area_mm2"):
        out["tensile_dumb_nmm2"] = round(bl_dumb / out["dumb_area_mm2"], 4)

    # ---- Elongation % ----
    init_g = num(d.get("initial_gauge"))
    tub_final = num(d.get("elong_tub_final"))
    dumb_final = num(d.get("elong_dumb_final"))
    if init_g and tub_final is not None and init_g != 0:
        out["elong_tub_pct"] = round((tub_final - init_g) / init_g * 100.0, 4)
    if init_g and dumb_final is not None and init_g != 0:
        out["elong_dumb_pct"] = round((dumb_final - init_g) / init_g * 100.0, 4)

    # ---- Hot Set & Permanent Set ----
    hs_final = num(d.get("hotset_elong_final"))
    ps_cooled = num(d.get("permset_cooled"))
    if init_g and hs_final is not None and init_g != 0:
        out["hotset_pct"] = round((hs_final - init_g) / init_g * 100.0, 4)
    if init_g and ps_cooled is not None and init_g != 0:
        out["permset_pct"] = round((ps_cooled - init_g) / init_g * 100.0, 4)

    return out


@app.get("/api/ctst/records")
def ctst_list(req: Request, limit: int = 200):
    u = need(req)
    if not check_tcode_access(u, "CTST"):
        raise HTTPException(403, "Access denied")
    rows = fetch_all("SELECT * FROM tcode_records WHERE tcode_code='CTST' AND status='active' ORDER BY id DESC LIMIT ?", (limit,))
    for r in rows:
        try: r["data"] = json.loads(r["data"])
        except: r["data"] = {}
        r["computed"] = _compute_ctst(r["data"])
    return rows


@app.post("/api/ctst/compute")
def ctst_compute(req: Request, payload: dict):
    """Live preview: compute derived values without saving."""
    need(req)
    return {"computed": _compute_ctst(payload)}


@app.post("/api/ctst/save")
def ctst_save(req: Request, payload: dict):
    u = need(req)
    if not check_tcode_access(u, "CTST"):
        raise HTTPException(403, "Access denied")
    plant = str(payload.pop("_plant", "") or (u.get("plants") or "1000").split(",")[0].strip() or "1000").strip()
    cur_time = now()
    exec_write(
        "INSERT INTO tcode_records (tcode_code, plant, data, status, created_by, "
        "created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?)",
        ("CTST", plant, json.dumps(payload, default=json_serial), "active",
         u["username"], u["full_name"], cur_time, cur_time))
    return {"ok": True, "computed": _compute_ctst(payload)}


@app.delete("/api/ctst/records/{rid}")
def ctst_delete(req: Request, rid: int):
    u = need(req)
    if not check_tcode_access(u, "CTST"):
        raise HTTPException(403, "Access denied")
    rec = fetch_one("SELECT * FROM tcode_records WHERE id=? AND tcode_code='CTST'", (rid,))
    if not rec:
        raise HTTPException(404, "Record not found")
    if rec["created_by"] != u["username"] and not is_master(u) and u["role"] != "admin":
        raise HTTPException(403, "Permission denied")
    exec_write("UPDATE tcode_records SET status='deleted', updated_at=? WHERE id=?", (now(), rid))
    return {"ok": True}


@app.get("/api/ctst/export")
def ctst_export(req: Request):
    u = need(req)
    if not check_tcode_access(u, "CTST"):
        raise HTTPException(403, "Export denied")
    rows = fetch_all("SELECT * FROM tcode_records WHERE tcode_code='CTST' AND status='active' ORDER BY id ASC")
    wb = Workbook()
    ws = wb.active
    ws.title = "CTST_DATA"
    headers = [
        "ID","Plant","Sample ID","Cable Type","Test Date",
        "OBD CR Reading (Ω)","Temp (°C)","Length (m)","CR 1m","CR Total (per km)","CR Warning",
        "Specified CR","Observed CR","CR Deviation %","Obs Max","Obs Min","CR Variation %",
        "HV Length","HV Fault Top (m)","HV Fault Bottom (m)","HV From Top (m)","HV From Bottom (m)",
        "Braid Wire Ø","Ø Under Braid","Wires/Carrier","Carriers","Lay Length","Total Wires",
        "Filling Factor","Coverage %",
        "Tubular Outer Ø","Tubular Wall","Tubular Area (mm²)","Tensile – Tubular (N/mm²)",
        "Dumbbell Width","Dumbbell Thk","Dumbbell Area (mm²)","Tensile – Dumbbell (N/mm²)",
        "Break Load – Tubular","Break Load – Dumbbell",
        "Initial Gauge","Tubular Final","Dumbbell Final","Tubular Elong %","Dumbbell Elong %",
        "Hot Set Final","Permanent Set Cooled","Hot Set %","Permanent Set %",
        "Result","Remarks","Created By","Created At"
    ]
    ws.append(headers)
    for r in rows:
        try: d = json.loads(r["data"])
        except: d = {}
        c = _compute_ctst(d)
        ws.append([
            r["id"], r["plant"], d.get("sample_id",""), d.get("cable_type",""), d.get("test_date",""),
            d.get("obd_cr_reading",""), d.get("temp_c",""), d.get("length_m",""),
            c.get("cr_1m",""), c.get("cr_total_per_km",""), c.get("cr_warning",""),
            d.get("specified_cr",""), d.get("observed_cr",""),
            c.get("cr_deviation_pct",""), d.get("cr_max",""), d.get("cr_min",""), c.get("cr_variation_pct",""),
            d.get("hv_length",""), d.get("hv_top_pos",""), d.get("hv_bottom_pos",""),
            c.get("hv_from_top_m",""), c.get("hv_from_bottom_m",""),
            d.get("braid_d",""), d.get("under_braid_D",""), d.get("wires_per_carrier",""),
            d.get("carriers",""), d.get("lay_length",""),
            c.get("braid_total_wires",""), c.get("braid_filling_factor",""), c.get("braid_coverage_pct",""),
            d.get("tub_outer_d",""), d.get("tub_wall",""), c.get("tub_area_mm2",""), c.get("tensile_tub_nmm2",""),
            d.get("dumb_width",""), d.get("dumb_thk",""), c.get("dumb_area_mm2",""), c.get("tensile_dumb_nmm2",""),
            d.get("break_load_tub",""), d.get("break_load_dumb",""),
            d.get("initial_gauge",""), d.get("elong_tub_final",""), d.get("elong_dumb_final",""),
            c.get("elong_tub_pct",""), c.get("elong_dumb_pct",""),
            d.get("hotset_elong_final",""), d.get("permset_cooled",""),
            c.get("hotset_pct",""), c.get("permset_pct",""),
            d.get("result",""), d.get("remarks",""),
            r["created_by_name"] or r["created_by"], r["created_at"]
        ])
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return StreamingResponse(bio, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             headers={"Content-Disposition": f"attachment; filename=CTST_Export_{datetime.date.today().strftime('%Y%m%d')}.xlsx"})


@app.get("/api/ctst/print/{rid}", response_class=HTMLResponse)
def ctst_print(req: Request, rid: int):
    u = need(req)
    if not check_tcode_access(u, "CTST"):
        return HTMLResponse("Access denied", 403)
    rec = fetch_one("SELECT * FROM tcode_records WHERE id=? AND tcode_code='CTST'", (rid,))
    if not rec:
        return HTMLResponse("Record not found", 404)
    try: d = json.loads(rec["data"])
    except: d = {}
    c = _compute_ctst(d)
    org = fetch_one("SELECT * FROM org WHERE id=1") or {}

    def cell(label, val, unit=""):
        return f"<tr><td style='padding:4px 8px;border:1px solid #cbd5e1;font-size:11px;background:#f8fafc'>{label}</td><td style='padding:4px 8px;border:1px solid #cbd5e1;font-size:11px;font-weight:600'>{val or ''} {unit}</td></tr>"

    def section(title, rows_html):
        return f"<div style='margin-bottom:16px'><h4 style='margin:0 0 6px;font-size:12px;color:#4f46e5;text-transform:uppercase'>{title}</h4><table style='width:100%;border-collapse:collapse'>{rows_html}</table></div>"

    html = f"""<!DOCTYPE html><html><head><meta charset="utf-8"><title>CTST {rec['id']}</title>
    <style>body{{font-family:sans-serif;margin:30px;color:#0f172a}}@media print{{.no-print{{display:none}}body{{margin:0}}}}</style></head><body>
    <div class="no-print" style="text-align:right;margin-bottom:12px">
      <button onclick="window.print()" style="background:#4f46e5;color:white;border:none;padding:8px 16px;border-radius:6px;cursor:pointer;font-weight:600">Print / Save PDF</button>
    </div>
    <h2 style="margin:0">{org.get('name','Thermocables Limited')}</h2>
    <h4 style="margin:4px 0 12px;color:#4f46e5">CABLE TYPE-TEST REPORT (CTST)</h4>
    <div style="font-size:11px;color:#475569;margin-bottom:16px">
      Record: #{rec['id']} &nbsp;|&nbsp; Plant: {rec['plant'] or '—'} &nbsp;|&nbsp; Tested by: {rec['created_by_name'] or rec['created_by']} &nbsp;|&nbsp; {rec['created_at']}
    </div>
    {section('Sample', cell('Sample / Batch ID', d.get('sample_id','')) + cell('Cable Type / Size', d.get('cable_type','')) + cell('Test Date', d.get('test_date','')))}
    {section('CR – Conductor Resistance', cell('OBD Reading', d.get('obd_cr_reading',''),'Ω') + cell('Temp', d.get('temp_c',''),'°C') + cell('Length', d.get('length_m',''),'m') + cell('CR (1 m)', c.get('cr_1m',''),'Ω') + cell('CR per km', c.get('cr_total_per_km',''),'Ω/km'))}
    {section('CR – Percentage vs Spec', cell('Specified CR', d.get('specified_cr','')) + cell('Observed CR', d.get('observed_cr','')) + cell('Deviation %', c.get('cr_deviation_pct',''),'%') + cell('Obs Max', d.get('cr_max','')) + cell('Obs Min', d.get('cr_min','')) + cell('Variation %', c.get('cr_variation_pct',''),'%'))}
    {section('HV Fault Location', cell('Test Length', d.get('hv_length',''),'m') + cell('Fault from Top', c.get('hv_from_top_m',''),'m') + cell('Fault from Bottom', c.get('hv_from_bottom_m',''),'m'))}
    {section('Braid / Shield Coverage', cell('Braid Wire Ø', d.get('braid_d',''),'mm') + cell('Ø Under Braid', d.get('under_braid_D',''),'mm') + cell('Wires/Carrier', d.get('wires_per_carrier','')) + cell('Carriers', d.get('carriers','')) + cell('Lay Length', d.get('lay_length',''),'mm') + cell('Total Wires', c.get('braid_total_wires','')) + cell('Filling Factor', c.get('braid_filling_factor','')) + cell('Coverage %', c.get('braid_coverage_pct',''),'%'))}
    {section('Tubular Mechanical', cell('Outer Ø', d.get('tub_outer_d',''),'mm') + cell('Wall', d.get('tub_wall',''),'mm') + cell('Area', c.get('tub_area_mm2',''),'mm²') + cell('Break Load', d.get('break_load_tub',''),'N') + cell('Tensile Strength', c.get('tensile_tub_nmm2',''),'N/mm²') + cell('Initial Gauge', d.get('initial_gauge',''),'mm') + cell('Final Gauge', d.get('elong_tub_final',''),'mm') + cell('Elongation', c.get('elong_tub_pct',''),'%'))}
    {section('Dumbbell Mechanical', cell('Width', d.get('dumb_width',''),'mm') + cell('Thickness', d.get('dumb_thk',''),'mm') + cell('Area', c.get('dumb_area_mm2',''),'mm²') + cell('Break Load', d.get('break_load_dumb',''),'N') + cell('Tensile Strength', c.get('tensile_dumb_nmm2',''),'N/mm²') + cell('Final Gauge', d.get('elong_dumb_final',''),'mm') + cell('Elongation', c.get('elong_dumb_pct',''),'%'))}
    {section('Hot Set & Permanent Set', cell('Initial Gauge', d.get('initial_gauge',''),'mm') + cell('Hot Set Final', d.get('hotset_elong_final',''),'mm') + cell('Hot Set %', c.get('hotset_pct',''),'%') + cell('Permanent Set Cooled', d.get('permset_cooled',''),'mm') + cell('Permanent Set %', c.get('permset_pct',''),'%'))}
    {section('Verdict', cell('Overall Result', d.get('result','')) + cell('Remarks', d.get('remarks','')))}
    </body></html>"""
    return HTMLResponse(html)


# --- TN01 ROUTES & PRESETS ---
TN01_PRESET_ROUTES = {
    "PROD_TO_QC": {"title": "Production to Final Testing", "route": "Production ➔ Final Testing", "from_dept": "Production", "to_dept": "Final Testing", "columns": ["Sr. No.", "Job No.", "Customer Name", "Material Description", "Specified Length", "Entry Length", "Drum Size", "Drum Code", "Batch No. "]},
    "QC_TO_COMM": {"title": "Final Testing to Commercial", "route": "Final Testing ➔ Commercial", "from_dept": "Final Testing", "to_dept": "Commercial", "columns": ["Sr. No.", "Job / Prod Order #", "Batch / Heat Lot", "Cable Description / Spec", "Drum #", "Tested Length (Mtrs)", "HV Test Result", "Insulation Resistance", "QC Inspector", "Release Status"]},
    "PROD_TO_STORE": {"title": "Production to Store", "route": "Production ➔ Store", "from_dept": "Production", "to_dept": "Store", "columns": ["Sr. No.", "Work Order #", "Item Code / Material", "Batch / Heat Lot", "Quantity (Mtrs/Kgs)", "Weight (Kgs)", "Drum / Pack #", "Handover Person", "Remarks"]},
    "QC_TO_PROD": {"title": "Final Testing to Production (Rework/Reject)", "route": "Final Testing ➔ Production", "from_dept": "Final Testing", "to_dept": "Production", "columns": ["Sr. No.", "Job Order #", "Batch / Heat Lot", "Cable Size", "Defect Observed", "Length (Mtrs)", "QC Engineer", "Rework Instructions"]},
    "COMM_TO_QC": {"title": "Commercial to Final Testing (Inspection Request)", "route": "Commercial ➔ Final Testing", "from_dept": "Commercial", "to_dept": "Final Testing", "columns": ["Sr. No.", "Customer Name", "Order / PO #", "Item Description", "Ordered Qty", "Inspection Agency", "Target Date", "Special Testing Notes"]},
    "STORE_TO_PROD": {"title": "Store to Production (Raw Material Issue)", "route": "Store ➔ Production", "from_dept": "Store", "to_dept": "Production", "columns": ["Sr. No.", "Material Code", "Material Description", "Lot / GRN #", "Issued Qty", "UOM", "Issued By", "Received For Work Order"]}
}

@app.get("/api/tn01/routes")
def tn01_get_routes(req: Request):
    need(req)
    return TN01_PRESET_ROUTES

@app.get("/api/tn01/suggest-users")
def tn01_suggest_users(req: Request, plant: str = "", dept: str = ""):
    need(req)
    sql, params = "SELECT username, full_name, department, plants, role FROM users WHERE status='approved'", []
    if dept:
        sql += " AND department = ?"
        params.append(dept)
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
        except: r["fields"] = []
    return rows

@app.post("/api/tn01/layouts/save")
def tn01_save_layout(req: Request, payload: dict):
    u = need_admin(req)
    name, key, columns = payload.get("name", "Custom Layout").strip(), payload.get("key", "").strip().upper(), payload.get("columns", [])
    if not key:
        cnt = fetch_one("SELECT COUNT(*) as c FROM tcode_layouts WHERE tcode_code='TN01'")["c"]
        key = f"TN01_VAR_{cnt+1:02d}"

    cur_time = now()
    existing = fetch_one("SELECT id FROM tcode_layouts WHERE tcode_code='TN01' AND layout_code=?", (key,))
    if existing:
        exec_write("UPDATE tcode_layouts SET name=?, title=?, fields=?, updated_at=? WHERE id=?", (name, name, json.dumps(columns), cur_time, existing["id"]))
    else:
        exec_write("INSERT INTO tcode_layouts (tcode_code, layout_code, name, title, is_default, fields, layout_config, description, created_by, created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                   ("TN01", key, name, name, 0, json.dumps(columns), "{}", "Custom Format", u["username"], u["full_name"], cur_time, cur_time))
    return {"ok": True, "key": key}

@app.delete("/api/tn01/layouts/{lid}")
def tn01_delete_layout(req: Request, lid: int):
    need_admin(req)
    exec_write("DELETE FROM tcode_layouts WHERE id=? AND tcode_code='TN01'", (lid,))
    return {"ok": True}

@app.post("/api/tn01/send")
def tn01_send(req: Request, payload: dict):
    u = need(req)
    cur_time = now()
    today_tag = datetime.date.today().strftime("%Y%m%d")
    cnt_row = fetch_one("SELECT COUNT(*) as c FROM transfer_notes WHERE created_at LIKE ?", (f"{datetime.date.today().strftime('%Y-%m-%d')}%",))
    note_no = f"TN-{today_tag}-{(cnt_row['c'] if cnt_row else 0) + 1:04d}"

    headers = payload.get("headers", [])
    rows_data = payload.get("rows", [])
    to_username = payload.get("to_user", "").strip().lower()
    target_user = fetch_one("SELECT username, full_name FROM users WHERE username=?", (to_username,))
    to_user_name = target_user["full_name"] if target_user else to_username
    from_plant = payload.get("from_plant", "").strip() or (u.get("plants") or "1000").split(",")[0].strip()

    exec_write('''
    INSERT INTO transfer_notes (
        note_no, layout_key, route, from_plant, from_dept, from_user, from_user_name,
        to_plant, to_dept, to_user, to_user_name, entry_count, sheet_headers, sheet_data,
        status, remarks, created_at
    ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    ''', (
        note_no, payload.get("layout_key", "STANDARD"), payload.get("route", "Transfer"),
        from_plant, u["department"], u["username"], u["full_name"],
        payload.get("to_plant", "").strip(), payload.get("to_dept", "").strip(),
        to_username, to_user_name, len(rows_data), json.dumps(headers),
        json.dumps(rows_data, default=json_serial), "sent", payload.get("remarks", "").strip(), cur_time
    ))
    return {"ok": True, "note_no": note_no, "entries": len(rows_data)}

@app.get("/api/tn01/inbox")
def tn01_get_inbox(req: Request):
    u = need(req)
    return fetch_all("SELECT * FROM transfer_notes WHERE to_user = ? OR (to_dept = ? AND to_user = '') ORDER BY id DESC", (u["username"].lower(), u["department"]))

@app.post("/api/tn01/receive/{nid}")
def tn01_receive(req: Request, nid: int):
    u = need(req)
    cur_time = now()
    exec_write("UPDATE transfer_notes SET status='received', received_at=?, received_by=?, received_by_name=? WHERE id=?",
               (cur_time, u["username"], u["full_name"], nid))
    return {"ok": True, "received_at": cur_time}

@app.get("/api/tn01/history")
def tn01_get_history(req: Request, plant: str = "", route: str = "", status: str = "",
                     from_dept: str = "", to_dept: str = "", date_from: str = "", date_to: str = ""):
    u = need(req)
    sql = "SELECT * FROM transfer_notes WHERE 1=1"
    params = []

    if not (is_master(u) or u["role"] == "admin"):
        sql += " AND (from_user = ? OR to_user = ? OR from_dept = ? OR to_dept = ?)"
        params.extend([u["username"], u["username"], u["department"], u["department"]])

    if plant:
        sql += " AND (from_plant = ? OR to_plant = ?)"
        params.extend([plant, plant])
    if route:
        sql += " AND route = ?"
        params.append(route)
    if status:
        sql += " AND status = ?"
        params.append(status)

    if (is_master(u) or u["role"] == "admin"):
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
    if not note: raise HTTPException(404, "Note not found")
    try: note["sheet_headers"] = json.loads(note["sheet_headers"])
    except: note["sheet_headers"] = []
    try: note["sheet_data"] = json.loads(note["sheet_data"])
    except: note["sheet_data"] = []
    return note

@app.get("/api/tn01/export/{nid}")
def tn01_export_excel(req: Request, nid: int):
    u = need(req)
    note = fetch_one("SELECT * FROM transfer_notes WHERE id=?", (nid,))
    if not note:
        raise HTTPException(404, "Note not found")

    is_sender = (u["username"] == note["from_user"]) or (u["department"] == note["from_dept"])
    is_receiver = (u["username"] == note["to_user"]) or (u["department"] == note["to_dept"])
    if not (is_sender or is_receiver or is_master(u) or u["role"] == "admin"):
        raise HTTPException(403, "You do not have permission to export this note.")

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
    ws.append([f"Created: {note['created_at']}  |  Received: {note['received_at'] or 'Pending'}"])
    ws.append([])
    ws.append(headers)
    for row in data_rows:
        ws.append(row)
    ws.append([])
    ws.append(["Handovered By: ________________", "", "Received By: ________________", "", "QC/Incharge: ________________"])

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
    for col_idx, h in enumerate(headers, start=1):
        ws.column_dimensions[ws.cell(row=7, column=col_idx).column_letter].width = max(14, min(40, len(str(h)) + 4))

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return StreamingResponse(bio, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             headers={"Content-Disposition": f"attachment; filename=TransferNote_{note['note_no']}.xlsx"})


@app.post("/api/tn01/export-grid")
def tn01_export_grid(req: Request, payload: dict):
    u = need(req)
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
    ws.append([f"Prepared by: {u['full_name']} ({u['username']}) · {u['department']}"])
    ws.append([f"Generated: {now()}"])
    ws.append([])
    ws.append(headers)
    for r in rows:
        ws.append(r)

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
    for col_idx, h in enumerate(headers, start=1):
        ws.column_dimensions[ws.cell(row=header_row, column=col_idx).column_letter].width = max(14, min(40, len(str(h)) + 4))

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return StreamingResponse(bio, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             headers={"Content-Disposition": f"attachment; filename=TN01_Draft_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"})


@app.get("/api/tn01/print/{nid}", response_class=HTMLResponse)
def tn01_print_preview(req: Request, nid: int):
    u = need(req)
    note = fetch_one("SELECT * FROM transfer_notes WHERE id=?", (nid,))
    if not note: return HTMLResponse("Note not found", 404)
    org = fetch_one("SELECT * FROM org WHERE id=1") or {}
    try: headers = json.loads(note["sheet_headers"])
    except: headers = []
    try: data_rows = json.loads(note["sheet_data"])
    except: data_rows = []

    th = "".join(f"<th style='border:1px solid #0f172a; padding:6px 8px; font-size:11px; background:#e0e7ff;'>{h}</th>" for h in headers)
    trs = "".join(f"<tr>{''.join(f'<td style=\"border:1px solid #cbd5e1; padding:6px 8px; font-size:11px;\">{c}</td>' for c in r)}</tr>" for r in data_rows)

    return HTMLResponse(f"""<!DOCTYPE html><html><head><meta charset="utf-8"><title>{note['note_no']} - Handover Note</title>
    <style>
      body {{ font-family: sans-serif; margin: 30px; color: #0f172a; }}
      .header {{ border-bottom: 2px solid #4f46e5; padding-bottom: 10px; margin-bottom: 15px; }}
      .meta {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; font-size: 12px; background: #f8fafc; padding: 12px; border: 1px solid #e2e8f0; margin-bottom: 20px; }}
      table {{ width: 100%; border-collapse: collapse; margin-bottom: 30px; }}
      .signs {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-top: 40px; text-align: center; font-size: 11px; font-weight: bold; }}
      .sign-box {{ border-top: 1px solid #0f172a; padding-top: 8px; }}
      @media print {{ .no-print {{ display: none; }} body {{ margin: 0; }} }}
    </style></head><body>
      <div class="no-print" style="margin-bottom:15px; text-align:right;">
        <button onclick="window.print()" style="background:#4f46e5; color:white; border:none; padding:8px 16px; border-radius:6px; cursor:pointer; font-weight:bold;">Print / Save PDF</button>
      </div>
      <div class="header">
        <h2 style="margin:0;">{org.get('name', 'Thermocables Limited')}</h2>
        <h4 style="margin:4px 0 0; color:#4f46e5;">INTER-DEPARTMENT MATERIAL TRANSFER & HANDOVER NOTE (TN01)</h4>
        <div style="font-family:monospace; font-weight:bold; margin-top:4px;">{note['note_no']} | Date: {note['created_at']}</div>
      </div>
      <div class="meta">
        <div><b>Route:</b><br>{note['route']}</div>
        <div><b>From:</b><br>{note['from_plant']} - {note['from_dept']}</div>
        <div><b>To:</b><br>{note['to_plant']} - {note['to_dept']}</div>
        <div><b>Entries:</b><br>{note['entry_count']} Batches</div>
        <div><b>Sender:</b><br>{note['from_user_name']}</div>
        <div><b>Recipient:</b><br>{note['to_user_name']}</div>
        <div><b>Status:</b><br><span style="text-transform:uppercase; font-weight:bold;">{note['status']}</span></div>
        <div><b>Receipt:</b><br>{note['received_at'] or 'Pending'}</div>
      </div>
      <table><thead><tr>{th}</tr></thead><tbody>{trs}</tbody></table>
      <div class="signs">
        <div class="sign-box">Handovered By ({note['from_user_name']})</div>
        <div class="sign-box">Production | Quality Incharge</div>
        <div class="sign-box">Received By ({note['received_by_name'] or note['to_user_name']})</div>
      </div>
    </body></html>""")

# --- STANDARD MODULE ALV & RECORDS ---
@app.get("/api/tcode/{code}/layouts")
def get_layouts(code: str, req: Request):
    u = need(req)
    if not check_tcode_access(u, code): raise HTTPException(403, "Access denied")
    rows = fetch_all("SELECT * FROM tcode_layouts WHERE tcode_code=? ORDER BY is_default DESC, id ASC", (code,))
    for r in rows:
        try: r["fields"] = json.loads(r["fields"])
        except: r["fields"] = []
    return rows

@app.get("/api/tcode/{code}/records")
def get_tcode_records(code: str, req: Request, plant: str = "", limit: int = 500):
    u = need(req)
    if not check_tcode_access(u, code): raise HTTPException(403, "Access denied")
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
        except: r["data"] = {}
    return rows

@app.post("/api/tcode/{code}/records")
def create_record(code: str, req: Request, payload: dict):
    u = need(req)
    if not check_tcode_access(u, code): raise HTTPException(403, "Access denied")
    plant = payload.pop("_plant", "").strip()
    exec_write("INSERT INTO tcode_records (tcode_code, plant, data, status, created_by, created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?)",
               (code, plant, json.dumps(payload, default=json_serial), "active", u["username"], u["full_name"], now(), now()))
    return {"ok": True}

@app.delete("/api/tcode/{code}/records/{rid}")
def delete_record(code: str, rid: int, req: Request):
    u = need(req)
    if not check_tcode_access(u, code): raise HTTPException(403, "Access denied")
    rec = fetch_one("SELECT * FROM tcode_records WHERE id=? AND tcode_code=?", (rid, code))
    if not rec: raise HTTPException(404, "Record not found")
    if rec["created_by"] != u["username"] and not is_master(u) and u["role"] != "admin":
        raise HTTPException(403, "Permission denied")
    exec_write("UPDATE tcode_records SET status='deleted', updated_at=? WHERE id=?", (now(), rid))
    return {"ok": True}

@app.get("/api/tcode/{code}/export")
def export_excel(code: str, req: Request):
    u = need(req)
    if not check_tcode_access(u, code): raise HTTPException(403, "Export denied")
    rows = fetch_all("SELECT * FROM tcode_records WHERE tcode_code=? AND status='active' ORDER BY id ASC", (code,))
    wb = Workbook()
    ws = wb.active
    ws.title = f"{code}_DATA"
    keys, dynamic_keys, parsed_rows = ["id", "plant", "created_by_name", "created_at"], set(), []
    for r in rows:
        try: d = json.loads(r["data"]) if isinstance(r["data"], str) else r["data"]
        except: d = {}
        dynamic_keys.update(d.keys())
        parsed_rows.append((r, d))

    all_cols = keys + sorted(list(dynamic_keys))
    ws.append(all_cols)
    for r, d in parsed_rows:
        ws.append([r["id"], r["plant"], r["created_by_name"], r["created_at"]] + [d.get(k, "") for k in sorted(list(dynamic_keys))])

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return StreamingResponse(bio, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f"attachment; filename=Export_{code}_{datetime.date.today().strftime('%Y%m%d')}.xlsx"})

# --- LIGHT PROFESSIONAL FRONTEND ---
HTML_MODERN_UI = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Nexus Core - Enterprise Operations Platform</title>
<script src="https://cdn.tailwindcss.com"></script>
<link href="https://cdn.jsdelivr.net/npm/remixicon@4.2.0/fonts/remixicon.css" rel="stylesheet">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">
<style>
  body { font-family: 'Inter', system-ui, sans-serif; background-color: #f8fafc; color: #0f172a; }
  .font-mono { font-family: 'JetBrains Mono', monospace; }
  .btn-primary { background: #4f46e5; color: white; transition: background .15s ease, box-shadow .15s ease; }
  .btn-primary:hover { background: #4338ca; box-shadow: 0 4px 12px rgba(79,70,229,.18); }
  .btn-primary:disabled { background: #a5b4fc; cursor: not-allowed; }
  .btn-secondary { background: #ffffff; border: 1px solid #e2e8f0; color: #475569; transition: all .15s ease; }
  .btn-secondary:hover { background: #f1f5f9; border-color: #cbd5e1; }
  .card { background: #ffffff; border: 1px solid #e2e8f0; }
  .custom-table th { background: #f1f5f9; color: #475569; font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: .02em; padding: 10px 14px; border-bottom: 1px solid #e2e8f0; }
  .custom-table td { padding: 12px 14px; border-bottom: 1px solid #f1f5f9; font-size: 13px; color: #334155; }
  .custom-table tbody tr:hover { background: #f8fafc; }
  .excel-grid { border-collapse: separate; border-spacing: 0; width: 100%; }
  .excel-grid th { background: #f1f5f9; color: #475569; font-weight: 600; font-size: 11px; text-transform: uppercase; padding: 8px 10px; border-right: 1px solid #e2e8f0; border-bottom: 1px solid #e2e8f0; position: sticky; top: 0; z-index: 2; }
  .excel-grid td { background: #ffffff; color: #0f172a; font-size: 12px; padding: 6px 10px; border-right: 1px solid #f1f5f9; border-bottom: 1px solid #f1f5f9; outline: none; min-width: 110px; }
  .excel-grid td.sr-col { background: #f8fafc; color: #64748b; font-weight: 600; text-align: center; width: 50px; user-select: none; }
  .excel-grid td:focus { background: #eef2ff; box-shadow: inset 0 0 0 2px #6366f1; }
  .form-input { background: #ffffff; border: 1px solid #cbd5e1; border-radius: .5rem; padding: .5rem .75rem; font-size: .75rem; color: #0f172a; width: 100%; outline: none; transition: border-color .15s ease, box-shadow .15s ease; }
  .form-input:focus { border-color: #6366f1; box-shadow: 0 0 0 3px rgba(99,102,241,.12); }
  .form-label { display: block; font-size: 11px; font-weight: 600; color: #475569; margin-bottom: 4px; text-transform: uppercase; letter-spacing: .02em; }
  .badge { display: inline-flex; align-items: center; font-size: 10px; font-weight: 600; padding: 2px 8px; border-radius: 999px; letter-spacing: .02em; text-transform: uppercase; }
  .badge-indigo { background: #eef2ff; color: #4338ca; }
  .badge-emerald { background: #ecfdf5; color: #047857; }
  .badge-amber { background: #fffbeb; color: #b45309; }
  .badge-red { background: #fef2f2; color: #b91c1c; }
  .badge-slate { background: #f1f5f9; color: #475569; }
  .sidebar-link { color: #475569; transition: background .15s ease, color .15s ease; }
  .sidebar-link:hover { background: #f1f5f9; color: #0f172a; }
  .sidebar-link.active { background: #eef2ff; color: #4338ca; border-left: 2px solid #4f46e5; }
  ::-webkit-scrollbar { width: 10px; height: 10px; }
  ::-webkit-scrollbar-track { background: #f1f5f9; }
  ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 5px; }
  ::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
</style>
</head>
<body class="h-screen flex flex-col select-none overflow-hidden">

<div id="app_root" class="h-full flex flex-col">
  <header class="h-14 bg-white border-b border-slate-200 px-4 flex items-center justify-between z-10 shrink-0 shadow-sm">
    <div class="flex items-center gap-4">
      <div class="flex items-center gap-2.5 cursor-pointer" onclick="goHome()">
        <div class="w-8 h-8 rounded-lg bg-indigo-600 flex items-center justify-center font-bold text-white shadow-sm"><i class="ri-pulse-fill text-lg"></i></div>
        <div class="font-bold text-sm tracking-tight text-slate-900 flex items-center gap-1.5">Nexus Core <span class="text-[10px] bg-indigo-50 text-indigo-600 border border-indigo-200 font-mono px-1.5 py-0.5 rounded">v3.7</span></div>
      </div>
      <div class="relative ml-4">
        <i class="ri-search-line absolute left-3 top-2.5 text-slate-400 text-xs"></i>
        <input id="cmd_box" placeholder="Jump to module (e.g. TN01, CTST, SU01, SPRO, STUDIO)..." class="bg-slate-50 border border-slate-200 rounded-lg pl-8 pr-16 py-1.5 text-xs text-slate-900 w-96 font-mono focus:outline-none focus:border-indigo-400 focus:bg-white" onkeydown="if(event.key==='Enter') execCmd()">
      </div>
    </div>
    <div id="nav_right" class="flex items-center gap-3">
      <div class="flex items-center gap-2 bg-slate-50 border border-slate-200 px-3 py-1 rounded-full text-xs">
        <div class="w-2 h-2 rounded-full bg-emerald-500"></div>
        <span id="user_display" class="font-medium text-slate-700">Connecting...</span>
        <span id="role_badge" class="badge badge-indigo"></span>
      </div>
      <button onclick="logout()" title="Sign Out" class="text-xs bg-white hover:bg-red-50 text-slate-500 hover:text-red-600 border border-slate-200 p-2 rounded-lg transition"><i class="ri-logout-box-r-line"></i></button>
    </div>
  </header>

  <div class="flex-1 flex overflow-hidden">
    <aside class="w-64 bg-white border-r border-slate-200 flex flex-col justify-between shrink-0">
      <div class="p-3 overflow-y-auto flex-1 space-y-4">
        <div>
          <div class="text-[10px] uppercase font-bold text-slate-400 tracking-wider px-2 mb-1.5">Special Modules</div>
          <div class="space-y-0.5">
            <button onclick="openTcode('CTST')" class="w-full text-left px-2.5 py-2 rounded-lg text-xs font-semibold bg-indigo-50 border border-indigo-200 text-indigo-700 hover:bg-indigo-100 flex items-center justify-between transition">
              <span class="flex items-center gap-2"><i class="ri-flask-line text-indigo-500 text-sm"></i> CTST - Type-Test</span>
              <span class="badge badge-indigo">NEW</span>
            </button>
            <button onclick="openTcode('TN01')" class="w-full text-left px-2.5 py-2 rounded-lg text-xs font-semibold bg-slate-50 border border-slate-200 text-slate-700 hover:bg-slate-100 flex items-center justify-between transition">
              <span class="flex items-center gap-2"><i class="ri-exchange-box-line text-slate-500 text-sm"></i> TN01 - Transfer Note</span>
              <span class="badge badge-slate">EXCEL</span>
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
              <span class="badge badge-emerald">NEW</span>
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
      <span id="stat_msg" class="text-slate-600">System Ready.</span>
    </div>
    <div class="flex items-center gap-4">
      <span id="stat_tcode" class="text-indigo-600 font-semibold font-mono">DASHBOARD</span>
      <span id="stat_resp">8ms</span>
      <span>READY</span>
    </div>
  </footer>
</div>

<div id="app_dialog" class="hidden fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-50 flex items-center justify-center p-4">
  <div id="app_dialog_body" class="bg-white text-slate-800 rounded-2xl shadow-2xl border border-slate-200 max-h-[90vh] overflow-y-auto w-full max-w-4xl"></div>
</div>

<script>
let ME = null, CURRENT_TCODE = null, ALL_TCODES = [], ALL_USERS_LIST = [];
const $ = s => document.querySelector(s);

function esc(str) {
  if (str === null || str === undefined) return '';
  return String(str).replace(/[&<>"']/g, s => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[s]));
}

function setMsg(msg, isErr = false) {
  const el = $('#stat_msg'), dot = $('#stat_dot');
  el.textContent = msg;
  dot.className = isErr ? "w-2 h-2 rounded-full bg-red-500" : "w-2 h-2 rounded-full bg-emerald-500";
  el.className = isErr ? "text-red-600 font-semibold" : "text-slate-600";
}

async function api(path, opts = {}) {
  const token = localStorage.getItem("nexus_token");
  const headers = {...(opts.headers || {})};
  if (token) headers["Authorization"] = `Bearer ${token}`;
  const t0 = performance.now();
  const res = await fetch(path, {...opts, headers, credentials: 'include'});
  $('#stat_resp').textContent = Math.round(performance.now() - t0) + 'ms';
  if (res.status === 401) { renderLogon(); throw 'Auth'; }
  if (!res.ok) {
    let t = await res.text();
    try { t = JSON.parse(t).detail || t; } catch {}
    throw t;
  }
  return await res.json();
}

function renderLogon(errorMsg = "") {
  $('#app_root').innerHTML = `
  <div class="min-h-screen bg-gradient-to-br from-slate-50 to-indigo-50 flex items-center justify-center p-4">
    <div class="w-full max-w-md bg-white border border-slate-200 rounded-2xl shadow-xl p-8 z-10">
      <div class="flex items-center gap-3 mb-6">
        <div class="w-10 h-10 rounded-xl bg-indigo-600 flex items-center justify-center font-bold text-white shadow-sm"><i class="ri-pulse-fill text-2xl"></i></div>
        <div><h1 class="text-xl font-bold text-slate-900">Nexus Core ERP</h1><p class="text-xs text-slate-500">Thermocables Industrial Suite</p></div>
      </div>
      ${errorMsg ? `<div class="mb-4 p-3 bg-red-50 border border-red-200 rounded-xl text-red-700 text-xs">${esc(errorMsg)}</div>` : ''}
      <div class="space-y-4">
        <div><label class="form-label">Username</label><input id="l_u" class="form-input" placeholder="admin" autofocus></div>
        <div><label class="form-label">Password</label><input id="l_p" type="password" class="form-input" placeholder="••••••••" onkeydown="if(event.key==='Enter') handleLogon()"></div>
        <button onclick="handleLogon()" class="w-full btn-primary py-2.5 rounded-lg font-semibold text-sm">Sign In</button>
      </div>
      <div class="mt-6 pt-4 border-t border-slate-200 text-[11px] text-slate-500 text-center">Default: admin / Admin@123</div>
    </div>
  </div>`;
}

async function handleLogon() {
  const u = $('#l_u').value.trim(), p = $('#l_p').value;
  if (!u || !p) return renderLogon("Enter username and password.");
  const fd = new FormData(); fd.append("username", u); fd.append("password", p);
  try {
    const res = await fetch('/api/login', {method: 'POST', body: fd});
    if (!res.ok) { let t = await res.text(); try { t = JSON.parse(t).detail || t; } catch {} return renderLogon(t); }
    const data = await res.json();
    if (data.token) localStorage.setItem("nexus_token", data.token);
    location.reload();
  } catch(e) { renderLogon("Connection error: " + e); }
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
    if (ME.is_master || ME.role === 'admin') $('#admin_nav_section').classList.remove('hidden');
    setMsg(`Connected as ${ME.username}`);
    await loadModules();
    goHome();
  } catch(e) { renderLogon(); }
}

function execCmd() {
  let val = $('#cmd_box').value.trim();
  if (!val) return;
  $('#cmd_box').value = '';
  if (val.startsWith('/n') || val.startsWith('/N')) val = val.substring(2).trim();
  if (!val) { goHome(); return; }
  openTcode(val.toUpperCase());
}

function goHome() {
  CURRENT_TCODE = null;
  $('#stat_tcode').textContent = "DASHBOARD";
  $('#main_screen').innerHTML = `
    <div class="flex-1 overflow-y-auto p-8 select-text">
      <div class="max-w-5xl mx-auto space-y-8">
        <div><h1 class="text-2xl font-bold text-slate-900">Operations Dashboard</h1><p class="text-sm text-slate-500 mt-1">Thermocables Industrial Management Platform</p></div>
        <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div class="card rounded-2xl p-5 shadow-sm">
            <div class="text-[10px] font-semibold text-slate-500 uppercase tracking-wider">Enterprise Org</div>
            <div class="text-lg font-bold text-slate-900 mt-1.5">${esc(ME.company)}</div>
            <div class="text-xs text-indigo-600 mt-1 font-medium">Dept: ${esc(ME.department)}</div>
          </div>
          <div class="card rounded-2xl p-5 shadow-sm">
            <div class="text-[10px] font-semibold text-slate-500 uppercase tracking-wider">Profile Role</div>
            <div class="text-lg font-bold text-emerald-600 mt-1.5 uppercase">${ME.role}</div>
            <div class="text-xs text-slate-500 mt-1 truncate">Plants: ${esc(ME.plants.join(', ') || 'Global')}</div>
          </div>
          <div class="card rounded-2xl p-5 shadow-sm">
            <div class="text-[10px] font-semibold text-slate-500 uppercase tracking-wider">Quick Jump</div>
            <div class="text-xs text-slate-600 mt-2 font-mono leading-relaxed">Launch <b class="text-indigo-600">CTST</b> for cable testing, <b class="text-indigo-600">TN01</b> for handover, or <b class="text-indigo-600">STUDIO</b> to paste AI modules.</div>
          </div>
        </div>
        <div>
          <h3 class="text-xs font-bold text-slate-500 uppercase tracking-wider mb-4">Operations Modules</h3>
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
    <button onclick="openTcode('${t.code}')" class="sidebar-link w-full text-left px-2.5 py-1.5 rounded-lg flex items-center justify-between">
      <span class="truncate"><b class="text-indigo-600 font-mono">${t.code}</b> <span class="text-slate-600">- ${esc(t.title)}</span></span>
    </button>
  `).join('');
}

function renderQuickGrid() {
  const container = $('#quick_grid');
  if (!container) return;
  const list = ALL_TCODES.filter(t => t.has_access);
  container.innerHTML = list.map(t => `
    <div onclick="openTcode('${t.code}')" class="card hover:border-indigo-300 hover:shadow-md p-5 rounded-2xl cursor-pointer transition-all">
      <div class="flex items-center justify-between mb-2">
        <span class="font-mono font-bold text-xs bg-indigo-50 text-indigo-600 border border-indigo-200 px-2 py-0.5 rounded-lg">${t.code}</span>
        <span class="badge badge-slate">${t.module || 'ERP'}</span>
      </div>
      <h4 class="font-semibold text-slate-900 text-sm mt-2">${esc(t.title)}</h4>
      <p class="text-xs text-slate-500 mt-1 line-clamp-2">${esc(t.description || 'Standard module flow')}</p>
    </div>
  `).join('');
}

async function openTcode(code) {
  code = code.toUpperCase().trim();
  CURRENT_TCODE = code;
  $('#stat_tcode').textContent = code;
  setMsg(`Module ${code} loaded`);

  if (code === 'CTST') return renderCTST();
  if (code === 'TN01') return renderTN01();
  if (code === 'SU01') return renderSU01();
  if (code === 'SPRO') return renderSPRO();
  if (code === 'SE93') return renderSE93();
  if (code === 'PFCG') return renderPFCG();
  if (code === 'STUDIO') return renderStudio();

  renderStandardModule(code);
}

// =========================================================================
// --- CTST : CABLE TYPE-TEST CALCULATOR ---
// =========================================================================
async function renderCTST() {
  const rows = await api('/api/ctst/records');

  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col p-6 overflow-hidden select-text">
      <div class="flex items-center justify-between pb-4 border-b border-slate-200 shrink-0">
        <div class="flex items-center gap-3">
          <span class="font-mono font-bold text-xs bg-indigo-600 text-white px-2.5 py-1 rounded-lg shadow-sm">CTST</span>
          <div>
            <h2 class="text-lg font-bold text-slate-900 tracking-tight">Cable Type-Test Calculator</h2>
            <p class="text-xs text-slate-500">CR / HV / Braid Coverage / Tubular & Dumbbell mechanical tests</p>
          </div>
        </div>
        <div class="flex items-center gap-2">
          <button onclick="showCTSTForm()" class="btn-primary text-xs font-bold px-4 py-2 rounded-lg flex items-center gap-1.5 shadow-sm"><i class="ri-add-line"></i> New Test</button>
          <a href="/api/ctst/export" class="btn-secondary text-xs font-semibold px-3 py-2 rounded-lg flex items-center gap-1"><i class="ri-download-2-line"></i> Export All</a>
          <button onclick="renderCTST()" class="btn-secondary text-xs font-semibold p-2 rounded-lg"><i class="ri-refresh-line"></i></button>
        </div>
      </div>

      <div class="flex-1 mt-4 card rounded-2xl overflow-auto">
        <table class="w-full text-left custom-table">
          <thead><tr>
            <th style="width:60px;">#</th><th>Sample / Batch</th><th>Cable Type</th><th>Date</th>
            <th>CR 1m</th><th>CR Deviation</th><th>Coverage %</th><th>Hot Set %</th><th>Result</th><th style="width:180px;">Actions</th>
          </tr></thead>
          <tbody>
            ${!rows.length ? `<tr><td colspan="10" class="text-center py-16 text-slate-400">No test records yet. Click <b>New Test</b> to begin.</td></tr>` : ''}
            ${rows.map(r => {
              const d = r.data || {}, c = r.computed || {};
              const verdictBadge = d.result === 'PASS' ? 'badge-emerald' : (d.result === 'FAIL' ? 'badge-red' : 'badge-amber');
              return `<tr>
                <td class="font-mono text-slate-400 font-semibold">${r.id}</td>
                <td class="font-semibold text-slate-900">${esc(d.sample_id||'—')}</td>
                <td>${esc(d.cable_type||'—')}</td>
                <td class="text-xs text-slate-500 font-mono">${esc(d.test_date||'—')}</td>
                <td class="font-mono text-xs">${c.cr_1m !== undefined ? c.cr_1m : '—'}</td>
                <td class="font-mono text-xs">${c.cr_deviation_pct !== undefined ? c.cr_deviation_pct + ' %' : '—'}</td>
                <td class="font-mono text-xs">${c.braid_coverage_pct !== undefined ? c.braid_coverage_pct + ' %' : '—'}</td>
                <td class="font-mono text-xs">${c.hotset_pct !== undefined ? c.hotset_pct + ' %' : '—'}</td>
                <td>${d.result ? `<span class="badge ${verdictBadge}">${esc(d.result)}</span>` : `<span class="text-slate-400 text-xs">—</span>`}</td>
                <td class="space-x-1">
                  <a href="/api/ctst/print/${r.id}" target="_blank" class="text-xs text-indigo-600 font-semibold">Print</a>
                  <button onclick="viewCTST(${r.id})" class="text-xs text-slate-600 font-semibold ml-2">View</button>
                  <button onclick="deleteCTST(${r.id})" class="text-xs text-red-600 font-semibold ml-2">Del</button>
                </td>
              </tr>`;
            }).join('')}
          </tbody>
        </table>
      </div>
    </div>`;
}

async function showCTSTForm() {
  openDialog(`
    <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl">
      <div>
        <h3 class="font-bold text-slate-900 text-sm">New Cable Type-Test</h3>
        <p class="text-xs text-slate-500 mt-0.5">Enter raw measurements — all derived values are calculated automatically</p>
      </div>
      <button onclick="closeDialog()"><i class="ri-close-line text-lg text-slate-500"></i></button>
    </div>

    <div class="p-6 space-y-6">
      <div class="grid grid-cols-3 gap-4">
        <div><label class="form-label">Sample / Batch ID *</label><input id="cf_sample_id" class="form-input font-mono" placeholder="e.g. LT-XLPE-4C-95-B12"></div>
        <div><label class="form-label">Cable Type / Size</label><input id="cf_cable_type" class="form-input" placeholder="e.g. 1.1kV XLPE 4Cx95"></div>
        <div><label class="form-label">Test Date</label><input id="cf_test_date" type="date" class="form-input"></div>
      </div>

      <div class="border border-slate-200 rounded-xl p-4 bg-slate-50/50">
        <h4 class="text-xs font-bold text-indigo-700 uppercase tracking-wider mb-3"><i class="ri-pulse-line"></i> CR — Conductor Resistance</h4>
        <div class="grid grid-cols-3 gap-3">
          <div><label class="form-label">OBD Reading (Ω)</label><input id="cf_obd_cr_reading" type="number" step="any" class="form-input"></div>
          <div><label class="form-label">Test Temp (°C)</label><input id="cf_temp_c" type="number" step="any" class="form-input" placeholder="e.g. 25"></div>
          <div><label class="form-label">Length (m)</label><input id="cf_length_m" type="number" step="any" class="form-input"></div>
        </div>
        <div class="mt-2 text-[10px] text-slate-500">Correction to 20 °C: <b>R20 = R × 250 / (230 + T)</b>. Temp must be ambient (−50 to 100 °C).</div>
      </div>

      <div class="border border-slate-200 rounded-xl p-4 bg-slate-50/50">
        <h4 class="text-xs font-bold text-indigo-700 uppercase tracking-wider mb-3"><i class="ri-bar-chart-line"></i> CR Percentage vs Spec</h4>
        <div class="grid grid-cols-4 gap-3">
          <div><label class="form-label">Specified CR</label><input id="cf_specified_cr" type="number" step="any" class="form-input"></div>
          <div><label class="form-label">Observed CR</label><input id="cf_observed_cr" type="number" step="any" class="form-input"></div>
          <div><label class="form-label">Obs. Max CR</label><input id="cf_cr_max" type="number" step="any" class="form-input"></div>
          <div><label class="form-label">Obs. Min CR</label><input id="cf_cr_min" type="number" step="any" class="form-input"></div>
        </div>
      </div>

      <div class="border border-slate-200 rounded-xl p-4 bg-slate-50/50">
        <h4 class="text-xs font-bold text-indigo-700 uppercase tracking-wider mb-3"><i class="ri-flashlight-line"></i> HV Fault Location</h4>
        <div class="grid grid-cols-3 gap-3">
          <div><label class="form-label">Test Length (m)</label><input id="cf_hv_length" type="number" step="any" class="form-input"></div>
          <div><label class="form-label">Fault from Top (m)</label><input id="cf_hv_top_pos" type="number" step="any" class="form-input"></div>
          <div><label class="form-label">Fault from Bottom (m)</label><input id="cf_hv_bottom_pos" type="number" step="any" class="form-input"></div>
        </div>
      </div>

      <div class="border border-slate-200 rounded-xl p-4 bg-slate-50/50">
        <h4 class="text-xs font-bold text-indigo-700 uppercase tracking-wider mb-3"><i class="ri-shield-check-line"></i> Braid / Shield Coverage</h4>
        <div class="grid grid-cols-5 gap-3">
          <div><label class="form-label">Braid Wire Ø (mm)</label><input id="cf_braid_d" type="number" step="any" class="form-input"></div>
          <div><label class="form-label">Ø Under Braid (mm)</label><input id="cf_under_braid_D" type="number" step="any" class="form-input"></div>
          <div><label class="form-label">Wires / Carrier</label><input id="cf_wires_per_carrier" type="number" step="1" class="form-input"></div>
          <div><label class="form-label">Carriers</label><input id="cf_carriers" type="number" step="1" class="form-input"></div>
          <div><label class="form-label">Lay Length (mm)</label><input id="cf_lay_length" type="number" step="any" class="form-input"></div>
        </div>
      </div>

      <div class="grid grid-cols-2 gap-4">
        <div class="border border-slate-200 rounded-xl p-4 bg-slate-50/50">
          <h4 class="text-xs font-bold text-indigo-700 uppercase tracking-wider mb-3"><i class="ri-circle-line"></i> Tubular Specimen</h4>
          <div class="grid grid-cols-2 gap-3">
            <div><label class="form-label">Outer Ø (mm)</label><input id="cf_tub_outer_d" type="number" step="any" class="form-input"></div>
            <div><label class="form-label">Wall (mm)</label><input id="cf_tub_wall" type="number" step="any" class="form-input"></div>
            <div><label class="form-label">Break Load (N)</label><input id="cf_break_load_tub" type="number" step="any" class="form-input"></div>
            <div><label class="form-label">Final Gauge (mm)</label><input id="cf_elong_tub_final" type="number" step="any" class="form-input"></div>
          </div>
        </div>
        <div class="border border-slate-200 rounded-xl p-4 bg-slate-50/50">
          <h4 class="text-xs font-bold text-indigo-700 uppercase tracking-wider mb-3"><i class="ri-shape-line"></i> Dumbbell Specimen</h4>
          <div class="grid grid-cols-2 gap-3">
            <div><label class="form-label">Width (mm)</label><input id="cf_dumb_width" type="number" step="any" class="form-input"></div>
            <div><label class="form-label">Thickness (mm)</label><input id="cf_dumb_thk" type="number" step="any" class="form-input"></div>
            <div><label class="form-label">Break Load (N)</label><input id="cf_break_load_dumb" type="number" step="any" class="form-input"></div>
            <div><label class="form-label">Final Gauge (mm)</label><input id="cf_elong_dumb_final" type="number" step="any" class="form-input"></div>
          </div>
        </div>
      </div>

      <div class="border border-slate-200 rounded-xl p-4 bg-slate-50/50">
        <h4 class="text-xs font-bold text-indigo-700 uppercase tracking-wider mb-3"><i class="ri-temp-hot-line"></i> Hot Set & Permanent Set</h4>
        <div class="grid grid-cols-3 gap-3">
          <div><label class="form-label">Initial Gauge (mm)</label><input id="cf_initial_gauge" type="number" step="any" class="form-input" value="20"></div>
          <div><label class="form-label">Hot Set Final (mm)</label><input id="cf_hotset_elong_final" type="number" step="any" class="form-input"></div>
          <div><label class="form-label">Perm. Set Cooled (mm)</label><input id="cf_permset_cooled" type="number" step="any" class="form-input"></div>
        </div>
      </div>

      <div class="grid grid-cols-2 gap-4">
        <div>
          <label class="form-label">Overall Result</label>
          <select id="cf_result" class="form-input">
            <option value="">— Not judged —</option>
            <option value="PASS">PASS</option>
            <option value="FAIL">FAIL</option>
            <option value="RETAIN">RETAIN</option>
          </select>
        </div>
        <div>
          <label class="form-label">Facility</label>
          <input id="cf_plant" class="form-input font-mono" value="${esc((ME.plants && ME.plants[0]) || '1000')}">
        </div>
      </div>

      <div>
        <label class="form-label">Remarks</label>
        <textarea id="cf_remarks" class="form-input" rows="2" placeholder="Any observations, deviations, re-test notes..."></textarea>
      </div>

      <div class="bg-indigo-50 border border-indigo-200 rounded-xl p-4">
        <div class="flex justify-between items-center mb-2">
          <div class="text-xs font-bold text-indigo-800 uppercase tracking-wider">Live Computed Results</div>
          <button onclick="previewCTST()" class="text-[11px] font-bold text-indigo-600 hover:text-indigo-800"><i class="ri-calculator-line"></i> Recompute</button>
        </div>
        <div id="cf_computed" class="grid grid-cols-3 gap-2 text-xs text-slate-700">
          <div class="text-slate-400">Fill fields above, then press Recompute.</div>
        </div>
      </div>
    </div>

    <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl">
      <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-lg text-xs font-semibold">Cancel</button>
      <button onclick="saveCTST()" class="btn-primary px-5 py-2 rounded-lg text-xs font-bold shadow-sm">Save Test Record</button>
    </div>`);

  const today = new Date().toISOString().slice(0,10);
  $('#cf_test_date').value = today;
}

const CTST_FIELD_IDS = [
  'sample_id','cable_type','test_date','obd_cr_reading','temp_c','length_m',
  'specified_cr','observed_cr','cr_max','cr_min',
  'hv_length','hv_top_pos','hv_bottom_pos',
  'braid_d','under_braid_D','wires_per_carrier','carriers','lay_length',
  'tub_outer_d','tub_wall','dumb_width','dumb_thk',
  'break_load_tub','break_load_dumb',
  'elong_tub_final','elong_dumb_final','initial_gauge',
  'hotset_elong_final','permset_cooled',
  'result','remarks'
];

function collectCTSTPayload() {
  const payload = {};
  for (const k of CTST_FIELD_IDS) {
    const el = document.getElementById('cf_' + k);
    if (el) payload[k] = el.value;
  }
  payload._plant = document.getElementById('cf_plant') ? document.getElementById('cf_plant').value : '';
  return payload;
}

async function previewCTST() {
  const payload = collectCTSTPayload();
  try {
    const res = await api('/api/ctst/compute', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(payload)});
    const c = res.computed || {};
    const rows = [
      ['CR 1 m', c.cr_1m, 'Ω'],
      ['CR per km', c.cr_total_per_km, 'Ω/km'],
      ['CR Deviation', c.cr_deviation_pct, '%'],
      ['CR Variation', c.cr_variation_pct, '%'],
      ['HV from Top', c.hv_from_top_m, 'm'],
      ['HV from Bottom', c.hv_from_bottom_m, 'm'],
      ['Total Wires', c.braid_total_wires, ''],
      ['Filling Factor', c.braid_filling_factor, ''],
      ['Coverage', c.braid_coverage_pct, '%'],
      ['Tubular Area', c.tub_area_mm2, 'mm²'],
      ['Tensile Tub', c.tensile_tub_nmm2, 'N/mm²'],
      ['Tubular Elong', c.elong_tub_pct, '%'],
      ['Dumbbell Area', c.dumb_area_mm2, 'mm²'],
      ['Tensile Dumb', c.tensile_dumb_nmm2, 'N/mm²'],
      ['Dumbbell Elong', c.elong_dumb_pct, '%'],
      ['Hot Set', c.hotset_pct, '%'],
      ['Permanent Set', c.permset_pct, '%']
    ];
    document.getElementById('cf_computed').innerHTML = rows.map(([l,v,u]) =>
      `<div class="bg-white border border-indigo-100 rounded-lg p-2"><div class="text-[9px] uppercase text-slate-500 font-semibold">${l}</div><div class="font-mono font-bold text-slate-900 text-sm">${v !== undefined && v !== null ? v : '—'}${v!==undefined && v!==null && u ? ' '+u : ''}</div></div>`
    ).join('') + (c.cr_warning ? `<div class="col-span-3 bg-amber-50 border border-amber-200 rounded-lg p-2 text-amber-700">⚠ ${esc(c.cr_warning)}</div>` : '');
  } catch(e) { setMsg("Preview failed: " + e, true); }
}

async function saveCTST() {
  const payload = collectCTSTPayload();
  if (!payload.sample_id || !payload.sample_id.trim()) return alert("Sample / Batch ID is required.");
  try {
    await api('/api/ctst/save', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(payload)});
    closeDialog();
    renderCTST();
    setMsg("Test record saved");
  } catch(e) { alert(e); }
}

async function deleteCTST(id) {
  if (!confirm(`Delete test record #${id}?`)) return;
  try {
    await api(`/api/ctst/records/${id}`, {method: 'DELETE'});
    renderCTST();
    setMsg("Record deleted");
  } catch(e) { alert(e); }
}

async function viewCTST(id) {
  window.open(`/api/ctst/print/${id}`, '_blank');
}

// =========================================================================
// --- SU01 ---
// =========================================================================
async function renderSU01() {
  if (!ME.is_master && ME.role !== 'admin') return setMsg("Permission denied", true);
  ALL_USERS_LIST = await api('/api/su01/users');

  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col p-6 overflow-hidden select-text">
      <div class="flex items-center justify-between pb-4 border-b border-slate-200 shrink-0">
        <div>
          <h2 class="text-lg font-bold text-slate-900 tracking-tight">User Directory & IAM (SU01)</h2>
          <p class="text-xs text-slate-500">Manage user credentials, plant access, roles, and status</p>
        </div>
        <div class="flex items-center gap-2">
          <input id="su01_search" oninput="filterUsersTable(this.value)" placeholder="Search users, plants..." class="form-input w-64">
          <button onclick="showCreateUserModal()" class="btn-primary text-xs font-bold px-4 py-2 rounded-lg flex items-center gap-1.5 shadow-sm">
            <i class="ri-user-add-line"></i> Add User
          </button>
          <button onclick="renderSU01()" class="btn-secondary text-xs p-2 rounded-lg"><i class="ri-refresh-line"></i></button>
        </div>
      </div>

      <div class="flex-1 mt-4 card rounded-2xl overflow-auto">
        <table class="w-full text-left custom-table">
          <thead><tr><th>Username</th><th>Full Name</th><th>Role</th><th>Department</th><th>Facilities</th><th>Status</th><th style="width:200px;">Actions</th></tr></thead>
          <tbody id="su01_tbody"></tbody>
        </table>
      </div>
    </div>`;
  renderUsersRows(ALL_USERS_LIST);
}

function filterUsersTable(q) {
  const term = q.toLowerCase();
  const filtered = ALL_USERS_LIST.filter(u =>
    u.username.toLowerCase().includes(term) ||
    u.full_name.toLowerCase().includes(term) ||
    (u.department||'').toLowerCase().includes(term) ||
    (u.plants||'').toLowerCase().includes(term)
  );
  renderUsersRows(filtered);
}

function renderUsersRows(users) {
  const tbody = $('#su01_tbody');
  if (!tbody) return;
  if (!users.length) {
    tbody.innerHTML = `<tr><td colspan="7" class="text-center py-10 text-slate-400">No users found.</td></tr>`;
    return;
  }
  tbody.innerHTML = users.map(u => {
    const isSelf = (u.username === ME.username);
    const isMaster = (u.role === 'master');
    const canDelete = !isSelf && (!isMaster || ME.is_master);
    const roleBadge = u.role === 'master' ? 'badge-red' : (u.role === 'admin' ? 'badge-amber' : 'badge-slate');
    const statusBadge = u.status === 'approved' ? 'badge-emerald' : 'badge-red';
    return `
      <tr>
        <td class="font-mono font-bold text-indigo-600">${esc(u.username)}</td>
        <td class="font-medium text-slate-900">${esc(u.full_name)}</td>
        <td><span class="badge ${roleBadge}">${u.role.toUpperCase()}</span></td>
        <td>${esc(u.department)}</td>
        <td class="font-mono text-xs">${esc(u.plants || 'ALL')}</td>
        <td><span class="badge ${statusBadge}">${u.status}</span></td>
        <td class="space-x-1">
          <button onclick="showEditUserModal('${esc(u.username)}')" class="text-xs text-indigo-600 hover:text-indigo-800 font-semibold px-1.5 py-1">Edit</button>
          <button onclick="showResetPassModal('${esc(u.username)}')" class="text-xs text-amber-600 hover:text-amber-800 font-semibold px-1.5 py-1">Reset PW</button>
          ${canDelete ? `<button onclick="deleteUser('${esc(u.username)}')" class="text-xs text-red-600 hover:text-red-800 font-semibold px-1.5 py-1">Delete</button>` : ''}
        </td>
      </tr>`;
  }).join('');
}

function showCreateUserModal() {
  api('/api/spro/overview').then(data => {
    openDialog(`
      <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl">
        <h3 class="font-bold text-slate-900 text-sm">Create New User</h3>
        <button onclick="closeDialog()"><i class="ri-close-line text-lg text-slate-500"></i></button>
      </div>
      <div class="p-6 space-y-4">
        <div class="grid grid-cols-2 gap-4">
          <div><label class="form-label">Username *</label><input id="nu_u" class="form-input font-mono" placeholder="e.g. jdoe"></div>
          <div><label class="form-label">Password *</label><input id="nu_p" type="password" class="form-input" placeholder="••••••••"></div>
        </div>
        <div><label class="form-label">Full Name *</label><input id="nu_fn" class="form-input" placeholder="e.g. John Doe"></div>
        <div class="grid grid-cols-2 gap-4">
          <div>
            <label class="form-label">Department</label>
            <select id="nu_d" class="form-input">
              ${data.departments.map(d=>`<option value="${esc(d.name)}">${esc(d.name)}</option>`).join('')}
            </select>
          </div>
          <div>
            <label class="form-label">Role</label>
            <select id="nu_r" class="form-input">
              <option value="user">Standard User</option>
              <option value="admin">Administrator</option>
              ${ME.is_master ? '<option value="master">Master Superuser</option>' : ''}
            </select>
          </div>
        </div>
        <div><label class="form-label">Plant Authorization (comma-separated, blank = ALL)</label><input id="nu_pl" class="form-input font-mono" placeholder="e.g. 1000, 1400"></div>
      </div>
      <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl">
        <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-lg text-xs font-semibold">Cancel</button>
        <button onclick="submitCreateUser()" class="btn-primary px-4 py-2 rounded-lg text-xs font-semibold">Save User</button>
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
  } catch(e) { alert(e); }
}

async function showEditUserModal(u) {
  const users = await api('/api/su01/users');
  const user = users.find(x => x.username === u);
  const data = await api('/api/spro/overview');
  openDialog(`
    <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl">
      <h3 class="font-bold text-slate-900 text-sm">Update User: ${esc(u)}</h3>
      <button onclick="closeDialog()"><i class="ri-close-line text-lg text-slate-500"></i></button>
    </div>
    <div class="p-6 space-y-4">
      <div><label class="form-label">Full Name</label><input id="eu_fn" value="${esc(user.full_name)}" class="form-input"></div>
      <div class="grid grid-cols-2 gap-4">
        <div>
          <label class="form-label">Department</label>
          <select id="eu_d" class="form-input">
            ${data.departments.map(d=>`<option value="${esc(d.name)}" ${d.name===user.department?'selected':''}>${esc(d.name)}</option>`).join('')}
          </select>
        </div>
        <div>
          <label class="form-label">Role</label>
          <select id="eu_r" class="form-input">
            <option value="user" ${user.role==='user'?'selected':''}>Standard User</option>
            <option value="admin" ${user.role==='admin'?'selected':''}>Administrator</option>
            ${ME.is_master ? `<option value="master" ${user.role==='master'?'selected':''}>Master</option>` : ''}
          </select>
        </div>
      </div>
      <div class="grid grid-cols-2 gap-4">
        <div><label class="form-label">Facility Access</label><input id="eu_pl" value="${esc(user.plants||'')}" class="form-input font-mono"></div>
        <div>
          <label class="form-label">Status</label>
          <select id="eu_st" class="form-input">
            <option value="approved" ${user.status==='approved'?'selected':''}>Approved / Active</option>
            <option value="locked" ${user.status==='locked'?'selected':''}>Locked</option>
          </select>
        </div>
      </div>
    </div>
    <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl">
      <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-lg text-xs font-semibold">Cancel</button>
      <button onclick="submitEditUser('${esc(u)}')" class="btn-primary px-4 py-2 rounded-lg text-xs font-semibold">Save Changes</button>
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
  } catch(e) { alert(e); }
}

function showResetPassModal(u) {
  const p = prompt(`Enter new password for ${u}:`);
  if (!p) return;
  const fd = new FormData(); fd.append("username", u); fd.append("new_password", p);
  api('/api/su01/reset-password', {method: 'POST', body: fd}).then(() => {
    setMsg(`Password reset for ${u}`);
    alert(`Password for ${u} updated.`);
  }).catch(e => alert(e));
}

async function deleteUser(u) {
  if (!confirm(`Permanently remove account ${u}?`)) return;
  const fd = new FormData(); fd.append("username", u);
  try {
    await api('/api/su01/delete', {method: 'POST', body: fd});
    renderSU01();
    setMsg(`User ${u} deleted`);
  } catch(e) { alert(e); }
}

// =========================================================================
// --- SPRO ---
// =========================================================================
async function renderSPRO() {
  if (!ME.is_master && ME.role !== 'admin') return setMsg("Permission denied", true);
  const data = await api('/api/spro/overview');

  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col p-6 overflow-hidden select-text">
      <div class="flex items-center justify-between pb-4 border-b border-slate-200 shrink-0">
        <div>
          <h2 class="text-lg font-bold text-slate-900 tracking-tight">Enterprise Structure & Customizing (SPRO)</h2>
          <p class="text-xs text-slate-500">Manage plants, departments, and corporate identity</p>
        </div>
        <button onclick="renderSPRO()" class="btn-secondary text-xs px-3 py-1.5 rounded-lg"><i class="ri-refresh-line"></i> Refresh</button>
      </div>

      <div class="mt-4 card p-4 rounded-2xl flex items-center justify-between shrink-0">
        <div>
          <div class="text-[10px] text-slate-500 uppercase font-semibold tracking-wider">Corporate Identity</div>
          <div class="text-base font-bold text-slate-900">${esc(data.org.name)} <span class="text-xs font-mono text-indigo-600">(${esc(data.org.system_id)} · ${esc(data.org.client)})</span></div>
        </div>
        ${ME.is_master ? `<button onclick="showEditOrgModal('${esc(data.org.name)}','${esc(data.org.system_id)}','${esc(data.org.client)}')" class="btn-secondary text-xs px-3 py-1.5 rounded-lg">Edit</button>` : ''}
      </div>

      <div class="grid grid-cols-1 md:grid-cols-2 gap-4 flex-1 mt-4 overflow-auto">
        <div class="card rounded-2xl flex flex-col overflow-hidden">
          <div class="p-3.5 border-b border-slate-200 flex justify-between items-center bg-slate-50">
            <span class="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
              <i class="ri-building-line text-indigo-500"></i> Plants (${data.plants.length})
            </span>
            ${ME.is_master ? `<button onclick="showCreatePlantModal()" class="btn-primary text-xs font-bold px-3 py-1.5 rounded-lg">+ Add</button>` : ''}
          </div>
          <div class="flex-1 overflow-auto">
            <table class="w-full text-left custom-table">
              <thead><tr><th>Code</th><th>Name</th><th>Status</th>${ME.is_master?'<th style="width:110px;">Actions</th>':''}</tr></thead>
              <tbody>
                ${data.plants.map(p => `
                  <tr>
                    <td class="font-mono font-bold text-indigo-600">${esc(p.code)}</td>
                    <td class="font-medium text-slate-900">${esc(p.name)}</td>
                    <td><span class="badge ${p.active?'badge-emerald':'badge-red'}">${p.active?'ACTIVE':'DISABLED'}</span></td>
                    ${ME.is_master ? `<td>
                      <button onclick="showEditPlantModal('${esc(p.code)}','${esc(p.name)}',${p.active})" class="text-xs text-indigo-600 hover:text-indigo-800 font-semibold">Edit</button>
                      <button onclick="deletePlant('${esc(p.code)}')" class="text-xs text-red-600 hover:text-red-800 font-semibold ml-2">Delete</button>
                    </td>` : ''}
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        </div>

        <div class="card rounded-2xl flex flex-col overflow-hidden">
          <div class="p-3.5 border-b border-slate-200 flex justify-between items-center bg-slate-50">
            <span class="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
              <i class="ri-team-line text-emerald-500"></i> Departments (${data.departments.length})
            </span>
            ${ME.is_master ? `<button onclick="showCreateDeptModal()" class="btn-primary text-xs font-bold px-3 py-1.5 rounded-lg">+ Add</button>` : ''}
          </div>
          <div class="flex-1 overflow-auto">
            <table class="w-full text-left custom-table">
              <thead><tr><th>Name</th><th>Status</th>${ME.is_master?'<th style="width:80px;">Action</th>':''}</tr></thead>
              <tbody>
                ${data.departments.map(d => `
                  <tr>
                    <td class="font-medium text-slate-900">${esc(d.name)}</td>
                    <td><span class="badge ${d.active?'badge-emerald':'badge-red'}">${d.active?'ACTIVE':'DISABLED'}</span></td>
                    ${ME.is_master ? `<td><button onclick="deleteDept('${esc(d.name)}')" class="text-xs text-red-600 hover:text-red-800 font-semibold">Delete</button></td>` : ''}
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>`;
}

function showEditOrgModal(n, sys, cl) {
  openDialog(`
    <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl">
      <h3 class="font-bold text-slate-900 text-sm">Edit Corporate Identity</h3>
      <button onclick="closeDialog()"><i class="ri-close-line text-lg text-slate-500"></i></button>
    </div>
    <div class="p-6 space-y-4">
      <div><label class="form-label">Company Name</label><input id="org_n" value="${esc(n)}" class="form-input"></div>
      <div class="grid grid-cols-2 gap-4">
        <div><label class="form-label">System Identifier</label><input id="org_sys" value="${esc(sys)}" class="form-input font-mono uppercase"></div>
        <div><label class="form-label">Client Code</label><input id="org_cl" value="${esc(cl)}" class="form-input font-mono"></div>
      </div>
    </div>
    <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl">
      <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-lg text-xs font-semibold">Cancel</button>
      <button onclick="submitEditOrg()" class="btn-primary px-4 py-2 rounded-lg text-xs font-semibold">Save</button>
    </div>`);
}

async function submitEditOrg() {
  const fd = new FormData();
  fd.append("name", $('#org_n').value.trim());
  fd.append("system_id", $('#org_sys').value.trim());
  fd.append("client", $('#org_cl').value.trim());
  try {
    await api('/api/admin/org/update', {method: 'POST', body: fd});
    closeDialog();
    renderSPRO();
    setMsg("Identity updated");
  } catch(e) { alert(e); }
}

function showCreatePlantModal() {
  openDialog(`
    <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl">
      <h3 class="font-bold text-slate-900 text-sm">Add Plant</h3>
      <button onclick="closeDialog()"><i class="ri-close-line text-lg text-slate-500"></i></button>
    </div>
    <div class="p-6 space-y-4">
      <div><label class="form-label">Plant Code *</label><input id="np_c" class="form-input font-mono uppercase" placeholder="e.g. 1500"></div>
      <div><label class="form-label">Plant Name *</label><input id="np_n" class="form-input" placeholder="e.g. Copper Drawing Unit"></div>
    </div>
    <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl">
      <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-lg text-xs font-semibold">Cancel</button>
      <button onclick="submitCreatePlant()" class="btn-primary px-4 py-2 rounded-lg text-xs font-semibold">Save</button>
    </div>`);
}

async function submitCreatePlant() {
  const c = $('#np_c').value.trim(), n = $('#np_n').value.trim();
  if (!c || !n) return alert("Code and Name are required.");
  const fd = new FormData(); fd.append("code", c); fd.append("name", n);
  try {
    await api('/api/admin/plants/create', {method: 'POST', body: fd});
    closeDialog();
    renderSPRO();
    setMsg(`Plant ${c} created`);
  } catch(e) { alert(e); }
}

function showEditPlantModal(code, name, active) {
  openDialog(`
    <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl">
      <h3 class="font-bold text-slate-900 text-sm">Edit Plant: ${esc(code)}</h3>
      <button onclick="closeDialog()"><i class="ri-close-line text-lg text-slate-500"></i></button>
    </div>
    <div class="p-6 space-y-4">
      <div><label class="form-label">Name</label><input id="ep_n" value="${esc(name)}" class="form-input"></div>
      <div>
        <label class="form-label">Status</label>
        <select id="ep_st" class="form-input">
          <option value="1" ${active?'selected':''}>Active</option>
          <option value="0" ${!active?'selected':''}>Disabled</option>
        </select>
      </div>
    </div>
    <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl">
      <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-lg text-xs font-semibold">Cancel</button>
      <button onclick="submitEditPlant('${esc(code)}')" class="btn-primary px-4 py-2 rounded-lg text-xs font-semibold">Save</button>
    </div>`);
}

async function submitEditPlant(code) {
  const fd = new FormData();
  fd.append("code", code);
  fd.append("name", $('#ep_n').value.trim());
  fd.append("active", $('#ep_st').value);
  try {
    await api('/api/admin/plants/update', {method: 'POST', body: fd});
    closeDialog();
    renderSPRO();
    setMsg(`Plant ${code} updated`);
  } catch(e) { alert(e); }
}

async function deletePlant(code) {
  if (!confirm(`Permanently remove plant ${code}?`)) return;
  const fd = new FormData(); fd.append("code", code);
  try {
    await api('/api/admin/plants/delete', {method: 'POST', body: fd});
    renderSPRO();
    setMsg(`Plant ${code} deleted`);
  } catch(e) { alert(e); }
}

function showCreateDeptModal() {
  const d = prompt("Enter New Department Name:");
  if (!d || !d.trim()) return;
  const fd = new FormData(); fd.append("name", d.trim());
  api('/api/admin/departments/create', {method: 'POST', body: fd}).then(() => {
    renderSPRO();
    setMsg("Department added");
  }).catch(e => alert(e));
}

async function deleteDept(name) {
  if (!confirm(`Delete department ${name}?`)) return;
  const fd = new FormData(); fd.append("name", name);
  try {
    await api('/api/admin/departments/delete', {method: 'POST', body: fd});
    renderSPRO();
    setMsg("Department removed");
  } catch(e) { alert(e); }
}

// =========================================================================
// --- PFCG ---
// =========================================================================
async function renderPFCG() {
  if (!ME.is_master && ME.role !== 'admin') return setMsg("Permission denied", true);
  const m = await api('/api/pfcg/matrix');
  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col p-6 overflow-hidden select-text">
      <div class="flex items-center justify-between pb-4 border-b border-slate-200 shrink-0">
        <div><h2 class="text-lg font-bold text-slate-900 tracking-tight">Access Control Matrix (PFCG)</h2></div>
        <div class="flex items-center gap-2">
          <select id="pfcg_u" class="form-input w-56">
            ${m.users.map(u=>`<option value="${esc(u.username)}">${esc(u.username)} (${esc(u.full_name)})</option>`).join('')}
          </select>
          <select id="pfcg_t" class="form-input w-56">
            ${m.tcodes.map(t=>`<option value="${esc(t.code)}">${esc(t.code)} - ${esc(t.title)}</option>`).join('')}
          </select>
          <button onclick="grantAccess()" class="btn-primary text-xs font-semibold px-3 py-2 rounded-lg">+ Grant</button>
        </div>
      </div>
      <div class="flex-1 mt-4 card rounded-2xl overflow-auto">
        <table class="w-full text-left custom-table">
          <thead><tr><th>User</th><th>Module</th><th>Granted By</th><th>Granted At</th><th>Action</th></tr></thead>
          <tbody>
            ${m.grants.map(g => `
              <tr>
                <td class="font-mono font-bold text-indigo-600">${esc(g.username)}</td>
                <td class="font-mono font-bold text-slate-900">${esc(g.tcode_code)}</td>
                <td>${esc(g.granted_by)}</td>
                <td class="text-slate-500 text-xs font-mono">${esc(g.granted_at)}</td>
                <td><button onclick="revokeAccess('${esc(g.username)}','${esc(g.tcode_code)}')" class="text-xs text-red-600 hover:text-red-800 font-semibold">Revoke</button></td>
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
  if (!confirm(`Revoke ${t} from ${u}?`)) return;
  const fd = new FormData(); fd.append("username", u); fd.append("tcode", t);
  await api('/api/pfcg/revoke', {method: 'POST', body: fd});
  renderPFCG();
  setMsg("Permission revoked");
}

// =========================================================================
// --- SE93 ---
// =========================================================================
async function renderSE93() {
  if (!ME.is_master && ME.role !== 'admin') return setMsg("Permission denied", true);
  const list = await api('/api/se93/list');
  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col p-6 overflow-hidden select-text">
      <div class="flex items-center justify-between pb-4 border-b border-slate-200 shrink-0">
        <div>
          <h2 class="text-lg font-bold text-slate-900 tracking-tight">Module Studio (SE93)</h2>
          <p class="text-xs text-slate-500">Upload Excel/CSV/JSON to auto-create a T-code</p>
        </div>
        <div class="flex gap-2">
          <button onclick="openTcode('STUDIO')" class="btn-secondary text-xs px-3 py-2 rounded-lg flex items-center gap-1.5"><i class="ri-robot-2-line text-emerald-500"></i> Use AI Studio</button>
          <button onclick="renderSE93()" class="btn-secondary text-xs px-3 py-2 rounded-lg"><i class="ri-refresh-line"></i></button>
        </div>
      </div>
      <div class="mt-4 grid grid-cols-1 lg:grid-cols-3 gap-4 flex-1 overflow-hidden">
        <div class="lg:col-span-1 card rounded-2xl p-4 overflow-auto">
          <h3 class="text-sm font-bold text-slate-900 mb-3 flex items-center gap-1.5"><i class="ri-upload-cloud-2-line text-indigo-500"></i> Upload Template</h3>
          <form id="se93_form" onsubmit="submitSE93(event)" class="space-y-3">
            <div><label class="form-label">T-Code</label><input name="code" required placeholder="e.g. DISP1" class="form-input font-mono uppercase"></div>
            <div><label class="form-label">Title</label><input name="title" required placeholder="Dispatch Log" class="form-input"></div>
            <div><label class="form-label">Description</label><input name="description" placeholder="Optional" class="form-input"></div>
            <div><label class="form-label">Module</label><input name="module" value="PROD" class="form-input font-mono uppercase"></div>
            <div><label class="form-label">Allowed Depts (CSV)</label><input name="allowed_depts" value="Production,Quality Control" class="form-input"></div>
            <div><label class="form-label">Import Sample Data?</label>
              <select name="import_sample_data" class="form-input">
                <option value="1">Yes</option><option value="0">No</option>
              </select>
            </div>
            <div><label class="form-label">Template File</label><input name="file" type="file" required class="w-full text-xs text-slate-600 file:mr-2 file:bg-indigo-600 file:text-white file:border-0 file:rounded file:px-2 file:py-1 file:text-xs"></div>
            <button type="submit" class="btn-primary w-full text-xs font-bold py-2 rounded-lg">Create Module</button>
          </form>
        </div>
        <div class="lg:col-span-2 card rounded-2xl overflow-hidden flex flex-col">
          <div class="p-3 border-b border-slate-200 bg-slate-50 flex justify-between items-center">
            <span class="text-xs font-bold text-slate-700 uppercase tracking-wider">Existing Modules (${list.length})</span>
          </div>
          <div class="flex-1 overflow-auto">
            <table class="w-full text-left custom-table">
              <thead><tr><th>Code</th><th>Title</th><th>Module</th><th>Layouts</th><th>Records</th><th>Action</th></tr></thead>
              <tbody>
                ${list.map(t => `
                  <tr>
                    <td class="font-mono font-bold text-indigo-600">${esc(t.code)}</td>
                    <td class="font-medium text-slate-900">${esc(t.title)}</td>
                    <td class="text-xs text-slate-500 uppercase">${esc(t.module||'-')}</td>
                    <td class="font-mono">${t.layout_count}</td>
                    <td class="font-mono">${t.record_count}</td>
                    <td class="space-x-2">
                      <button onclick="openTcode('${esc(t.code)}')" class="text-xs text-indigo-600 font-semibold">Open</button>
                      <button onclick="loadStudioExisting('${esc(t.code)}')" class="text-xs text-emerald-600 font-semibold">Edit in Studio</button>
                      ${ME.is_master ? `<button onclick="deleteSE93('${esc(t.code)}')" class="text-xs text-red-600 font-semibold">Delete</button>` : ''}
                    </td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>`;
}

async function submitSE93(ev) {
  ev.preventDefault();
  const fd = new FormData(ev.target);
  try {
    const res = await fetch('/api/se93/auto-create', {method: 'POST', body: fd, headers: {'Authorization': 'Bearer ' + (localStorage.getItem("nexus_token")||'')}});
    const data = await res.json();
    if (!res.ok) throw (data.detail || "Failed");
    alert(`Module ${data.code} created with ${data.columns_count} columns and ${data.imported_count} sample rows.`);
    renderSE93();
    loadModules();
  } catch(e) { alert("Error: " + e); }
}

async function deleteSE93(code) {
  if (!confirm(`Delete module ${code} and all its data?`)) return;
  const fd = new FormData(); fd.append("code", code);
  try {
    await api('/api/se93/delete', {method: 'POST', body: fd});
    renderSE93();
    loadModules();
    setMsg(`Module ${code} deleted`);
  } catch(e) { alert(e); }
}

// =========================================================================
// --- AI CODE STUDIO ---
// =========================================================================
const STUDIO_EXAMPLE = {
  action: "upsert",
  code: "DISP0",
  title: "Dispatch & Delivery Challan",
  description: "Outward material dispatch log",
  module: "SD",
  allowed_depts: "Commercial,Store,Planning",
  fields: [
    { key: "sr_no",        label: "Sr. No.",             type: "text",   width: 1 },
    { key: "challan_no",   label: "Challan No.",         type: "text",   required: true },
    { key: "customer",     label: "Customer",            type: "text" },
    { key: "invoice_no",   label: "Invoice No.",         type: "text" },
    { key: "dispatch_date",label: "Dispatch Date",       type: "date" },
    { key: "qty",          label: "Quantity (Mtrs/Kgs)", type: "number" },
    { key: "vehicle_no",   label: "Vehicle No.",         type: "text" },
    { key: "transporter",  label: "Transporter",         type: "text" },
    { key: "remarks",      label: "Remarks",             type: "textarea", width: 2 }
  ],
  seed_rows: [], presets: []
};

async function renderStudio() {
  if (!ME.is_master) return setMsg("Master Administrator only", true);
  const hist = await api('/api/studio/history');

  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col p-6 overflow-hidden select-text">
      <div class="flex items-center justify-between pb-4 border-b border-slate-200 shrink-0">
        <div>
          <h2 class="text-lg font-bold text-slate-900 tracking-tight flex items-center gap-2">
            <i class="ri-robot-2-line text-emerald-500"></i> AI Code Studio — Paste-to-Build
          </h2>
          <p class="text-xs text-slate-500">Paste a JSON module spec. Create, update, or delete T-codes live.</p>
        </div>
        <div class="flex items-center gap-2">
          <button onclick="loadStudioExample()" class="btn-secondary text-xs px-3 py-1.5 rounded-lg"><i class="ri-magic-line"></i> Example</button>
          <button onclick="renderStudio()" class="btn-secondary text-xs px-3 py-1.5 rounded-lg"><i class="ri-refresh-line"></i></button>
        </div>
      </div>

      <div class="grid grid-cols-1 lg:grid-cols-2 gap-4 flex-1 mt-4 overflow-hidden">
        <div class="flex flex-col card rounded-2xl overflow-hidden">
          <div class="p-3 border-b border-slate-200 bg-slate-50 flex justify-between items-center">
            <span class="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
              <i class="ri-code-s-slash-line text-indigo-500"></i> Module Spec (JSON)
            </span>
            <div class="flex gap-1.5">
              <button onclick="studioValidate()" class="text-[11px] font-bold px-2.5 py-1 rounded-md bg-amber-50 border border-amber-200 text-amber-700">Validate</button>
              <button onclick="studioCopyCurrent()" class="text-[11px] font-semibold px-2.5 py-1 rounded-md bg-white border border-slate-200 text-slate-600">Copy</button>
              <button onclick="studioFormat()" class="text-[11px] font-semibold px-2.5 py-1 rounded-md bg-white border border-slate-200 text-slate-600">Prettify</button>
            </div>
          </div>
          <textarea id="studio_json" spellcheck="false"
            class="flex-1 bg-slate-50 text-emerald-700 font-mono text-[11.5px] leading-relaxed p-3 outline-none resize-none"
            placeholder='Paste module JSON here...'></textarea>
          <div class="p-3 border-t border-slate-200 bg-slate-50 flex items-center justify-between">
            <span id="studio_status" class="text-[11px] font-mono text-slate-500">Idle</span>
            <button onclick="studioApply()" class="btn-primary text-xs font-bold px-4 py-1.5 rounded-md flex items-center gap-1">
              <i class="ri-play-fill"></i> Apply Spec
            </button>
          </div>
        </div>

        <div class="flex flex-col gap-4 overflow-hidden">
          <div id="studio_preview" class="card rounded-2xl p-4 overflow-auto flex-1 min-h-[200px]">
            <div class="text-xs text-slate-400 text-center pt-10">Validate a spec to see live preview.</div>
          </div>
          <div class="card rounded-2xl overflow-hidden" style="max-height:44%;">
            <div class="p-2.5 border-b border-slate-200 bg-slate-50 text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
              <i class="ri-history-line text-indigo-500"></i> Version History
            </div>
            <div class="overflow-auto" style="max-height:200px;">
              <table class="w-full text-left custom-table">
                <thead><tr><th>v</th><th>Module</th><th>Action</th><th>By</th><th>When</th><th></th></tr></thead>
                <tbody>
                  ${!hist.length ? `<tr><td colspan="6" class="text-center py-6 text-slate-400 text-xs">No installs yet.</td></tr>` :
                    hist.map(h => `<tr>
                      <td class="font-mono text-indigo-600 font-bold">v${h.version}</td>
                      <td class="font-mono font-bold text-slate-900">${esc(h.tcode_code)}</td>
                      <td><span class="badge ${h.action==='delete'?'badge-red':'badge-emerald'}">${h.action}</span></td>
                      <td class="text-xs">${esc(h.applied_by)}</td>
                      <td class="text-xs font-mono text-slate-500">${esc(h.applied_at)}</td>
                      <td>
                        <button onclick="studioRollback(${h.id})" class="text-[11px] text-amber-600 font-semibold">Rollback</button>
                        <button onclick="studioLoadExisting('${esc(h.tcode_code)}')" class="text-[11px] text-indigo-600 font-semibold ml-2">Load</button>
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

function loadStudioExample() {
  $('#studio_json').value = JSON.stringify(STUDIO_EXAMPLE, null, 2);
  setStudioStatus("Example loaded — tweak and Apply.", false);
}

function setStudioStatus(msg, isErr) {
  const el = $('#studio_status');
  if (!el) return;
  el.textContent = msg;
  el.className = "text-[11px] font-mono " + (isErr ? "text-red-600" : "text-emerald-600");
}

function studioFormat() {
  try {
    const v = JSON.parse($('#studio_json').value);
    $('#studio_json').value = JSON.stringify(v, null, 2);
    setStudioStatus("Formatted OK.", false);
  } catch(e) { setStudioStatus("Invalid JSON: " + e.message, true); }
}

function studioCopyCurrent() {
  navigator.clipboard.writeText($('#studio_json').value);
  setStudioStatus("Copied to clipboard.", false);
}

async function studioValidate() {
  let spec;
  try { spec = JSON.parse($('#studio_json').value); }
  catch(e) { setStudioStatus("Invalid JSON: " + e.message, true); return null; }

  try {
    const res = await api('/api/studio/validate', {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({spec})
    });
    if (!res.ok) {
      setStudioStatus("Errors: " + res.errors.join(" | "), true);
      renderStudioPreview(res.clean, res.errors);
      return res;
    }
    setStudioStatus(res.will_overwrite ? "Valid — will UPDATE existing module." : "Valid — will CREATE new module.", false);
    renderStudioPreview(res.clean, []);
    return res;
  } catch(e) { setStudioStatus("Error: " + e, true); return null; }
}

function renderStudioPreview(clean, errors) {
  const box = $('#studio_preview');
  if (!box) return;
  if (!clean || errors.length) {
    box.innerHTML = `<div class="text-xs text-red-600 whitespace-pre-wrap">${esc(errors.join("\\n"))}</div>`;
    return;
  }
  if (clean.action === "delete") {
    box.innerHTML = `<div class="text-sm"><b class="text-red-600">DELETE MODULE:</b> <span class="font-mono text-slate-900">${esc(clean.code)}</span></div>
      <div class="text-xs text-slate-500 mt-2">All schemas, layouts, records, and grants will be removed.</div>`;
    return;
  }
  box.innerHTML = `
    <div class="mb-3">
      <div class="text-[10px] uppercase font-bold text-slate-400">Preview</div>
      <div class="text-base font-bold text-slate-900">${esc(clean.title)} <span class="font-mono text-indigo-600 text-xs">[${esc(clean.code)}]</span></div>
      <div class="text-xs text-slate-500 mt-0.5">${esc(clean.description || '—')}</div>
      <div class="text-[11px] text-slate-500 mt-1">Module: <b class="text-slate-700">${esc(clean.module)}</b> · Depts: <span class="text-slate-700">${esc(clean.allowed_depts || 'none')}</span></div>
    </div>
    <div class="text-[10px] uppercase font-bold text-slate-400 mb-1">Columns (${clean.fields.length})</div>
    <table class="w-full text-left text-[11px]">
      <thead><tr class="bg-slate-50 text-slate-500"><th class="p-1.5">Key</th><th class="p-1.5">Label</th><th class="p-1.5">Type</th></tr></thead>
      <tbody>${clean.fields.map(f => `<tr class="border-b border-slate-100"><td class="p-1.5 font-mono text-indigo-600">${esc(f.key)}</td><td class="p-1.5 text-slate-900">${esc(f.label)}</td><td class="p-1.5 text-slate-500">${esc(f.type)}</td></tr>`).join('')}</tbody>
    </table>`;
}

async function studioApply() {
  let spec;
  try { spec = JSON.parse($('#studio_json').value); }
  catch(e) { setStudioStatus("Invalid JSON: " + e.message, true); return; }

  if (!confirm(`Apply spec for module ${spec.code}? This will ${spec.action==='delete'?'DELETE':'install/update'} it now.`)) return;

  try {
    const res = await api('/api/studio/apply', {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({spec})
    });
    setStudioStatus(`OK — ${res.action} ${res.code} (v${res.version})`, false);
    await loadModules();
    renderStudio();
  } catch(e) { setStudioStatus("Error: " + e, true); alert(e); }
}

async function studioRollback(msid) {
  if (!confirm(`Rollback this version? A new version will be created from the stored spec.`)) return;
  try {
    const res = await api(`/api/studio/rollback/${msid}`, {method: 'POST'});
    setStudioStatus(`Rolled back — ${res.action} ${res.code} (v${res.version})`, false);
    await loadModules();
    renderStudio();
  } catch(e) { alert(e); }
}

async function studioLoadExisting(code) {
  try {
    const spec = await api(`/api/studio/spec/${code}`);
    $('#studio_json').value = JSON.stringify(spec, null, 2);
    setStudioStatus(`Loaded current spec of ${code}. Edit and re-apply to update.`, false);
    await studioValidate();
  } catch(e) { alert(e); }
}

// =========================================================================
// --- TN01 WORKSPACE ---
// =========================================================================
let TN01_ROUTES = {}, TN01_ACTIVE_ROUTE = "PROD_TO_QC", TN01_HEADERS = [], TN01_ROWS = [], TN01_UNDO_STACK = [], TN01_REDO_STACK = [], TN01_ACTIVE_CELL = { r: 0, c: 1 };

async function renderTN01() {
  TN01_ROUTES = await api('/api/tn01/routes');
  renderTN01Main();
}

function renderTN01Main(activeTab = "sheet") {
  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col overflow-hidden select-text">
      <div class="p-3 bg-white border-b border-slate-200 flex items-center justify-between shrink-0">
        <div class="flex items-center gap-3">
          <div class="flex items-center gap-2">
            <span class="font-mono font-bold text-xs bg-indigo-600 text-white px-2.5 py-1 rounded-lg shadow-sm">TN01</span>
            <div><h2 class="text-sm font-bold text-slate-900">Transfer Note & Handover System</h2><p class="text-[10px] text-slate-500">Inter-department batch transfer slips</p></div>
          </div>
          <div class="h-6 w-[1px] bg-slate-200 mx-2"></div>
          <div class="flex items-center bg-slate-100 border border-slate-200 rounded-lg p-1 gap-1 text-xs font-semibold">
            <button onclick="switchTN01Tab('sheet')" class="px-3 py-1.5 rounded-md flex items-center gap-1.5 transition ${activeTab==='sheet'?'bg-white text-indigo-600 shadow-sm':'text-slate-500 hover:text-slate-700'}"><i class="ri-table-line"></i> Excel Sheet</button>
            <button onclick="switchTN01Tab('receive')" class="px-3 py-1.5 rounded-md flex items-center gap-1.5 transition ${activeTab==='receive'?'bg-white text-indigo-600 shadow-sm':'text-slate-500 hover:text-slate-700'}"><i class="ri-inbox-archive-line"></i> Receive Notes</button>
            <button onclick="switchTN01Tab('history')" class="px-3 py-1.5 rounded-md flex items-center gap-1.5 transition ${activeTab==='history'?'bg-white text-indigo-600 shadow-sm':'text-slate-500 hover:text-slate-700'}"><i class="ri-history-line"></i> History & Log</button>
            ${ME.is_master || ME.role==='admin' ? `<button onclick="switchTN01Tab('format')" class="px-3 py-1.5 rounded-md flex items-center gap-1.5 transition ${activeTab==='format'?'bg-white text-indigo-600 shadow-sm':'text-slate-500 hover:text-slate-700'}"><i class="ri-paint-brush-line"></i> Format Studio</button>` : ''}
          </div>
        </div>
        <div>${activeTab === 'sheet' ? `<button onclick="openSendModal()" class="btn-primary text-xs font-bold px-4 py-2 rounded-lg flex items-center gap-1.5 shadow-sm"><i class="ri-send-plane-fill"></i> Send Transfer Note</button>` : ''}</div>
      </div>
      <div id="tn01_content" class="flex-1 overflow-hidden flex flex-col"></div>
    </div>`;

  if (activeTab === 'sheet') renderTN01SheetTab();
  else if (activeTab === 'receive') renderTN01ReceiveTab();
  else if (activeTab === 'history') renderTN01HistoryTab();
  else if (activeTab === 'format') renderTN01FormatTab();
}

function switchTN01Tab(t) { renderTN01Main(t); }

function renderTN01SheetTab() {
  if (!TN01_HEADERS.length) applyRouteLayout(TN01_ACTIVE_ROUTE, false);
  $('#tn01_content').innerHTML = `
    <div class="p-2 bg-white border-b border-slate-200 flex items-center justify-between text-xs shrink-0">
      <div class="flex items-center gap-2">
        <label class="font-bold text-slate-600 text-xs flex items-center gap-1.5">Preset Route:
          <select id="route_selector" onchange="applyRouteLayout(this.value, true)" class="form-input w-auto py-1 font-semibold">
            ${Object.keys(TN01_ROUTES).map(k => `<option value="${k}" ${k===TN01_ACTIVE_ROUTE?'selected':''}>${TN01_ROUTES[k].route}</option>`).join('')}
          </select>
        </label>
        <div class="h-4 w-[1px] bg-slate-200 mx-1"></div>
        <button onclick="undoGrid()" class="bg-white text-slate-600 px-2 py-1 rounded-md border border-slate-200 hover:bg-slate-50"><i class="ri-arrow-go-back-line"></i></button>
        <button onclick="redoGrid()" class="bg-white text-slate-600 px-2 py-1 rounded-md border border-slate-200 hover:bg-slate-50"><i class="ri-arrow-go-forward-line"></i></button>
        <button onclick="addGridRow()" class="bg-white text-slate-600 px-2.5 py-1 rounded-md border border-slate-200 hover:bg-slate-50">+ Row</button>
        <button onclick="deleteGridRow()" class="bg-white text-red-500 px-2.5 py-1 rounded-md border border-slate-200 hover:bg-red-50">− Row</button>
        <button onclick="clearGrid()" class="bg-white text-slate-500 px-2 py-1 rounded-md border border-slate-200 hover:bg-slate-50">Clear</button>
      </div>
      <div class="flex items-center gap-2">
        <button onclick="exportGridAsExcel()" class="bg-white text-emerald-700 border border-emerald-200 px-3 py-1 rounded-md flex items-center gap-1 font-semibold hover:bg-emerald-50"><i class="ri-file-excel-2-line"></i> Export as Excel</button>
        <span class="text-[11px] bg-slate-50 border border-slate-200 px-2 py-0.5 rounded text-slate-500 font-mono">Press <b>Ctrl+V</b> to paste from Excel</span>
      </div>
    </div>
    <div id="grid_table_container" class="flex-1 overflow-auto p-2 select-text" tabindex="0"></div>
    <div class="p-2 bg-white border-t border-slate-200 text-slate-500 text-[11px] flex justify-between items-center font-mono shrink-0">
      <span id="sheet_count_lbl">0 Batches loaded</span><span>Click cell to edit • Auto-serial prefilled</span>
    </div>`;
  renderGridTable();
  setupPasteHandler();
}

function applyRouteLayout(routeKey, resetData = true) {
  TN01_ACTIVE_ROUTE = routeKey;
  TN01_HEADERS = [...TN01_ROUTES[routeKey].columns];
  if (resetData || !TN01_ROWS.length) {
    TN01_ROWS = [];
    for (let i = 1; i <= 8; i++) {
      const row = new Array(TN01_HEADERS.length).fill("");
      row[0] = i;
      TN01_ROWS.push(row);
    }
  }
  renderGridTable();
}

function renderGridTable() {
  const c = $('#grid_table_container');
  if (!c) return;
  let html = `<table class="excel-grid"><thead><tr>${TN01_HEADERS.map(h=>`<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>`;
  TN01_ROWS.forEach((row, rIdx) => {
    html += `<tr>${row.map((cell, cIdx) => `<td contenteditable="${cIdx!==0}" class="${cIdx===0?'sr-col':''}" data-r="${rIdx}" data-c="${cIdx}" onfocus="TN01_ACTIVE_CELL={r:${rIdx},c:${cIdx}}" onblur="updateGridCell(${rIdx},${cIdx},this.innerText)">${esc(cell)}</td>`).join('')}</tr>`;
  });
  html += `</tbody></table>`;
  c.innerHTML = html;
  const countFilled = TN01_ROWS.filter(r => r.slice(1).some(x => String(x).trim() !== '')).length;
  if ($('#sheet_count_lbl')) $('#sheet_count_lbl').textContent = `${countFilled} Filled Batches (${TN01_ROWS.length} Total Rows)`;
}

function updateGridCell(r, col, val) {
  if (col === 0) return;
  if (TN01_ROWS[r][col] !== val.trim()) {
    TN01_UNDO_STACK.push(JSON.stringify(TN01_ROWS));
    TN01_ROWS[r][col] = val.trim();
  }
}

function undoGrid() { if (!TN01_UNDO_STACK.length) return; TN01_REDO_STACK.push(JSON.stringify(TN01_ROWS)); TN01_ROWS = JSON.parse(TN01_UNDO_STACK.pop()); renderGridTable(); }
function redoGrid() { if (!TN01_REDO_STACK.length) return; TN01_UNDO_STACK.push(JSON.stringify(TN01_ROWS)); TN01_ROWS = JSON.parse(TN01_REDO_STACK.pop()); renderGridTable(); }

function addGridRow() {
  TN01_UNDO_STACK.push(JSON.stringify(TN01_ROWS));
  const r = new Array(TN01_HEADERS.length).fill("");
  r[0] = TN01_ROWS.length + 1;
  TN01_ROWS.push(r);
  renderGridTable();
}
function deleteGridRow() { if (TN01_ROWS.length <= 1) return; TN01_UNDO_STACK.push(JSON.stringify(TN01_ROWS)); TN01_ROWS.pop(); renderGridTable(); }
function clearGrid() {
  if (!confirm("Clear all rows in the sheet?")) return;
  TN01_UNDO_STACK.push(JSON.stringify(TN01_ROWS));
  TN01_ROWS = [];
  for (let i = 1; i <= 6; i++) { const r = new Array(TN01_HEADERS.length).fill(""); r[0] = i; TN01_ROWS.push(r); }
  renderGridTable();
}

function setupPasteHandler() {
  const container = $('#grid_table_container');
  if (!container) return;
  container.addEventListener('paste', function(e) {
    e.preventDefault();
    const clipData = (e.clipboardData || window.clipboardData).getData('text');
    if (!clipData) return;
    TN01_UNDO_STACK.push(JSON.stringify(TN01_ROWS));
    const lines = clipData.replace(/\\r\\n/g, '\\n').replace(/\\r/g, '\\n').split('\\n');
    let startR = TN01_ACTIVE_CELL.r || 0;
    let startC = TN01_ACTIVE_CELL.c || 1;
    if (startC === 0) startC = 1;
    lines.forEach((line, lIdx) => {
      if (!line && lIdx === lines.length - 1) return;
      const cells = line.split('\\t');
      const targetR = startR + lIdx;
      while (TN01_ROWS.length <= targetR) { const nr = new Array(TN01_HEADERS.length).fill(""); nr[0] = TN01_ROWS.length + 1; TN01_ROWS.push(nr); }
      cells.forEach((val, cIdx) => {
        const targetC = startC + cIdx;
        if (targetC < TN01_HEADERS.length) TN01_ROWS[targetR][targetC] = val.trim();
      });
    });
    TN01_ROWS.forEach((r, idx) => { r[0] = idx + 1; });
    renderGridTable();
    setMsg("Excel data pasted.");
  });
}

async function exportGridAsExcel() {
  const filledRows = TN01_ROWS.filter(r => r.slice(1).some(x => String(x).trim() !== ''));
  if (!filledRows.length) return alert("No entries in the grid to export.");
  const routeTitle = TN01_ROUTES[TN01_ACTIVE_ROUTE] ? TN01_ROUTES[TN01_ACTIVE_ROUTE].route : "Transfer Note";
  const payload = { headers: TN01_HEADERS, rows: filledRows, title: `TN01 Draft — ${routeTitle}` };
  try {
    const token = localStorage.getItem("nexus_token");
    const res = await fetch('/api/tn01/export-grid', {
      method: 'POST',
      headers: {'Content-Type': 'application/json', 'Authorization': token ? `Bearer ${token}` : ''},
      body: JSON.stringify(payload)
    });
    if (!res.ok) { let t = await res.text(); try { t = JSON.parse(t).detail || t; } catch {} throw t; }
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const cd = res.headers.get('Content-Disposition') || '';
    const m = cd.match(/filename=([^;]+)/);
    a.download = m ? m[1].replace(/["']/g, '') : 'TN01_Draft.xlsx';
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(url);
    setMsg("Draft exported as Excel");
  } catch(e) { alert("Export failed: " + e); }
}

async function openSendModal() {
  const activeRoute = TN01_ROUTES[TN01_ACTIVE_ROUTE];
  const filledRows = TN01_ROWS.filter(r => r.slice(1).some(x => String(x).trim() !== ''));
  if (!filledRows.length) return alert("Please enter at least one entry before sending.");
  const spro = await api('/api/spro/overview');
  const myPlant = (ME.plants && ME.plants[0]) ? ME.plants[0] : "1000";

  openDialog(`
    <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl">
      <div><h3 class="font-bold text-slate-900 text-sm">Transfer & Handover Note</h3><p class="text-xs text-slate-500">Route: ${activeRoute.route}</p></div>
      <button onclick="closeDialog()"><i class="ri-close-line text-lg text-slate-500"></i></button>
    </div>
    <div class="p-6 space-y-4">
      <div class="grid grid-cols-3 gap-3">
        <div><label class="form-label">Target Plant</label>
          <select id="send_plant" onchange="updateSuggestedUsers()" class="form-input">
            ${spro.plants.map(p=>`<option value="${esc(p.code)}" ${p.code===myPlant?'selected':''}>${esc(p.code)} - ${esc(p.name)}</option>`).join('')}
          </select>
        </div>
        <div><label class="form-label">Target Department</label>
          <select id="send_dept" onchange="updateSuggestedUsers()" class="form-input">
            ${spro.departments.map(d=>`<option value="${esc(d.name)}" ${d.name===activeRoute.to_dept?'selected':''}>${esc(d.name)}</option>`).join('')}
          </select>
        </div>
        <div><label class="form-label">Assigned Receiver</label>
          <select id="send_user" class="form-input"></select>
        </div>
      </div>
      <div><label class="form-label">Remarks</label>
        <textarea id="send_remarks" class="form-input" rows="2" placeholder="Write something here..."></textarea>
      </div>
    </div>
    <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl">
      <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-lg text-xs font-semibold">Cancel</button>
      <button onclick="submitSendNote()" class="btn-primary px-5 py-2 rounded-lg text-xs font-bold">Confirm & Send</button>
    </div>`);
  updateSuggestedUsers();
}

async function updateSuggestedUsers() {
  const pl = $('#send_plant').value, dept = $('#send_dept').value;
  const sel = $('#send_user');
  const users = await api(`/api/tn01/suggest-users?plant=${pl}&dept=${encodeURIComponent(dept)}`);
  sel.innerHTML = users.length ? users.map(u => `<option value="${esc(u.username)}">${esc(u.full_name)} (${esc(u.username)})</option>`).join('') : `<option value="">(Department Pool)</option>`;
}

async function submitSendNote() {
  const filledRows = TN01_ROWS.filter(r => r.slice(1).some(c => String(c).trim() !== ''));
  const payload = {
    layout_key: TN01_ACTIVE_ROUTE, route: TN01_ROUTES[TN01_ACTIVE_ROUTE].route,
    to_plant: $('#send_plant').value, to_dept: $('#send_dept').value,
    to_user: $('#send_user').value, remarks: $('#send_remarks').value,
    headers: TN01_HEADERS, rows: filledRows
  };
  try {
    const res = await api('/api/tn01/send', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(payload)});
    closeDialog();
    alert(`Transfer note shared! Note #: ${res.note_no}`);
    switchTN01Tab('history');
  } catch(e) { alert(e); }
}

async function renderTN01ReceiveTab() {
  const notes = await api('/api/tn01/inbox');
  $('#tn01_content').innerHTML = `
    <div class="p-6 flex-1 flex flex-col overflow-hidden">
      <div class="flex items-center justify-between pb-4 border-b border-slate-200 shrink-0">
        <div><h2 class="text-base font-bold text-slate-900">Incoming Transfer Notes</h2></div>
        <button onclick="renderTN01ReceiveTab()" class="btn-secondary text-xs px-3 py-1.5 rounded-lg"><i class="ri-refresh-line"></i> Refresh</button>
      </div>
      <div class="flex-1 mt-4 card rounded-2xl overflow-auto">
        <table class="w-full text-left custom-table">
          <thead><tr><th>Note #</th><th>Route</th><th>Origin</th><th>Sender</th><th>Batches</th><th>Time</th><th>Status</th><th>Actions</th></tr></thead>
          <tbody>
            ${!notes.length ? `<tr><td colspan="8" class="text-center py-12 text-slate-400">Inbox is empty.</td></tr>` : ''}
            ${notes.map(n => `
              <tr>
                <td class="font-mono font-bold text-indigo-600">${esc(n.note_no)}</td>
                <td class="font-semibold text-slate-900">${esc(n.route)}</td>
                <td>${esc(n.from_plant)} - ${esc(n.from_dept)}</td>
                <td>${esc(n.from_user_name)}</td>
                <td class="font-mono font-bold">${n.entry_count}</td>
                <td class="font-mono text-xs text-slate-500">${esc(n.created_at)}</td>
                <td><span class="badge ${n.status==='received'?'badge-emerald':'badge-amber'}">${n.status.toUpperCase()}</span></td>
                <td class="space-x-1">
                  <button onclick="viewTransferNote(${n.id})" class="btn-primary text-xs font-semibold px-2.5 py-1 rounded-md">Open</button>
                  ${n.status==='sent' ? `<button onclick="acknowledgeNote(${n.id})" class="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold px-2.5 py-1 rounded-md">Acknowledge</button>` : ''}
                </td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    </div>`;
}

async function acknowledgeNote(nid) {
  await api(`/api/tn01/receive/${nid}`, {method: 'POST'});
  setMsg("Acknowledged");
  renderTN01ReceiveTab();
}

async function viewTransferNote(nid) {
  const note = await api(`/api/tn01/note/${nid}`);
  openDialog(`
    <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl">
      <div><span class="text-xs bg-indigo-50 text-indigo-700 border border-indigo-200 font-mono font-bold px-2 py-0.5 rounded">${esc(note.note_no)}</span><h3 class="font-bold text-slate-900 text-sm mt-1">${esc(note.route)}</h3></div>
      <button onclick="closeDialog()"><i class="ri-close-line text-lg text-slate-500"></i></button>
    </div>
    <div class="p-6 space-y-4">
      <div class="flex justify-between items-center">
        <span class="text-xs font-bold text-slate-700">Batches (${note.entry_count} rows):</span>
        <div class="flex gap-2">
          <a href="/api/tn01/export/${nid}" class="btn-secondary text-xs px-3 py-1.5 rounded-md flex items-center gap-1 text-emerald-700 font-semibold border-emerald-200 hover:bg-emerald-50"><i class="ri-file-excel-2-line"></i> Export as Excel</a>
          <button onclick="loadNoteIntoWorkspace(${nid})" class="btn-primary text-xs px-3 py-1.5 rounded-md flex items-center gap-1 font-semibold"><i class="ri-download-cloud-2-line"></i> Load into Workspace</button>
        </div>
      </div>
      <div class="max-h-72 overflow-auto border border-slate-200 rounded-xl">
        <table class="w-full text-left text-xs">
          <thead><tr class="bg-slate-50 text-slate-700 font-bold">${note.sheet_headers.map(h => `<th class="p-2 border-r border-slate-200">${esc(h)}</th>`).join('')}</tr></thead>
          <tbody>${note.sheet_data.map(r => `<tr class="border-b border-slate-100">${r.map(c => `<td class="p-2 border-r border-slate-100">${esc(c)}</td>`).join('')}</tr>`).join('')}</tbody>
        </table>
      </div>
    </div>
    <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-between items-center rounded-b-2xl">
      <a href="/api/tn01/print/${note.id}" target="_blank" class="btn-secondary text-xs px-3 py-2 rounded-md flex items-center gap-1 font-semibold"><i class="ri-printer-line"></i> Print / PDF</a>
      <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-md text-xs font-semibold">Close</button>
    </div>`);
}

async function loadNoteIntoWorkspace(nid) {
  const note = await api(`/api/tn01/note/${nid}`);
  TN01_HEADERS = [...note.sheet_headers];
  TN01_ROWS = JSON.parse(JSON.stringify(note.sheet_data));
  closeDialog();
  switchTN01Tab('sheet');
  setMsg("Loaded note into active sheet.");
}

async function renderTN01HistoryTab() {
  const isAdmin = (ME.is_master || ME.role === 'admin');
  const hist = await api('/api/tn01/history');
  const spro = isAdmin ? await api('/api/spro/overview') : null;

  $('#tn01_content').innerHTML = `
    <div class="p-6 flex-1 flex flex-col overflow-hidden">
      <div class="flex items-center justify-between pb-3 border-b border-slate-200 shrink-0">
        <div>
          <h2 class="text-base font-bold text-slate-900">${isAdmin ? "All Transfer Notes (Admin View)" : "My Transfer Notes"}</h2>
          <p class="text-xs text-slate-500">${isAdmin ? "Full history across departments — apply filters below" : "Showing notes sent by or addressed to your department"}</p>
        </div>
        <button onclick="renderTN01HistoryTab()" class="btn-secondary text-xs px-3 py-1.5 rounded-lg"><i class="ri-refresh-line"></i> Refresh</button>
      </div>

      ${isAdmin ? `
      <div class="mt-3 mb-3 card rounded-xl p-3 grid grid-cols-1 md:grid-cols-6 gap-2 shrink-0">
        <div>
          <label class="form-label">From Dept</label>
          <select id="f_from" class="form-input">
            <option value="">— Any —</option>
            ${spro.departments.map(d => `<option value="${esc(d.name)}">${esc(d.name)}</option>`).join('')}
          </select>
        </div>
        <div>
          <label class="form-label">To Dept</label>
          <select id="f_to" class="form-input">
            <option value="">— Any —</option>
            ${spro.departments.map(d => `<option value="${esc(d.name)}">${esc(d.name)}</option>`).join('')}
          </select>
        </div>
        <div>
          <label class="form-label">Status</label>
          <select id="f_status" class="form-input">
            <option value="">— Any —</option>
            <option value="sent">Sent</option>
            <option value="received">Received</option>
          </select>
        </div>
        <div>
          <label class="form-label">Date From</label>
          <input type="date" id="f_from_date" class="form-input">
        </div>
        <div>
          <label class="form-label">Date To</label>
          <input type="date" id="f_to_date" class="form-input">
        </div>
        <div class="flex items-end gap-1">
          <button onclick="applyHistoryFilters()" class="btn-primary text-xs font-bold px-3 py-2 rounded-md flex-1"><i class="ri-filter-line"></i> Filter</button>
          <button onclick="renderTN01HistoryTab()" class="btn-secondary text-xs font-semibold px-2 py-2 rounded-md">Reset</button>
        </div>
      </div>` : ''}

      <div class="flex-1 mt-2 card rounded-2xl overflow-auto">
        <table class="w-full text-left custom-table">
          <thead><tr><th>Note #</th><th>Route</th><th>From</th><th>To</th><th>Batches</th><th>Sender</th><th>Timestamp</th><th>Status</th><th>Actions</th></tr></thead>
          <tbody id="hist_tbody">${renderHistoryRows(hist)}</tbody>
        </table>
      </div>
    </div>`;
}

function renderHistoryRows(hist) {
  if (!hist.length) return `<tr><td colspan="9" class="text-center py-10 text-slate-400">No records found.</td></tr>`;
  return hist.map(h => `
    <tr>
      <td class="font-mono font-bold text-indigo-600">${esc(h.note_no)}</td>
      <td class="font-semibold text-slate-900">${esc(h.route)}</td>
      <td>${esc(h.from_plant)} - ${esc(h.from_dept)}</td>
      <td>${esc(h.to_plant)} - ${esc(h.to_dept)}</td>
      <td class="font-mono font-bold">${h.entry_count}</td>
      <td>${esc(h.from_user_name)}</td>
      <td class="font-mono text-xs text-slate-500">${esc(h.created_at)}</td>
      <td><span class="badge ${h.status==='received'?'badge-emerald':'badge-amber'}">${h.status.toUpperCase()}</span></td>
      <td class="space-x-1">
        <button onclick="viewTransferNote(${h.id})" class="text-xs text-indigo-600 font-semibold">View</button>
        <a href="/api/tn01/print/${h.id}" target="_blank" class="text-xs text-slate-600 font-semibold ml-2">PDF</a>
        <a href="/api/tn01/export/${h.id}" class="text-xs text-emerald-600 font-semibold ml-2">Excel</a>
      </td>
    </tr>
  `).join('');
}

async function applyHistoryFilters() {
  const qs = new URLSearchParams();
  const fd = $('#f_from'); if (fd && fd.value) qs.set('from_dept', fd.value);
  const td = $('#f_to'); if (td && td.value) qs.set('to_dept', td.value);
  const st = $('#f_status'); if (st && st.value) qs.set('status', st.value);
  const df = $('#f_from_date'); if (df && df.value) qs.set('date_from', df.value);
  const dt = $('#f_to_date'); if (dt && dt.value) qs.set('date_to', dt.value);
  const hist = await api('/api/tn01/history?' + qs.toString());
  const tbody = $('#hist_tbody');
  if (tbody) tbody.innerHTML = renderHistoryRows(hist);
  setMsg(`Filtered: ${hist.length} note(s)`);
}

async function renderTN01FormatTab() {
  const layouts = await api('/api/tn01/layouts');
  $('#tn01_content').innerHTML = `
    <div class="p-6 flex-1 flex flex-col overflow-hidden">
      <div class="flex items-center justify-between pb-4 border-b border-slate-200 shrink-0">
        <div><h2 class="text-base font-bold text-slate-900">Transfer Note Formats (Admin)</h2></div>
        <button onclick="openCreateLayoutModal()" class="btn-primary text-xs font-semibold px-4 py-2 rounded-lg">+ Create Format</button>
      </div>
      <div class="grid grid-cols-1 md:grid-cols-3 gap-4 mt-4 overflow-auto">
        ${layouts.map(l => `
          <div class="card p-4 rounded-2xl flex flex-col justify-between">
            <div>
              <div class="flex justify-between"><span class="font-mono text-xs text-indigo-600 font-bold">${esc(l.layout_code)}</span><button onclick="deleteLayoutTN01(${l.id})" class="text-red-600 text-xs font-semibold">Delete</button></div>
              <h4 class="font-bold text-slate-900 text-sm mt-2">${esc(l.name)}</h4>
              <p class="text-xs text-slate-500 mt-1">${(l.fields || []).length} Columns</p>
            </div>
            <button onclick="loadFormatIntoWorkspace(${l.id})" class="btn-primary px-3 py-1.5 rounded-md text-xs mt-3 font-semibold">Use in Grid</button>
          </div>
        `).join('')}
      </div>
    </div>`;
}

function openCreateLayoutModal() {
  openDialog(`
    <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl">
      <h3 class="font-bold text-slate-900 text-sm">Design TN01 Custom Format</h3>
      <button onclick="closeDialog()"><i class="ri-close-line text-lg text-slate-500"></i></button>
    </div>
    <div class="p-6 space-y-4">
      <div><label class="form-label">Layout Name</label><input id="nl_n" class="form-input" placeholder="e.g. 11kV Armoured Cable Format"></div>
      <div><label class="form-label">Columns (comma-separated)</label>
        <textarea id="nl_c" class="form-input font-mono" rows="3">Sr. No., Production Order, Heat Lot, Conductor Size, Length (Mtrs), Weight (Kg), Remarks</textarea>
      </div>
    </div>
    <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl">
      <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-lg text-xs font-semibold">Cancel</button>
      <button onclick="submitSaveFormat()" class="btn-primary px-4 py-2 rounded-lg text-xs font-semibold">Save</button>
    </div>`);
}

async function submitSaveFormat() {
  const name = $('#nl_n').value.trim();
  const cols = $('#nl_c').value.split(',').map(c => c.trim()).filter(Boolean);
  if (!cols.includes("Sr. No.")) cols.unshift("Sr. No.");
  await api('/api/tn01/layouts/save', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({name, columns: cols})});
  closeDialog();
  renderTN01FormatTab();
}

async function loadFormatIntoWorkspace(lid) {
  const layouts = await api('/api/tn01/layouts');
  const l = layouts.find(x => x.id === lid);
  if (!l) return;
  TN01_HEADERS = [...l.fields];
  TN01_ROWS = [];
  for (let i = 1; i <= 6; i++) { const r = new Array(TN01_HEADERS.length).fill(""); r[0] = i; TN01_ROWS.push(r); }
  switchTN01Tab('sheet');
}

async function deleteLayoutTN01(lid) {
  if (!confirm("Delete format?")) return;
  await api(`/api/tn01/layouts/${lid}`, {method: 'DELETE'});
  renderTN01FormatTab();
}

// =========================================================================
// --- STANDARD MODULE ---
// =========================================================================
async function renderStandardModule(code) {
  const layouts = await api(`/api/tcode/${code}/layouts`);
  const activeLayout = layouts.find(l => l.is_default) || layouts[0];
  const fields = activeLayout ? activeLayout.fields : [];
  const recs = await api(`/api/tcode/${code}/records`);

  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col overflow-hidden select-text">
      <div class="p-4 border-b border-slate-200 flex items-center justify-between shrink-0 bg-white">
        <div class="flex items-center gap-3">
          <span class="font-mono font-bold text-xs bg-indigo-600 text-white px-2 py-0.5 rounded-lg shadow-sm">${code}</span>
          <h2 class="text-base font-bold text-slate-900">${activeLayout ? esc(activeLayout.title) : code}</h2>
          <div class="h-5 w-[1px] bg-slate-200 mx-2"></div>
          <button onclick="showStandardRecordModal('${code}')" class="btn-primary text-xs font-semibold px-3 py-1.5 rounded-lg flex items-center gap-1.5 shadow-sm"><i class="ri-add-line"></i> New Entry</button>
        </div>
        <div class="flex items-center gap-2">
          <a href="/api/tcode/${code}/export" class="btn-secondary text-xs font-semibold px-2.5 py-1.5 rounded-lg flex items-center gap-1"><i class="ri-download-2-line"></i> Export</a>
          <button onclick="renderStandardModule('${code}')" class="btn-secondary text-xs font-semibold p-2 rounded-lg"><i class="ri-refresh-line"></i></button>
        </div>
      </div>
      <div class="flex-1 overflow-auto">
        <table class="w-full text-left custom-table">
          <thead><tr><th style="width:60px;">#</th><th style="width:90px;">Plant</th>${fields.map(f=>`<th>${esc(f.label)}</th>`).join('')}<th style="width:130px;">Logged By</th><th style="width:140px;">Timestamp</th><th style="width:80px;">Action</th></tr></thead>
          <tbody>
            ${!recs.length ? `<tr><td colspan="${fields.length + 5}" class="text-center py-10 text-slate-400">No active records found.</td></tr>` : ''}
            ${recs.map(r => `
              <tr>
                <td class="font-mono text-slate-400 font-semibold">${r.id}</td>
                <td><span class="font-mono text-xs bg-slate-100 px-2 py-0.5 rounded text-slate-700">${esc(r.plant || '1000')}</span></td>
                ${fields.map(f=>`<td>${esc(r.data[f.key]||'')}</td>`).join('')}
                <td class="text-xs text-slate-500">${esc(r.created_by_name || r.created_by)}</td>
                <td class="text-xs text-slate-500 font-mono">${esc(r.created_at)}</td>
                <td><button onclick="deleteStandardRecord('${code}', ${r.id})" class="text-xs text-red-600 hover:text-red-800 font-semibold">Del</button></td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    </div>`;
}

function showStandardRecordModal(code) {
  api(`/api/tcode/${code}/layouts`).then(layouts => {
    const activeLayout = layouts.find(l => l.is_default) || layouts[0];
    const fields = activeLayout ? activeLayout.fields : [];
    const myPlant = (ME.plants && ME.plants[0]) ? ME.plants[0] : "1000";

    openDialog(`
      <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl">
        <h3 class="font-bold text-slate-900 text-sm">Post Entry: ${code}</h3>
        <button onclick="closeDialog()"><i class="ri-close-line text-lg text-slate-500"></i></button>
      </div>
      <div class="p-6 space-y-4">
        <div><label class="form-label">Facility Code</label><input id="std_f__plant" class="form-input font-mono" value="${esc(myPlant)}"></div>
        <div class="grid grid-cols-2 gap-4">
          ${fields.map(f => `
            <div class="col-span-${f.width||1}">
              <label class="form-label">${esc(f.label)}</label>
              ${f.type==='textarea' ? `<textarea id="std_f_${f.key}" class="form-input" rows="2"></textarea>` : (f.type==='select' && f.options ? `<select id="std_f_${f.key}" class="form-input">${f.options.map(o=>`<option value="${esc(o)}">${esc(o)}</option>`).join('')}</select>` : `<input type="${f.type==='number'?'number':(f.type==='date'?'date':'text')}" id="std_f_${f.key}" class="form-input">`)}
            </div>
          `).join('')}
        </div>
      </div>
      <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl">
        <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-lg text-xs font-semibold">Cancel</button>
        <button onclick="submitStandardRecord('${code}')" class="btn-primary px-4 py-2 rounded-lg text-xs font-semibold">Save</button>
      </div>`);
  });
}

async function submitStandardRecord(code) {
  const layouts = await api(`/api/tcode/${code}/layouts`);
  const activeLayout = layouts.find(l => l.is_default) || layouts[0];
  const fields = activeLayout ? activeLayout.fields : [];
  const payload = { _plant: $('#std_f__plant').value };
  for (let f of fields) {
    const el = $('#std_f_' + f.key);
    payload[f.key] = el ? el.value.trim() : "";
  }
  try {
    await api(`/api/tcode/${code}/records`, {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(payload)});
    closeDialog();
    renderStandardModule(code);
    setMsg("Entry saved");
  } catch(e) { alert(e); }
}

async function deleteStandardRecord(code, rid) {
  if (!confirm(`Delete entry #${rid}?`)) return;
  try {
    await api(`/api/tcode/${code}/records/${rid}`, {method: 'DELETE'});
    renderStandardModule(code);
    setMsg("Record deleted");
  } catch(e) { alert(e); }
}

function openDialog(content) { $('#app_dialog_body').innerHTML = content; $('#app_dialog').classList.remove('hidden'); }
function closeDialog() { $('#app_dialog').classList.add('hidden'); $('#app_dialog_body').innerHTML = ''; }

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
    return {"status": "ACTIVE", "system": "NEXUS-CORE", "version": "3.7.0"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
