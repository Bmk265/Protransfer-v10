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
                r = cur.fetchone()
                return dict(r) if r else None
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

def esc_py(s):
    return str(s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

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
        if v is None or v == "": continue
        if isinstance(v, (datetime.datetime, datetime.date)): is_num = False; continue
        if isinstance(v, (int, float)): is_date = False; continue
        s = str(v).strip()
        if len(s) > 70 or '\n' in s: is_long, is_num, is_date = True, False, False; break
        if len(s) > 1 and s[0] == '0' and s.isdigit(): is_num = False
        else:
            try: float(s.replace(',', ''))
            except ValueError: is_num = False
        if not any(re.match(p, s) for p in date_patterns): is_date = False
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
    if USE_PG: cur.execute(schema_sql)
    else: cur.executescript(schema_sql)
    conn.commit()

    if USE_PG:
        try:
            cur.execute("SELECT setval(pg_get_serial_sequence('tcode_records', 'id'), COALESCE((SELECT MAX(id)+1 FROM tcode_records), 1), false);")
            cur.execute("SELECT setval(pg_get_serial_sequence('transfer_notes', 'id'), COALESCE((SELECT MAX(id)+1 FROM transfer_notes), 1), false);")
            conn.commit()
        except Exception: conn.rollback()

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

    for su_name, su_full, su_dept, su_pl in [
        ("prod_user", "Production Supervisor", "Production", "1000"),
        ("testing_lead", "Final Testing Incharge", "Final Testing", "1000"),
        ("comm_lead", "Commercial Manager", "Commercial", "1000"),
        ("store_head", "Store Incharge", "Store", "1000"),
        ("qc_lead", "Quality Control Incharge", "Quality Control", "1000")
    ]:
        cur.execute(q("SELECT 1 FROM users WHERE username=?"), (su_name,))
        if not cur.fetchone():
            cur.execute(q("INSERT INTO users (username, password_hash, full_name, role, department, plants, status, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?)"),
                        (su_name, ph, su_full, "user", su_dept, su_pl, "approved", cur_time, cur_time))

    default_modules = [
        ("TN01", "Transfer Note & Handover Sheet", "Inter-department material handover & batch transfer note", "MM", "Production,Final Testing,Commercial,Store,Quality Control,Planning"),
        ("CTST", "Cable Type-Test Calculator", "CR, Braid Coverage, Tubular & Dumbbell mechanical tests", "QC", "Quality Control,Final Testing,Engineering"),
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

    # CTST is handled by dedicated endpoints — remove any schema row from earlier versions
    cur.execute(q("DELETE FROM tcode_schemas WHERE code='CTST'"))
    cur.execute(q("DELETE FROM tcode_layouts WHERE tcode_code='CTST'"))

    conn.commit()
    conn.close()

init_db()

# --- AUTH HELPERS ---
def get_user(req: Request):
    token = None
    ah = req.headers.get("Authorization")
    if ah and ah.startswith("Bearer "): token = ah.split(" ")[1].strip()
    if not token: token = req.cookies.get("nexus_session") or req.cookies.get("sap_session")
    if not token: return None
    s = fetch_one("SELECT * FROM sessions WHERE token=?", (token,))
    if not s: return None
    u = fetch_one("SELECT * FROM users WHERE username=?", (s["username"],))
    return u if (u and u["status"] == "approved") else None

def need(req: Request):
    u = get_user(req)
    if not u: raise HTTPException(401, "Authentication required. Please sign in.")
    return u

def is_master(u): return u["role"] == "master" or u["username"] == "admin"

def need_admin(req: Request):
    u = need(req)
    if u["role"] not in ["admin", "master"]: raise HTTPException(403, "Access restricted to Administrators.")
    return u

def need_master(req: Request):
    u = need(req)
    if not is_master(u): raise HTTPException(403, "Master Administrator clearance required.")
    return u

def check_tcode_access(u, code: str):
    code = code.upper().strip()
    if is_master(u) or u["role"] == "admin": return True
    tc = fetch_one("SELECT * FROM tcodes WHERE code=? AND active=1", (code,))
    if not tc: return False
    allowed = [x.strip() for x in (tc["allowed_depts"] or "").split(",") if x.strip()]
    if u["department"] in allowed: return True
    return bool(fetch_one("SELECT 1 FROM access_grants WHERE username=? AND tcode_code=?", (u["username"], code)))

# --- FASTAPI APP ---
app = FastAPI(title="Nexus Core ERP", version="3.8.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True, allow_methods=["*"], allow_headers=["*"])

# --- AUTH APIS ---
@app.post("/api/login")
def api_login(res: Response, username: str = Form(...), password: str = Form(...)):
    username = username.lower().strip()
    u = fetch_one("SELECT * FROM users WHERE username=?", (username,))
    if not u or not bcrypt.checkpw(password.encode('utf-8'), u["password_hash"].encode('utf-8')):
        raise HTTPException(401, "Invalid username or password.")
    if u["status"] != "approved": raise HTTPException(403, "Account is locked or pending approval.")
    token = secrets.token_hex(32)
    exp = (datetime.datetime.now() + datetime.timedelta(days=7)).strftime("%Y-%m-%d %H:%M:%S")
    exec_write("INSERT INTO sessions (token, username, expires) VALUES (?,?,?)", (token, username, exp))
    res.set_cookie("nexus_session", token, httponly=True, samesite="lax", max_age=86400 * 7)
    return {"ok": True, "token": token, "user": {"username": u["username"], "full_name": u["full_name"], "role": u["role"]}}

@app.post("/api/logout")
def api_logout(res: Response, req: Request):
    token = req.headers.get("Authorization", "").replace("Bearer ", "").strip() or req.cookies.get("nexus_session")
    if token: exec_write("DELETE FROM sessions WHERE token=?", (token,))
    res.delete_cookie("nexus_session")
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

# --- SPRO ---
@app.get("/api/spro/overview")
def spro_overview(req: Request):
    need(req)
    return {"org": fetch_one("SELECT * FROM org WHERE id=1"),
            "departments": fetch_all("SELECT * FROM departments ORDER BY name"),
            "plants": fetch_all("SELECT * FROM plants ORDER BY code")}

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

# --- SU01 ---
@app.get("/api/su01/users")
def su01_list_users(req: Request):
    need_admin(req)
    return fetch_all("SELECT username, full_name, role, department, plants, status, created_at, updated_at FROM users ORDER BY username")

@app.post("/api/su01/create")
def su01_create_user(req: Request, username: str = Form(...), password: str = Form(...), full_name: str = Form(...),
                     department: str = Form(...), plants: str = Form(""), role: str = Form("user")):
    caller = need_admin(req)
    username = username.lower().strip()
    if role == "master" and not is_master(caller): raise HTTPException(403, "Only Master Admins can create superuser accounts.")
    if fetch_one("SELECT 1 FROM users WHERE username=?", (username,)): raise HTTPException(400, f"User '{username}' already exists.")
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
    if not target: raise HTTPException(404, "User not found")
    if is_master(target) and not is_master(caller): raise HTTPException(403, "Cannot modify a Master Administrator.")
    if role == "master" and not is_master(caller): raise HTTPException(403, "Cannot grant Master privileges.")
    exec_write("UPDATE users SET full_name=?, department=?, plants=?, role=?, status=?, updated_at=? WHERE username=?",
               (full_name.strip(), department.strip(), plants.strip(), role, status, now(), username))
    return {"ok": True}

@app.post("/api/su01/delete")
def su01_delete_user(req: Request, username: str = Form(...)):
    caller = need_admin(req)
    username = username.lower().strip()
    target = fetch_one("SELECT * FROM users WHERE username=?", (username,))
    if not target: raise HTTPException(404, "User not found")
    if is_master(target): raise HTTPException(400, "Master Administrator account is protected from deletion.")
    if target["role"] == "admin" and not is_master(caller): raise HTTPException(403, "Only Master Admins can delete administrative accounts.")
    if username == caller["username"]: raise HTTPException(400, "You cannot delete your own session.")
    exec_write("DELETE FROM users WHERE username=?", (username,))
    exec_write("DELETE FROM sessions WHERE username=?", (username,))
    exec_write("DELETE FROM access_grants WHERE username=?", (username,))
    return {"ok": True}

@app.post("/api/su01/reset-password")
def su01_reset_password(req: Request, username: str = Form(...), new_password: str = Form(...)):
    caller = need_admin(req)
    username = username.lower().strip()
    target = fetch_one("SELECT * FROM users WHERE username=?", (username,))
    if not target: raise HTTPException(404, "User not found")
    if is_master(target) and not is_master(caller): raise HTTPException(403, "Cannot reset password for Master Administrator.")
    ph = bcrypt.hashpw(new_password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
    exec_write("UPDATE users SET password_hash=?, updated_at=? WHERE username=?", (ph, now(), username))
    exec_write("DELETE FROM sessions WHERE username=?", (username,))
    return {"ok": True}

# --- PFCG ---
@app.get("/api/pfcg/matrix")
def pfcg_matrix(req: Request):
    need_admin(req)
    return {"users": fetch_all("SELECT username, full_name, role, department FROM users ORDER BY username"),
            "tcodes": fetch_all("SELECT code, title, module, allowed_depts FROM tcodes WHERE active=1 ORDER BY code"),
            "grants": fetch_all("SELECT * FROM access_grants")}

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

# --- SE93 ---
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
async def se93_auto_create(req: Request, code: str = Form(...), title: str = Form(...), description: str = Form(""),
    module: str = Form("PROD"), allowed_depts: str = Form("Production,Quality Control"),
    import_sample_data: int = Form(1), file: UploadFile = File(...)):
    u = need_admin(req)
    code = code.upper().strip()
    if not code: raise HTTPException(400, "T-Code identifier is required.")
    if fetch_one("SELECT 1 FROM tcodes WHERE code=?", (code,)): raise HTTPException(400, f"Module '{code}' already exists.")
    contents = await file.read()
    filename = file.filename.lower()
    headers, data_rows = [], []
    try:
        if filename.endswith(".xlsx") or filename.endswith(".xls"):
            wb = load_workbook(io.BytesIO(contents), data_only=True)
            rows = list(wb.active.iter_rows(values_only=True))
            if rows:
                headers = [str(h).strip() for h in rows[0] if h is not None and str(h).strip()]
                data_rows = rows[1:]
        elif filename.endswith(".csv"):
            rows = list(csv.reader(io.StringIO(contents.decode("utf-8", errors="ignore"))))
            if rows:
                headers = [str(h).strip() for h in rows[0] if str(h).strip()]
                data_rows = rows[1:]
        elif filename.endswith(".json"):
            parsed = json.loads(contents.decode("utf-8", errors="ignore"))
            if isinstance(parsed, list) and parsed and isinstance(parsed[0], dict):
                headers = list(parsed[0].keys())
                for item in parsed: data_rows.append([item.get(k, "") for k in headers])
        else: raise HTTPException(400, "Please upload .xlsx/.xls/.csv/.json")
    except Exception as e: raise HTTPException(400, f"Template error: {str(e)}")
    if not headers: raise HTTPException(400, "No valid column headers found.")
    col_samples = {i: [] for i in range(len(headers))}
    for r in data_rows[:25]:
        for ci in range(len(headers)):
            if ci < len(r) and r[ci] is not None: col_samples[ci].append(r[ci])
    fields = []
    for idx, h in enumerate(headers):
        ck = re.sub(r'[^a-zA-Z0-9_]', '_', h.lower().strip())
        ck = re.sub(r'_+', '_', ck).strip('_') or f"col_{idx+1}"
        ft = infer_field_type(col_samples[idx])
        fields.append({"key": ck, "label": h, "type": ft, "required": False, "width": 2 if ft == "textarea" else 1})
    cur_time = now()
    exec_write("INSERT INTO tcodes (code, title, description, module, allowed_depts, active, created_at, updated_at, created_by) VALUES (?,?,?,?,?,?,?,?,?)",
               (code, title.strip() or f"{code} Module", description.strip(), module.strip(), allowed_depts.strip(), 1, cur_time, cur_time, u["username"]))
    fields_json = json.dumps(fields)
    exec_write("INSERT INTO tcode_schemas (code, fields, layout, updated_at, updated_by) VALUES (?,?,?,?,?)",
               (code, fields_json, json.dumps({"columns": 3}), cur_time, u["username"]))
    exec_write("INSERT INTO tcode_layouts (tcode_code, layout_code, name, title, is_default, fields, layout_config, description, created_by, created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
               (code, f"GRID_{code}_01", "Default Grid", f"{title or code} Data View", 1, fields_json, json.dumps({"columns": 3}), "Auto-generated from file", u["username"], u["full_name"], cur_time, cur_time))
    imported = 0
    if import_sample_data and data_rows:
        conn = get_conn(); cur = conn.cursor()
        try:
            insert_sql = q("INSERT INTO tcode_records (tcode_code, plant, data, status, created_by, created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?)")
            def_plant = (u.get("plants") or "1000").split(",")[0].strip() or "1000"
            for r in data_rows:
                if not any(r): continue
                rec, rp = {}, def_plant
                for idx in range(len(headers)):
                    if idx < len(r):
                        val = r[idx]
                        if val is None: val = ""
                        elif isinstance(val, (datetime.datetime, datetime.date)):
                            val = val.strftime("%Y-%m-%d %H:%M:%S") if isinstance(val, datetime.datetime) else val.strftime("%Y-%m-%d")
                        k = fields[idx]["key"]
                        if k in ["plant", "facility"]: rp = str(val or def_plant).strip()
                        else: rec[k] = val
                cur.execute(insert_sql, (code, rp, json.dumps(rec, default=json_serial), "active", u["username"], u["full_name"], cur_time, cur_time))
                imported += 1
            conn.commit()
        finally: conn.close()
    return {"ok": True, "code": code, "title": title, "columns_count": len(fields), "imported_count": imported}

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

# --- AI CODE STUDIO ---
VALID_TYPES = {"text", "number", "date", "textarea", "select", "email", "tel"}
RESERVED_CODES = {"TN01", "SU01", "SPRO", "SE93", "PFCG", "STUDIO", "CTST"}

def _validate_spec(spec):
    errors = []
    if not isinstance(spec, dict): return None, ["Spec must be a JSON object."]
    action = (spec.get("action") or "upsert").lower()
    if action not in ("upsert", "delete"): errors.append("action must be 'upsert' or 'delete'.")
    code = str(spec.get("code") or "").upper().strip()
    if not re.match(r'^[A-Z][A-Z0-9_]{1,10}$', code): errors.append("code must be 2-11 chars, start with a letter, A-Z/0-9/_ only.")
    if code in RESERVED_CODES: errors.append(f"'{code}' is a reserved system module.")
    clean = {"action": action, "code": code}
    if action == "delete": return clean, errors
    clean["title"] = str(spec.get("title") or code).strip()[:120]
    clean["description"] = str(spec.get("description") or "").strip()[:500]
    clean["module"] = str(spec.get("module") or "CUSTOM").strip()[:16].upper()
    depts = spec.get("allowed_depts") or ""
    if isinstance(depts, list): depts = ",".join(str(d).strip() for d in depts)
    clean["allowed_depts"] = str(depts).strip()
    fields = spec.get("fields") or []
    if not isinstance(fields, list) or not fields:
        errors.append("fields must be a non-empty array.")
        return clean, errors
    clean_fields, seen = [], set()
    for i, f in enumerate(fields):
        if not isinstance(f, dict): errors.append(f"fields[{i}] must be an object."); continue
        key = re.sub(r'[^a-z0-9_]', '_', str(f.get("key") or f.get("label") or f"col_{i+1}").lower().strip())
        key = re.sub(r'_+', '_', key).strip('_') or f"col_{i+1}"
        if key in seen: errors.append(f"duplicate field key '{key}'."); continue
        seen.add(key)
        ft = str(f.get("type") or "text").lower()
        if ft not in VALID_TYPES: errors.append(f"field '{key}' invalid type '{ft}'."); ft = "text"
        clean_fields.append({"key": key, "label": str(f.get("label") or key).strip()[:120], "type": ft,
                             "required": bool(f.get("required", False)),
                             "width": 2 if ft == "textarea" else int(f.get("width", 1) or 1),
                             "options": f.get("options") if ft == "select" else None})
    clean["fields"] = clean_fields
    seeds = spec.get("seed_rows") or []
    clean["seed_rows"] = [r for r in seeds if isinstance(r, dict)][:500] if isinstance(seeds, list) else []
    return clean, errors

def _apply_spec(clean, user, note=""):
    code, action = clean["code"], clean["action"]
    cur_time = now()
    existing = fetch_one("SELECT code FROM tcodes WHERE code=?", (code,))
    if action == "delete":
        if not existing: raise HTTPException(404, f"Module '{code}' does not exist.")
        for tbl, col in [("tcodes","code"),("tcode_schemas","code"),("tcode_layouts","tcode_code"),("tcode_records","tcode_code"),("access_grants","tcode_code")]:
            exec_write(f"DELETE FROM {tbl} WHERE {col}=?", (code,))
        summary = {"action": "delete", "code": code}
    else:
        fields_json = json.dumps(clean["fields"])
        lc = json.dumps({"columns": 3})
        if existing:
            exec_write("UPDATE tcodes SET title=?, description=?, module=?, allowed_depts=?, updated_at=?, created_by=? WHERE code=?",
                       (clean["title"], clean["description"], clean["module"], clean["allowed_depts"], cur_time, user["username"], code))
            exec_write("UPDATE tcode_schemas SET fields=?, layout=?, updated_at=?, updated_by=? WHERE code=?",
                       (fields_json, lc, cur_time, user["username"], code))
            dl = fetch_one("SELECT id FROM tcode_layouts WHERE tcode_code=? AND is_default=1", (code,))
            if dl: exec_write("UPDATE tcode_layouts SET fields=?, layout_config=?, updated_at=? WHERE id=?", (fields_json, lc, cur_time, dl["id"]))
            else: exec_write("INSERT INTO tcode_layouts (tcode_code, layout_code, name, title, is_default, fields, layout_config, description, created_by, created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                             (code, f"GRID_{code}_01", "Default Grid", clean["title"], 1, fields_json, lc, "Managed by AI Studio", user["username"], user["full_name"], cur_time, cur_time))
            summary = {"action": "updated", "code": code}
        else:
            exec_write("INSERT INTO tcodes (code, title, description, module, allowed_depts, active, created_at, updated_at, created_by) VALUES (?,?,?,?,?,?,?,?,?)",
                       (code, clean["title"], clean["description"], clean["module"], clean["allowed_depts"], 1, cur_time, cur_time, user["username"]))
            exec_write("INSERT INTO tcode_schemas (code, fields, layout, updated_at, updated_by) VALUES (?,?,?,?,?)",
                       (code, fields_json, lc, cur_time, user["username"]))
            exec_write("INSERT INTO tcode_layouts (tcode_code, layout_code, name, title, is_default, fields, layout_config, description, created_by, created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                       (code, f"GRID_{code}_01", "Default Grid", clean["title"], 1, fields_json, lc, "Created via AI Studio", user["username"], user["full_name"], cur_time, cur_time))
            summary = {"action": "created", "code": code}
        seeded = 0
        fkeys = {f["key"] for f in clean["fields"]}
        for row in clean["seed_rows"]:
            rec = {k: row[k] for k in row if k in fkeys}
            if not rec: continue
            plant = str(row.get("plant") or (user.get("plants") or "1000").split(",")[0].strip() or "1000")
            exec_write("INSERT INTO tcode_records (tcode_code, plant, data, status, created_by, created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?)",
                       (code, plant, json.dumps(rec, default=json_serial), "active", user["username"], user["full_name"], cur_time, cur_time))
            seeded += 1
        summary["seeded_rows"] = seeded
    prev = fetch_one("SELECT COALESCE(MAX(version),0) as v FROM module_specs WHERE tcode_code=?", (code,))
    ver = (prev["v"] if prev else 0) + 1
    exec_write("INSERT INTO module_specs (tcode_code, version, action, spec, applied_by, applied_at, note) VALUES (?,?,?,?,?,?,?)",
               (code, ver, action, json.dumps(clean), user["username"], cur_time, note or ""))
    summary["version"] = ver
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
    if errors: raise HTTPException(400, "Spec invalid: " + "; ".join(errors))
    return {"ok": True, **_apply_spec(clean, user, note=str(payload.get("note", "") or ""))}

@app.get("/api/studio/history")
def studio_history(req: Request, code: str = ""):
    need_master(req)
    sql = "SELECT id, tcode_code, version, action, applied_by, applied_at, note FROM module_specs"
    params = []
    if code: sql += " WHERE tcode_code=?"; params.append(code.upper().strip())
    sql += " ORDER BY id DESC LIMIT 200"
    return fetch_all(sql, tuple(params))

@app.get("/api/studio/spec/{code}")
def studio_get_spec(req: Request, code: str):
    need_master(req)
    code = code.upper().strip()
    tc = fetch_one("SELECT * FROM tcodes WHERE code=?", (code,))
    if not tc: raise HTTPException(404, "Module not found.")
    schema = fetch_one("SELECT fields FROM tcode_schemas WHERE code=?", (code,))
    try: fields = json.loads(schema["fields"]) if schema else []
    except: fields = []
    return {"action": "upsert", "code": code, "title": tc["title"], "description": tc["description"] or "",
            "module": tc["module"] or "CUSTOM", "allowed_depts": tc["allowed_depts"] or "",
            "fields": [{"key": f.get("key"), "label": f.get("label"), "type": f.get("type", "text"),
                        "required": f.get("required", False), "width": f.get("width", 1)} for f in fields],
            "seed_rows": []}

@app.post("/api/studio/rollback/{msid}")
def studio_rollback(req: Request, msid: int):
    user = need_master(req)
    row = fetch_one("SELECT * FROM module_specs WHERE id=?", (msid,))
    if not row: raise HTTPException(404, "History entry not found.")
    try: prev = json.loads(row["spec"])
    except: raise HTTPException(400, "Stored spec corrupt.")
    prev["action"] = "upsert" if row["action"] != "delete" else "delete"
    clean, errors = _validate_spec(prev)
    if errors: raise HTTPException(400, "Cannot rollback: " + "; ".join(errors))
    return {"ok": True, **_apply_spec(clean, user, note=f"Rollback to v{row['version']}")}

# =========================================================================
# --- CTST : CABLE TYPE-TEST CALCULATOR (PICK-YOUR-TEST) ---
# =========================================================================
CTST_TESTS = {
    "cr":       {"title": "CR — Conductor Resistance", "icon": "ri-pulse-line",        "desc": "Temperature-corrected conductor resistance per metre & per km"},
    "cr_pct":   {"title": "CR Percentage vs Spec",     "icon": "ri-bar-chart-line",    "desc": "Deviation and length-wise variation vs specified CR"},
    "braid":    {"title": "Braid / Shield Coverage",   "icon": "ri-shield-check-line", "desc": "Fill factor and % coverage for wire braid screens"},
    "tubular":  {"title": "Tubular Mechanical Test",   "icon": "ri-circle-line",       "desc": "Area, tensile strength, elongation for tube specimen"},
    "dumbbell": {"title": "Dumbbell Mechanical Test",  "icon": "ri-shape-line",        "desc": "Area, tensile strength, elongation for dumbbell specimen"},
    "hotset":   {"title": "Hot Set & Permanent Set",   "icon": "ri-temp-hot-line",     "desc": "Hot-set and permanent-set elongation percentages"},
}

def _n(v):
    try:
        if v is None or str(v).strip() == "": return None
        return float(str(v).replace(",", ""))
    except Exception: return None

def _compute_by_test(tt: str, d: dict) -> dict:
    out = {}
    if tt == "cr":
        obd = _n(d.get("obd_cr_reading")); t = _n(d.get("temp_c")); L = _n(d.get("length_m"))
        if obd is not None and t is not None:
            if -50 <= t <= 100:
                factor = 250.0 / (230.0 + t)
                out["cr_1m"] = round(factor * obd, 6)
                if L and L > 0: out["cr_total_per_km"] = round(factor * obd / L * 1000.0, 6)
            else:
                out["cr_warning"] = "Temp outside −50 to 100 °C — check entry"
    elif tt == "cr_pct":
        spec = _n(d.get("specified_cr")); obs = _n(d.get("observed_cr"))
        mx = _n(d.get("cr_max")); mn = _n(d.get("cr_min"))
        if spec and obs is not None and spec != 0: out["cr_deviation_pct"] = round((obs - spec) / spec * 100.0, 4)
        if mx and mn and (mx + mn) != 0: out["cr_variation_pct"] = round((mx - mn) / (mx + mn) * 100.0, 4)
    elif tt == "braid":
        bd = _n(d.get("braid_d")); uD = _n(d.get("under_braid_D"))
        wpc = _n(d.get("wires_per_carrier")); carr = _n(d.get("carriers")); lay = _n(d.get("lay_length"))
        if all(v is not None for v in (bd, uD, wpc, carr, lay)) and uD > 0 and lay > 0:
            N = wpc * carr
            F = (N * bd) / (2 * math.pi * uD) * math.sqrt(1 + (math.pi * uD / lay) ** 2)
            out["braid_total_wires"] = int(N)
            out["braid_filling_factor"] = round(F, 6)
            out["braid_coverage_pct"] = round((2 * F - F * F) * 100.0, 3)
    elif tt in ("tubular", "dumbbell"):
        if tt == "tubular":
            D = _n(d.get("tub_outer_d")); wall = _n(d.get("tub_wall"))
            if D and wall and D > wall: out["area_mm2"] = round(math.pi * (D - wall) * wall, 4)
            load = _n(d.get("break_load_tub")); final = _n(d.get("elong_tub_final"))
        else:
            w = _n(d.get("dumb_width")); th = _n(d.get("dumb_thk"))
            if w and th: out["area_mm2"] = round(w * th, 4)
            load = _n(d.get("break_load_dumb")); final = _n(d.get("elong_dumb_final"))
        init = _n(d.get("initial_gauge"))
        if load is not None and out.get("area_mm2"): out["tensile_nmm2"] = round(load / out["area_mm2"], 4)
        if init and final is not None and init != 0: out["elongation_pct"] = round((final - init) / init * 100.0, 4)
    elif tt == "hotset":
        init = _n(d.get("initial_gauge")); hs = _n(d.get("hotset_elong_final")); ps = _n(d.get("permset_cooled"))
        if init and hs is not None and init != 0: out["hotset_pct"] = round((hs - init) / init * 100.0, 4)
        if init and ps is not None and init != 0: out["permset_pct"] = round((ps - init) / init * 100.0, 4)
    return out

def _filter_by_test(sql_base: str, test_type: str):
    """Cross-DB way to filter JSON by _test_type."""
    if not test_type: return sql_base, []
    if USE_PG: return sql_base + " AND data::json->>'_test_type' = %s", [test_type]
    return sql_base + " AND json_extract(data, '$._test_type') = ?", [test_type]

@app.get("/api/ctst/tests")
def ctst_tests(req: Request):
    need(req)
    return CTST_TESTS

@app.get("/api/ctst/records")
def ctst_list(req: Request, test_type: str = "", limit: int = 200):
    u = need(req)
    if not check_tcode_access(u, "CTST"): raise HTTPException(403, "Access denied")
    base = "SELECT * FROM tcode_records WHERE tcode_code='CTST' AND status='active'"
    sql, params = _filter_by_test(base, test_type)
    sql += " ORDER BY id DESC LIMIT ?" if not USE_PG else " ORDER BY id DESC LIMIT %s"
    params.append(limit)
    rows = fetch_all(sql, tuple(params))
    for r in rows:
        try: r["data"] = json.loads(r["data"])
        except: r["data"] = {}
        tt = r["data"].get("_test_type", "")
        r["computed"] = _compute_by_test(tt, r["data"]) if tt else {}
    return rows

@app.post("/api/ctst/compute")
def ctst_compute(req: Request, payload: dict):
    need(req)
    tt = str(payload.get("_test_type", "")).strip()
    if tt not in CTST_TESTS: raise HTTPException(400, "Unknown test type.")
    return {"computed": _compute_by_test(tt, payload)}

@app.post("/api/ctst/save")
def ctst_save(req: Request, payload: dict):
    u = need(req)
    if not check_tcode_access(u, "CTST"): raise HTTPException(403, "Access denied")
    tt = str(payload.get("_test_type", "")).strip()
    if tt not in CTST_TESTS: raise HTTPException(400, "Unknown test type.")
    plant = str(payload.pop("_plant", "") or (u.get("plants") or "1000").split(",")[0].strip() or "1000").strip()
    cur_time = now()
    exec_write("INSERT INTO tcode_records (tcode_code, plant, data, status, created_by, created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?)",
               ("CTST", plant, json.dumps(payload, default=json_serial), "active", u["username"], u["full_name"], cur_time, cur_time))
    return {"ok": True, "computed": _compute_by_test(tt, payload)}

@app.delete("/api/ctst/records/{rid}")
def ctst_delete(req: Request, rid: int):
    u = need(req)
    if not check_tcode_access(u, "CTST"): raise HTTPException(403, "Access denied")
    rec = fetch_one("SELECT * FROM tcode_records WHERE id=? AND tcode_code='CTST'", (rid,))
    if not rec: raise HTTPException(404, "Record not found")
    if rec["created_by"] != u["username"] and not is_master(u) and u["role"] != "admin":
        raise HTTPException(403, "Permission denied")
    exec_write("UPDATE tcode_records SET status='deleted', updated_at=? WHERE id=?", (now(), rid))
    return {"ok": True}

@app.get("/api/ctst/export")
def ctst_export(req: Request, test_type: str = ""):
    u = need(req)
    if not check_tcode_access(u, "CTST"): raise HTTPException(403, "Export denied")
    base = "SELECT * FROM tcode_records WHERE tcode_code='CTST' AND status='active'"
    sql, params = _filter_by_test(base, test_type)
    sql += " ORDER BY id ASC"
    rows = fetch_all(sql, tuple(params))
    wb = Workbook(); ws = wb.active
    ws.title = (test_type.upper() if test_type else "CTST") + "_DATA"
    ws.append(["ID","Plant","Test Type","Sample ID","Cable Type","Date","Result","Computed","Remarks","Created By","Created At"])
    for r in rows:
        try: d = json.loads(r["data"])
        except: d = {}
        tt = d.get("_test_type", "")
        c = _compute_by_test(tt, d)
        computed_str = "; ".join(f"{k}={v}" for k, v in c.items())
        ws.append([r["id"], r["plant"], tt, d.get("sample_id",""), d.get("cable_type",""),
                   d.get("test_date",""), d.get("result",""), computed_str,
                   d.get("remarks",""), r["created_by_name"] or r["created_by"], r["created_at"]])
    bio = io.BytesIO(); wb.save(bio); bio.seek(0)
    tag = test_type.upper() if test_type else "ALL"
    return StreamingResponse(bio, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             headers={"Content-Disposition": f"attachment; filename=CTST_{tag}_{datetime.date.today().strftime('%Y%m%d')}.xlsx"})

@app.get("/api/ctst/print/{rid}", response_class=HTMLResponse)
def ctst_print(req: Request, rid: int):
    u = need(req)
    if not check_tcode_access(u, "CTST"): return HTMLResponse("Access denied", 403)
    rec = fetch_one("SELECT * FROM tcode_records WHERE id=? AND tcode_code='CTST'", (rid,))
    if not rec: return HTMLResponse("Record not found", 404)
    try: d = json.loads(rec["data"])
    except: d = {}
    tt = d.get("_test_type", "")
    c = _compute_by_test(tt, d)
    org = fetch_one("SELECT * FROM org WHERE id=1") or {}
    title = CTST_TESTS.get(tt, {}).get("title", tt.upper())
    def cell(label, val, unit=""):
        return f"<tr><td style='padding:4px 8px;border:1px solid #cbd5e1;font-size:11px;background:#f8fafc'>{label}</td><td style='padding:4px 8px;border:1px solid #cbd5e1;font-size:11px;font-weight:600'>{val or ''} {unit}</td></tr>"
    labels = {"obd_cr_reading":("OBD Reading","Ω"),"temp_c":("Temp","°C"),"length_m":("Length","m"),
              "specified_cr":("Specified CR",""),"observed_cr":("Observed CR",""),"cr_max":("Obs. Max",""),"cr_min":("Obs. Min",""),
              "braid_d":("Braid Wire Ø","mm"),"under_braid_D":("Ø Under Braid","mm"),"wires_per_carrier":("Wires/Carrier",""),
              "carriers":("Carriers",""),"lay_length":("Lay Length","mm"),
              "tub_outer_d":("Tubular Outer Ø","mm"),"tub_wall":("Tubular Wall","mm"),
              "dumb_width":("Dumbbell Width","mm"),"dumb_thk":("Dumbbell Thickness","mm"),
              "break_load_tub":("Break Load (Tub)","N"),"break_load_dumb":("Break Load (Dumb)","N"),
              "elong_tub_final":("Tubular Final Gauge","mm"),"elong_dumb_final":("Dumbbell Final Gauge","mm"),
              "initial_gauge":("Initial Gauge","mm"),"hotset_elong_final":("Hot Set Final","mm"),"permset_cooled":("Perm Set Cooled","mm")}
    comp_l = {"cr_1m":("CR 1 m","Ω"),"cr_total_per_km":("CR per km","Ω/km"),
              "cr_deviation_pct":("CR Deviation","%"),"cr_variation_pct":("CR Variation","%"),
              "braid_total_wires":("Total Wires",""),"braid_filling_factor":("Filling Factor",""),"braid_coverage_pct":("Coverage","%"),
              "area_mm2":("Area","mm²"),"tensile_nmm2":("Tensile Strength","N/mm²"),"elongation_pct":("Elongation","%"),
              "hotset_pct":("Hot Set","%"),"permset_pct":("Permanent Set","%")}
    irows = "".join(cell(l, d.get(k, ""), u_) for k, (l, u_) in labels.items() if k in d)
    orows = "".join(cell(l, c.get(k, ""), u_) for k, (l, u_) in comp_l.items() if k in c)
    if c.get("cr_warning"):
        orows += f"<tr><td colspan='2' style='padding:6px 8px;background:#fef3c7;border:1px solid #fcd34d;font-size:11px;color:#92400e'>⚠ {esc_py(c['cr_warning'])}</td></tr>"
    html = f"""<!DOCTYPE html><html><head><meta charset="utf-8"><title>CTST #{rec['id']}</title>
    <style>body{{font-family:sans-serif;margin:30px;color:#0f172a}}table{{width:100%;border-collapse:collapse;margin-bottom:16px}}@media print{{.no-print{{display:none}}body{{margin:0}}}}</style></head><body>
    <div class="no-print" style="text-align:right;margin-bottom:12px">
      <button onclick="window.print()" style="background:#4f46e5;color:white;border:none;padding:8px 16px;border-radius:6px;cursor:pointer;font-weight:600">Print / Save PDF</button>
    </div>
    <h2 style="margin:0">{esc_py(org.get('name','Thermocables Limited'))}</h2>
    <h4 style="margin:4px 0 12px;color:#4f46e5">{esc_py(title)}</h4>
    <div style="font-size:11px;color:#475569;margin-bottom:16px">
      Record #{rec['id']} | Plant {esc_py(rec['plant'] or '—')} | Sample {esc_py(d.get('sample_id','—'))} | {esc_py(d.get('cable_type','—'))} | Tested by {esc_py(rec['created_by_name'] or rec['created_by'])} | {esc_py(rec['created_at'])}
    </div>
    <h4 style="font-size:12px;color:#4338ca;text-transform:uppercase;margin-bottom:6px">Inputs</h4>
    <table>{irows}</table>
    <h4 style="font-size:12px;color:#4338ca;text-transform:uppercase;margin-bottom:6px">Computed Results</h4>
    <table>{orows or '<tr><td style="padding:6px;font-size:11px;color:#64748b">No computed values</td></tr>'}</table>
    <h4 style="font-size:12px;color:#4338ca;text-transform:uppercase;margin-bottom:6px">Verdict</h4>
    <table>{cell('Result', d.get('result',''))}{cell('Remarks', d.get('remarks',''))}</table>
    </body></html>"""
    return HTMLResponse(html)

# --- TN01 ---
TN01_PRESET_ROUTES = {
    "PROD_TO_QC": {"title": "Production to Final Testing", "route": "Production ➔ Final Testing", "from_dept": "Production", "to_dept": "Final Testing", "columns": ["Sr. No.", "Job No.", "Customer Name", "Material Description", "Specified Length", "Entry Length", "Drum Size", "Drum Code", "Batch No. "]},
    "QC_TO_COMM": {"title": "Final Testing to Commercial", "route": "Final Testing ➔ Commercial", "from_dept": "Final Testing", "to_dept": "Commercial", "columns": ["Sr. No.", "Job / Prod Order #", "Batch / Heat Lot", "Cable Description / Spec", "Drum #", "Tested Length (Mtrs)", "HV Test Result", "Insulation Resistance", "QC Inspector", "Release Status"]},
    "PROD_TO_STORE": {"title": "Production to Store", "route": "Production ➔ Store", "from_dept": "Production", "to_dept": "Store", "columns": ["Sr. No.", "Work Order #", "Item Code / Material", "Batch / Heat Lot", "Quantity (Mtrs/Kgs)", "Weight (Kgs)", "Drum / Pack #", "Handover Person", "Remarks"]},
    "QC_TO_PROD": {"title": "Final Testing to Production (Rework/Reject)", "route": "Final Testing ➔ Production", "from_dept": "Final Testing", "to_dept": "Production", "columns": ["Sr. No.", "Job Order #", "Batch / Heat Lot", "Cable Size", "Defect Observed", "Length (Mtrs)", "QC Engineer", "Rework Instructions"]},
    "COMM_TO_QC": {"title": "Commercial to Final Testing (Inspection Request)", "route": "Commercial ➔ Final Testing", "from_dept": "Commercial", "to_dept": "Final Testing", "columns": ["Sr. No.", "Customer Name", "Order / PO #", "Item Description", "Ordered Qty", "Inspection Agency", "Target Date", "Special Testing Notes"]},
    "STORE_TO_PROD": {"title": "Store to Production (Raw Material Issue)", "route": "Store ➔ Production", "from_dept": "Store", "to_dept": "Production", "columns": ["Sr. No.", "Material Code", "Material Description", "Lot / GRN #", "Issued Qty", "UOM", "Issued By", "Received For Work Order"]}
}

@app.get("/api/tn01/routes")
def tn01_get_routes(req: Request): need(req); return TN01_PRESET_ROUTES

@app.get("/api/tn01/suggest-users")
def tn01_suggest_users(req: Request, plant: str = "", dept: str = ""):
    need(req)
    sql, params = "SELECT username, full_name, department, plants, role FROM users WHERE status='approved'", []
    if dept: sql += " AND department = ?"; params.append(dept)
    users = fetch_all(sql, tuple(params))
    if plant: return [u for u in users if not (u.get("plants") or "").strip() or plant in [p.strip() for p in (u.get("plants") or "").split(",")]]
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
    name = payload.get("name", "Custom Layout").strip()
    key = payload.get("key", "").strip().upper()
    cols = payload.get("columns", [])
    if not key:
        cnt = fetch_one("SELECT COUNT(*) as c FROM tcode_layouts WHERE tcode_code='TN01'")["c"]
        key = f"TN01_VAR_{cnt+1:02d}"
    cur_time = now()
    existing = fetch_one("SELECT id FROM tcode_layouts WHERE tcode_code='TN01' AND layout_code=?", (key,))
    if existing: exec_write("UPDATE tcode_layouts SET name=?, title=?, fields=?, updated_at=? WHERE id=?", (name, name, json.dumps(cols), cur_time, existing["id"]))
    else: exec_write("INSERT INTO tcode_layouts (tcode_code, layout_code, name, title, is_default, fields, layout_config, description, created_by, created_by_name, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                     ("TN01", key, name, name, 0, json.dumps(cols), "{}", "Custom Format", u["username"], u["full_name"], cur_time, cur_time))
    return {"ok": True, "key": key}

@app.delete("/api/tn01/layouts/{lid}")
def tn01_delete_layout(req: Request, lid: int): need_admin(req); exec_write("DELETE FROM tcode_layouts WHERE id=? AND tcode_code='TN01'", (lid,)); return {"ok": True}

@app.post("/api/tn01/send")
def tn01_send(req: Request, payload: dict):
    u = need(req); cur_time = now()
    today_tag = datetime.date.today().strftime("%Y%m%d")
    cnt_row = fetch_one("SELECT COUNT(*) as c FROM transfer_notes WHERE created_at LIKE ?", (f"{datetime.date.today().strftime('%Y-%m-%d')}%",))
    note_no = f"TN-{today_tag}-{(cnt_row['c'] if cnt_row else 0) + 1:04d}"
    headers = payload.get("headers", []); rows_data = payload.get("rows", [])
    to_username = payload.get("to_user", "").strip().lower()
    target = fetch_one("SELECT username, full_name FROM users WHERE username=?", (to_username,))
    to_user_name = target["full_name"] if target else to_username
    from_plant = payload.get("from_plant", "").strip() or (u.get("plants") or "1000").split(",")[0].strip()
    exec_write("INSERT INTO transfer_notes (note_no, layout_key, route, from_plant, from_dept, from_user, from_user_name, to_plant, to_dept, to_user, to_user_name, entry_count, sheet_headers, sheet_data, status, remarks, created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
               (note_no, payload.get("layout_key", "STANDARD"), payload.get("route", "Transfer"),
                from_plant, u["department"], u["username"], u["full_name"],
                payload.get("to_plant", "").strip(), payload.get("to_dept", "").strip(),
                to_username, to_user_name, len(rows_data), json.dumps(headers),
                json.dumps(rows_data, default=json_serial), "sent", payload.get("remarks", "").strip(), cur_time))
    return {"ok": True, "note_no": note_no, "entries": len(rows_data)}

@app.get("/api/tn01/inbox")
def tn01_get_inbox(req: Request):
    u = need(req)
    return fetch_all("SELECT * FROM transfer_notes WHERE to_user = ? OR (to_dept = ? AND to_user = '') ORDER BY id DESC", (u["username"].lower(), u["department"]))

@app.post("/api/tn01/receive/{nid}")
def tn01_receive(req: Request, nid: int):
    u = need(req); cur_time = now()
    exec_write("UPDATE transfer_notes SET status='received', received_at=?, received_by=?, received_by_name=? WHERE id=?",
               (cur_time, u["username"], u["full_name"], nid))
    return {"ok": True, "received_at": cur_time}

@app.get("/api/tn01/history")
def tn01_get_history(req: Request, plant: str = "", route: str = "", status: str = "",
                     from_dept: str = "", to_dept: str = "", date_from: str = "", date_to: str = ""):
    u = need(req); sql, params = "SELECT * FROM transfer_notes WHERE 1=1", []
    if not (is_master(u) or u["role"] == "admin"):
        sql += " AND (from_user = ? OR to_user = ? OR from_dept = ? OR to_dept = ?)"
        params.extend([u["username"], u["username"], u["department"], u["department"]])
    if plant: sql += " AND (from_plant = ? OR to_plant = ?)"; params.extend([plant, plant])
    if route: sql += " AND route = ?"; params.append(route)
    if status: sql += " AND status = ?"; params.append(status)
    if (is_master(u) or u["role"] == "admin"):
        if from_dept: sql += " AND from_dept = ?"; params.append(from_dept)
        if to_dept: sql += " AND to_dept = ?"; params.append(to_dept)
        if date_from: sql += " AND created_at >= ?"; params.append(date_from + " 00:00:00")
        if date_to: sql += " AND created_at <= ?"; params.append(date_to + " 23:59:59")
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
    if not note: raise HTTPException(404, "Note not found")
    is_sender = (u["username"] == note["from_user"]) or (u["department"] == note["from_dept"])
    is_receiver = (u["username"] == note["to_user"]) or (u["department"] == note["to_dept"])
    if not (is_sender or is_receiver or is_master(u) or u["role"] == "admin"):
        raise HTTPException(403, "You do not have permission to export this note.")
    try: headers = json.loads(note["sheet_headers"])
    except: headers = []
    try: data_rows = json.loads(note["sheet_data"])
    except: data_rows = []
    wb = Workbook(); ws = wb.active; ws.title = "Transfer_Note"
    org = fetch_one("SELECT * FROM org WHERE id=1") or {}
    ws.append([org.get("name", "Thermocables Limited")])
    ws.append([f"TRANSFER NOTE {note['note_no']} — {note['route']}"])
    ws.append([f"From: {note['from_plant']} - {note['from_dept']}  |  To: {note['to_plant']} - {note['to_dept']}"])
    ws.append([f"Sender: {note['from_user_name']}  |  Receiver: {note['to_user_name']}  |  Status: {note['status'].upper()}"])
    ws.append([f"Created: {note['created_at']}  |  Received: {note['received_at'] or 'Pending'}"])
    ws.append([]); ws.append(headers)
    for row in data_rows: ws.append(row)
    ws.append([]); ws.append(["Handovered By: ________________", "", "Received By: ________________", "", "QC/Incharge: ________________"])
    from openpyxl.styles import Font, Alignment, PatternFill
    hf = PatternFill("solid", fgColor="4F46E5"); hfont = Font(bold=True, color="FFFFFF")
    ws["A1"].font = Font(bold=True, size=14); ws["A2"].font = Font(bold=True)
    for ci in range(1, len(headers) + 1):
        c = ws.cell(row=7, column=ci); c.fill = hf; c.font = hfont; c.alignment = Alignment(horizontal="center")
    for ci, h in enumerate(headers, 1):
        ws.column_dimensions[ws.cell(row=7, column=ci).column_letter].width = max(14, min(40, len(str(h)) + 4))
    bio = io.BytesIO(); wb.save(bio); bio.seek(0)
    return StreamingResponse(bio, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             headers={"Content-Disposition": f"attachment; filename=TransferNote_{note['note_no']}.xlsx"})

@app.post("/api/tn01/export-grid")
def tn01_export_grid(req: Request, payload: dict):
    u = need(req)
    headers = payload.get("headers", []); rows = payload.get("rows", [])
    title = payload.get("title", "Transfer Note (Draft)")
    if not headers: raise HTTPException(400, "No columns to export.")
    wb = Workbook(); ws = wb.active; ws.title = "Draft_Note"
    org = fetch_one("SELECT * FROM org WHERE id=1") or {}
    ws.append([org.get("name", "Thermocables Limited")])
    ws.append([title])
    ws.append([f"Prepared by: {u['full_name']} ({u['username']}) · {u['department']}"])
    ws.append([f"Generated: {now()}"]); ws.append([]); ws.append(headers)
    for r in rows: ws.append(r)
    from openpyxl.styles import Font, Alignment, PatternFill
    hf = PatternFill("solid", fgColor="4F46E5"); hfont = Font(bold=True, color="FFFFFF")
    ws["A1"].font = Font(bold=True, size=14); ws["A2"].font = Font(bold=True)
    for ci in range(1, len(headers) + 1):
        c = ws.cell(row=6, column=ci); c.fill = hf; c.font = hfont; c.alignment = Alignment(horizontal="center")
    for ci, h in enumerate(headers, 1):
        ws.column_dimensions[ws.cell(row=6, column=ci).column_letter].width = max(14, min(40, len(str(h)) + 4))
    bio = io.BytesIO(); wb.save(bio); bio.seek(0)
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
    return HTMLResponse(f"""<!DOCTYPE html><html><head><meta charset="utf-8"><title>{note['note_no']}</title>
    <style>body{{font-family:sans-serif;margin:30px;color:#0f172a}}.header{{border-bottom:2px solid #4f46e5;padding-bottom:10px;margin-bottom:15px}}.meta{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;font-size:12px;background:#f8fafc;padding:12px;border:1px solid #e2e8f0;margin-bottom:20px}}table{{width:100%;border-collapse:collapse;margin-bottom:30px}}.signs{{display:grid;grid-template-columns:repeat(3,1fr);gap:20px;margin-top:40px;text-align:center;font-size:11px;font-weight:bold}}.sign-box{{border-top:1px solid #0f172a;padding-top:8px}}@media print{{.no-print{{display:none}}body{{margin:0}}}}</style></head><body>
      <div class="no-print" style="margin-bottom:15px;text-align:right"><button onclick="window.print()" style="background:#4f46e5;color:white;border:none;padding:8px 16px;border-radius:6px;cursor:pointer;font-weight:bold">Print / Save PDF</button></div>
      <div class="header"><h2 style="margin:0">{org.get('name','Thermocables Limited')}</h2><h4 style="margin:4px 0 0;color:#4f46e5">INTER-DEPARTMENT MATERIAL TRANSFER & HANDOVER NOTE (TN01)</h4><div style="font-family:monospace;font-weight:bold;margin-top:4px">{note['note_no']} | Date: {note['created_at']}</div></div>
      <div class="meta">
        <div><b>Route:</b><br>{note['route']}</div><div><b>From:</b><br>{note['from_plant']} - {note['from_dept']}</div>
        <div><b>To:</b><br>{note['to_plant']} - {note['to_dept']}</div><div><b>Entries:</b><br>{note['entry_count']} Batches</div>
        <div><b>Sender:</b><br>{note['from_user_name']}</div><div><b>Recipient:</b><br>{note['to_user_name']}</div>
        <div><b>Status:</b><br><span style="text-transform:uppercase;font-weight:bold">{note['status']}</span></div><div><b>Receipt:</b><br>{note['received_at'] or 'Pending'}</div>
      </div>
      <table><thead><tr>{th}</tr></thead><tbody>{trs}</tbody></table>
      <div class="signs"><div class="sign-box">Handovered By ({note['from_user_name']})</div><div class="sign-box">Production | Quality Incharge</div><div class="sign-box">Received By ({note['received_by_name'] or note['to_user_name']})</div></div>
    </body></html>""")

# --- STANDARD MODULES ---
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
    if plant: sql += " AND plant=?"; params.append(plant)
    sql += " ORDER BY id DESC LIMIT ?"; params.append(limit)
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
    wb = Workbook(); ws = wb.active; ws.title = f"{code}_DATA"
    keys, dyn, parsed = ["id", "plant", "created_by_name", "created_at"], set(), []
    for r in rows:
        try: d = json.loads(r["data"]) if isinstance(r["data"], str) else r["data"]
        except: d = {}
        dyn.update(d.keys()); parsed.append((r, d))
    cols = keys + sorted(dyn)
    ws.append(cols)
    for r, d in parsed:
        ws.append([r["id"], r["plant"], r["created_by_name"], r["created_at"]] + [d.get(k, "") for k in sorted(dyn)])
    bio = io.BytesIO(); wb.save(bio); bio.seek(0)
    return StreamingResponse(bio, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             headers={"Content-Disposition": f"attachment; filename=Export_{code}_{datetime.date.today().strftime('%Y%m%d')}.xlsx"})

# =========================================================================
# --- LIGHT PROFESSIONAL FRONTEND ---
# =========================================================================
HTML_MODERN_UI = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Nexus Core — Enterprise Operations</title>
<script src="https://cdn.tailwindcss.com"></script>
<link href="https://cdn.jsdelivr.net/npm/remixicon@4.2.0/fonts/remixicon.css" rel="stylesheet">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">
<style>
  :root {
    --bg: #f8fafc;
    --surface: #ffffff;
    --border: #e2e8f0;
    --border-strong: #cbd5e1;
    --text: #0f172a;
    --text-muted: #64748b;
    --accent: #4f46e5;
    --accent-hover: #4338ca;
    --accent-soft: #eef2ff;
    --success: #059669;
    --warn: #d97706;
    --danger: #dc2626;
  }
  * { -webkit-font-smoothing: antialiased; }
  html, body { height: 100%; margin: 0; }
  body { font-family: 'Inter', system-ui, -apple-system, sans-serif; background: var(--bg); color: var(--text); font-size: 13px; }
  .font-mono { font-family: 'JetBrains Mono', monospace; }

  /* Buttons */
  .btn-primary { background: var(--accent); color: #fff; border: 1px solid var(--accent); transition: background .15s ease, box-shadow .15s ease; }
  .btn-primary:hover { background: var(--accent-hover); box-shadow: 0 4px 12px rgba(79,70,229,.20); }
  .btn-primary:disabled { background: #a5b4fc; border-color: #a5b4fc; cursor: not-allowed; box-shadow: none; }
  .btn-secondary { background: var(--surface); color: #475569; border: 1px solid var(--border); transition: background .15s ease, border-color .15s ease; }
  .btn-secondary:hover { background: #f1f5f9; border-color: var(--border-strong); }

  /* Cards & containers */
  .card { background: var(--surface); border: 1px solid var(--border); border-radius: 14px; box-shadow: 0 1px 2px rgba(15,23,42,.03); }

  /* Tables */
  .custom-table { width: 100%; border-collapse: collapse; }
  .custom-table thead th { background: #f1f5f9; color: #475569; font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: .03em; padding: 10px 14px; border-bottom: 1px solid var(--border); text-align: left; }
  .custom-table tbody td { padding: 11px 14px; border-bottom: 1px solid #f1f5f9; font-size: 13px; color: #334155; }
  .custom-table tbody tr:hover { background: #f8fafc; }

  /* Excel grid */
  .excel-grid { border-collapse: separate; border-spacing: 0; width: 100%; }
  .excel-grid th { background: #f1f5f9; color: #475569; font-weight: 600; font-size: 11px; text-transform: uppercase; padding: 8px 10px; border-right: 1px solid var(--border); border-bottom: 1px solid var(--border); position: sticky; top: 0; z-index: 2; text-align: left; }
  .excel-grid td { background: #fff; color: var(--text); font-size: 12px; padding: 6px 10px; border-right: 1px solid #f1f5f9; border-bottom: 1px solid #f1f5f9; outline: none; min-width: 110px; }
  .excel-grid td.sr-col { background: #f8fafc; color: #64748b; font-weight: 600; text-align: center; width: 50px; user-select: none; }
  .excel-grid td:focus { background: #eef2ff; box-shadow: inset 0 0 0 2px #6366f1; }

  /* Forms */
  .form-input { background: var(--surface); border: 1px solid var(--border-strong); border-radius: 8px; padding: 8px 12px; font-size: 13px; color: var(--text); width: 100%; outline: none; transition: border-color .15s ease, box-shadow .15s ease; font-family: inherit; }
  .form-input:focus { border-color: var(--accent); box-shadow: 0 0 0 3px rgba(99,102,241,.12); }
  .form-label { display: block; font-size: 11px; font-weight: 600; color: #475569; margin-bottom: 5px; letter-spacing: .01em; }

  /* Badges */
  .badge { display: inline-flex; align-items: center; font-size: 10px; font-weight: 600; padding: 2px 8px; border-radius: 999px; letter-spacing: .02em; text-transform: uppercase; }
  .badge-indigo { background: #eef2ff; color: #4338ca; border: 1px solid #c7d2fe; }
  .badge-emerald { background: #ecfdf5; color: #047857; border: 1px solid #a7f3d0; }
  .badge-amber { background: #fffbeb; color: #b45309; border: 1px solid #fde68a; }
  .badge-red { background: #fef2f2; color: #b91c1c; border: 1px solid #fecaca; }
  .badge-slate { background: #f1f5f9; color: #475569; border: 1px solid #e2e8f0; }

  /* Sidebar */
  .sidebar-link { color: #475569; border-left: 2px solid transparent; transition: all .15s ease; }
  .sidebar-link:hover { background: #f1f5f9; color: #0f172a; }
  .sidebar-link.active { background: var(--accent-soft); color: var(--accent); border-left-color: var(--accent); }

  /* Scrollbar */
  ::-webkit-scrollbar { width: 10px; height: 10px; }
  ::-webkit-scrollbar-track { background: #f1f5f9; }
  ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 5px; }
  ::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
</style>
</head>
<body class="h-screen flex flex-col overflow-hidden">

<div id="app_root" class="h-full flex flex-col">
  <header class="h-14 bg-white border-b border-slate-200 px-4 flex items-center justify-between z-10 shrink-0 shadow-sm">
    <div class="flex items-center gap-4">
      <div class="flex items-center gap-2.5 cursor-pointer" onclick="goHome()">
        <div class="w-8 h-8 rounded-lg bg-indigo-600 flex items-center justify-center font-bold text-white shadow-sm"><i class="ri-pulse-fill text-lg"></i></div>
        <div class="font-bold text-sm tracking-tight text-slate-900 flex items-center gap-1.5">Nexus Core <span class="text-[10px] bg-indigo-50 text-indigo-600 border border-indigo-200 font-mono px-1.5 py-0.5 rounded">v3.8</span></div>
      </div>
      <div class="relative ml-4">
        <i class="ri-search-line absolute left-3 top-2.5 text-slate-400 text-xs"></i>
        <input id="cmd_box" placeholder="Jump to module (TN01, CTST, SU01, SPRO, STUDIO)..." class="bg-slate-50 border border-slate-200 rounded-lg pl-8 pr-4 py-1.5 text-xs text-slate-900 w-96 font-mono focus:outline-none focus:border-indigo-400 focus:bg-white" onkeydown="if(event.key==='Enter') execCmd()">
      </div>
    </div>
    <div class="flex items-center gap-3">
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
function esc(str) { if (str === null || str === undefined) return ''; return String(str).replace(/[&<>"']/g, s => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[s])); }
function setMsg(msg, isErr=false) { const el=$('#stat_msg'),d=$('#stat_dot'); el.textContent=msg; d.className=isErr?"w-2 h-2 rounded-full bg-red-500":"w-2 h-2 rounded-full bg-emerald-500"; el.className=isErr?"text-red-600 font-semibold":"text-slate-600"; }

async function api(path, opts = {}) {
  const token = localStorage.getItem("nexus_token");
  const headers = {...(opts.headers || {})};
  if (token) headers["Authorization"] = `Bearer ${token}`;
  const t0 = performance.now();
  const res = await fetch(path, {...opts, headers, credentials: 'include'});
  $('#stat_resp').textContent = Math.round(performance.now() - t0) + 'ms';
  if (res.status === 401) { renderLogon(); throw 'Auth'; }
  if (!res.ok) { let t = await res.text(); try { t = JSON.parse(t).detail || t; } catch {} throw t; }
  return await res.json();
}

function renderLogon(errorMsg="") {
  $('#app_root').innerHTML = `
  <div class="min-h-screen bg-gradient-to-br from-slate-50 to-indigo-50 flex items-center justify-center p-4">
    <div class="w-full max-w-md bg-white border border-slate-200 rounded-2xl shadow-xl p-8">
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

async function logout() { await api('/api/logout', {method: 'POST'}); localStorage.removeItem("nexus_token"); location.reload(); }

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
          <div class="card p-5">
            <div class="text-[10px] font-semibold text-slate-500 uppercase tracking-wider">Enterprise Org</div>
            <div class="text-lg font-bold text-slate-900 mt-1.5">${esc(ME.company)}</div>
            <div class="text-xs text-indigo-600 mt-1 font-medium">Dept: ${esc(ME.department)}</div>
          </div>
          <div class="card p-5">
            <div class="text-[10px] font-semibold text-slate-500 uppercase tracking-wider">Profile Role</div>
            <div class="text-lg font-bold text-emerald-600 mt-1.5 uppercase">${ME.role}</div>
            <div class="text-xs text-slate-500 mt-1 truncate">Plants: ${esc(ME.plants.join(', ') || 'Global')}</div>
          </div>
          <div class="card p-5">
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
    <div onclick="openTcode('${t.code}')" class="card hover:border-indigo-300 hover:shadow-md p-5 cursor-pointer transition-all">
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
// --- CTST : PICK-YOUR-TEST HUB ---
// =========================================================================
const CTST_TESTS = {
  cr:       {title: "CR — Conductor Resistance", icon: "ri-pulse-line",        desc: "Temperature-corrected conductor resistance per metre & per km"},
  cr_pct:   {title: "CR Percentage vs Spec",     icon: "ri-bar-chart-line",    desc: "Deviation and length-wise variation vs specified CR"},
  braid:    {title: "Braid / Shield Coverage",   icon: "ri-shield-check-line", desc: "Fill factor and % coverage for wire braid screens"},
  tubular:  {title: "Tubular Mechanical Test",   icon: "ri-circle-line",       desc: "Area, tensile strength, elongation for tube specimen"},
  dumbbell: {title: "Dumbbell Mechanical Test",  icon: "ri-shape-line",        desc: "Area, tensile strength, elongation for dumbbell specimen"},
  hotset:   {title: "Hot Set & Permanent Set",   icon: "ri-temp-hot-line",     desc: "Hot-set and permanent-set elongation percentages"},
};

async function renderCTST(testType = null) {
  if (testType) return renderCTSTForm(testType);
  const allRows = await api('/api/ctst/records');

  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col p-6 overflow-hidden select-text">
      <div class="flex items-center justify-between pb-4 border-b border-slate-200 shrink-0">
        <div class="flex items-center gap-3">
          <span class="font-mono font-bold text-xs bg-indigo-600 text-white px-2.5 py-1 rounded-lg shadow-sm">CTST</span>
          <div>
            <h2 class="text-lg font-bold text-slate-900">Cable Type-Test Calculator</h2>
            <p class="text-xs text-slate-500">Pick a test — only fill what you need</p>
          </div>
        </div>
        <div class="flex items-center gap-2">
          <a href="/api/ctst/export" class="btn-secondary text-xs font-semibold px-3 py-2 rounded-lg flex items-center gap-1"><i class="ri-download-2-line"></i> Export All</a>
          <button onclick="renderCTST()" class="btn-secondary text-xs font-semibold p-2 rounded-lg"><i class="ri-refresh-line"></i></button>
        </div>
      </div>

      <div class="mt-4 grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 shrink-0">
        ${Object.entries(CTST_TESTS).map(([k, t]) => `
          <div onclick="renderCTST('${k}')" class="card hover:border-indigo-300 hover:shadow-md p-5 cursor-pointer transition-all group">
            <div class="flex items-start justify-between mb-2">
              <div class="w-10 h-10 rounded-xl bg-indigo-50 border border-indigo-200 flex items-center justify-center text-indigo-600 group-hover:bg-indigo-600 group-hover:text-white transition">
                <i class="${t.icon} text-lg"></i>
              </div>
              <span class="badge badge-indigo">${k.toUpperCase()}</span>
            </div>
            <h4 class="font-bold text-slate-900 text-sm mt-2">${t.title}</h4>
            <p class="text-xs text-slate-500 mt-1">${t.desc}</p>
            <div class="mt-3 flex items-center text-xs font-semibold text-indigo-600 gap-1 group-hover:gap-2 transition-all">
              Open calculator <i class="ri-arrow-right-line"></i>
            </div>
          </div>
        `).join('')}
      </div>

      <div class="flex-1 mt-6 card overflow-hidden flex flex-col">
        <div class="p-3 border-b border-slate-200 bg-slate-50">
          <span class="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5">
            <i class="ri-history-line text-indigo-500"></i> Recent Test Records (${allRows.length})
          </span>
        </div>
        <div class="flex-1 overflow-auto">
          <table class="custom-table">
            <thead><tr><th>#</th><th>Type</th><th>Sample</th><th>Cable</th><th>Date</th><th>Key Result</th><th>Verdict</th><th style="width:160px;">Actions</th></tr></thead>
            <tbody>
              ${!allRows.length ? `<tr><td colspan="8" class="text-center py-10 text-slate-400">No tests saved yet. Pick a test above to begin.</td></tr>` : ''}
              ${allRows.map(r => {
                const d = r.data || {}, c = r.computed || {};
                const tt = d._test_type || '';
                const t = CTST_TESTS[tt] || {title: tt};
                const vb = d.result === 'PASS' ? 'badge-emerald' : (d.result === 'FAIL' ? 'badge-red' : 'badge-amber');
                let key = '—';
                if (tt === 'cr') key = c.cr_1m !== undefined ? `CR ${c.cr_1m} Ω/m` : '—';
                else if (tt === 'cr_pct') key = c.cr_deviation_pct !== undefined ? `Dev ${c.cr_deviation_pct}%` : '—';
                else if (tt === 'braid') key = c.braid_coverage_pct !== undefined ? `Cov ${c.braid_coverage_pct}%` : '—';
                else if (tt === 'tubular' || tt === 'dumbbell') key = c.tensile_nmm2 !== undefined ? `${c.tensile_nmm2} N/mm²` : '—';
                else if (tt === 'hotset') key = c.hotset_pct !== undefined ? `Hot ${c.hotset_pct}%` : '—';
                return `<tr>
                  <td class="font-mono text-slate-400 font-semibold">${r.id}</td>
                  <td><span class="badge badge-indigo">${esc((t.title||'').split('—')[0].trim() || tt)}</span></td>
                  <td class="font-semibold text-slate-900">${esc(d.sample_id || '—')}</td>
                  <td class="text-xs">${esc(d.cable_type || '—')}</td>
                  <td class="text-xs text-slate-500 font-mono">${esc(d.test_date || '—')}</td>
                  <td class="font-mono text-xs">${key}</td>
                  <td>${d.result ? `<span class="badge ${vb}">${esc(d.result)}</span>` : `<span class="text-slate-400 text-xs">—</span>`}</td>
                  <td><a href="/api/ctst/print/${r.id}" target="_blank" class="text-xs text-indigo-600 font-semibold">Print</a><button onclick="deleteCTST(${r.id})" class="text-xs text-red-600 font-semibold ml-2">Del</button></td>
                </tr>`;
              }).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>`;
}

const CTST_FIELDS = {
  cr: [
    {key: "obd_cr_reading", label: "OBD CR Reading (Ω)", step: "any"},
    {key: "temp_c",         label: "Test Temp (°C)",     step: "any", placeholder: "e.g. 25"},
    {key: "length_m",       label: "Length (m)",         step: "any"},
  ],
  cr_pct: [
    {key: "specified_cr", label: "Specified CR", step: "any"},
    {key: "observed_cr",  label: "Observed CR",  step: "any"},
    {key: "cr_max",       label: "Obs. Max CR",  step: "any"},
    {key: "cr_min",       label: "Obs. Min CR",  step: "any"},
  ],
  braid: [
    {key: "braid_d",           label: "Braid Wire Ø (mm)", step: "any"},
    {key: "under_braid_D",     label: "Ø Under Braid (mm)", step: "any"},
    {key: "wires_per_carrier", label: "Wires / Carrier", step: "1"},
    {key: "carriers",          label: "Carriers", step: "1"},
    {key: "lay_length",        label: "Lay Length (mm)", step: "any"},
  ],
  tubular: [
    {key: "tub_outer_d",    label: "Outer Ø (mm)", step: "any"},
    {key: "tub_wall",       label: "Wall Thickness (mm)", step: "any"},
    {key: "break_load_tub", label: "Break Load (N)", step: "any"},
    {key: "initial_gauge",  label: "Initial Gauge (mm)", step: "any", value: "20"},
    {key: "elong_tub_final",label: "Final Gauge (mm)", step: "any"},
  ],
  dumbbell: [
    {key: "dumb_width",     label: "Width (mm)", step: "any"},
    {key: "dumb_thk",       label: "Thickness (mm)", step: "any"},
    {key: "break_load_dumb",label: "Break Load (N)", step: "any"},
    {key: "initial_gauge",  label: "Initial Gauge (mm)", step: "any", value: "20"},
    {key: "elong_dumb_final",label: "Final Gauge (mm)", step: "any"},
  ],
  hotset: [
    {key: "initial_gauge",       label: "Initial Gauge (mm)", step: "any", value: "20"},
    {key: "hotset_elong_final",  label: "Hot Set Final (mm)", step: "any"},
    {key: "permset_cooled",      label: "Perm. Set Cooled (mm)", step: "any"},
  ],
};

function renderCTSTForm(testType) {
  const t = CTST_TESTS[testType];
  const fields = CTST_FIELDS[testType] || [];

  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col p-6 overflow-hidden select-text">
      <div class="flex items-center justify-between pb-4 border-b border-slate-200 shrink-0">
        <div class="flex items-center gap-3">
          <button onclick="renderCTST()" class="btn-secondary text-xs p-2 rounded-lg"><i class="ri-arrow-left-line"></i></button>
          <div class="w-10 h-10 rounded-xl bg-indigo-600 flex items-center justify-center text-white shadow-sm"><i class="${t.icon} text-lg"></i></div>
          <div>
            <h2 class="text-lg font-bold text-slate-900">${t.title}</h2>
            <p class="text-xs text-slate-500">${t.desc}</p>
          </div>
        </div>
        <a href="/api/ctst/export?test_type=${testType}" class="btn-secondary text-xs font-semibold px-3 py-2 rounded-lg flex items-center gap-1"><i class="ri-download-2-line"></i> Export ${testType.toUpperCase()}</a>
      </div>

      <div class="flex-1 mt-4 overflow-auto">
        <div class="max-w-3xl mx-auto">
          <div class="card p-6 space-y-5">
            <div class="grid grid-cols-3 gap-4">
              <div><label class="form-label">Sample / Batch ID *</label><input id="cf_sample_id" class="form-input font-mono" placeholder="e.g. LT-XLPE-4C-95-B12"></div>
              <div><label class="form-label">Cable Type / Size</label><input id="cf_cable_type" class="form-input" placeholder="e.g. 1.1kV XLPE 4Cx95"></div>
              <div><label class="form-label">Test Date</label><input id="cf_test_date" type="date" class="form-input"></div>
            </div>

            <div class="pt-4 border-t border-slate-100">
              <h3 class="text-xs font-bold text-indigo-700 uppercase tracking-wider mb-3">${t.title} — Inputs</h3>
              <div class="grid grid-cols-2 gap-4">
                ${fields.map(f => `
                  <div>
                    <label class="form-label">${f.label}</label>
                    <input id="cf_${f.key}" type="number" step="${f.step||'any'}" value="${f.value||''}" placeholder="${f.placeholder||''}" class="form-input">
                  </div>
                `).join('')}
              </div>
            </div>

            <div class="pt-4 border-t border-slate-100 grid grid-cols-2 gap-4">
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
              <textarea id="cf_remarks" class="form-input" rows="2" placeholder="Observations, deviations, re-test notes..."></textarea>
            </div>

            <div class="bg-indigo-50 border border-indigo-200 rounded-xl p-4">
              <div class="flex justify-between items-center mb-3">
                <div class="text-xs font-bold text-indigo-800 uppercase tracking-wider">Live Computed Results</div>
                <button onclick="previewCTSTForm('${testType}')" class="text-[11px] font-bold text-indigo-600 hover:text-indigo-800 flex items-center gap-1"><i class="ri-calculator-line"></i> Recompute</button>
              </div>
              <div id="cf_computed" class="grid grid-cols-3 gap-2 text-xs">
                <div class="text-slate-400 col-span-3">Fill fields above, then press Recompute.</div>
              </div>
            </div>

            <div class="flex justify-end gap-2 pt-2 border-t border-slate-100">
              <button onclick="renderCTST()" class="btn-secondary px-4 py-2 rounded-lg text-xs font-semibold">Cancel</button>
              <button onclick="saveCTST('${testType}')" class="btn-primary px-5 py-2 rounded-lg text-xs font-bold shadow-sm">Save Test Record</button>
            </div>
          </div>
        </div>
      </div>
    </div>`;

  $('#cf_test_date').value = new Date().toISOString().slice(0,10);
}

async function previewCTSTForm(testType) {
  const fields = CTST_FIELDS[testType] || [];
  const payload = { _test_type: testType };
  payload.sample_id = ($('#cf_sample_id') || {}).value || '';
  payload.cable_type = ($('#cf_cable_type') || {}).value || '';
  payload.test_date = ($('#cf_test_date') || {}).value || '';
  for (const f of fields) {
    const el = $('#cf_' + f.key);
    if (el) payload[f.key] = el.value;
  }
  try {
    const res = await api('/api/ctst/compute', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(payload)});
    const c = res.computed || {};
    const fmtKey = {
      cr_1m: ['CR 1 m','Ω'], cr_total_per_km: ['CR per km','Ω/km'], cr_warning: ['Warning',''],
      cr_deviation_pct: ['CR Deviation','%'], cr_variation_pct: ['CR Variation','%'],
      braid_total_wires: ['Total Wires',''], braid_filling_factor: ['Filling Factor',''], braid_coverage_pct: ['Coverage','%'],
      area_mm2: ['Area','mm²'], tensile_nmm2: ['Tensile Strength','N/mm²'], elongation_pct: ['Elongation','%'],
      hotset_pct: ['Hot Set','%'], permset_pct: ['Permanent Set','%']
    };
    const keys = Object.keys(c);
    if (!keys.length) {
      $('#cf_computed').innerHTML = '<div class="text-slate-500 col-span-3">No results yet — fill the inputs above.</div>';
      return;
    }
    $('#cf_computed').innerHTML = keys.map(k => {
      const [lbl, unit] = fmtKey[k] || [k, ''];
      const isWarn = k === 'cr_warning';
      const cls = isWarn ? 'col-span-3 bg-amber-50 border-amber-200 text-amber-800' : 'bg-white border-indigo-100';
      return `<div class="${cls} border rounded-lg p-2"><div class="text-[9px] uppercase text-slate-500 font-semibold">${lbl}</div><div class="font-mono font-bold text-slate-900 text-sm">${c[k]}${unit ? ' '+unit : ''}</div></div>`;
    }).join('');
  } catch(e) { setMsg("Preview failed: " + e, true); }
}

async function saveCTST(testType) {
  const fields = CTST_FIELDS[testType] || [];
  const payload = { _test_type: testType };
  payload.sample_id = ($('#cf_sample_id') || {}).value || '';
  if (!payload.sample_id.trim()) return alert("Sample / Batch ID is required.");
  payload.cable_type = ($('#cf_cable_type') || {}).value || '';
  payload.test_date = ($('#cf_test_date') || {}).value || '';
  payload.result = ($('#cf_result') || {}).value || '';
  payload.remarks = ($('#cf_remarks') || {}).value || '';
  payload._plant = ($('#cf_plant') || {}).value || '';
  for (const f of fields) {
    const el = $('#cf_' + f.key);
    if (el) payload[f.key] = el.value;
  }
  try {
    await api('/api/ctst/save', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(payload)});
    setMsg("Test record saved");
    renderCTST();
  } catch(e) { alert(e); }
}

async function deleteCTST(id) {
  if (!confirm(`Delete test record #${id}?`)) return;
  try { await api(`/api/ctst/records/${id}`, {method: 'DELETE'}); renderCTST(); setMsg("Record deleted"); }
  catch(e) { alert(e); }
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
        <div><h2 class="text-lg font-bold text-slate-900">User Directory & IAM (SU01)</h2><p class="text-xs text-slate-500">Manage users, plant access, roles, status</p></div>
        <div class="flex items-center gap-2">
          <input id="su01_search" oninput="filterUsersTable(this.value)" placeholder="Search users..." class="form-input w-64">
          <button onclick="showCreateUserModal()" class="btn-primary text-xs font-bold px-4 py-2 rounded-lg flex items-center gap-1.5 shadow-sm"><i class="ri-user-add-line"></i> Add User</button>
          <button onclick="renderSU01()" class="btn-secondary text-xs p-2 rounded-lg"><i class="ri-refresh-line"></i></button>
        </div>
      </div>
      <div class="flex-1 mt-4 card overflow-auto">
        <table class="custom-table"><thead><tr><th>Username</th><th>Full Name</th><th>Role</th><th>Department</th><th>Facilities</th><th>Status</th><th style="width:200px;">Actions</th></tr></thead><tbody id="su01_tbody"></tbody></table>
      </div>
    </div>`;
  renderUsersRows(ALL_USERS_LIST);
}

function filterUsersTable(q) {
  const term = q.toLowerCase();
  renderUsersRows(ALL_USERS_LIST.filter(u =>
    u.username.toLowerCase().includes(term) || u.full_name.toLowerCase().includes(term) ||
    (u.department||'').toLowerCase().includes(term) || (u.plants||'').toLowerCase().includes(term)));
}

function renderUsersRows(users) {
  const tbody = $('#su01_tbody');
  if (!tbody) return;
  if (!users.length) { tbody.innerHTML = `<tr><td colspan="7" class="text-center py-10 text-slate-400">No users found.</td></tr>`; return; }
  tbody.innerHTML = users.map(u => {
    const isSelf = (u.username === ME.username);
    const isM = (u.role === 'master');
    const canDel = !isSelf && (!isM || ME.is_master);
    const rb = u.role === 'master' ? 'badge-red' : (u.role === 'admin' ? 'badge-amber' : 'badge-slate');
    const sb = u.status === 'approved' ? 'badge-emerald' : 'badge-red';
    return `<tr>
      <td class="font-mono font-bold text-indigo-600">${esc(u.username)}</td>
      <td class="font-medium text-slate-900">${esc(u.full_name)}</td>
      <td><span class="badge ${rb}">${u.role.toUpperCase()}</span></td>
      <td>${esc(u.department)}</td>
      <td class="font-mono text-xs">${esc(u.plants || 'ALL')}</td>
      <td><span class="badge ${sb}">${u.status}</span></td>
      <td>
        <button onclick="showEditUserModal('${esc(u.username)}')" class="text-xs text-indigo-600 font-semibold">Edit</button>
        <button onclick="showResetPassModal('${esc(u.username)}')" class="text-xs text-amber-600 font-semibold ml-2">Reset PW</button>
        ${canDel ? `<button onclick="deleteUser('${esc(u.username)}')" class="text-xs text-red-600 font-semibold ml-2">Delete</button>` : ''}
      </td>
    </tr>`;
  }).join('');
}

async function showCreateUserModal() {
  const data = await api('/api/spro/overview');
  openDialog(`
    <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl">
      <h3 class="font-bold text-slate-900 text-sm">Create New User</h3>
      <button onclick="closeDialog()"><i class="ri-close-line text-lg text-slate-500"></i></button>
    </div>
    <div class="p-6 space-y-4">
      <div class="grid grid-cols-2 gap-4">
        <div><label class="form-label">Username *</label><input id="nu_u" class="form-input font-mono"></div>
        <div><label class="form-label">Password *</label><input id="nu_p" type="password" class="form-input"></div>
      </div>
      <div><label class="form-label">Full Name *</label><input id="nu_fn" class="form-input"></div>
      <div class="grid grid-cols-2 gap-4">
        <div><label class="form-label">Department</label><select id="nu_d" class="form-input">${data.departments.map(d=>`<option value="${esc(d.name)}">${esc(d.name)}</option>`).join('')}</select></div>
        <div><label class="form-label">Role</label><select id="nu_r" class="form-input"><option value="user">User</option><option value="admin">Admin</option>${ME.is_master?'<option value="master">Master</option>':''}</select></div>
      </div>
      <div><label class="form-label">Plants (CSV, blank = ALL)</label><input id="nu_pl" class="form-input font-mono" placeholder="1000, 1400"></div>
    </div>
    <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl">
      <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-lg text-xs font-semibold">Cancel</button>
      <button onclick="submitCreateUser()" class="btn-primary px-4 py-2 rounded-lg text-xs font-semibold">Save</button>
    </div>`);
}

async function submitCreateUser() {
  const fd = new FormData();
  fd.append("username", $('#nu_u').value); fd.append("password", $('#nu_p').value);
  fd.append("full_name", $('#nu_fn').value); fd.append("department", $('#nu_d').value);
  fd.append("role", $('#nu_r').value); fd.append("plants", $('#nu_pl').value);
  try { await api('/api/su01/create', {method: 'POST', body: fd}); closeDialog(); renderSU01(); setMsg("User created"); }
  catch(e) { alert(e); }
}

async function showEditUserModal(u) {
  const users = await api('/api/su01/users');
  const user = users.find(x => x.username === u);
  const data = await api('/api/spro/overview');
  openDialog(`
    <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl">
      <h3 class="font-bold text-slate-900 text-sm">Update: ${esc(u)}</h3>
      <button onclick="closeDialog()"><i class="ri-close-line text-lg text-slate-500"></i></button>
    </div>
    <div class="p-6 space-y-4">
      <div><label class="form-label">Full Name</label><input id="eu_fn" value="${esc(user.full_name)}" class="form-input"></div>
      <div class="grid grid-cols-2 gap-4">
        <div><label class="form-label">Department</label><select id="eu_d" class="form-input">${data.departments.map(d=>`<option value="${esc(d.name)}" ${d.name===user.department?'selected':''}>${esc(d.name)}</option>`).join('')}</select></div>
        <div><label class="form-label">Role</label><select id="eu_r" class="form-input"><option value="user" ${user.role==='user'?'selected':''}>User</option><option value="admin" ${user.role==='admin'?'selected':''}>Admin</option>${ME.is_master?`<option value="master" ${user.role==='master'?'selected':''}>Master</option>`:''}</select></div>
      </div>
      <div class="grid grid-cols-2 gap-4">
        <div><label class="form-label">Facilities</label><input id="eu_pl" value="${esc(user.plants||'')}" class="form-input font-mono"></div>
        <div><label class="form-label">Status</label><select id="eu_st" class="form-input"><option value="approved" ${user.status==='approved'?'selected':''}>Approved</option><option value="locked" ${user.status==='locked'?'selected':''}>Locked</option></select></div>
      </div>
    </div>
    <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl">
      <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-lg text-xs font-semibold">Cancel</button>
      <button onclick="submitEditUser('${esc(u)}')" class="btn-primary px-4 py-2 rounded-lg text-xs font-semibold">Save</button>
    </div>`);
}

async function submitEditUser(u) {
  const fd = new FormData();
  fd.append("username", u); fd.append("full_name", $('#eu_fn').value);
  fd.append("department", $('#eu_d').value); fd.append("role", $('#eu_r').value);
  fd.append("plants", $('#eu_pl').value); fd.append("status", $('#eu_st').value);
  try { await api('/api/su01/update', {method: 'POST', body: fd}); closeDialog(); renderSU01(); setMsg("User updated"); }
  catch(e) { alert(e); }
}

function showResetPassModal(u) {
  const p = prompt(`New password for ${u}:`); if (!p) return;
  const fd = new FormData(); fd.append("username", u); fd.append("new_password", p);
  api('/api/su01/reset-password', {method: 'POST', body: fd}).then(()=>{setMsg(`Password reset for ${u}`);}).catch(e=>alert(e));
}

async function deleteUser(u) {
  if (!confirm(`Remove account ${u}?`)) return;
  const fd = new FormData(); fd.append("username", u);
  try { await api('/api/su01/delete', {method: 'POST', body: fd}); renderSU01(); setMsg(`User ${u} deleted`); }
  catch(e) { alert(e); }
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
        <div><h2 class="text-lg font-bold text-slate-900">Enterprise Structure (SPRO)</h2><p class="text-xs text-slate-500">Manage plants, departments, corporate identity</p></div>
        <button onclick="renderSPRO()" class="btn-secondary text-xs px-3 py-1.5 rounded-lg"><i class="ri-refresh-line"></i> Refresh</button>
      </div>
      <div class="mt-4 card p-4 flex items-center justify-between shrink-0">
        <div>
          <div class="text-[10px] text-slate-500 uppercase font-semibold">Corporate Identity</div>
          <div class="text-base font-bold text-slate-900">${esc(data.org.name)} <span class="text-xs font-mono text-indigo-600">(${esc(data.org.system_id)} · ${esc(data.org.client)})</span></div>
        </div>
        ${ME.is_master ? `<button onclick="showEditOrgModal('${esc(data.org.name)}','${esc(data.org.system_id)}','${esc(data.org.client)}')" class="btn-secondary text-xs px-3 py-1.5 rounded-lg">Edit</button>` : ''}
      </div>
      <div class="grid grid-cols-1 md:grid-cols-2 gap-4 flex-1 mt-4 overflow-auto">
        <div class="card flex flex-col overflow-hidden">
          <div class="p-3 border-b border-slate-200 flex justify-between items-center bg-slate-50">
            <span class="text-xs font-bold text-slate-700 uppercase tracking-wider"><i class="ri-building-line text-indigo-500"></i> Plants (${data.plants.length})</span>
            ${ME.is_master ? `<button onclick="showCreatePlantModal()" class="btn-primary text-xs font-bold px-3 py-1.5 rounded-lg">+ Add</button>` : ''}
          </div>
          <div class="flex-1 overflow-auto">
            <table class="custom-table"><thead><tr><th>Code</th><th>Name</th><th>Status</th>${ME.is_master?'<th>Actions</th>':''}</tr></thead>
              <tbody>${data.plants.map(p => `<tr><td class="font-mono font-bold text-indigo-600">${esc(p.code)}</td><td class="font-medium">${esc(p.name)}</td><td><span class="badge ${p.active?'badge-emerald':'badge-red'}">${p.active?'ACTIVE':'OFF'}</span></td>${ME.is_master?`<td><button onclick="showEditPlantModal('${esc(p.code)}','${esc(p.name)}',${p.active})" class="text-xs text-indigo-600 font-semibold">Edit</button><button onclick="deletePlant('${esc(p.code)}')" class="text-xs text-red-600 font-semibold ml-2">Delete</button></td>`:''}</tr>`).join('')}</tbody>
            </table>
          </div>
        </div>
        <div class="card flex flex-col overflow-hidden">
          <div class="p-3 border-b border-slate-200 flex justify-between items-center bg-slate-50">
            <span class="text-xs font-bold text-slate-700 uppercase tracking-wider"><i class="ri-team-line text-emerald-500"></i> Departments (${data.departments.length})</span>
            ${ME.is_master ? `<button onclick="showCreateDeptModal()" class="btn-primary text-xs font-bold px-3 py-1.5 rounded-lg">+ Add</button>` : ''}
          </div>
          <div class="flex-1 overflow-auto">
            <table class="custom-table"><thead><tr><th>Name</th><th>Status</th>${ME.is_master?'<th>Action</th>':''}</tr></thead>
              <tbody>${data.departments.map(d => `<tr><td class="font-medium">${esc(d.name)}</td><td><span class="badge ${d.active?'badge-emerald':'badge-red'}">${d.active?'ACTIVE':'OFF'}</span></td>${ME.is_master?`<td><button onclick="deleteDept('${esc(d.name)}')" class="text-xs text-red-600 font-semibold">Delete</button></td>`:''}</tr>`).join('')}</tbody>
            </table>
          </div>
        </div>
      </div>
    </div>`;
}

function showEditOrgModal(n, sys, cl) {
  openDialog(`<div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl"><h3 class="font-bold text-sm">Edit Corporate Identity</h3><button onclick="closeDialog()"><i class="ri-close-line text-lg text-slate-500"></i></button></div>
    <div class="p-6 space-y-4"><div><label class="form-label">Company Name</label><input id="org_n" value="${esc(n)}" class="form-input"></div>
    <div class="grid grid-cols-2 gap-4"><div><label class="form-label">System ID</label><input id="org_sys" value="${esc(sys)}" class="form-input font-mono uppercase"></div><div><label class="form-label">Client</label><input id="org_cl" value="${esc(cl)}" class="form-input font-mono"></div></div></div>
    <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl"><button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-lg text-xs font-semibold">Cancel</button><button onclick="submitEditOrg()" class="btn-primary px-4 py-2 rounded-lg text-xs font-semibold">Save</button></div>`);
}

async function submitEditOrg() {
  const fd = new FormData();
  fd.append("name", $('#org_n').value.trim()); fd.append("system_id", $('#org_sys').value.trim()); fd.append("client", $('#org_cl').value.trim());
  try { await api('/api/admin/org/update', {method: 'POST', body: fd}); closeDialog(); renderSPRO(); setMsg("Identity updated"); }
  catch(e) { alert(e); }
}

function showCreatePlantModal() {
  openDialog(`<div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl"><h3 class="font-bold text-sm">Add Plant</h3><button onclick="closeDialog()"><i class="ri-close-line text-lg text-slate-500"></i></button></div>
    <div class="p-6 space-y-4"><div><label class="form-label">Code *</label><input id="np_c" class="form-input font-mono uppercase"></div><div><label class="form-label">Name *</label><input id="np_n" class="form-input"></div></div>
    <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl"><button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-lg text-xs font-semibold">Cancel</button><button onclick="submitCreatePlant()" class="btn-primary px-4 py-2 rounded-lg text-xs font-semibold">Save</button></div>`);
}

async function submitCreatePlant() {
  const c = $('#np_c').value.trim(), n = $('#np_n').value.trim();
  if (!c || !n) return alert("Code and Name required.");
  const fd = new FormData(); fd.append("code", c); fd.append("name", n);
  try { await api('/api/admin/plants/create', {method: 'POST', body: fd}); closeDialog(); renderSPRO(); setMsg(`Plant ${c} created`); }
  catch(e) { alert(e); }
}

function showEditPlantModal(code, name, active) {
  openDialog(`<div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl"><h3 class="font-bold text-sm">Edit Plant: ${esc(code)}</h3><button onclick="closeDialog()"><i class="ri-close-line text-lg text-slate-500"></i></button></div>
    <div class="p-6 space-y-4"><div><label class="form-label">Name</label><input id="ep_n" value="${esc(name)}" class="form-input"></div><div><label class="form-label">Status</label><select id="ep_st" class="form-input"><option value="1" ${active?'selected':''}>Active</option><option value="0" ${!active?'selected':''}>Off</option></select></div></div>
    <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl"><button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-lg text-xs font-semibold">Cancel</button><button onclick="submitEditPlant('${esc(code)}')" class="btn-primary px-4 py-2 rounded-lg text-xs font-semibold">Save</button></div>`);
}

async function submitEditPlant(code) {
  const fd = new FormData(); fd.append("code", code); fd.append("name", $('#ep_n').value.trim()); fd.append("active", $('#ep_st').value);
  try { await api('/api/admin/plants/update', {method: 'POST', body: fd}); closeDialog(); renderSPRO(); setMsg(`Plant ${code} updated`); }
  catch(e) { alert(e); }
}

async function deletePlant(code) {
  if (!confirm(`Remove plant ${code}?`)) return;
  const fd = new FormData(); fd.append("code", code);
  try { await api('/api/admin/plants/delete', {method: 'POST', body: fd}); renderSPRO(); setMsg(`Plant ${code} deleted`); }
  catch(e) { alert(e); }
}

function showCreateDeptModal() {
  const d = prompt("Department name:"); if (!d || !d.trim()) return;
  const fd = new FormData(); fd.append("name", d.trim());
  api('/api/admin/departments/create', {method: 'POST', body: fd}).then(()=>{renderSPRO(); setMsg("Department added");}).catch(e=>alert(e));
}

async function deleteDept(name) {
  if (!confirm(`Delete ${name}?`)) return;
  const fd = new FormData(); fd.append("name", name);
  try { await api('/api/admin/departments/delete', {method: 'POST', body: fd}); renderSPRO(); setMsg("Removed"); }
  catch(e) { alert(e); }
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
        <div><h2 class="text-lg font-bold text-slate-900">Access Control (PFCG)</h2></div>
        <div class="flex items-center gap-2">
          <select id="pfcg_u" class="form-input w-56">${m.users.map(u=>`<option value="${esc(u.username)}">${esc(u.username)} (${esc(u.full_name)})</option>`).join('')}</select>
          <select id="pfcg_t" class="form-input w-56">${m.tcodes.map(t=>`<option value="${esc(t.code)}">${esc(t.code)} - ${esc(t.title)}</option>`).join('')}</select>
          <button onclick="grantAccess()" class="btn-primary text-xs font-semibold px-3 py-2 rounded-lg">+ Grant</button>
        </div>
      </div>
      <div class="flex-1 mt-4 card overflow-auto">
        <table class="custom-table"><thead><tr><th>User</th><th>Module</th><th>Granted By</th><th>At</th><th>Action</th></tr></thead>
          <tbody>${m.grants.map(g => `<tr><td class="font-mono font-bold text-indigo-600">${esc(g.username)}</td><td class="font-mono font-bold">${esc(g.tcode_code)}</td><td>${esc(g.granted_by)}</td><td class="text-xs font-mono text-slate-500">${esc(g.granted_at)}</td><td><button onclick="revokeAccess('${esc(g.username)}','${esc(g.tcode_code)}')" class="text-xs text-red-600 font-semibold">Revoke</button></td></tr>`).join('')}</tbody>
        </table>
      </div>
    </div>`;
}
async function grantAccess() {
  const fd = new FormData(); fd.append("username", $('#pfcg_u').value); fd.append("tcode", $('#pfcg_t').value);
  await api('/api/pfcg/grant', {method: 'POST', body: fd}); renderPFCG(); setMsg("Granted");
}
async function revokeAccess(u, t) {
  if (!confirm(`Revoke ${t} from ${u}?`)) return;
  const fd = new FormData(); fd.append("username", u); fd.append("tcode", t);
  await api('/api/pfcg/revoke', {method: 'POST', body: fd}); renderPFCG(); setMsg("Revoked");
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
        <div><h2 class="text-lg font-bold text-slate-900">Module Studio (SE93)</h2><p class="text-xs text-slate-500">Upload template to auto-create a T-code</p></div>
        <div class="flex gap-2"><button onclick="openTcode('STUDIO')" class="btn-secondary text-xs px-3 py-2 rounded-lg flex items-center gap-1.5"><i class="ri-robot-2-line text-emerald-500"></i> Use AI Studio</button><button onclick="renderSE93()" class="btn-secondary text-xs px-3 py-2 rounded-lg"><i class="ri-refresh-line"></i></button></div>
      </div>
      <div class="mt-4 grid grid-cols-1 lg:grid-cols-3 gap-4 flex-1 overflow-hidden">
        <div class="lg:col-span-1 card p-4 overflow-auto">
          <h3 class="text-sm font-bold text-slate-900 mb-3 flex items-center gap-1.5"><i class="ri-upload-cloud-2-line text-indigo-500"></i> Upload Template</h3>
          <form id="se93_form" onsubmit="submitSE93(event)" class="space-y-3">
            <div><label class="form-label">T-Code</label><input name="code" required placeholder="e.g. DISP1" class="form-input font-mono uppercase"></div>
            <div><label class="form-label">Title</label><input name="title" required class="form-input"></div>
            <div><label class="form-label">Description</label><input name="description" class="form-input"></div>
            <div><label class="form-label">Module</label><input name="module" value="PROD" class="form-input font-mono uppercase"></div>
            <div><label class="form-label">Allowed Depts (CSV)</label><input name="allowed_depts" value="Production,Quality Control" class="form-input"></div>
            <div><label class="form-label">Import Sample Data?</label><select name="import_sample_data" class="form-input"><option value="1">Yes</option><option value="0">No</option></select></div>
            <div><label class="form-label">File</label><input name="file" type="file" required class="w-full text-xs text-slate-600 file:mr-2 file:bg-indigo-600 file:text-white file:border-0 file:rounded file:px-2 file:py-1 file:text-xs"></div>
            <button type="submit" class="btn-primary w-full text-xs font-bold py-2 rounded-lg">Create Module</button>
          </form>
        </div>
        <div class="lg:col-span-2 card overflow-hidden flex flex-col">
          <div class="p-3 border-b border-slate-200 bg-slate-50"><span class="text-xs font-bold text-slate-700 uppercase tracking-wider">Modules (${list.length})</span></div>
          <div class="flex-1 overflow-auto">
            <table class="custom-table"><thead><tr><th>Code</th><th>Title</th><th>Module</th><th>Layouts</th><th>Records</th><th>Action</th></tr></thead>
              <tbody>${list.map(t => `<tr><td class="font-mono font-bold text-indigo-600">${esc(t.code)}</td><td class="font-medium">${esc(t.title)}</td><td class="text-xs text-slate-500 uppercase">${esc(t.module||'-')}</td><td class="font-mono">${t.layout_count}</td><td class="font-mono">${t.record_count}</td><td><button onclick="openTcode('${esc(t.code)}')" class="text-xs text-indigo-600 font-semibold">Open</button><button onclick="loadStudioExisting('${esc(t.code)}')" class="text-xs text-emerald-600 font-semibold ml-2">Edit</button>${ME.is_master?`<button onclick="deleteSE93('${esc(t.code)}')" class="text-xs text-red-600 font-semibold ml-2">Delete</button>`:''}</td></tr>`).join('')}</tbody>
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
    alert(`Module ${data.code} created (${data.columns_count} cols, ${data.imported_count} rows).`);
    renderSE93(); loadModules();
  } catch(e) { alert("Error: " + e); }
}
async function deleteSE93(code) {
  if (!confirm(`Delete module ${code}?`)) return;
  const fd = new FormData(); fd.append("code", code);
  try { await api('/api/se93/delete', {method: 'POST', body: fd}); renderSE93(); loadModules(); setMsg(`Deleted ${code}`); }
  catch(e) { alert(e); }
}

// =========================================================================
// --- AI CODE STUDIO ---
// =========================================================================
const STUDIO_EXAMPLE = {
  action: "upsert", code: "DISP0", title: "Dispatch Challan",
  description: "Outward dispatch log", module: "SD", allowed_depts: "Commercial,Store",
  fields: [
    { key: "challan_no", label: "Challan No.", type: "text", required: true },
    { key: "customer", label: "Customer", type: "text" },
    { key: "dispatch_date", label: "Date", type: "date" },
    { key: "qty", label: "Quantity", type: "number" },
    { key: "remarks", label: "Remarks", type: "textarea", width: 2 }
  ], seed_rows: []
};

async function renderStudio() {
  if (!ME.is_master) return setMsg("Master Admin only", true);
  const hist = await api('/api/studio/history');
  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col p-6 overflow-hidden select-text">
      <div class="flex items-center justify-between pb-4 border-b border-slate-200 shrink-0">
        <div>
          <h2 class="text-lg font-bold text-slate-900 flex items-center gap-2"><i class="ri-robot-2-line text-emerald-500"></i> AI Code Studio</h2>
          <p class="text-xs text-slate-500">Paste JSON to create/update/delete modules live</p>
        </div>
        <div class="flex items-center gap-2">
          <button onclick="loadStudioExample()" class="btn-secondary text-xs px-3 py-1.5 rounded-lg"><i class="ri-magic-line"></i> Example</button>
          <button onclick="renderStudio()" class="btn-secondary text-xs px-3 py-1.5 rounded-lg"><i class="ri-refresh-line"></i></button>
        </div>
      </div>
      <div class="grid grid-cols-1 lg:grid-cols-2 gap-4 flex-1 mt-4 overflow-hidden">
        <div class="flex flex-col card overflow-hidden">
          <div class="p-3 border-b border-slate-200 bg-slate-50 flex justify-between items-center">
            <span class="text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5"><i class="ri-code-s-slash-line text-indigo-500"></i> Module Spec (JSON)</span>
            <div class="flex gap-1.5">
              <button onclick="studioValidate()" class="text-[11px] font-bold px-2.5 py-1 rounded-md bg-amber-50 border border-amber-200 text-amber-700">Validate</button>
              <button onclick="studioCopyCurrent()" class="text-[11px] font-semibold px-2.5 py-1 rounded-md bg-white border border-slate-200 text-slate-600">Copy</button>
              <button onclick="studioFormat()" class="text-[11px] font-semibold px-2.5 py-1 rounded-md bg-white border border-slate-200 text-slate-600">Prettify</button>
            </div>
          </div>
          <textarea id="studio_json" spellcheck="false" class="flex-1 bg-slate-50 text-emerald-700 font-mono text-[11.5px] leading-relaxed p-3 outline-none resize-none" placeholder='Paste module JSON here...'></textarea>
          <div class="p-3 border-t border-slate-200 bg-slate-50 flex items-center justify-between">
            <span id="studio_status" class="text-[11px] font-mono text-slate-500">Idle</span>
            <button onclick="studioApply()" class="btn-primary text-xs font-bold px-4 py-1.5 rounded-md flex items-center gap-1"><i class="ri-play-fill"></i> Apply Spec</button>
          </div>
        </div>
        <div class="flex flex-col gap-4 overflow-hidden">
          <div id="studio_preview" class="card p-4 overflow-auto flex-1 min-h-[200px]"><div class="text-xs text-slate-400 text-center pt-10">Validate a spec to see live preview.</div></div>
          <div class="card overflow-hidden" style="max-height:44%;">
            <div class="p-2.5 border-b border-slate-200 bg-slate-50 text-xs font-bold text-slate-700 uppercase tracking-wider flex items-center gap-1.5"><i class="ri-history-line text-indigo-500"></i> Version History</div>
            <div class="overflow-auto" style="max-height:200px;">
              <table class="custom-table"><thead><tr><th>v</th><th>Module</th><th>Action</th><th>By</th><th>When</th><th></th></tr></thead>
                <tbody>${!hist.length ? `<tr><td colspan="6" class="text-center py-6 text-slate-400 text-xs">No installs yet.</td></tr>` :
                  hist.map(h => `<tr><td class="font-mono text-indigo-600 font-bold">v${h.version}</td><td class="font-mono font-bold">${esc(h.tcode_code)}</td><td><span class="badge ${h.action==='delete'?'badge-red':'badge-emerald'}">${h.action}</span></td><td class="text-xs">${esc(h.applied_by)}</td><td class="text-xs font-mono text-slate-500">${esc(h.applied_at)}</td><td><button onclick="studioRollback(${h.id})" class="text-[11px] text-amber-600 font-semibold">Rollback</button><button onclick="studioLoadExisting('${esc(h.tcode_code)}')" class="text-[11px] text-indigo-600 font-semibold ml-2">Load</button></td></tr>`).join('')}</tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    </div>`;
}
function loadStudioExample() { $('#studio_json').value = JSON.stringify(STUDIO_EXAMPLE, null, 2); setStudioStatus("Example loaded.", false); }
function setStudioStatus(m, e) { const el = $('#studio_status'); if (!el) return; el.textContent = m; el.className = "text-[11px] font-mono " + (e ? "text-red-600" : "text-emerald-600"); }
function studioFormat() { try { const v = JSON.parse($('#studio_json').value); $('#studio_json').value = JSON.stringify(v, null, 2); setStudioStatus("Formatted.", false); } catch(e) { setStudioStatus("Invalid JSON: " + e.message, true); } }
function studioCopyCurrent() { navigator.clipboard.writeText($('#studio_json').value); setStudioStatus("Copied.", false); }
async function studioValidate() {
  let spec; try { spec = JSON.parse($('#studio_json').value); } catch(e) { setStudioStatus("Invalid JSON: " + e.message, true); return null; }
  try {
    const res = await api('/api/studio/validate', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({spec})});
    if (!res.ok) { setStudioStatus("Errors: " + res.errors.join(" | "), true); renderStudioPreview(res.clean, res.errors); return res; }
    setStudioStatus(res.will_overwrite ? "Valid — will UPDATE." : "Valid — will CREATE.", false);
    renderStudioPreview(res.clean, []); return res;
  } catch(e) { setStudioStatus("Error: " + e, true); return null; }
}
function renderStudioPreview(clean, errors) {
  const box = $('#studio_preview'); if (!box) return;
  if (!clean || errors.length) { box.innerHTML = `<div class="text-xs text-red-600 whitespace-pre-wrap">${esc(errors.join("\\n"))}</div>`; return; }
  if (clean.action === "delete") { box.innerHTML = `<div class="text-sm"><b class="text-red-600">DELETE:</b> <span class="font-mono">${esc(clean.code)}</span></div>`; return; }
  box.innerHTML = `<div class="mb-3"><div class="text-[10px] uppercase font-bold text-slate-400">Preview</div><div class="text-base font-bold text-slate-900">${esc(clean.title)} <span class="font-mono text-indigo-600 text-xs">[${esc(clean.code)}]</span></div><div class="text-xs text-slate-500 mt-0.5">${esc(clean.description || '—')}</div><div class="text-[11px] text-slate-500 mt-1">Module: <b>${esc(clean.module)}</b> · Depts: <span>${esc(clean.allowed_depts || 'none')}</span></div></div>
    <div class="text-[10px] uppercase font-bold text-slate-400 mb-1">Columns (${clean.fields.length})</div>
    <table class="w-full text-left text-[11px]"><thead><tr class="bg-slate-50 text-slate-500"><th class="p-1.5">Key</th><th class="p-1.5">Label</th><th class="p-1.5">Type</th></tr></thead>
    <tbody>${clean.fields.map(f => `<tr class="border-b border-slate-100"><td class="p-1.5 font-mono text-indigo-600">${esc(f.key)}</td><td class="p-1.5 text-slate-900">${esc(f.label)}</td><td class="p-1.5 text-slate-500">${esc(f.type)}</td></tr>`).join('')}</tbody></table>`;
}
async function studioApply() {
  let spec; try { spec = JSON.parse($('#studio_json').value); } catch(e) { setStudioStatus("Invalid JSON: " + e.message, true); return; }
  if (!confirm(`Apply ${spec.action} on ${spec.code}?`)) return;
  try { const res = await api('/api/studio/apply', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({spec})}); setStudioStatus(`OK — ${res.action} ${res.code} (v${res.version})`, false); await loadModules(); renderStudio(); }
  catch(e) { setStudioStatus("Error: " + e, true); alert(e); }
}
async function studioRollback(id) {
  if (!confirm("Rollback this version?")) return;
  try { const r = await api(`/api/studio/rollback/${id}`, {method: 'POST'}); setStudioStatus(`Rolled back (v${r.version})`, false); await loadModules(); renderStudio(); }
  catch(e) { alert(e); }
}
async function studioLoadExisting(code) {
  try { const spec = await api(`/api/studio/spec/${code}`); $('#studio_json').value = JSON.stringify(spec, null, 2); setStudioStatus(`Loaded ${code}.`, false); await studioValidate(); }
  catch(e) { alert(e); }
}

// =========================================================================
// --- TN01 ---
// =========================================================================
let TN01_ROUTES = {}, TN01_ACTIVE_ROUTE = "PROD_TO_QC", TN01_HEADERS = [], TN01_ROWS = [], TN01_UNDO_STACK = [], TN01_REDO_STACK = [], TN01_ACTIVE_CELL = { r: 0, c: 1 };
async function renderTN01() { TN01_ROUTES = await api('/api/tn01/routes'); renderTN01Main(); }
function renderTN01Main(activeTab = "sheet") {
  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col overflow-hidden select-text">
      <div class="p-3 bg-white border-b border-slate-200 flex items-center justify-between shrink-0">
        <div class="flex items-center gap-3">
          <div class="flex items-center gap-2">
            <span class="font-mono font-bold text-xs bg-indigo-600 text-white px-2.5 py-1 rounded-lg shadow-sm">TN01</span>
            <div><h2 class="text-sm font-bold text-slate-900">Transfer Note & Handover</h2><p class="text-[10px] text-slate-500">Inter-department transfer slips</p></div>
          </div>
          <div class="h-6 w-[1px] bg-slate-200 mx-2"></div>
          <div class="flex items-center bg-slate-100 border border-slate-200 rounded-lg p-1 gap-1 text-xs font-semibold">
            <button onclick="switchTN01Tab('sheet')" class="px-3 py-1.5 rounded-md transition ${activeTab==='sheet'?'bg-white text-indigo-600 shadow-sm':'text-slate-500 hover:text-slate-700'}"><i class="ri-table-line"></i> Excel Sheet</button>
            <button onclick="switchTN01Tab('receive')" class="px-3 py-1.5 rounded-md transition ${activeTab==='receive'?'bg-white text-indigo-600 shadow-sm':'text-slate-500 hover:text-slate-700'}"><i class="ri-inbox-archive-line"></i> Receive</button>
            <button onclick="switchTN01Tab('history')" class="px-3 py-1.5 rounded-md transition ${activeTab==='history'?'bg-white text-indigo-600 shadow-sm':'text-slate-500 hover:text-slate-700'}"><i class="ri-history-line"></i> History</button>
            ${ME.is_master || ME.role==='admin' ? `<button onclick="switchTN01Tab('format')" class="px-3 py-1.5 rounded-md transition ${activeTab==='format'?'bg-white text-indigo-600 shadow-sm':'text-slate-500 hover:text-slate-700'}"><i class="ri-paint-brush-line"></i> Formats</button>` : ''}
          </div>
        </div>
        <div>${activeTab === 'sheet' ? `<button onclick="openSendModal()" class="btn-primary text-xs font-bold px-4 py-2 rounded-lg flex items-center gap-1.5 shadow-sm"><i class="ri-send-plane-fill"></i> Send</button>` : ''}</div>
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
        <label class="font-bold text-slate-600 text-xs flex items-center gap-1.5">Route:
          <select id="route_selector" onchange="applyRouteLayout(this.value, true)" class="form-input w-auto py-1 font-semibold">${Object.keys(TN01_ROUTES).map(k => `<option value="${k}" ${k===TN01_ACTIVE_ROUTE?'selected':''}>${TN01_ROUTES[k].route}</option>`).join('')}</select>
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
        <span class="text-[11px] bg-slate-50 border border-slate-200 px-2 py-0.5 rounded text-slate-500 font-mono">Ctrl+V to paste</span>
      </div>
    </div>
    <div id="grid_table_container" class="flex-1 overflow-auto p-2 select-text" tabindex="0"></div>
    <div class="p-2 bg-white border-t border-slate-200 text-slate-500 text-[11px] flex justify-between items-center font-mono shrink-0">
      <span id="sheet_count_lbl">0 Batches</span><span>Click cell to edit</span>
    </div>`;
  renderGridTable(); setupPasteHandler();
}
function applyRouteLayout(routeKey, resetData = true) {
  TN01_ACTIVE_ROUTE = routeKey;
  TN01_HEADERS = [...TN01_ROUTES[routeKey].columns];
  if (resetData || !TN01_ROWS.length) {
    TN01_ROWS = [];
    for (let i = 1; i <= 8; i++) { const row = new Array(TN01_HEADERS.length).fill(""); row[0] = i; TN01_ROWS.push(row); }
  }
  renderGridTable();
}
function renderGridTable() {
  const c = $('#grid_table_container'); if (!c) return;
  let html = `<table class="excel-grid"><thead><tr>${TN01_HEADERS.map(h=>`<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>`;
  TN01_ROWS.forEach((row, rIdx) => { html += `<tr>${row.map((cell, cIdx) => `<td contenteditable="${cIdx!==0}" class="${cIdx===0?'sr-col':''}" data-r="${rIdx}" data-c="${cIdx}" onfocus="TN01_ACTIVE_CELL={r:${rIdx},c:${cIdx}}" onblur="updateGridCell(${rIdx},${cIdx},this.innerText)">${esc(cell)}</td>`).join('')}</tr>`; });
  html += `</tbody></table>`;
  c.innerHTML = html;
  const filled = TN01_ROWS.filter(r => r.slice(1).some(x => String(x).trim() !== '')).length;
  if ($('#sheet_count_lbl')) $('#sheet_count_lbl').textContent = `${filled} Filled (${TN01_ROWS.length} Rows)`;
}
function updateGridCell(r, col, val) { if (col === 0) return; if (TN01_ROWS[r][col] !== val.trim()) { TN01_UNDO_STACK.push(JSON.stringify(TN01_ROWS)); TN01_ROWS[r][col] = val.trim(); } }
function undoGrid() { if (!TN01_UNDO_STACK.length) return; TN01_REDO_STACK.push(JSON.stringify(TN01_ROWS)); TN01_ROWS = JSON.parse(TN01_UNDO_STACK.pop()); renderGridTable(); }
function redoGrid() { if (!TN01_REDO_STACK.length) return; TN01_UNDO_STACK.push(JSON.stringify(TN01_ROWS)); TN01_ROWS = JSON.parse(TN01_REDO_STACK.pop()); renderGridTable(); }
function addGridRow() { TN01_UNDO_STACK.push(JSON.stringify(TN01_ROWS)); const r = new Array(TN01_HEADERS.length).fill(""); r[0] = TN01_ROWS.length + 1; TN01_ROWS.push(r); renderGridTable(); }
function deleteGridRow() { if (TN01_ROWS.length <= 1) return; TN01_UNDO_STACK.push(JSON.stringify(TN01_ROWS)); TN01_ROWS.pop(); renderGridTable(); }
function clearGrid() { if (!confirm("Clear all rows?")) return; TN01_UNDO_STACK.push(JSON.stringify(TN01_ROWS)); TN01_ROWS = []; for (let i = 1; i <= 6; i++) { const r = new Array(TN01_HEADERS.length).fill(""); r[0] = i; TN01_ROWS.push(r); } renderGridTable(); }
function setupPasteHandler() {
  const c = $('#grid_table_container'); if (!c) return;
  c.addEventListener('paste', function(e) {
    e.preventDefault();
    const clip = (e.clipboardData || window.clipboardData).getData('text'); if (!clip) return;
    TN01_UNDO_STACK.push(JSON.stringify(TN01_ROWS));
    const lines = clip.replace(/\\r\\n/g, '\\n').replace(/\\r/g, '\\n').split('\\n');
    let startR = TN01_ACTIVE_CELL.r || 0, startC = TN01_ACTIVE_CELL.c || 1;
    if (startC === 0) startC = 1;
    lines.forEach((line, lIdx) => {
      if (!line && lIdx === lines.length - 1) return;
      const cells = line.split('\\t');
      const tr = startR + lIdx;
      while (TN01_ROWS.length <= tr) { const nr = new Array(TN01_HEADERS.length).fill(""); nr[0] = TN01_ROWS.length + 1; TN01_ROWS.push(nr); }
      cells.forEach((v, cIdx) => { const tc = startC + cIdx; if (tc < TN01_HEADERS.length) TN01_ROWS[tr][tc] = v.trim(); });
    });
    TN01_ROWS.forEach((r, i) => { r[0] = i + 1; });
    renderGridTable(); setMsg("Pasted.");
  });
}
async function exportGridAsExcel() {
  const filled = TN01_ROWS.filter(r => r.slice(1).some(x => String(x).trim() !== ''));
  if (!filled.length) return alert("No rows.");
  const title = TN01_ROUTES[TN01_ACTIVE_ROUTE] ? TN01_ROUTES[TN01_ACTIVE_ROUTE].route : "Draft";
  try {
    const token = localStorage.getItem("nexus_token");
    const res = await fetch('/api/tn01/export-grid', {method: 'POST', headers: {'Content-Type': 'application/json', 'Authorization': token ? `Bearer ${token}` : ''}, body: JSON.stringify({headers: TN01_HEADERS, rows: filled, title: `TN01 Draft — ${title}`})});
    if (!res.ok) { let t = await res.text(); try { t = JSON.parse(t).detail || t; } catch {} throw t; }
    const blob = await res.blob(); const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url;
    const cd = res.headers.get('Content-Disposition') || '';
    const m = cd.match(/filename=([^;]+)/);
    a.download = m ? m[1].replace(/["']/g, '') : 'TN01_Draft.xlsx';
    document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
    setMsg("Exported as Excel");
  } catch(e) { alert("Export failed: " + e); }
}
async function openSendModal() {
  const active = TN01_ROUTES[TN01_ACTIVE_ROUTE];
  const filled = TN01_ROWS.filter(r => r.slice(1).some(x => String(x).trim() !== ''));
  if (!filled.length) return alert("Enter at least one row.");
  const spro = await api('/api/spro/overview');
  const myPlant = (ME.plants && ME.plants[0]) ? ME.plants[0] : "1000";
  openDialog(`
    <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl">
      <div><h3 class="font-bold text-sm">Send Transfer Note</h3><p class="text-xs text-slate-500">Route: ${active.route}</p></div>
      <button onclick="closeDialog()"><i class="ri-close-line text-lg text-slate-500"></i></button>
    </div>
    <div class="p-6 space-y-4">
      <div class="grid grid-cols-3 gap-3">
        <div><label class="form-label">Plant</label><select id="send_plant" onchange="updateSuggestedUsers()" class="form-input">${spro.plants.map(p=>`<option value="${esc(p.code)}" ${p.code===myPlant?'selected':''}>${esc(p.code)} - ${esc(p.name)}</option>`).join('')}</select></div>
        <div><label class="form-label">Dept</label><select id="send_dept" onchange="updateSuggestedUsers()" class="form-input">${spro.departments.map(d=>`<option value="${esc(d.name)}" ${d.name===active.to_dept?'selected':''}>${esc(d.name)}</option>`).join('')}</select></div>
        <div><label class="form-label">Receiver</label><select id="send_user" class="form-input"></select></div>
      </div>
      <div><label class="form-label">Remarks</label><textarea id="send_remarks" class="form-input" rows="2"></textarea></div>
    </div>
    <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl">
      <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-lg text-xs font-semibold">Cancel</button>
      <button onclick="submitSendNote()" class="btn-primary px-5 py-2 rounded-lg text-xs font-bold">Send</button>
    </div>`);
  updateSuggestedUsers();
}
async function updateSuggestedUsers() {
  const pl = $('#send_plant').value, dept = $('#send_dept').value;
  const users = await api(`/api/tn01/suggest-users?plant=${pl}&dept=${encodeURIComponent(dept)}`);
  $('#send_user').innerHTML = users.length ? users.map(u => `<option value="${esc(u.username)}">${esc(u.full_name)} (${esc(u.username)})</option>`).join('') : `<option value="">(Dept Pool)</option>`;
}
async function submitSendNote() {
  const filled = TN01_ROWS.filter(r => r.slice(1).some(c => String(c).trim() !== ''));
  const payload = {layout_key: TN01_ACTIVE_ROUTE, route: TN01_ROUTES[TN01_ACTIVE_ROUTE].route, to_plant: $('#send_plant').value, to_dept: $('#send_dept').value, to_user: $('#send_user').value, remarks: $('#send_remarks').value, headers: TN01_HEADERS, rows: filled};
  try { const r = await api('/api/tn01/send', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(payload)}); closeDialog(); alert(`Sent. Note #: ${r.note_no}`); switchTN01Tab('history'); }
  catch(e) { alert(e); }
}
async function renderTN01ReceiveTab() {
  const notes = await api('/api/tn01/inbox');
  $('#tn01_content').innerHTML = `
    <div class="p-6 flex-1 flex flex-col overflow-hidden">
      <div class="flex items-center justify-between pb-4 border-b border-slate-200 shrink-0">
        <h2 class="text-base font-bold text-slate-900">Incoming Notes</h2>
        <button onclick="renderTN01ReceiveTab()" class="btn-secondary text-xs px-3 py-1.5 rounded-lg"><i class="ri-refresh-line"></i> Refresh</button>
      </div>
      <div class="flex-1 mt-4 card overflow-auto">
        <table class="custom-table"><thead><tr><th>Note #</th><th>Route</th><th>Origin</th><th>Sender</th><th>Batches</th><th>Time</th><th>Status</th><th>Actions</th></tr></thead>
          <tbody>${!notes.length ? `<tr><td colspan="8" class="text-center py-12 text-slate-400">Inbox empty.</td></tr>` : ''}${notes.map(n => `<tr><td class="font-mono font-bold text-indigo-600">${esc(n.note_no)}</td><td class="font-semibold">${esc(n.route)}</td><td>${esc(n.from_plant)} - ${esc(n.from_dept)}</td><td>${esc(n.from_user_name)}</td><td class="font-mono font-bold">${n.entry_count}</td><td class="font-mono text-xs text-slate-500">${esc(n.created_at)}</td><td><span class="badge ${n.status==='received'?'badge-emerald':'badge-amber'}">${n.status.toUpperCase()}</span></td><td><button onclick="viewTransferNote(${n.id})" class="btn-primary text-xs font-semibold px-2.5 py-1 rounded-md">Open</button>${n.status==='sent'?`<button onclick="acknowledgeNote(${n.id})" class="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold px-2.5 py-1 rounded-md ml-1">Ack</button>`:''}</td></tr>`).join('')}</tbody>
        </table>
      </div>
    </div>`;
}
async function acknowledgeNote(nid) { await api(`/api/tn01/receive/${nid}`, {method: 'POST'}); setMsg("Acknowledged"); renderTN01ReceiveTab(); }
async function viewTransferNote(nid) {
  const note = await api(`/api/tn01/note/${nid}`);
  openDialog(`
    <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl">
      <div><span class="text-xs bg-indigo-50 text-indigo-700 border border-indigo-200 font-mono font-bold px-2 py-0.5 rounded">${esc(note.note_no)}</span><h3 class="font-bold text-sm mt-1">${esc(note.route)}</h3></div>
      <button onclick="closeDialog()"><i class="ri-close-line text-lg text-slate-500"></i></button>
    </div>
    <div class="p-6 space-y-4">
      <div class="flex justify-between items-center">
        <span class="text-xs font-bold text-slate-700">Batches (${note.entry_count}):</span>
        <div class="flex gap-2">
          <a href="/api/tn01/export/${nid}" class="btn-secondary text-xs px-3 py-1.5 rounded-md flex items-center gap-1 text-emerald-700 font-semibold border-emerald-200"><i class="ri-file-excel-2-line"></i> Export as Excel</a>
          <button onclick="loadNoteIntoWorkspace(${nid})" class="btn-primary text-xs px-3 py-1.5 rounded-md flex items-center gap-1 font-semibold"><i class="ri-download-cloud-2-line"></i> Load</button>
        </div>
      </div>
      <div class="max-h-72 overflow-auto border border-slate-200 rounded-xl">
        <table class="w-full text-left text-xs"><thead><tr class="bg-slate-50 font-bold">${note.sheet_headers.map(h => `<th class="p-2 border-r border-slate-200">${esc(h)}</th>`).join('')}</tr></thead>
        <tbody>${note.sheet_data.map(r => `<tr class="border-b border-slate-100">${r.map(c => `<td class="p-2 border-r border-slate-100">${esc(c)}</td>`).join('')}</tr>`).join('')}</tbody></table>
      </div>
    </div>
    <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-between items-center rounded-b-2xl">
      <a href="/api/tn01/print/${note.id}" target="_blank" class="btn-secondary text-xs px-3 py-2 rounded-md flex items-center gap-1 font-semibold"><i class="ri-printer-line"></i> Print</a>
      <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-md text-xs font-semibold">Close</button>
    </div>`);
}
async function loadNoteIntoWorkspace(nid) { const n = await api(`/api/tn01/note/${nid}`); TN01_HEADERS = [...n.sheet_headers]; TN01_ROWS = JSON.parse(JSON.stringify(n.sheet_data)); closeDialog(); switchTN01Tab('sheet'); setMsg("Loaded"); }
async function renderTN01HistoryTab() {
  const isAdmin = (ME.is_master || ME.role === 'admin');
  const hist = await api('/api/tn01/history');
  const spro = isAdmin ? await api('/api/spro/overview') : null;
  $('#tn01_content').innerHTML = `
    <div class="p-6 flex-1 flex flex-col overflow-hidden">
      <div class="flex items-center justify-between pb-3 border-b border-slate-200 shrink-0">
        <div><h2 class="text-base font-bold text-slate-900">${isAdmin ? "All Transfer Notes" : "My Transfer Notes"}</h2><p class="text-xs text-slate-500">${isAdmin ? "Full history — filters below" : "Notes for your department"}</p></div>
        <button onclick="renderTN01HistoryTab()" class="btn-secondary text-xs px-3 py-1.5 rounded-lg"><i class="ri-refresh-line"></i> Refresh</button>
      </div>
      ${isAdmin ? `<div class="mt-3 mb-3 card p-3 grid grid-cols-1 md:grid-cols-6 gap-2 shrink-0">
        <div><label class="form-label">From Dept</label><select id="f_from" class="form-input"><option value="">— Any —</option>${spro.departments.map(d => `<option value="${esc(d.name)}">${esc(d.name)}</option>`).join('')}</select></div>
        <div><label class="form-label">To Dept</label><select id="f_to" class="form-input"><option value="">— Any —</option>${spro.departments.map(d => `<option value="${esc(d.name)}">${esc(d.name)}</option>`).join('')}</select></div>
        <div><label class="form-label">Status</label><select id="f_status" class="form-input"><option value="">— Any —</option><option value="sent">Sent</option><option value="received">Received</option></select></div>
        <div><label class="form-label">Date From</label><input type="date" id="f_from_date" class="form-input"></div>
        <div><label class="form-label">Date To</label><input type="date" id="f_to_date" class="form-input"></div>
        <div class="flex items-end gap-1"><button onclick="applyHistoryFilters()" class="btn-primary text-xs font-bold px-3 py-2 rounded-md flex-1"><i class="ri-filter-line"></i> Filter</button><button onclick="renderTN01HistoryTab()" class="btn-secondary text-xs font-semibold px-2 py-2 rounded-md">Reset</button></div>
      </div>` : ''}
      <div class="flex-1 mt-2 card overflow-auto">
        <table class="custom-table"><thead><tr><th>Note #</th><th>Route</th><th>From</th><th>To</th><th>Batches</th><th>Sender</th><th>Time</th><th>Status</th><th>Actions</th></tr></thead>
          <tbody id="hist_tbody">${renderHistoryRows(hist)}</tbody></table>
      </div>
    </div>`;
}
function renderHistoryRows(hist) {
  if (!hist.length) return `<tr><td colspan="9" class="text-center py-10 text-slate-400">No records.</td></tr>`;
  return hist.map(h => `<tr>
    <td class="font-mono font-bold text-indigo-600">${esc(h.note_no)}</td>
    <td class="font-semibold">${esc(h.route)}</td>
    <td>${esc(h.from_plant)} - ${esc(h.from_dept)}</td>
    <td>${esc(h.to_plant)} - ${esc(h.to_dept)}</td>
    <td class="font-mono font-bold">${h.entry_count}</td>
    <td>${esc(h.from_user_name)}</td>
    <td class="font-mono text-xs text-slate-500">${esc(h.created_at)}</td>
    <td><span class="badge ${h.status==='received'?'badge-emerald':'badge-amber'}">${h.status.toUpperCase()}</span></td>
    <td><button onclick="viewTransferNote(${h.id})" class="text-xs text-indigo-600 font-semibold">View</button><a href="/api/tn01/print/${h.id}" target="_blank" class="text-xs text-slate-600 font-semibold ml-2">PDF</a><a href="/api/tn01/export/${h.id}" class="text-xs text-emerald-600 font-semibold ml-2">Excel</a></td>
  </tr>`).join('');
}
async function applyHistoryFilters() {
  const qs = new URLSearchParams();
  const fd = $('#f_from'); if (fd && fd.value) qs.set('from_dept', fd.value);
  const td = $('#f_to'); if (td && td.value) qs.set('to_dept', td.value);
  const st = $('#f_status'); if (st && st.value) qs.set('status', st.value);
  const df = $('#f_from_date'); if (df && df.value) qs.set('date_from', df.value);
  const dt = $('#f_to_date'); if (dt && dt.value) qs.set('date_to', dt.value);
  const hist = await api('/api/tn01/history?' + qs.toString());
  const tbody = $('#hist_tbody'); if (tbody) tbody.innerHTML = renderHistoryRows(hist);
  setMsg(`Filtered: ${hist.length}`);
}
async function renderTN01FormatTab() {
  const layouts = await api('/api/tn01/layouts');
  $('#tn01_content').innerHTML = `
    <div class="p-6 flex-1 flex flex-col overflow-hidden">
      <div class="flex items-center justify-between pb-4 border-b border-slate-200 shrink-0"><h2 class="text-base font-bold text-slate-900">Formats (Admin)</h2><button onclick="openCreateLayoutModal()" class="btn-primary text-xs font-semibold px-4 py-2 rounded-lg">+ Create</button></div>
      <div class="grid grid-cols-1 md:grid-cols-3 gap-4 mt-4 overflow-auto">
        ${layouts.map(l => `<div class="card p-4 flex flex-col justify-between"><div><div class="flex justify-between"><span class="font-mono text-xs text-indigo-600 font-bold">${esc(l.layout_code)}</span><button onclick="deleteLayoutTN01(${l.id})" class="text-red-600 text-xs font-semibold">Delete</button></div><h4 class="font-bold text-slate-900 text-sm mt-2">${esc(l.name)}</h4><p class="text-xs text-slate-500 mt-1">${(l.fields || []).length} Columns</p></div><button onclick="loadFormatIntoWorkspace(${l.id})" class="btn-primary px-3 py-1.5 rounded-md text-xs mt-3 font-semibold">Use</button></div>`).join('')}
      </div>
    </div>`;
}
function openCreateLayoutModal() {
  openDialog(`<div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl"><h3 class="font-bold text-sm">Custom Format</h3><button onclick="closeDialog()"><i class="ri-close-line text-lg text-slate-500"></i></button></div>
    <div class="p-6 space-y-4"><div><label class="form-label">Name</label><input id="nl_n" class="form-input"></div><div><label class="form-label">Columns (comma-separated)</label><textarea id="nl_c" class="form-input font-mono" rows="3">Sr. No., Production Order, Heat Lot, Size, Length, Weight, Remarks</textarea></div></div>
    <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl"><button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-lg text-xs font-semibold">Cancel</button><button onclick="submitSaveFormat()" class="btn-primary px-4 py-2 rounded-lg text-xs font-semibold">Save</button></div>`);
}
async function submitSaveFormat() {
  const name = $('#nl_n').value.trim();
  const cols = $('#nl_c').value.split(',').map(c => c.trim()).filter(Boolean);
  if (!cols.includes("Sr. No.")) cols.unshift("Sr. No.");
  await api('/api/tn01/layouts/save', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({name, columns: cols})});
  closeDialog(); renderTN01FormatTab();
}
async function loadFormatIntoWorkspace(lid) {
  const ls = await api('/api/tn01/layouts'); const l = ls.find(x => x.id === lid); if (!l) return;
  TN01_HEADERS = [...l.fields]; TN01_ROWS = [];
  for (let i = 1; i <= 6; i++) { const r = new Array(TN01_HEADERS.length).fill(""); r[0] = i; TN01_ROWS.push(r); }
  switchTN01Tab('sheet');
}
async function deleteLayoutTN01(lid) { if (!confirm("Delete format?")) return; await api(`/api/tn01/layouts/${lid}`, {method: 'DELETE'}); renderTN01FormatTab(); }

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
        <div class="flex items-center gap-2"><a href="/api/tcode/${code}/export" class="btn-secondary text-xs font-semibold px-2.5 py-1.5 rounded-lg flex items-center gap-1"><i class="ri-download-2-line"></i> Export</a><button onclick="renderStandardModule('${code}')" class="btn-secondary text-xs p-2 rounded-lg"><i class="ri-refresh-line"></i></button></div>
      </div>
      <div class="flex-1 overflow-auto">
        <table class="custom-table"><thead><tr><th style="width:60px;">#</th><th style="width:90px;">Plant</th>${fields.map(f=>`<th>${esc(f.label)}</th>`).join('')}<th style="width:130px;">Logged By</th><th style="width:140px;">Time</th><th style="width:80px;">Action</th></tr></thead>
          <tbody>${!recs.length ? `<tr><td colspan="${fields.length + 5}" class="text-center py-10 text-slate-400">No records.</td></tr>` : ''}${recs.map(r => `<tr><td class="font-mono text-slate-400 font-semibold">${r.id}</td><td><span class="font-mono text-xs bg-slate-100 px-2 py-0.5 rounded">${esc(r.plant || '1000')}</span></td>${fields.map(f=>`<td>${esc(r.data[f.key]||'')}</td>`).join('')}<td class="text-xs text-slate-500">${esc(r.created_by_name || r.created_by)}</td><td class="text-xs text-slate-500 font-mono">${esc(r.created_at)}</td><td><button onclick="deleteStandardRecord('${code}', ${r.id})" class="text-xs text-red-600 font-semibold">Del</button></td></tr>`).join('')}</tbody>
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
      <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl"><h3 class="font-bold text-sm">New Entry: ${code}</h3><button onclick="closeDialog()"><i class="ri-close-line text-lg text-slate-500"></i></button></div>
      <div class="p-6 space-y-4">
        <div><label class="form-label">Plant</label><input id="std_f__plant" class="form-input font-mono" value="${esc(myPlant)}"></div>
        <div class="grid grid-cols-2 gap-4">
          ${fields.map(f => `<div class="col-span-${f.width||1}"><label class="form-label">${esc(f.label)}</label>${f.type==='textarea' ? `<textarea id="std_f_${f.key}" class="form-input" rows="2"></textarea>` : (f.type==='select' && f.options ? `<select id="std_f_${f.key}" class="form-input">${f.options.map(o=>`<option value="${esc(o)}">${esc(o)}</option>`).join('')}</select>` : `<input type="${f.type==='number'?'number':(f.type==='date'?'date':'text')}" id="std_f_${f.key}" class="form-input">`)}</div>`).join('')}
        </div>
      </div>
      <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl"><button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-lg text-xs font-semibold">Cancel</button><button onclick="submitStandardRecord('${code}')" class="btn-primary px-4 py-2 rounded-lg text-xs font-semibold">Save</button></div>`);
  });
}
async function submitStandardRecord(code) {
  const layouts = await api(`/api/tcode/${code}/layouts`);
  const activeLayout = layouts.find(l => l.is_default) || layouts[0];
  const fields = activeLayout ? activeLayout.fields : [];
  const payload = { _plant: $('#std_f__plant').value };
  for (const f of fields) { const el = $('#std_f_' + f.key); payload[f.key] = el ? el.value.trim() : ""; }
  try { await api(`/api/tcode/${code}/records`, {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(payload)}); closeDialog(); renderStandardModule(code); setMsg("Saved"); }
  catch(e) { alert(e); }
}
async function deleteStandardRecord(code, rid) {
  if (!confirm(`Delete #${rid}?`)) return;
  try { await api(`/api/tcode/${code}/records/${rid}`, {method: 'DELETE'}); renderStandardModule(code); setMsg("Deleted"); }
  catch(e) { alert(e); }
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
    return {"status": "ACTIVE", "system": "NEXUS-CORE", "version": "3.8.0"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
