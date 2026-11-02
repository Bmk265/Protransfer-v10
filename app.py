import io, json, sqlite3, secrets, datetime, os, re, csv
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

    # Seed Master Org
    cur.execute(q("SELECT 1 FROM org WHERE id=1"))
    if not cur.fetchone():
        cur.execute(q("INSERT INTO org (id, name, system_id, client) VALUES (?,?,?,?)"), (1, "Thermocables Limited", "NX-01", "GLOBAL"))

    # Seed Departments
    for d in ["Production", "Final Testing", "Quality Control", "Commercial", "Store", "Planning", "Engineering"]:
        cur.execute(q("SELECT 1 FROM departments WHERE name=?"), (d,))
        if not cur.fetchone():
            cur.execute(q("INSERT INTO departments (name, active, created_at) VALUES (?,?,?)"), (d, 1, cur_time))

    # Seed Plants
    for p_code, p_name in [("1000", "Gooty Central Plant"), ("1400", "Extrusion Unit 1400"), ("2000", "Hyderabad Works"), ("CORP", "Corporate Headquarters")]:
        cur.execute(q("SELECT 1 FROM plants WHERE code=?"), (p_code,))
        if not cur.fetchone():
            cur.execute(q("INSERT INTO plants (code, name, active, created_at) VALUES (?,?,?,?)"), (p_code, p_name, 1, cur_time))

    # Master Admin Guarantee
    ph = bcrypt.hashpw("Admin@123".encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
    cur.execute(q("SELECT username FROM users WHERE username='admin'"))
    if cur.fetchone():
        cur.execute(q("UPDATE users SET password_hash=?, status='approved', role='master' WHERE username='admin'"), (ph,))
    else:
        cur.execute(q("INSERT INTO users (username, password_hash, full_name, role, department, plants, status, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?)"),
                    ("admin", ph, "Master System Administrator", "master", "IT & Systems", "1000,1400,2000,CORP", "approved", cur_time, cur_time))

    # Sample Users
    sample_users = [
        ("prod_user", "Production Supervisor", "Production", "1000"),
        ("testing_lead", "Final Testing Incharge", "Final Testing", "1000"),
        ("comm_lead", "Commercial Manager", "Commercial", "1000"),
        ("store_head", "Store Incharge", "Store", "1000")
    ]
    for su_name, su_full, su_dept, su_pl in sample_users:
        cur.execute(q("SELECT 1 FROM users WHERE username=?"), (su_name,))
        if not cur.fetchone():
            cur.execute(q("INSERT INTO users (username, password_hash, full_name, role, department, plants, status, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?)"),
                        (su_name, ph, su_full, "user", su_dept, su_pl, "approved", cur_time, cur_time))

    # Seed Default Modules
    default_modules = [
        ("TN01", "Transfer Note & Handover Sheet", "Inter-department material handover & batch transfer note", "MM", "Production,Final Testing,Commercial,Store,Quality Control,Planning"),
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
app = FastAPI(title="Nexus Core ERP", version="3.5.0")
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
    exec_write("DELETE FROM tcode_layouts WHERE code=?", (code,))
    exec_write("DELETE FROM tcode_records WHERE code=?", (code,))
    exec_write("DELETE FROM access_grants WHERE code=?", (code,))
    return {"ok": True}

# --- TN01 ROUTES & PRESETS ---
TN01_PRESET_ROUTES = {
    "PROD_TO_QC": {"title": "Production to Final Testing", "route": "Production ➔ Final Testing", "from_dept": "Production", "to_dept": "Final Testing", "columns": ["Sr. No.", "Job / Prod Order #", "Batch / Heat Lot", "Cable Description / Spec", "Drum / Coil #", "Length (Mtrs)", "Operator Name", "Production Remarks"]},
    "QC_TO_COMM": {"title": "Final Testing to Commercial", "route": "Final Testing ➔ Commercial", "from_dept": "Final Testing", "to_dept": "Commercial", "columns": ["Sr. No.", "Job / Prod Order #", "Batch / Heat Lot", "Cable Description / Spec", "Drum #", "Tested Length (Mtrs)", "HV Test Result", "Insulation Resistance", "QC Inspector", "Release Status"]},
    "PROD_TO_STORE": {"title": "Production to Store", "route": "Production ➔ Store", "from_dept": "Production", "to_dept": "Store", "columns": ["Sr. No.", "Work Order #", "Item Code / Material", "Batch / Heat Lot", "Quantity (Mtrs/Kgs)", "Weight (Kgs)", "Drum / Pack #", "Handover Person", "Remarks"]},
    "QC_TO_PROD": {"title": "Final Testing to Production (Rework/Reject)", "route": "Final Testing ➔ Production", "from_dept": "Final Testing", "to_dept": "Production", "columns": ["Sr. No.", "Job Order #", "Batch / Heat Lot", "Cable Size", "Defect Observed", "Length (Mtrs)", "QC Engineer", "Rework Instructions"]},
    "COMM_TO_QC": {"title": "Commercial to Final Testing (Inspection Request)", "route": "Commercial ➔ Final Testing", "from_dept": "Commercial", "to_dept": "Final Testing", "columns": ["Sr. No.", "Customer Name", "Order / PO #", "Item Description", "Ordered Qty", "Inspection Agency", "Target Dispatch Date", "Special Testing Notes"]},
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
def tn01_get_history(req: Request, plant: str = "", route: str = "", status: str = ""):
    need(req)
    sql, params = "SELECT * FROM transfer_notes WHERE 1=1", []
    if plant:
        sql += " AND (from_plant = ? OR to_plant = ?)"
        params.extend([plant, plant])
    if route:
        sql += " AND route = ?"
        params.append(route)
    if status:
        sql += " AND status = ?"
        params.append(status)
    sql += " ORDER BY id DESC LIMIT 150"
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

    th = "".join(f"<th style='border:1px solid #0f172a; padding:6px 8px; font-size:11px; background:#f1f5f9;'>{h}</th>" for h in headers)
    trs = "".join(f"<tr>{''.join(f'<td style=\"border:1px solid #cbd5e1; padding:6px 8px; font-size:11px;\">{c}</td>' for c in r)}</tr>" for r in data_rows)

    return HTMLResponse(f"""<!DOCTYPE html><html><head><meta charset="utf-8"><title>{note['note_no']} - Handover Note</title>
    <style>
      body {{ font-family: sans-serif; margin: 30px; color: #0f172a; }}
      .header {{ border-bottom: 2px solid #0f172a; padding-bottom: 10px; margin-bottom: 15px; }}
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
        <div class="sign-box">Sent By ({note['from_user_name']})</div>
        <div class="sign-box">Production Incharge | Quality Incharge</div>
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

# --- FRONTEND CLIENT (NORDIC SLATE UI) ---
HTML_MODERN_UI = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Nexus Core - Enterprise Operations Platform</title>
<script src="https://cdn.tailwindcss.com"></script>
<link href="https://cdn.jsdelivr.net/npm/remixicon@4.2.0/fonts/remixicon.css" rel="stylesheet">
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">
<style>
  body { font-family: 'Plus Jakarta Sans', sans-serif; background-color: #0b0f17; color: #1e293b; }
  .font-mono { font-family: 'JetBrains Mono', monospace; }
  .btn-primary { background: #4f46e5; color: white; transition: all 0.15s ease; }
  .btn-primary:hover { background: #4338ca; }
  .btn-secondary { background: #1e293b; border: 1px solid #334155; color: #cbd5e1; }
  .btn-secondary:hover { background: #334155; }
  .custom-table th { background: #f8fafc; color: #64748b; font-size: 11px; font-weight: 600; text-transform: uppercase; padding: 10px 14px; border-bottom: 1px solid #e2e8f0; }
  .custom-table td { padding: 12px 14px; border-bottom: 1px solid #f1f5f9; font-size: 13px; }
  .excel-grid { border-collapse: separate; border-spacing: 0; width: 100%; }
  .excel-grid th { background: #1e293b; color: #cbd5e1; font-weight: 600; font-size: 11px; text-transform: uppercase; padding: 8px 10px; border-right: 1px solid #334155; border-bottom: 1px solid #334155; position: sticky; top: 0; z-index: 2; }
  .excel-grid td { background: #0f172a; color: #e2e8f0; font-size: 12px; padding: 6px 10px; border-right: 1px solid #1e293b; border-bottom: 1px solid #1e293b; outline: none; min-width: 110px; }
  .excel-grid td.sr-col { background: #1e293b; color: #94a3b8; font-weight: bold; text-align: center; width: 50px; user-select: none; }
  .excel-grid td:focus { background: #1e1b4b; box-shadow: inset 0 0 0 2px #6366f1; }
</style>
</head>
<body class="h-screen flex flex-col select-none overflow-hidden bg-slate-900 text-slate-100">

<div id="app_root" class="h-full flex flex-col">
  <header class="h-14 bg-slate-900/90 backdrop-blur border-b border-slate-800 px-4 flex items-center justify-between z-10 shrink-0">
    <div class="flex items-center gap-4">
      <div class="flex items-center gap-2.5 cursor-pointer" onclick="goHome()">
        <div class="w-8 h-8 rounded-lg bg-indigo-600 flex items-center justify-center font-bold text-white shadow-md shadow-indigo-600/30"><i class="ri-pulse-fill text-lg"></i></div>
        <div class="font-bold text-sm tracking-tight text-white flex items-center gap-1.5">Nexus Core <span class="text-[10px] bg-indigo-500/20 text-indigo-300 border border-indigo-500/30 font-mono px-1.5 py-0.5 rounded">v3.5</span></div>
      </div>
      <div class="relative ml-4">
        <i class="ri-search-line absolute left-3 top-2.5 text-slate-400 text-xs"></i>
        <input id="cmd_box" placeholder="Jump to module (e.g. TN01, SU01, SPRO, SE93)..." class="bg-slate-800/80 border border-slate-700 rounded-lg pl-8 pr-16 py-1.5 text-xs text-slate-100 w-80 font-mono focus:outline-none focus:border-indigo-500" onkeydown="if(event.key==='Enter') execCmd()">
      </div>
    </div>
    <div id="nav_right" class="flex items-center gap-3">
      <div class="flex items-center gap-2 bg-slate-800 border border-slate-700 px-3 py-1 rounded-full text-xs">
        <div class="w-2 h-2 rounded-full bg-emerald-400 animate-pulse"></div>
        <span id="user_display" class="font-medium text-slate-200">Connecting...</span>
        <span id="role_badge" class="bg-indigo-950 text-indigo-300 border border-indigo-700/40 text-[10px] px-1.5 py-0.5 rounded-full font-semibold uppercase"></span>
      </div>
      <button onclick="logout()" title="Sign Out" class="text-xs bg-slate-800 hover:bg-red-950/40 text-slate-300 hover:text-red-400 border border-slate-700 p-2 rounded-lg"><i class="ri-logout-box-r-line"></i></button>
    </div>
  </header>

  <div class="flex-1 flex overflow-hidden">
    <aside class="w-64 bg-slate-950 border-r border-slate-800 flex flex-col justify-between shrink-0">
      <div class="p-3 overflow-y-auto flex-1 space-y-4">
        <div>
          <div class="text-[11px] uppercase font-bold text-slate-500 tracking-wider px-2 mb-1.5">Special Modules</div>
          <div class="space-y-0.5">
            <button onclick="openTcode('TN01')" class="w-full text-left px-2.5 py-2 rounded-xl text-xs font-semibold text-white bg-indigo-600/20 border border-indigo-500/30 hover:bg-indigo-600/30 flex items-center justify-between">
              <span class="flex items-center gap-2"><i class="ri-exchange-box-line text-indigo-400 text-sm"></i> TN01 - Transfer Note</span>
              <span class="text-[9px] bg-indigo-500 text-white px-1.5 py-0.5 rounded-full font-mono">EXCEL</span>
            </button>
          </div>
        </div>

        <div id="admin_nav_section" class="hidden">
          <div class="text-[11px] uppercase font-bold text-indigo-400 tracking-wider px-2 mb-1.5 flex items-center justify-between">
            <span>Administration</span><i class="ri-shield-keyhole-line text-xs"></i>
          </div>
          <div class="space-y-0.5">
            <button onclick="openTcode('SU01')" class="w-full text-left px-2.5 py-1.5 rounded-lg text-xs font-medium text-slate-300 hover:bg-slate-800 hover:text-white flex items-center justify-between">
              <span class="flex items-center gap-2"><i class="ri-user-settings-line text-indigo-400"></i> User Directory (SU01)</span>
              <span class="text-[9px] bg-indigo-900 text-indigo-300 px-1 rounded">IAM</span>
            </button>
            <button onclick="openTcode('SPRO')" class="w-full text-left px-2.5 py-1.5 rounded-lg text-xs font-medium text-slate-300 hover:bg-slate-800 hover:text-white flex items-center justify-between">
              <span class="flex items-center gap-2"><i class="ri-building-line text-slate-400"></i> Units & Plants (SPRO)</span>
              <span class="text-[9px] bg-indigo-900 text-indigo-300 px-1 rounded">MGMT</span>
            </button>
            <button onclick="openTcode('SE93')" class="w-full text-left px-2.5 py-1.5 rounded-lg text-xs font-medium text-slate-300 hover:bg-slate-800 hover:text-white flex items-center justify-between">
              <span class="flex items-center gap-2"><i class="ri-apps-2-line text-slate-400"></i> Module Studio (SE93)</span>
              <span class="text-[9px] bg-indigo-900 text-indigo-300 px-1 rounded">AUTO</span>
            </button>
            <button onclick="openTcode('PFCG')" class="w-full text-left px-2.5 py-1.5 rounded-lg text-xs font-medium text-slate-300 hover:bg-slate-800 hover:text-white flex items-center gap-2">
              <i class="ri-shield-user-line text-slate-400"></i> Access Policies (PFCG)
            </button>
          </div>
        </div>

        <div>
          <div class="text-[11px] uppercase font-bold text-slate-500 tracking-wider px-2 mb-1.5">Operations Modules</div>
          <div id="modules_list" class="space-y-0.5 text-xs"></div>
        </div>
      </div>
      <div class="p-3 border-t border-slate-800 bg-slate-950/60 text-[11px]">
        <div class="font-bold text-slate-300 truncate" id="side_company">Thermocables Limited</div>
        <div class="text-slate-500 truncate" id="side_plant">Facilities: Authorized</div>
      </div>
    </aside>
    <main id="main_screen" class="flex-1 bg-slate-900 overflow-hidden flex flex-col"></main>
  </div>

  <footer class="h-7 bg-slate-950 border-t border-slate-800 px-4 flex items-center justify-between text-[11px] text-slate-400 font-mono shrink-0">
    <div class="flex items-center gap-2 truncate">
      <span id="stat_dot" class="w-2 h-2 rounded-full bg-emerald-400"></span>
      <span id="stat_msg" class="text-slate-300">System Ready.</span>
    </div>
    <div class="flex items-center gap-4 text-slate-500">
      <span id="stat_tcode" class="text-indigo-400 font-semibold font-mono">DASHBOARD</span>
      <span id="stat_resp">8ms</span>
      <span>READY</span>
    </div>
  </footer>
</div>

<div id="app_dialog" class="hidden fixed inset-0 bg-slate-950/70 backdrop-blur-sm z-50 flex items-center justify-center p-4">
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
  dot.className = isErr ? "w-2 h-2 rounded-full bg-red-400" : "w-2 h-2 rounded-full bg-emerald-400";
  el.className = isErr ? "text-red-400 font-semibold" : "text-slate-300";
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
  <div class="min-h-screen bg-slate-950 flex items-center justify-center p-4">
    <div class="w-full max-w-md bg-slate-900 border border-slate-800 rounded-2xl shadow-2xl p-8 z-10">
      <div class="flex items-center gap-3 mb-6">
        <div class="w-10 h-10 rounded-xl bg-indigo-600 flex items-center justify-center font-bold text-white"><i class="ri-pulse-fill text-2xl"></i></div>
        <div><h1 class="text-xl font-bold text-white">Nexus Core ERP</h1><p class="text-xs text-slate-400">Thermocables Limited</p></div>
      </div>
      ${errorMsg ? `<div class="mb-4 p-3 bg-red-950/40 border border-red-800/60 rounded-xl text-red-300 text-xs">${esc(errorMsg)}</div>` : ''}
      <div class="space-y-4">
        <div><label class="block text-xs font-semibold text-slate-300 mb-1.5">Username</label><input id="l_u" class="w-full bg-slate-800 border border-slate-700 rounded-xl px-3.5 py-2.5 text-sm text-white" placeholder="admin" autofocus></div>
        <div><label class="block text-xs font-semibold text-slate-300 mb-1.5">Password</label><input id="l_p" type="password" class="w-full bg-slate-800 border border-slate-700 rounded-xl px-3.5 py-2.5 text-sm text-white" placeholder="••••••••" onkeydown="if(event.key==='Enter') handleLogon()"></div>
        <button onclick="handleLogon()" class="w-full btn-primary py-2.5 rounded-xl font-semibold text-sm">Sign In</button>
      </div>
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
    <div class="flex-1 overflow-y-auto p-8 bg-slate-900 select-text">
      <div class="max-w-5xl mx-auto space-y-8">
        <div><h1 class="text-2xl font-bold text-white">Operations Dashboard</h1><p class="text-sm text-slate-400 mt-1">Thermocables Ltd Management Platform</p></div>
        <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div class="bg-slate-800/60 border border-slate-700 rounded-2xl p-5">
            <div class="text-xs font-semibold text-slate-400 uppercase">Enterprise Org</div>
            <div class="text-xl font-bold text-white mt-1.5">${esc(ME.company)}</div>
            <div class="text-xs text-indigo-400 mt-1 font-medium">Dept: ${esc(ME.department)}</div>
          </div>
          <div class="bg-slate-800/60 border border-slate-700 rounded-2xl p-5">
            <div class="text-xs font-semibold text-slate-400 uppercase">Profile Role</div>
            <div class="text-xl font-bold text-emerald-400 mt-1.5 uppercase">${ME.role}</div>
            <div class="text-xs text-slate-400 mt-1 truncate">Plants: ${esc(ME.plants.join(', ') || 'Global')}</div>
          </div>
          <div class="bg-slate-800/60 border border-slate-700 rounded-2xl p-5">
            <div class="text-xs font-semibold text-slate-400 uppercase">Quick Jump</div>
            <div class="text-xs text-slate-300 mt-2 font-mono">Launch <b class="text-indigo-400">SU01</b> for Users, <b class="text-indigo-400">SPRO</b> for Plants, or <b class="text-indigo-400">TN01</b> for Handover.</div>
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
    <button onclick="openTcode('${t.code}')" class="w-full text-left px-2.5 py-1.5 rounded-lg text-slate-300 hover:bg-slate-800 hover:text-white flex items-center justify-between">
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
    </div>
  `).join('');
}

async function openTcode(code) {
  code = code.toUpperCase().trim();
  CURRENT_TCODE = code;
  $('#stat_tcode').textContent = code;
  setMsg(`Module ${code} loaded`);

  if (code === 'TN01') return renderTN01();
  if (code === 'SU01') return renderSU01();
  if (code === 'SPRO') return renderSPRO();
  if (code === 'SE93') return renderSE93();
  if (code === 'PFCG') return renderPFCG();

  renderStandardModule(code);
}

// =========================================================================
// --- SU01: USER DIRECTORY & IAM (COMPLETE WITH ALL ACTIONS) ---
// =========================================================================
async function renderSU01() {
  if (!ME.is_master && ME.role !== 'admin') return setMsg("Permission denied", true);
  ALL_USERS_LIST = await api('/api/su01/users');

  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col bg-slate-900 p-6 overflow-hidden select-text">
      <div class="flex items-center justify-between pb-5 border-b border-slate-800 shrink-0">
        <div>
          <h2 class="text-lg font-bold text-white tracking-tight">User Directory & IAM (SU01)</h2>
          <p class="text-xs text-slate-400">Manage user credentials, plant access, roles, and status</p>
        </div>
        <div class="flex items-center gap-2">
          <input id="su01_search" oninput="filterUsersTable(this.value)" placeholder="Search users, plants..." class="bg-slate-800 border border-slate-700 text-slate-200 text-xs rounded-xl px-3 py-1.5 focus:outline-none focus:border-indigo-500">
          <button onclick="showCreateUserModal()" class="btn-primary text-xs font-bold px-4 py-2 rounded-xl flex items-center gap-1.5 shadow-lg shadow-indigo-600/30">
            <i class="ri-user-add-line"></i> + Add New User
          </button>
          <button onclick="renderSU01()" title="Refresh" class="btn-secondary text-xs px-3 py-2 rounded-xl"><i class="ri-refresh-line"></i></button>
        </div>
      </div>

      <div class="flex-1 mt-4 bg-slate-950 border border-slate-800 rounded-2xl overflow-auto">
        <table class="w-full text-left custom-table">
          <thead>
            <tr>
              <th>Username</th><th>Full Name</th><th>Role</th><th>Department</th><th>Facilities</th><th>Status</th><th style="width:200px;">Actions</th>
            </tr>
          </thead>
          <tbody id="su01_tbody" class="text-slate-300"></tbody>
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
    tbody.innerHTML = `<tr><td colspan="7" class="text-center py-10 text-slate-500">No users found.</td></tr>`;
    return;
  }
  tbody.innerHTML = users.map(u => {
    const isSelf = (u.username === ME.username);
    const isMaster = (u.role === 'master');
    const canDelete = !isSelf && (!isMaster || ME.is_master);

    return `
      <tr class="border-b border-slate-800/60 hover:bg-slate-900/60">
        <td class="font-mono font-bold text-indigo-400">${esc(u.username)}</td>
        <td class="font-medium text-white">${esc(u.full_name)}</td>
        <td><span class="text-[10px] font-semibold px-2 py-0.5 rounded-full ${u.role==='master'?'bg-red-950 text-red-300 border border-red-800':(u.role==='admin'?'bg-amber-950 text-amber-300 border border-amber-800':'bg-slate-800 text-slate-300')}">${u.role.toUpperCase()}</span></td>
        <td>${esc(u.department)}</td>
        <td class="font-mono text-xs">${esc(u.plants || 'ALL')}</td>
        <td><span class="text-[10px] font-semibold px-2 py-0.5 rounded-full ${u.status==='approved'?'bg-emerald-950 text-emerald-300 border border-emerald-800':'bg-red-950 text-red-300 border border-red-800'}">${u.status}</span></td>
        <td class="space-x-1">
          <button onclick="showEditUserModal('${esc(u.username)}')" class="text-xs text-indigo-400 hover:text-indigo-300 font-semibold px-1.5 py-1">Edit</button>
          <button onclick="showResetPassModal('${esc(u.username)}')" class="text-xs text-amber-400 hover:text-amber-300 font-semibold px-1.5 py-1">Reset PW</button>
          ${canDelete ? `<button onclick="deleteUser('${esc(u.username)}')" class="text-xs text-red-400 hover:text-red-300 font-semibold px-1.5 py-1">Delete</button>` : ''}
        </td>
      </tr>`;
  }).join('');
}

function showCreateUserModal() {
  api('/api/spro/overview').then(data => {
    openDialog(`
      <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl">
        <h3 class="font-bold text-slate-800 text-sm">Create New Operator / User</h3>
        <button onclick="closeDialog()"><i class="ri-close-line text-lg"></i></button>
      </div>
      <div class="p-6 space-y-4">
        <div class="grid grid-cols-2 gap-4">
          <div><label class="block text-xs font-semibold text-slate-600 mb-1">Username *</label><input id="nu_u" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs font-mono" placeholder="e.g. jdoe"></div>
          <div><label class="block text-xs font-semibold text-slate-600 mb-1">Password *</label><input id="nu_p" type="password" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs" placeholder="••••••••"></div>
        </div>
        <div><label class="block text-xs font-semibold text-slate-600 mb-1">Full Name *</label><input id="nu_fn" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs" placeholder="e.g. John Doe"></div>
        <div class="grid grid-cols-2 gap-4">
          <div>
            <label class="block text-xs font-semibold text-slate-600 mb-1">Department</label>
            <select id="nu_d" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs bg-white">
              ${data.departments.map(d=>`<option value="${esc(d.name)}">${esc(d.name)}</option>`).join('')}
            </select>
          </div>
          <div>
            <label class="block text-xs font-semibold text-slate-600 mb-1">Role</label>
            <select id="nu_r" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs bg-white">
              <option value="user">Standard User</option>
              <option value="admin">Administrator</option>
              ${ME.is_master ? '<option value="master">Master Superuser</option>' : ''}
            </select>
          </div>
        </div>
        <div>
          <label class="block text-xs font-semibold text-slate-600 mb-1">Plant Authorization (comma-separated, leave blank for ALL)</label>
          <input id="nu_pl" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs font-mono" placeholder="e.g. 1000, 1400">
        </div>
      </div>
      <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl">
        <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-xl text-xs font-semibold">Cancel</button>
        <button onclick="submitCreateUser()" class="btn-primary px-4 py-2 rounded-xl text-xs font-semibold">Save User</button>
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
            <option value="approved" ${user.status==='approved'?'selected':''}>Approved / Active</option>
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
// --- SPRO: COMPLETE PLANTS, DEPTS & ENTERPRISE MANAGEMENT ---
// =========================================================================
async function renderSPRO() {
  if (!ME.is_master && ME.role !== 'admin') return setMsg("Permission denied", true);
  const data = await api('/api/spro/overview');

  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col bg-slate-900 p-6 overflow-hidden select-text">
      <div class="flex items-center justify-between pb-5 border-b border-slate-800 shrink-0">
        <div>
          <h2 class="text-lg font-bold text-white tracking-tight">Enterprise Structure & Customizing (SPRO)</h2>
          <p class="text-xs text-slate-400">Manage manufacturing plants, facilities, departments, and corporate identity</p>
        </div>
        <button onclick="renderSPRO()" class="btn-secondary text-xs px-3 py-1.5 rounded-xl"><i class="ri-refresh-line"></i> Refresh</button>
      </div>

      <!-- ORG INFO STRIP -->
      <div class="mt-4 bg-slate-950 border border-slate-800 p-4 rounded-2xl flex items-center justify-between shrink-0">
        <div>
          <div class="text-[10px] text-slate-400 uppercase font-semibold">Corporate Enterprise Identity</div>
          <div class="text-base font-bold text-white">${esc(data.org.name)} <span class="text-xs font-mono text-indigo-400">(${esc(data.org.system_id)} - Client: ${esc(data.org.client)})</span></div>
        </div>
        ${ME.is_master ? `<button onclick="showEditOrgModal('${esc(data.org.name)}','${esc(data.org.system_id)}','${esc(data.org.client)}')" class="btn-secondary text-xs px-3 py-1.5 rounded-xl">Edit Identity</button>` : ''}
      </div>

      <!-- PLANTS & DEPARTMENTS SPLIT VIEW -->
      <div class="grid grid-cols-1 md:grid-cols-2 gap-6 flex-1 mt-4 overflow-auto">
        <!-- PLANTS / FACILITIES -->
        <div class="bg-slate-950 border border-slate-800 rounded-2xl flex flex-col overflow-hidden">
          <div class="p-3.5 border-b border-slate-800 flex justify-between items-center bg-slate-900/60">
            <div>
              <span class="text-xs font-bold text-white uppercase tracking-wider flex items-center gap-1.5">
                <i class="ri-building-line text-indigo-400"></i> Manufacturing Plants (${data.plants.length})
              </span>
            </div>
            ${ME.is_master ? `<button onclick="showCreatePlantModal()" class="btn-primary text-xs font-bold px-3 py-1.5 rounded-lg">+ Add Facility</button>` : ''}
          </div>
          <div class="flex-1 overflow-auto">
            <table class="w-full text-left custom-table">
              <thead><tr><th>Code</th><th>Plant Name</th><th>Status</th>${ME.is_master?'<th style="width:110px;">Actions</th>':''}</tr></thead>
              <tbody class="text-slate-300">
                ${data.plants.map(p => `
                  <tr class="border-b border-slate-800/60">
                    <td class="font-mono font-bold text-indigo-400">${esc(p.code)}</td>
                    <td class="font-medium text-white">${esc(p.name)}</td>
                    <td><span class="text-[10px] font-semibold px-2 py-0.5 rounded-full ${p.active?'bg-emerald-950 text-emerald-300':'bg-red-950 text-red-300'}">${p.active?'ACTIVE':'DISABLED'}</span></td>
                    ${ME.is_master ? `
                    <td class="space-x-1">
                      <button onclick="showEditPlantModal('${esc(p.code)}','${esc(p.name)}',${p.active})" class="text-xs text-indigo-400 hover:text-indigo-300 font-semibold">Edit</button>
                      <button onclick="deletePlant('${esc(p.code)}')" class="text-xs text-red-400 hover:text-red-300 font-semibold ml-2">Delete</button>
                    </td>` : ''}
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        </div>

        <!-- DEPARTMENTS -->
        <div class="bg-slate-950 border border-slate-800 rounded-2xl flex flex-col overflow-hidden">
          <div class="p-3.5 border-b border-slate-800 flex justify-between items-center bg-slate-900/60">
            <div>
              <span class="text-xs font-bold text-white uppercase tracking-wider flex items-center gap-1.5">
                <i class="ri-team-line text-emerald-400"></i> Operational Departments (${data.departments.length})
              </span>
            </div>
            ${ME.is_master ? `<button onclick="showCreateDeptModal()" class="btn-primary text-xs font-bold px-3 py-1.5 rounded-lg">+ Add Dept</button>` : ''}
          </div>
          <div class="flex-1 overflow-auto">
            <table class="w-full text-left custom-table">
              <thead><tr><th>Department Name</th><th>Status</th>${ME.is_master?'<th style="width:80px;">Action</th>':''}</tr></thead>
              <tbody class="text-slate-300">
                ${data.departments.map(d => `
                  <tr class="border-b border-slate-800/60">
                    <td class="font-medium text-white">${esc(d.name)}</td>
                    <td><span class="text-[10px] font-semibold px-2 py-0.5 rounded-full ${d.active?'bg-emerald-950 text-emerald-300':'bg-red-950 text-red-300'}">${d.active?'ACTIVE':'DISABLED'}</span></td>
                    ${ME.is_master ? `<td><button onclick="deleteDept('${esc(d.name)}')" class="text-xs text-red-400 hover:text-red-300 font-semibold">Delete</button></td>` : ''}
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
      <h3 class="font-bold text-slate-800 text-sm">Edit Enterprise Organization Identity</h3>
      <button onclick="closeDialog()"><i class="ri-close-line text-lg"></i></button>
    </div>
    <div class="p-6 space-y-4">
      <div><label class="block text-xs font-semibold text-slate-700 mb-1">Company / Group Name</label><input id="org_n" value="${esc(n)}" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs"></div>
      <div class="grid grid-cols-2 gap-4">
        <div><label class="block text-xs font-semibold text-slate-700 mb-1">System Identifier</label><input id="org_sys" value="${esc(sys)}" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs font-mono uppercase"></div>
        <div><label class="block text-xs font-semibold text-slate-700 mb-1">Client Code</label><input id="org_cl" value="${esc(cl)}" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs font-mono"></div>
      </div>
    </div>
    <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl">
      <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-xl text-xs font-semibold">Cancel</button>
      <button onclick="submitEditOrg()" class="btn-primary px-4 py-2 rounded-xl text-xs font-semibold">Save Identity</button>
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
      <h3 class="font-bold text-slate-800 text-sm">Add Manufacturing Plant / Facility</h3>
      <button onclick="closeDialog()"><i class="ri-close-line text-lg"></i></button>
    </div>
    <div class="p-6 space-y-4">
      <div><label class="block text-xs font-semibold text-slate-700 mb-1">Plant Code (e.g. 1500, HYD1) *</label><input id="np_c" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs font-mono uppercase"></div>
      <div><label class="block text-xs font-semibold text-slate-700 mb-1">Facility Name / Description *</label><input id="np_n" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs" placeholder="e.g. Copper Drawing Unit"></div>
    </div>
    <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl">
      <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-xl text-xs font-semibold">Cancel</button>
      <button onclick="submitCreatePlant()" class="btn-primary px-4 py-2 rounded-xl text-xs font-semibold">Save Plant</button>
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
      <h3 class="font-bold text-slate-800 text-sm">Edit Facility: ${esc(code)}</h3>
      <button onclick="closeDialog()"><i class="ri-close-line text-lg"></i></button>
    </div>
    <div class="p-6 space-y-4">
      <div><label class="block text-xs font-semibold text-slate-700 mb-1">Facility Name</label><input id="ep_n" value="${esc(name)}" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs"></div>
      <div>
        <label class="block text-xs font-semibold text-slate-700 mb-1">Status</label>
        <select id="ep_st" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs bg-white">
          <option value="1" ${active?'selected':''}>Active</option>
          <option value="0" ${!active?'selected':''}>Disabled</option>
        </select>
      </div>
    </div>
    <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl">
      <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-xl text-xs font-semibold">Cancel</button>
      <button onclick="submitEditPlant('${esc(code)}')" class="btn-primary px-4 py-2 rounded-xl text-xs font-semibold">Save Changes</button>
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
  if (!confirm(`Permanently remove facility ${code}?`)) return;
  const fd = new FormData(); fd.append("code", code);
  try {
    await api('/api/admin/plants/delete', {method: 'POST', body: fd});
    renderSPRO();
    setMsg(`Plant ${code} deleted`);
  } catch(e) { alert(e); }
}

function showCreateDeptModal() {
  const d = prompt("Enter New Department Name (e.g. Maintenance, R&D):");
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
// --- PFCG & TN01 FUNCTIONS ---
// =========================================================================
async function renderPFCG() {
  if (!ME.is_master && ME.role !== 'admin') return setMsg("Permission denied", true);
  const m = await api('/api/pfcg/matrix');
  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col bg-slate-900 p-6 overflow-hidden select-text">
      <div class="flex items-center justify-between pb-5 border-b border-slate-800 shrink-0">
        <div><h2 class="text-lg font-bold text-white tracking-tight">Access Control & Security Matrix (PFCG)</h2></div>
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
  if (!confirm(`Revoke ${t} from ${u}?`)) return;
  const fd = new FormData(); fd.append("username", u); fd.append("tcode", t);
  await api('/api/pfcg/revoke', {method: 'POST', body: fd});
  renderPFCG();
  setMsg("Permission revoked");
}

// TN01 WORKSPACE
let TN01_ROUTES = {}, TN01_ACTIVE_ROUTE = "PROD_TO_QC", TN01_HEADERS = [], TN01_ROWS = [], TN01_UNDO_STACK = [], TN01_REDO_STACK = [], TN01_ACTIVE_CELL = { r: 0, c: 1 };

async function renderTN01() {
  TN01_ROUTES = await api('/api/tn01/routes');
  renderTN01Main();
}

function renderTN01Main(activeTab = "sheet") {
  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col bg-slate-900 overflow-hidden select-text">
      <div class="p-3 bg-slate-950 border-b border-slate-800 flex items-center justify-between shrink-0">
        <div class="flex items-center gap-3">
          <div class="flex items-center gap-2">
            <span class="font-mono font-bold text-xs bg-indigo-500/20 text-indigo-300 border border-indigo-500/30 px-2.5 py-1 rounded-lg">TN01</span>
            <div><h2 class="text-sm font-bold text-white">Transfer Note & Handover System</h2><p class="text-[10px] text-slate-400">Inter-department batch inspection & dispatch slips</p></div>
          </div>
          <div class="h-6 w-[1px] bg-slate-800 mx-2"></div>
          <div class="flex items-center bg-slate-900 border border-slate-800 rounded-xl p-1 gap-1 text-xs font-semibold">
            <button onclick="switchTN01Tab('sheet')" class="px-3 py-1.5 rounded-lg flex items-center gap-1.5 ${activeTab==='sheet'?'bg-indigo-600 text-white shadow':'text-slate-400 hover:text-white'}"><i class="ri-table-line"></i> Excel Sheet</button>
            <button onclick="switchTN01Tab('receive')" class="px-3 py-1.5 rounded-lg flex items-center gap-1.5 ${activeTab==='receive'?'bg-indigo-600 text-white shadow':'text-slate-400 hover:text-white'}"><i class="ri-inbox-archive-line"></i> Receive Notes</button>
            <button onclick="switchTN01Tab('history')" class="px-3 py-1.5 rounded-lg flex items-center gap-1.5 ${activeTab==='history'?'bg-indigo-600 text-white shadow':'text-slate-400 hover:text-white'}"><i class="ri-history-line"></i> History & Log</button>
            ${ME.is_master || ME.role==='admin' ? `<button onclick="switchTN01Tab('format')" class="px-3 py-1.5 rounded-lg flex items-center gap-1.5 ${activeTab==='format'?'bg-indigo-600 text-white shadow':'text-slate-400 hover:text-white'}"><i class="ri-paint-brush-line"></i> Format Studio (Admin)</button>` : ''}
          </div>
        </div>
        <div>${activeTab === 'sheet' ? `<button onclick="openSendModal()" class="btn-primary text-xs font-bold px-4 py-2 rounded-xl flex items-center gap-1.5 shadow-lg shadow-indigo-600/40"><i class="ri-send-plane-fill"></i> Send Transfer Note</button>` : ''}</div>
      </div>
      <div id="tn01_content" class="flex-1 overflow-hidden flex flex-col bg-slate-900"></div>
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
    <div class="p-2 bg-slate-900 border-b border-slate-800 flex items-center justify-between text-xs shrink-0">
      <div class="flex items-center gap-2">
        <label class="font-bold text-slate-400 text-xs flex items-center gap-1.5">Preset Route:
          <select id="route_selector" onchange="applyRouteLayout(this.value, true)" class="bg-slate-800 border border-slate-700 text-indigo-300 font-semibold rounded-lg px-2.5 py-1 text-xs">
            ${Object.keys(TN01_ROUTES).map(k => `<option value="${k}" ${k===TN01_ACTIVE_ROUTE?'selected':''}>${TN01_ROUTES[k].route}</option>`).join('')}
          </select>
        </label>
        <div class="h-4 w-[1px] bg-slate-800 mx-1"></div>
        <button onclick="undoGrid()" class="bg-slate-800 text-slate-300 px-2 py-1 rounded border border-slate-700"><i class="ri-arrow-go-back-line"></i></button>
        <button onclick="redoGrid()" class="bg-slate-800 text-slate-300 px-2 py-1 rounded border border-slate-700"><i class="ri-arrow-go-forward-line"></i></button>
        <button onclick="addGridRow()" class="bg-slate-800 text-slate-300 px-2.5 py-1 rounded border border-slate-700">+ Row</button>
        <button onclick="deleteGridRow()" class="bg-slate-800 text-red-400 px-2.5 py-1 rounded border border-slate-700">- Row</button>
        <button onclick="clearGrid()" class="bg-slate-800 text-slate-400 px-2 py-1 rounded border border-slate-700">Clear</button>
      </div>
      <div class="flex items-center gap-2 text-slate-400 text-xs">
        <button onclick="copyGridForExcel()" class="bg-slate-800 text-emerald-400 border border-slate-700 px-3 py-1 rounded-lg flex items-center gap-1 font-semibold"><i class="ri-file-copy-line"></i> Copy for Excel</button>
        <span class="text-[11px] bg-slate-800 border border-slate-700 px-2 py-0.5 rounded text-slate-400 font-mono">Press <b>Ctrl+V</b> to paste from Excel</span>
      </div>
    </div>
    <div id="grid_table_container" class="flex-1 overflow-auto bg-slate-950 p-2 select-text" tabindex="0"></div>
    <div class="p-2 bg-slate-950 border-t border-slate-800 text-slate-400 text-[11px] flex justify-between items-center font-mono shrink-0">
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

function undoGrid() {
  if (!TN01_UNDO_STACK.length) return;
  TN01_REDO_STACK.push(JSON.stringify(TN01_ROWS));
  TN01_ROWS = JSON.parse(TN01_UNDO_STACK.pop());
  renderGridTable();
}

function redoGrid() {
  if (!TN01_REDO_STACK.length) return;
  TN01_UNDO_STACK.push(JSON.stringify(TN01_ROWS));
  TN01_ROWS = JSON.parse(TN01_REDO_STACK.pop());
  renderGridTable();
}

function addGridRow() {
  TN01_UNDO_STACK.push(JSON.stringify(TN01_ROWS));
  const r = new Array(TN01_HEADERS.length).fill("");
  r[0] = TN01_ROWS.length + 1;
  TN01_ROWS.push(r);
  renderGridTable();
}

function deleteGridRow() {
  if (TN01_ROWS.length <= 1) return;
  TN01_UNDO_STACK.push(JSON.stringify(TN01_ROWS));
  TN01_ROWS.pop();
  renderGridTable();
}

function clearGrid() {
  if (!confirm("Clear all rows in the sheet?")) return;
  TN01_UNDO_STACK.push(JSON.stringify(TN01_ROWS));
  TN01_ROWS = [];
  for (let i = 1; i <= 6; i++) {
    const r = new Array(TN01_HEADERS.length).fill("");
    r[0] = i;
    TN01_ROWS.push(r);
  }
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
      while (TN01_ROWS.length <= targetR) {
        const nr = new Array(TN01_HEADERS.length).fill("");
        nr[0] = TN01_ROWS.length + 1;
        TN01_ROWS.push(nr);
      }
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

function copyGridForExcel() {
  let tsv = TN01_HEADERS.join('\\t') + '\\n';
  TN01_ROWS.forEach(r => { tsv += r.join('\\t') + '\\n'; });
  navigator.clipboard.writeText(tsv).then(() => { alert("Copied! Paste into Excel with Ctrl+V."); });
}

async function openSendModal() {
  const activeRoute = TN01_ROUTES[TN01_ACTIVE_ROUTE];
  const filledRows = TN01_ROWS.filter(r => r.slice(1).some(x => String(x).trim() !== ''));
  if (!filledRows.length) return alert("Please enter at least one entry before sending.");
  const spro = await api('/api/spro/overview');
  const myPlant = (ME.plants && ME.plants[0]) ? ME.plants[0] : "1000";

  openDialog(`
    <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl">
      <div><h3 class="font-bold text-slate-800 text-sm">Transfer Dispatch & Handover Note</h3><p class="text-xs text-slate-500">Route: ${activeRoute.route}</p></div>
      <button onclick="closeDialog()"><i class="ri-close-line text-lg"></i></button>
    </div>
    <div class="p-6 space-y-4">
      <div class="grid grid-cols-3 gap-3">
        <div><label class="block text-xs font-semibold text-slate-700 mb-1">Target Plant</label>
          <select id="send_plant" onchange="updateSuggestedUsers()" class="w-full border border-slate-300 rounded-xl p-2 text-xs bg-white">
            ${spro.plants.map(p=>`<option value="${esc(p.code)}" ${p.code===myPlant?'selected':''}>${esc(p.code)} - ${esc(p.name)}</option>`).join('')}
          </select>
        </div>
        <div><label class="block text-xs font-semibold text-slate-700 mb-1">Target Department</label>
          <select id="send_dept" onchange="updateSuggestedUsers()" class="w-full border border-slate-300 rounded-xl p-2 text-xs bg-white">
            ${spro.departments.map(d=>`<option value="${esc(d.name)}" ${d.name===activeRoute.to_dept?'selected':''}>${esc(d.name)}</option>`).join('')}
          </select>
        </div>
        <div><label class="block text-xs font-semibold text-slate-700 mb-1">Assigned Receiver</label>
          <select id="send_user" class="w-full border border-slate-300 rounded-xl p-2 text-xs bg-white font-medium text-indigo-900"></select>
        </div>
      </div>
      <div><label class="block text-xs font-semibold text-slate-700 mb-1">Remarks</label>
        <textarea id="send_remarks" class="w-full border border-slate-300 rounded-xl p-2 text-xs" rows="2" placeholder="Dispatch notes..."></textarea>
      </div>
    </div>
    <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl">
      <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-xl text-xs font-semibold">Cancel</button>
      <button onclick="submitSendNote()" class="btn-primary px-5 py-2 rounded-xl text-xs font-bold">Confirm & Send</button>
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
    alert(`Transfer Dispatched! Note #: ${res.note_no}`);
    switchTN01Tab('history');
  } catch(e) { alert(e); }
}

async function renderTN01ReceiveTab() {
  const notes = await api('/api/tn01/inbox');
  $('#tn01_content').innerHTML = `
    <div class="p-6 flex-1 flex flex-col overflow-hidden">
      <div class="flex items-center justify-between pb-4 border-b border-slate-800 shrink-0">
        <div><h2 class="text-base font-bold text-white">Incoming Transfer Notes (Receive Inbox)</h2></div>
        <button onclick="renderTN01ReceiveTab()" class="btn-secondary text-xs px-3 py-1.5 rounded-lg"><i class="ri-refresh-line"></i> Refresh</button>
      </div>
      <div class="flex-1 mt-4 bg-slate-950 border border-slate-800 rounded-2xl overflow-auto">
        <table class="w-full text-left custom-table">
          <thead><tr><th>Note #</th><th>Route</th><th>Origin</th><th>Sender</th><th>Batches</th><th>Time</th><th>Status</th><th>Actions</th></tr></thead>
          <tbody class="text-slate-300">
            ${!notes.length ? `<tr><td colspan="8" class="text-center py-12 text-slate-500">Inbox is empty.</td></tr>` : ''}
            ${notes.map(n => `
              <tr class="border-b border-slate-800/60">
                <td class="font-mono font-bold text-indigo-400">${esc(n.note_no)}</td>
                <td class="font-semibold text-white">${esc(n.route)}</td>
                <td>${esc(n.from_plant)} - ${esc(n.from_dept)}</td>
                <td>${esc(n.from_user_name)}</td>
                <td class="font-mono font-bold">${n.entry_count}</td>
                <td class="font-mono text-xs text-slate-400">${esc(n.created_at)}</td>
                <td><span class="text-[10px] font-bold px-2 py-0.5 rounded-full ${n.status==='received'?'bg-emerald-950 text-emerald-300':'bg-amber-950 text-amber-300 animate-pulse'}">${n.status.toUpperCase()}</span></td>
                <td class="space-x-1">
                  <button onclick="viewTransferNote(${n.id})" class="btn-primary text-xs font-semibold px-2.5 py-1 rounded-lg">Open Sheet</button>
                  ${n.status==='sent' ? `<button onclick="acknowledgeNote(${n.id})" class="bg-emerald-600 text-white text-xs font-semibold px-2.5 py-1 rounded-lg">Acknowledge</button>` : ''}
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
      <div><span class="text-xs bg-indigo-100 text-indigo-800 font-mono font-bold px-2 py-0.5 rounded">${esc(note.note_no)}</span><h3 class="font-bold text-slate-800 text-sm mt-1">${esc(note.route)}</h3></div>
      <button onclick="closeDialog()"><i class="ri-close-line text-lg"></i></button>
    </div>
    <div class="p-6 space-y-4">
      <div class="flex justify-between items-center">
        <span class="text-xs font-bold text-slate-700">Batches (${note.entry_count} rows):</span>
        <div class="flex gap-2">
          <button onclick="copyReceivedForExcel(${nid})" class="btn-secondary text-xs px-3 py-1.5 rounded-lg flex items-center gap-1 text-emerald-700 font-semibold border-emerald-300"><i class="ri-file-copy-line"></i> Copy for Excel</button>
          <button onclick="loadNoteIntoWorkspace(${nid})" class="btn-primary text-xs px-3 py-1.5 rounded-lg flex items-center gap-1 font-semibold"><i class="ri-download-cloud-2-line"></i> Load into Workspace</button>
        </div>
      </div>
      <div class="max-h-72 overflow-auto border border-slate-200 rounded-xl">
        <table class="w-full text-left text-xs">
          <thead><tr class="bg-slate-100 text-slate-700 font-bold">${note.sheet_headers.map(h => `<th class="p-2 border-r">${esc(h)}</th>`).join('')}</tr></thead>
          <tbody>${note.sheet_data.map(r => `<tr class="border-b">${r.map(c => `<td class="p-2 border-r">${esc(c)}</td>`).join('')}</tr>`).join('')}</tbody>
        </table>
      </div>
    </div>
    <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-between items-center rounded-b-2xl">
      <a href="/api/tn01/print/${note.id}" target="_blank" class="btn-secondary text-xs px-3 py-2 rounded-xl flex items-center gap-1 font-semibold"><i class="ri-printer-line"></i> Print / PDF</a>
      <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-xl text-xs font-semibold">Close</button>
    </div>`);
}

async function copyReceivedForExcel(nid) {
  const note = await api(`/api/tn01/note/${nid}`);
  let tsv = note.sheet_headers.join('\\t') + '\\n';
  note.sheet_data.forEach(r => { tsv += r.join('\\t') + '\\n'; });
  navigator.clipboard.writeText(tsv).then(() => { alert("Copied! Paste into Excel with Ctrl+V."); });
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
  const p = $('#hist_p') ? $('#hist_p').value : '';
  const history = await api(`/api/tn01/history?plant=${p}`);
  $('#tn01_content').innerHTML = `
    <div class="p-6 flex-1 flex flex-col overflow-hidden">
      <div class="flex-1 mt-2 bg-slate-950 border border-slate-800 rounded-2xl overflow-auto">
        <table class="w-full text-left custom-table">
          <thead><tr><th>Note #</th><th>Route</th><th>From</th><th>To</th><th>Batches</th><th>Sender</th><th>Timestamp</th><th>Status</th><th>Actions</th></tr></thead>
          <tbody class="text-slate-300">
            ${!history.length ? `<tr><td colspan="9" class="text-center py-10 text-slate-500">No records found.</td></tr>` : ''}
            ${history.map(h => `
              <tr class="border-b border-slate-800/60">
                <td class="font-mono font-bold text-indigo-400">${esc(h.note_no)}</td>
                <td class="font-semibold text-white">${esc(h.route)}</td>
                <td>${esc(h.from_plant)} - ${esc(h.from_dept)}</td>
                <td>${esc(h.to_plant)} - ${esc(h.to_dept)}</td>
                <td class="font-mono font-bold">${h.entry_count}</td>
                <td>${esc(h.from_user_name)}</td>
                <td class="font-mono text-xs text-slate-400">${esc(h.created_at)}</td>
                <td><span class="text-[10px] font-bold px-2 py-0.5 rounded-full ${h.status==='received'?'bg-emerald-950 text-emerald-300':'bg-amber-950 text-amber-300'}">${h.status.toUpperCase()}</span></td>
                <td class="space-x-1">
                  <button onclick="viewTransferNote(${h.id})" class="text-xs text-indigo-400 font-semibold">View</button>
                  <a href="/api/tn01/print/${h.id}" target="_blank" class="text-xs text-emerald-400 font-semibold ml-2">PDF</a>
                </td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    </div>`;
}

async function renderTN01FormatTab() {
  const layouts = await api('/api/tn01/layouts');
  $('#tn01_content').innerHTML = `
    <div class="p-6 flex-1 flex flex-col overflow-hidden">
      <div class="flex items-center justify-between pb-4 border-b border-slate-800 shrink-0">
        <div><h2 class="text-base font-bold text-white">Transfer Note Formats (Admin)</h2></div>
        <button onclick="openCreateLayoutModal()" class="btn-primary text-xs font-semibold px-4 py-2 rounded-xl">+ Create Format</button>
      </div>
      <div class="grid grid-cols-1 md:grid-cols-3 gap-4 mt-4 overflow-auto">
        ${layouts.map(l => `
          <div class="bg-slate-950 border border-slate-800 p-4 rounded-2xl flex flex-col justify-between">
            <div>
              <div class="flex justify-between"><span class="font-mono text-xs text-indigo-400 font-bold">${esc(l.layout_code)}</span><button onclick="deleteLayoutTN01(${l.id})" class="text-red-400 text-xs">Delete</button></div>
              <h4 class="font-bold text-white text-sm mt-2">${esc(l.name)}</h4>
              <p class="text-xs text-slate-400 mt-1">${(l.fields || []).length} Columns</p>
            </div>
            <button onclick="loadFormatIntoWorkspace(${l.id})" class="btn-primary px-3 py-1 rounded-lg text-xs mt-3">Use in Grid</button>
          </div>
        `).join('')}
      </div>
    </div>`;
}

function openCreateLayoutModal() {
  openDialog(`
    <div class="px-6 py-4 border-b border-slate-200 flex justify-between items-center bg-slate-50 rounded-t-2xl">
      <h3 class="font-bold text-slate-800 text-sm">Design TN01 Custom Format</h3>
      <button onclick="closeDialog()"><i class="ri-close-line text-lg"></i></button>
    </div>
    <div class="p-6 space-y-4">
      <div><label class="block text-xs font-semibold text-slate-700 mb-1">Layout Name</label><input id="nl_n" placeholder="e.g. 11kV Armoured Cable Format" class="w-full border border-slate-300 rounded-xl p-2 text-xs"></div>
      <div><label class="block text-xs font-semibold text-slate-700 mb-1">Columns (comma-separated)</label>
        <textarea id="nl_c" class="w-full border border-slate-300 rounded-xl p-2 text-xs font-mono" rows="3">Sr. No., Production Order, Heat Lot, Conductor Size, Length (Mtrs), Weight (Kg), Remarks</textarea>
      </div>
    </div>
    <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl">
      <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-xl text-xs font-semibold">Cancel</button>
      <button onclick="submitSaveFormat()" class="btn-primary px-4 py-2 rounded-xl text-xs font-semibold">Save Layout</button>
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
  for (let i = 1; i <= 6; i++) {
    const r = new Array(TN01_HEADERS.length).fill("");
    r[0] = i;
    TN01_ROWS.push(r);
  }
  switchTN01Tab('sheet');
}

async function deleteLayoutTN01(lid) {
  if (!confirm("Delete format?")) return;
  await api(`/api/tn01/layouts/${lid}`, {method: 'DELETE'});
  renderTN01FormatTab();
}

// STANDARD MODULE (PROD, GCH0, AUTO-CREATED MODULES)
async function renderStandardModule(code) {
  const layouts = await api(`/api/tcode/${code}/layouts`);
  const activeLayout = layouts.find(l => l.is_default) || layouts[0];
  const fields = activeLayout ? activeLayout.fields : [];
  const recs = await api(`/api/tcode/${code}/records`);

  $('#main_screen').innerHTML = `
    <div class="flex-1 flex flex-col bg-slate-900 overflow-hidden select-text">
      <div class="p-4 border-b border-slate-800 flex items-center justify-between shrink-0">
        <div class="flex items-center gap-3">
          <span class="font-mono font-bold text-xs bg-indigo-500/10 text-indigo-400 border border-indigo-500/30 px-2 py-0.5 rounded-lg">${code}</span>
          <h2 class="text-base font-bold text-white">${activeLayout ? esc(activeLayout.title) : code}</h2>
          <div class="h-5 w-[1px] bg-slate-800 mx-2"></div>
          <button onclick="showStandardRecordModal('${code}')" class="btn-primary text-xs font-semibold px-3 py-1.5 rounded-xl flex items-center gap-1.5"><i class="ri-add-line"></i> New Entry</button>
        </div>
        <div class="flex items-center gap-2">
          <a href="/api/tcode/${code}/export" class="btn-secondary text-xs font-semibold px-2.5 py-1.5 rounded-xl flex items-center gap-1"><i class="ri-download-2-line"></i> Export ALV</a>
          <button onclick="renderStandardModule('${code}')" class="btn-secondary text-xs font-semibold p-2 rounded-xl"><i class="ri-refresh-line"></i></button>
        </div>
      </div>
      <div class="flex-1 bg-slate-950 border-t border-slate-800 overflow-auto">
        <table class="w-full text-left custom-table">
          <thead><tr><th style="width:60px;">Doc #</th><th style="width:90px;">Facility</th>${fields.map(f=>`<th>${esc(f.label)}</th>`).join('')}<th style="width:130px;">Logged By</th><th style="width:140px;">Timestamp</th><th style="width:80px;">Action</th></tr></thead>
          <tbody class="text-slate-300">
            ${!recs.length ? `<tr><td colspan="${fields.length + 5}" class="text-center py-10 text-slate-500">No active records found.</td></tr>` : ''}
            ${recs.map(r => `
              <tr class="border-b border-slate-800/60">
                <td class="font-mono text-slate-500 font-semibold">${r.id}</td>
                <td><span class="font-mono text-xs bg-slate-800 px-2 py-0.5 rounded text-indigo-300">${esc(r.plant || '1000')}</span></td>
                ${fields.map(f=>`<td>${esc(r.data[f.key]||'')}</td>`).join('')}
                <td class="text-xs text-slate-400">${esc(r.created_by_name || r.created_by)}</td>
                <td class="text-xs text-slate-500 font-mono">${esc(r.created_at)}</td>
                <td><button onclick="deleteStandardRecord('${code}', ${r.id})" class="text-xs text-red-400 hover:text-red-300">Del</button></td>
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
        <h3 class="font-bold text-slate-800 text-sm">Post Entry: ${code}</h3>
        <button onclick="closeDialog()"><i class="ri-close-line text-lg"></i></button>
      </div>
      <div class="p-6 space-y-4">
        <div><label class="block text-xs font-semibold text-slate-600 mb-1">Facility Code</label><input id="std_f__plant" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs font-mono" value="${esc(myPlant)}"></div>
        <div class="grid grid-cols-2 gap-4">
          ${fields.map(f => `
            <div class="col-span-${f.width||1}">
              <label class="block text-xs font-semibold text-slate-600 mb-1">${esc(f.label)}</label>
              ${f.type==='textarea' ? `<textarea id="std_f_${f.key}" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs" rows="2"></textarea>` : `<input type="${f.type==='number'?'number':(f.type==='date'?'date':'text')}" id="std_f_${f.key}" class="w-full border border-slate-300 rounded-xl p-2.5 text-xs">`}
            </div>
          `).join('')}
        </div>
      </div>
      <div class="px-6 py-4 bg-slate-50 border-t border-slate-200 flex justify-end gap-2 rounded-b-2xl">
        <button onclick="closeDialog()" class="btn-secondary px-4 py-2 rounded-xl text-xs font-semibold">Cancel</button>
        <button onclick="submitStandardRecord('${code}')" class="btn-primary px-4 py-2 rounded-xl text-xs font-semibold">Save Entry</button>
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

function openDialog(content) {
  $('#app_dialog_body').innerHTML = content;
  $('#app_dialog').classList.remove('hidden');
}

function closeDialog() {
  $('#app_dialog').classList.add('hidden');
  $('#app_dialog_body').innerHTML = '';
}

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
    return {"status": "ACTIVE", "system": "NEXUS-CORE", "version": "3.5.0"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
