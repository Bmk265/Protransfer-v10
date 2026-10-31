import io, json, sqlite3, secrets, datetime, os
from pathlib import Path
import bcrypt
from fastapi import FastAPI, Request, Response, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook, Workbook

DATABASE_URL = os.environ.get("DATABASE_URL")
USE_PG = bool(DATABASE_URL and DATABASE_URL.startswith("postgres"))
if USE_PG:
    import psycopg2
    import psycopg2.extras
    if DATABASE_URL.startswith("postgres://"):
        DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

DB_PATH = Path(__file__).parent / "protransfer.db"

def get_conn():
    if USE_PG:
        return psycopg2.connect(DATABASE_URL, sslmode='require')
    else:
        con = sqlite3.connect(DB_PATH)
        con.row_factory = sqlite3.Row
        return con

def q(sql):
    return sql.replace("?", "%s") if USE_PG else sql

def fetch_all(sql, params=()):
    conn = get_conn()
    try:
        if USE_PG:
            cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            cur.execute(q(sql), params)
            return cur.fetchall()
        else:
            return [dict(r) for r in conn.execute(sql, params).fetchall()]
    finally:
        conn.close()

def fetch_one(sql, params=()):
    conn = get_conn()
    try:
        if USE_PG:
            cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            cur.execute(q(sql), params)
            return cur.fetchone()
        else:
            r = conn.execute(sql, params).fetchone()
            return dict(r) if r else None
    finally:
        conn.close()

def exec_write(sql, params=()):
    conn = get_conn()
    try:
        if USE_PG:
            cur = conn.cursor()
            cur.execute(q(sql), params)
        else:
            conn.execute(sql, params)
        conn.commit()
    finally:
        conn.close()

def init_db():
    conn = get_conn()
    cur = conn.cursor()
    if USE_PG:
        cur.execute('''
        CREATE TABLE IF NOT EXISTS users(username TEXT PRIMARY KEY,password_hash TEXT,full_name TEXT,role TEXT,department TEXT,sub_department TEXT,plant TEXT,status TEXT DEFAULT 'pending',created_at TEXT,updated_at TEXT);
        CREATE TABLE IF NOT EXISTS sessions(token TEXT PRIMARY KEY,username TEXT,expires TEXT);
        CREATE TABLE IF NOT EXISTS departments(name TEXT PRIMARY KEY,active INT DEFAULT 1,created_at TEXT);
        CREATE TABLE IF NOT EXISTS plants(name TEXT PRIMARY KEY,active INT DEFAULT 1,created_at TEXT);
        CREATE TABLE IF NOT EXISTS org(id INT PRIMARY KEY,name TEXT);
        CREATE TABLE IF NOT EXISTS tcodes(code TEXT PRIMARY KEY,title TEXT,description TEXT,color TEXT,icon TEXT,allowed_depts TEXT,active INT DEFAULT 1,created_at TEXT,updated_at TEXT,created_by TEXT);
        CREATE TABLE IF NOT EXISTS tcode_schemas(code TEXT PRIMARY KEY,fields TEXT,layout TEXT,updated_at TEXT,updated_by TEXT);
        CREATE TABLE IF NOT EXISTS tcode_records(id SERIAL PRIMARY KEY,tcode_code TEXT,data TEXT,status TEXT DEFAULT 'active',created_by TEXT,created_by_name TEXT,created_at TEXT,updated_at TEXT);
        CREATE TABLE IF NOT EXISTS tcode_layouts(id SERIAL PRIMARY KEY,tcode_code TEXT,layout_code TEXT UNIQUE,name TEXT,title TEXT,is_default INT DEFAULT 0,fields TEXT,layout_config TEXT,description TEXT,created_by TEXT,created_by_name TEXT,created_at TEXT,updated_at TEXT);
        CREATE TABLE IF NOT EXISTS access_grants(id SERIAL PRIMARY KEY,username TEXT,tcode_code TEXT,layout_code TEXT,granted_by TEXT,granted_at TEXT,UNIQUE(username,tcode_code,layout_code));
        ''')
    else:
        cur.executescript('''
        CREATE TABLE IF NOT EXISTS users(username TEXT PRIMARY KEY,password_hash TEXT,full_name TEXT,role TEXT,department TEXT,sub_department TEXT,plant TEXT,status TEXT DEFAULT 'pending',created_at TEXT,updated_at TEXT);
        CREATE TABLE IF NOT EXISTS sessions(token TEXT PRIMARY KEY,username TEXT,expires TEXT);
        CREATE TABLE IF NOT EXISTS departments(name TEXT PRIMARY KEY,active INT DEFAULT 1,created_at TEXT);
        CREATE TABLE IF NOT EXISTS plants(name TEXT PRIMARY KEY,active INT DEFAULT 1,created_at TEXT);
        CREATE TABLE IF NOT EXISTS org(id INT PRIMARY KEY,name TEXT DEFAULT 'TCL Pvt Ltd');
        CREATE TABLE IF NOT EXISTS tcodes(code TEXT PRIMARY KEY,title TEXT,description TEXT,color TEXT,icon TEXT,allowed_depts TEXT,active INT DEFAULT 1,created_at TEXT,updated_at TEXT,created_by TEXT);
        CREATE TABLE IF NOT EXISTS tcode_schemas(code TEXT PRIMARY KEY,fields TEXT,layout TEXT,updated_at TEXT,updated_by TEXT);
        CREATE TABLE IF NOT EXISTS tcode_records(id INTEGER PRIMARY KEY AUTOINCREMENT,tcode_code TEXT,data TEXT,status TEXT DEFAULT 'active',created_by TEXT,created_by_name TEXT,created_at TEXT,updated_at TEXT);
        CREATE TABLE IF NOT EXISTS tcode_layouts(id INTEGER PRIMARY KEY AUTOINCREMENT,tcode_code TEXT,layout_code TEXT UNIQUE,name TEXT,title TEXT,is_default INT DEFAULT 0,fields TEXT,layout_config TEXT,description TEXT,created_by TEXT,created_by_name TEXT,created_at TEXT,updated_at TEXT);
        CREATE TABLE IF NOT EXISTS access_grants(id INTEGER PRIMARY KEY AUTOINCREMENT,username TEXT,tcode_code TEXT,layout_code TEXT,granted_by TEXT,granted_at TEXT,UNIQUE(username,tcode_code,layout_code));
        ''')
    conn.commit()
    now_ = datetime.datetime.now().isoformat()

    def safe(sql, vals):
        try:
            if USE_PG:
                sql2 = sql.replace("INSERT OR IGNORE INTO", "INSERT INTO").replace("?", "%s")
                if "ON CONFLICT" not in sql2:
                    if "departments" in sql2:
                        sql2 += " ON CONFLICT (name) DO NOTHING"
                    elif "plants" in sql2:
                        sql2 += " ON CONFLICT (name) DO NOTHING"
                    elif "org" in sql2:
                        sql2 += " ON CONFLICT (id) DO NOTHING"
                    elif "tcodes" in sql2:
                        sql2 += " ON CONFLICT (code) DO NOTHING"
                    elif "users" in sql2:
                        sql2 += " ON CONFLICT (username) DO NOTHING"
                cur.execute(sql2, vals)
            else:
                cur.execute(sql, vals)
            conn.commit()
        except Exception as e:
            print("safe error", e)

    safe("INSERT OR IGNORE INTO departments VALUES(?,?,?)", ("Production", 1, now_))
    safe("INSERT OR IGNORE INTO departments VALUES(?,?,?)", ("Quality Control", 1, now_))
    safe("INSERT OR IGNORE INTO departments VALUES(?,?,?)", ("Final QC", 1, now_))
    safe("INSERT OR IGNORE INTO departments VALUES(?,?,?)", ("Commercial", 1, now_))
    safe("INSERT OR IGNORE INTO departments VALUES(?,?,?)", ("Planning", 1, now_))
    safe("INSERT OR IGNORE INTO departments VALUES(?,?,?)", ("Store", 1, now_))
    safe("INSERT OR IGNORE INTO plants VALUES(?,?,?)", ("Gooty", 1, now_))
    safe("INSERT OR IGNORE INTO plants VALUES(?,?,?)", ("Plant 1400", 1, now_))
    safe("INSERT OR IGNORE INTO plants VALUES(?,?,?)", ("HYD", 1, now_))
    safe("INSERT OR IGNORE INTO plants VALUES(?,?,?)", ("TCL Corporate", 1, now_))
    safe("INSERT OR IGNORE INTO org VALUES(?,?)", (1, "Thermocables Limited"))
    safe("INSERT OR IGNORE INTO tcodes VALUES(?,?,?,?,?,?,?,?,?,?)", ("GCH0", "Green Card Records", "Green card inspection", "#22c55e", "ri-checkbox-circle-line", "Quality Control,Production", 1, now_, now_, "system"))
    safe("INSERT OR IGNORE INTO tcodes VALUES(?,?,?,?,?,?,?,?,?,?)", ("NCR0", "Non Conformance Records", "Quality issues NCR", "#ef4444", "ri-error-warning-line", "Quality Control,Final QC,Production", 1, now_, now_, "system"))
    safe("INSERT OR IGNORE INTO tcodes VALUES(?,?,?,?,?,?,?,?,?,?)", ("PLN0", "Planning Details", "Master planning & Scheduling", "#06b6d4", "ri-calendar-check-line", "Planning,Production,Commercial,Quality Control,Final QC,Store", 1, now_, now_, "system"))
    safe("INSERT OR IGNORE INTO tcodes VALUES(?,?,?,?,?,?,?,?,?,?)", ("PROD", "Production Handover", "Outer Sheathing Handover", "#3b82f6", "ri-send-plane-line", "Production,Quality Control,Final QC", 1, now_, now_, "system"))
    safe("INSERT OR IGNORE INTO tcodes VALUES(?,?,?,?,?,?,?,?,?,?)", ("QCH0", "QC Handover", "QC to Final QC", "#f59e0b", "ri-shield-check-line", "Quality Control,Final QC", 1, now_, now_, "system"))
    safe("INSERT OR IGNORE INTO tcodes VALUES(?,?,?,?,?,?,?,?,?,?)", ("STD0", "Store Transfer", "Store material transfer", "#8b5cf6", "ri-store-3-line", "Store,Planning,Production", 1, now_, now_, "system"))

    ph = bcrypt.hashpw("Admin@123".encode(), bcrypt.gensalt()).decode()
    safe("INSERT OR IGNORE INTO users VALUES(?,?,?,?,?,?,?,?,?,?)", ("admin", ph, "System Admin", "master", "Planning", "", "Gooty", "approved", now_, now_))
    conn.close()

init_db()

def now():
    return datetime.datetime.now().isoformat()

def get_user(req: Request):
    t = req.cookies.get("pt")
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
        raise HTTPException(401, "login")
    return u

def is_master(u):
    return u["username"] == "admin" or u["role"] == "master"

def need_admin(req: Request):
    u = need(req)
    if u["role"] not in ["admin", "master"]:
        raise HTTPException(403, "admin required")
    return u

def need_master(req: Request):
    u = need(req)
    if not is_master(u):
        raise HTTPException(403, "MASTER ADMIN ONLY")
    return u

app = FastAPI()
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

@app.post("/api/login")
def api_login(res: Response, username: str = Form(...), password: str = Form(...)):
    username = username.lower().strip()
    u = fetch_one("SELECT * FROM users WHERE username=?", (username,))
    if not u or not bcrypt.checkpw(password.encode(), u["password_hash"].encode()):
        raise HTTPException(401, "invalid credentials")
    if u["status"] != "approved":
        raise HTTPException(403, "not approved")
    tok = secrets.token_urlsafe(32)
    exp = (datetime.datetime.now() + datetime.timedelta(days=30)).isoformat()
    exec_write("INSERT INTO sessions VALUES(?,?,?)", (tok, username, exp))
    res.set_cookie("pt", tok, httponly=True, samesite="lax", max_age=2592000)
    return {"ok": True}

@app.post("/api/logout")
def api_logout(res: Response, req: Request):
    exec_write("DELETE FROM sessions WHERE token=?", (req.cookies.get("pt"),))
    res.delete_cookie("pt")
    return {"ok": True}

@app.get("/api/me")
def api_me(req: Request):
    return need(req)

@app.get("/api/org")
def get_org(req: Request):
    need(req)
    return fetch_one("SELECT * FROM org WHERE id=?", (1,))

@app.post("/api/org")
def set_org(req: Request, name: str = Form(...)):
    need_master(req)
    exec_write("UPDATE org SET name=? WHERE id=1", (name,))
    return {"ok": True}

@app.get("/api/admin/users")
def admin_list_users(req: Request):
    need_admin(req)
    return fetch_all("SELECT * FROM users ORDER BY created_at DESC")

@app.post("/api/admin/create-user")
def admin_create_user(req: Request, username: str = Form(...), password: str = Form(...), full_name: str = Form(...), department: str = Form(...), plant: str = Form(...), role: str = Form(...)):
    need_admin(req)
    caller = need(req)
    if role == "master" and not is_master(caller):
        raise HTTPException(403, "Only MASTER can create master")
    username = username.lower().strip()
    if fetch_one("SELECT 1 FROM users WHERE username=?", (username,)):
        raise HTTPException(400, "User exists")
    ph = bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()
    exec_write("INSERT INTO users VALUES(?,?,?,?,?,?,?,?,?,?)", (username, ph, full_name, role, department, "", plant, "approved", now(), now()))
    return {"ok": True}

@app.post("/api/admin/update-user")
def admin_update_user(req: Request, username: str = Form(...), full_name: str = Form(...), department: str = Form(...), plant: str = Form(...), role: str = Form(...), status: str = Form(...)):
    caller = need_admin(req)
    target = fetch_one("SELECT * FROM users WHERE username=?", (username.lower().strip(),))
    if not target:
        raise HTTPException(404, "not found")
    if is_master(target) and not is_master(caller):
        raise HTTPException(403, "Cannot edit MASTER")
    if role == "master" and not is_master(caller):
        raise HTTPException(403, "Only MASTER can set master role")
    exec_write("UPDATE users SET full_name=?,department=?,plant=?,role=?,status=?,updated_at=? WHERE username=?", (full_name, department, plant, role, status, now(), username.lower().strip()))
    return {"ok": True}

@app.post("/api/admin/delete-user")
def admin_delete_user(req: Request, username: str = Form(...)):
    caller = need_admin(req)
    username = username.lower().strip()
    target = fetch_one("SELECT * FROM users WHERE username=?", (username,))
    if not target:
        raise HTTPException(404, "not found")
    if is_master(target):
        raise HTTPException(400, "Cannot delete MASTER ADMIN - Protected")
    if target["role"] == "admin" and not is_master(caller):
        raise HTTPException(403, "Only MASTER can delete admins")
    if username == caller["username"]:
        raise HTTPException(400, "Cannot delete self")
    exec_write("DELETE FROM users WHERE username=?", (username,))
    exec_write("DELETE FROM sessions WHERE username=?", (username,))
    exec_write("DELETE FROM access_grants WHERE username=?", (username,))
    return {"ok": True}

@app.post("/api/admin/reset-password")
def admin_reset_password(req: Request, username: str = Form(...), new_password: str = Form(...)):
    need_admin(req)
    ph = bcrypt.hashpw(new_password.encode(), bcrypt.gensalt()).decode()
    exec_write("UPDATE users SET password_hash=? WHERE username=?", (ph, username.lower().strip()))
    return {"ok": True}

@app.get("/api/departments")
def depts(req: Request):
    need(req)
    return fetch_all("SELECT * FROM departments ORDER BY name")

@app.get("/api/admin/departments")
def admin_depts(req: Request):
    need_admin(req)
    return fetch_all("SELECT * FROM departments ORDER BY name")

@app.post("/api/admin/departments/create")
def create_dept(req: Request, name: str = Form(...)):
    need_master(req)
    exec_write("INSERT INTO departments(name,active,created_at) VALUES(?,?,?) ON CONFLICT (name) DO NOTHING" if USE_PG else "INSERT OR REPLACE INTO departments VALUES(?,?,?)", (name.strip(), 1, now()))
    return {"ok": True}

@app.post("/api/admin/departments/delete")
def delete_dept(req: Request, name: str = Form(...)):
    need_master(req)
    exec_write("DELETE FROM departments WHERE name=?", (name,))
    return {"ok": True}

@app.post("/api/admin/departments/toggle")
def toggle_dept(req: Request, name: str = Form(...)):
    need_master(req)
    r = fetch_one("SELECT active FROM departments WHERE name=?", (name,))
    if r:
        exec_write("UPDATE departments SET active=? WHERE name=?", (0 if r["active"] == 1 else 1, name))
    return {"ok": True}

@app.get("/api/plants")
def plants(req: Request):
    need(req)
    return fetch_all("SELECT * FROM plants ORDER BY name")

@app.get("/api/admin/plants")
def admin_plants(req: Request):
    need_admin(req)
    return fetch_all("SELECT * FROM plants ORDER BY name")

@app.post("/api/admin/plants/create")
def create_plant(req: Request, name: str = Form(...)):
    need_master(req)
    exec_write("INSERT INTO plants(name,active,created_at) VALUES(?,?,?) ON CONFLICT (name) DO NOTHING" if USE_PG else "INSERT OR REPLACE INTO plants VALUES(?,?,?)", (name.strip(), 1, now()))
    return {"ok": True}

@app.post("/api/admin/plants/delete")
def delete_plant(req: Request, name: str = Form(...)):
    need_master(req)
    exec_write("DELETE FROM plants WHERE name=?", (name,))
    return {"ok": True}

@app.post("/api/admin/plants/toggle")
def toggle_plant(req: Request, name: str = Form(...)):
    need_master(req)
    r = fetch_one("SELECT active FROM plants WHERE name=?", (name,))
    if r:
        exec_write("UPDATE plants SET active=? WHERE name=?", (0 if r["active"] == 1 else 1, name))
    return {"ok": True}

@app.get("/api/tcodes")
def list_tcodes(req: Request):
    u = need(req)
    all_c = fetch_all("SELECT * FROM tcodes WHERE active=1 ORDER BY code")
    for tc in all_c:
        try:
            tc["layout_count"] = fetch_one("SELECT COUNT(*) as c FROM tcode_layouts WHERE tcode_code=?", (tc["code"],))["c"]
        except:
            tc["layout_count"] = 0
    if is_master(u) or u["role"] == "admin":
        return all_c
    grants = fetch_all("SELECT tcode_code FROM access_grants WHERE username=?", (u["username"],))
    granted = set(g["tcode_code"] for g in grants)
    return [tc for tc in all_c if u["department"] in [x.strip() for x in tc["allowed_depts"].split(",")] or tc["code"] in granted]

@app.get("/api/tcodes/all")
def list_tcodes_all(req: Request):
    need_admin(req)
    all_c = fetch_all("SELECT * FROM tcodes ORDER BY code")
    for tc in all_c:
        try:
            cnt = fetch_one("SELECT COUNT(*) as c FROM tcode_layouts WHERE tcode_code=?", (tc["code"],))
            tc["layout_count"] = cnt["c"] if cnt else 0
        except:
            tc["layout_count"] = 0
    return all_c

@app.post("/api/admin/tcodes/create")
def create_tcode(req: Request, code: str = Form(...), title: str = Form(...), description: str = Form(...), color: str = Form(...), icon: str = Form(...), allowed_depts: str = Form(...)):
    need_admin(req)
    code = code.upper().strip()
    if fetch_one("SELECT 1 FROM tcodes WHERE code=?", (code,)):
        raise HTTPException(400, "T-Code exists")
    exec_write("INSERT INTO tcodes VALUES(?,?,?,?,?,?,?,?,?,?)", (code, title, description, color or "#3b82f6", icon or "ri-apps-line", allowed_depts, 1, now(), now(), need(req)["username"]))
    default_fields = json.dumps([
        {"key": "job_no", "label": "Job No / SAP No", "type": "text", "required": True, "width": 1},
        {"key": "customer", "label": "Customer", "type": "text", "width": 1},
        {"key": "cable_size", "label": "Cable Size", "type": "text", "width": 1},
        {"key": "remarks", "label": "Remarks", "type": "textarea", "width": 3}
    ])
    exec_write("INSERT INTO tcode_schemas VALUES(?,?,?,?,?)", (code, default_fields, json.dumps({"columns": 3}), now(), need(req)["username"]))
    exec_write("INSERT INTO tcode_layouts(tcode_code,layout_code,name,title,is_default,fields,layout_config,description,created_by,created_by_name,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)", (code, f"{code}-001", "Default", f"Default {title} View", 1, default_fields, json.dumps({"columns": 3}), f"Default for {code}", need(req)["username"], need(req)["full_name"], now(), now()))
    return {"ok": True}

@app.post("/api/admin/tcodes/update")
def update_tcode(req: Request, code: str = Form(...), title: str = Form(...), description: str = Form(...), color: str = Form(...), icon: str = Form(...), allowed_depts: str = Form(...), active: int = Form(...)):
    need_admin(req)
    exec_write("UPDATE tcodes SET title=?,description=?,color=?,icon=?,allowed_depts=?,active=?,updated_at=? WHERE code=?", (title, description, color, icon, allowed_depts, active, now(), code.upper().strip()))
    return {"ok": True}

@app.post("/api/admin/tcodes/delete")
def delete_tcode(req: Request, code: str = Form(...)):
    u = need(req)
    code = code.upper().strip()
    if not is_master(u):
        raise HTTPException(403, "Only MASTER ADMIN can delete T-Codes")
    exec_write("DELETE FROM tcodes WHERE code=?", (code,))
    exec_write("DELETE FROM tcode_schemas WHERE code=?", (code,))
    exec_write("DELETE FROM tcode_layouts WHERE tcode_code=?", (code,))
    exec_write("DELETE FROM tcode_records WHERE tcode_code=?", (code,))
    return {"ok": True}

@app.get("/api/tcode/{code}/layouts")
def list_layouts(code: str, req: Request):
    need(req)
    rows = fetch_all("SELECT * FROM tcode_layouts WHERE tcode_code=? ORDER BY is_default DESC, layout_code ASC", (code,))
    for r in rows:
        try:
            r["fields"] = json.loads(r["fields"])
        except:
            r["fields"] = []
        try:
            r["layout_config"] = json.loads(r["layout_config"]) if isinstance(r["layout_config"], str) else (r["layout_config"] or {})
        except:
            r["layout_config"] = {}
    return rows

@app.post("/api/tcode/{code}/layouts")
async def create_layout(code: str, req: Request):
    u = need(req)
    b = await req.json()
    layout_code = b.get("layout_code", "").upper().strip()
    if not layout_code:
        cnt = fetch_one("SELECT COUNT(*) as c FROM tcode_layouts WHERE tcode_code=?", (code,))["c"]
        layout_code = f"{code}-{cnt+1:03d}"
    exec_write("INSERT INTO tcode_layouts(tcode_code,layout_code,name,title,is_default,fields,layout_config,description,created_by,created_by_name,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)", (code, layout_code, b.get("name", ""), b.get("title", ""), 0, json.dumps(b["fields"]), json.dumps(b.get("layout_config", {"columns": 3})), b.get("description", ""), u["username"], u["full_name"], now(), now()))
    return {"ok": True, "layout_code": layout_code}

@app.delete("/api/tcode/{code}/layouts/{lid}")
def del_layout(code: str, lid: int, req: Request):
    u = need(req)
    r = fetch_one("SELECT * FROM tcode_layouts WHERE id=?", (lid,))
    if not r:
        raise HTTPException(404, "not found")
    if r["is_default"] == 1:
        raise HTTPException(400, "Cannot delete default layout")
    if r["created_by"] != u["username"] and not is_master(u):
        raise HTTPException(403, "not owner")
    exec_write("DELETE FROM tcode_layouts WHERE id=?", (lid,))
    return {"ok": True}

@app.post("/api/tcode/{code}/layouts/{lid}/set-default")
def set_default_layout(code: str, lid: int, req: Request):
    need_admin(req)
    exec_write("UPDATE tcode_layouts SET is_default=0 WHERE tcode_code=?", (code,))
    exec_write("UPDATE tcode_layouts SET is_default=1,updated_at=? WHERE id=?", (now(), lid))
    return {"ok": True}

@app.get("/api/tcode/{code}/schema")
def get_schema(code: str, req: Request):
    need(req)
    s = fetch_one("SELECT * FROM tcode_schemas WHERE code=?", (code,))
    if not s:
        return {"fields": [], "layout": {"columns": 3}}
    try:
        fields = json.loads(s["fields"])
    except:
        fields = []
    try:
        layout = json.loads(s["layout"]) if isinstance(s["layout"], str) else s["layout"]
    except:
        layout = {"columns": 3}
    return {"fields": fields, "layout": layout}

@app.post("/api/tcode/{code}/schema")
async def save_schema(code: str, req: Request):
    need_admin(req)
    b = await req.json()
    if USE_PG:
        exec_write("INSERT INTO tcode_schemas VALUES(?,?,?,?,?) ON CONFLICT (code) DO UPDATE SET fields=?,layout=?,updated_at=?,updated_by=?", (code, json.dumps(b["fields"]), json.dumps(b.get("layout", {"columns": 3})), now(), need(req)["username"], json.dumps(b["fields"]), json.dumps(b.get("layout", {"columns": 3})), now(), need(req)["username"]))
    else:
        exec_write("INSERT OR REPLACE INTO tcode_schemas VALUES(?,?,?,?,?)", (code, json.dumps(b["fields"]), json.dumps(b.get("layout", {"columns": 3})), now(), need(req)["username"]))
    return {"ok": True}

@app.get("/api/tcode/{code}/records")
def get_records(code: str, req: Request, layout_code: str = ""):
    need(req)
    rows = fetch_all("SELECT * FROM tcode_records WHERE tcode_code=? AND status='active' ORDER BY id DESC", (code,))
    for r in rows:
        try:
            r["data"] = json.loads(r["data"])
        except:
            r["data"] = {}
    return rows

@app.post("/api/tcode/{code}/records")
async def create_record(code: str, req: Request):
    u = need(req)
    b = await req.json()
    if USE_PG:
        exec_write("INSERT INTO tcode_records(tcode_code,data,created_by,created_by_name,created_at,updated_at) VALUES(%s,%s,%s,%s,%s,%s)", (code, json.dumps(b), u["username"], u["full_name"], now(), now()))
    else:
        exec_write("INSERT INTO tcode_records(tcode_code,data,created_by,created_by_name,created_at,updated_at) VALUES(?,?,?,?,?,?)", (code, json.dumps(b), u["username"], u["full_name"], now(), now()))
    return {"ok": True}

@app.put("/api/tcode/{code}/records/{rid}")
async def update_record(code: str, rid: int, req: Request):
    u = need(req)
    b = await req.json()
    rec = fetch_one("SELECT * FROM tcode_records WHERE id=?", (rid,))
    if not rec:
        raise HTTPException(404, "not found")
    if rec["created_by"] != u["username"] and not is_master(u) and u["role"] != "admin":
        raise HTTPException(403, "not owner")
    exec_write("UPDATE tcode_records SET data=?,updated_at=? WHERE id=?", (json.dumps(b), now(), rid))
    return {"ok": True}

@app.delete("/api/tcode/{code}/records/{rid}")
def delete_record(code: str, rid: int, req: Request):
    u = need(req)
    rec = fetch_one("SELECT * FROM tcode_records WHERE id=?", (rid,))
    if not rec:
        raise HTTPException(404, "not found")
    if rec["created_by"] != u["username"] and not is_master(u):
        raise HTTPException(403, "not owner")
    exec_write("UPDATE tcode_records SET status='deleted',updated_at=? WHERE id=?", (now(), rid))
    return {"ok": True}

@app.post("/api/tcode/{code}/import")
async def import_excel(code: str, req: Request, file: UploadFile = File(...)):
    need(req)
    data = await file.read()
    wb = load_workbook(io.BytesIO(data))
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return {"ok": True, "imported": 0}
    headers = [str(h).strip() if h else "" for h in rows[0]]
    count = 0
    for row in rows[1:]:
        d = {}
        for i, h in enumerate(headers):
            if h:
                d[h] = row[i] if i < len(row) and row[i] is not None else ""
        u = need(req)
        if USE_PG:
            exec_write("INSERT INTO tcode_records(tcode_code,data,created_by,created_by_name,created_at,updated_at) VALUES(%s,%s,%s,%s,%s,%s)", (code, json.dumps(d), u["username"], u["full_name"], now(), now()))
        else:
            exec_write("INSERT INTO tcode_records(tcode_code,data,created_by,created_by_name,created_at,updated_at) VALUES(?,?,?,?,?,?)", (code, json.dumps(d), u["username"], u["full_name"], now(), now()))
        count += 1
    return {"ok": True, "imported": count}

@app.get("/api/tcode/{code}/export")
def export_excel(code: str, req: Request):
    need(req)
    rows = fetch_all("SELECT * FROM tcode_records WHERE tcode_code=? AND status='active' ORDER BY id", (code,))
    wb = Workbook()
    ws = wb.active
    keys = set()
    parsed = []
    for r in rows:
        try:
            d = json.loads(r["data"]) if isinstance(r["data"], str) else r["data"]
        except:
            d = {}
        parsed.append(d)
        keys.update(d.keys())
    keys = sorted(list(keys))
    if not keys:
        ws.append(["No Data"])
    else:
        ws.append(keys)
        for d in parsed:
            ws.append([d.get(k, "") for k in keys])
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return StreamingResponse(bio, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f"attachment; filename={code}_export.xlsx"})

@app.get("/api/admin/access")
def list_access(req: Request):
    need_admin(req)
    return fetch_all("SELECT * FROM access_grants ORDER BY granted_at DESC")

@app.post("/api/admin/access/grant")
def grant_access(req: Request, username: str = Form(...), tcode_code: str = Form(...), layout_code: str = Form(...)):
    need_admin(req)
    if USE_PG:
        exec_write("INSERT INTO access_grants(username,tcode_code,layout_code,granted_by,granted_at) VALUES(?,?,?,?,?) ON CONFLICT (username,tcode_code,layout_code) DO NOTHING", (username.lower().strip(), tcode_code.upper().strip(), layout_code.upper().strip(), need(req)["username"], now()))
    else:
        exec_write("INSERT OR IGNORE INTO access_grants(username,tcode_code,layout_code,granted_by,granted_at) VALUES(?,?,?,?,?)", (username.lower().strip(), tcode_code.upper().strip(), layout_code.upper().strip(), need(req)["username"], now()))
    return {"ok": True}

@app.post("/api/admin/access/revoke")
def revoke_access(req: Request, id: int = Form(...)):
    need_admin(req)
    exec_write("DELETE FROM access_grants WHERE id=?", (id,))
    return {"ok": True}

# HTML Frontend - SECURE LOGIN (NO CREDENTIALS SHOWN) + COMPLETE FEATURES
HTML = """
<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>ProTransfer V15 COMPLETE - Secure</title>
<script src="https://cdn.tailwindcss.com"></script><link href="https://cdn.jsdelivr.net/npm/remixicon@4.2.0/fonts/remixicon.css" rel="stylesheet">
</head><body class="bg-[#0f172a] min-h-screen"><div id="root"></div>
<script>
let ME=null, TCODE=null, LAYOUTS=[], CUR_LAYOUT=null;
const $=s=>document.querySelector(s);
async function api(path, opts={}){ const r=await fetch(path, {...opts, credentials:'include'}); if(r.status==401){ auth(); throw 'login'; } if(!r.ok){ let t=await r.text(); throw t; } try{ return await r.json(); } catch{ return {}; } }
function auth(){
 root.innerHTML=`<div class="min-h-screen flex items-center justify-center p-4 bg-[#0f172a]"><div class="bg-white rounded-2xl p-8 w-full max-w-md shadow-2xl"><div class="flex items-center gap-3 mb-6"><div class="w-10 h-10 bg-black rounded-xl flex items-center justify-center text-white"><i class="ri-shield-keyhole-line text-xl"></i></div><div><h2 class="text-xl font-bold">Secure Login</h2><div class="text-[11px] text-gray-500">Thermocables Limited - Authorized Access Only</div></div></div><div><input id="lu" placeholder="Username" class="w-full border p-3 rounded-xl mb-3"><input id="lp" type="password" placeholder="Password" class="w-full border p-3 rounded-xl mb-4"><button onclick="login()" class="w-full bg-black text-white py-3 rounded-xl font-semibold">Login</button><div class="mt-4 text-[11px] text-gray-400 bg-gray-50 p-3 rounded-xl">Contact Admin for account. No self-registration.</div></div></div></div>`;
}
async function login(){
 const u=$("#lu").value.trim(), p=$("#lp").value.trim(); if(!u||!p) return alert("Enter both");
 const fd=new FormData(); fd.append("username",u); fd.append("password",p);
 try{ await fetch("/api/login",{method:"POST",body:fd,credentials:"include"}).then(async r=>{ if(!r.ok) throw await r.text(); }); location.reload(); } catch(e){ alert("Login failed: "+e); }
}
async function load(){
 try{ ME=await api("/api/me"); } catch{ return auth(); }
 root.innerHTML=`<div class="min-h-screen bg-[#f8fafc]"><header class="bg-white border-b px-6 py-3 flex items-center justify-between sticky top-0 z-10"><div class="flex items-center gap-4"><div class="font-bold text-lg"><i class="ri-shield-check-line"></i> ProTransfer V15 COMPLETE</div><span class="text-xs bg-black text-white px-2 py-1 rounded-full">${ME.role.toUpperCase()}</span><span class="text-xs text-gray-500">${ME.username} - ${ME.department} - ${ME.plant}</span></div><div class="flex items-center gap-2"><button onclick="showAdmin()" class="text-xs bg-gray-900 text-white px-3 py-2 rounded-xl ${ME.role!='admin' && ME.role!='master' ? 'hidden':''}"><i class="ri-settings-3-line"></i> Admin Panel - Master</button><button onclick="logout()" class="text-xs border px-3 py-2 rounded-xl">Logout</button></div></header><div class="p-6"><div id="tcodes" class="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-4 gap-4"></div><div id="main" class="mt-6"></div></div></div>`;
 loadTcodes();
}
async function loadTcodes(){
 const list=await api("/api/tcodes");
 $("#tcodes").innerHTML=list.map(t=>`<div onclick="openTcode('${t.code}')" class="bg-white rounded-2xl p-5 border hover:shadow-lg cursor-pointer"><div class="flex items-center gap-3"><div class="w-10 h-10 rounded-xl flex items-center justify-center text-white" style="background:${t.color}"><i class="${t.icon}"></i></div><div><div class="font-bold">${t.code}</div><div class="text-xs text-gray-500">${t.title}</div></div></div><div class="mt-3 text-[11px] text-gray-400">${t.description||''} • ${t.layout_count||0} layouts • ${t.allowed_depts}</div></div>`).join("") || "<div class='text-gray-400'>No T-Codes assigned</div>";
}
async function openTcode(code){
 TCODE=code; LAYOUTS=await api(`/api/tcode/${code}/layouts`);
 $("#main").innerHTML=`<div class="bg-white rounded-2xl border p-5"><div class="flex items-center justify-between"><h2 class="font-bold text-lg">${code} - Layouts</h2><button onclick="newLayout()" class="text-xs bg-black text-white px-3 py-2 rounded-xl">+ New Layout</button></div><div class="mt-4 grid grid-cols-1 md:grid-cols-3 gap-3">${LAYOUTS.map(l=>`<div class="border rounded-xl p-4 ${l.is_default?'bg-gray-900 text-white':''}"><div class="font-bold text-sm">${l.layout_code} ${l.is_default?'<span class="text-[10px] bg-white text-black px-2 py-0.5 rounded-full">DEFAULT</span>':''}</div><div class="text-xs mt-1">${l.title}</div><div class="text-[11px] mt-1 text-gray-400">${l.description||''}</div><div class="mt-3 flex gap-2"><button onclick="openLayout('${l.layout_code}')" class="text-[11px] ${l.is_default?'bg-white text-black':'bg-black text-white'} px-3 py-1.5 rounded-full">Open</button><button onclick="setDefault(${l.id})" class="text-[11px] border px-3 py-1.5 rounded-full">Set Default</button><button onclick="deleteLayout(${l.id})" class="text-[11px] text-red-500">Del</button></div></div>`).join("")}</div><div id="records" class="mt-6"></div></div>`;
 if(LAYOUTS.length) openLayout(LAYOUTS[0].layout_code);
}
async function openLayout(lc){
 CUR_LAYOUT=LAYOUTS.find(x=>x.layout_code==lc); if(!CUR_LAYOUT) return;
 const sch=await api(`/api/tcode/${TCODE}/schema`);
 const recs=await api(`/api/tcode/${TCODE}/records?layout_code=${lc}`);
 let fields=CUR_LAYOUT.fields||sch.fields;
 $("#records").innerHTML=`<div class="border-t pt-5"><div class="flex items-center justify-between"><h3 class="font-bold">${CUR_LAYOUT.title} - Records (${recs.length})</h3><div class="flex gap-2"><button onclick="showAdd()" class="text-xs bg-black text-white px-3 py-2 rounded-xl">+ Add Record</button><a href="/api/tcode/${TCODE}/export" class="text-xs border px-3 py-2 rounded-xl">Export Excel</a><label class="text-xs border px-3 py-2 rounded-xl cursor-pointer">Import Excel<input type="file" class="hidden" onchange="importFile(event)"></label></div></div><div class="mt-4 overflow-x-auto"><table class="w-full text-xs"><thead><tr class="bg-gray-50">${fields.map(f=>`<th class="text-left p-2 border">${f.label}</th>`).join("")}<th class="p-2 border">By</th><th class="p-2 border">Actions</th></tr></thead><tbody>${recs.map(r=>`<tr>${fields.map(f=>`<td class="p-2 border">${(r.data[f.key]||'')}</td>`).join("")}<td class="p-2 border">${r.created_by_name||r.created_by}</td><td class="p-2 border"><button onclick="editRec(${r.id})" class="text-blue-600">Edit</button> <button onclick="delRec(${r.id})" class="text-red-600 ml-2">Del</button></td></tr>`).join("")}</tbody></table></div></div>
 <div id="formModal" class="hidden fixed inset-0 bg-black/50 flex items-center justify-center p-4 z-20"><div class="bg-white rounded-2xl p-6 w-full max-w-3xl max-h-[90vh] overflow-auto"><h3 class="font-bold mb-4" id="formTitle">Add Record</h3><div id="formFields" class="grid grid-cols-3 gap-3"></div><div class="mt-4 flex gap-2"><button onclick="saveRec()" class="bg-black text-white px-4 py-2 rounded-xl text-sm">Save</button><button onclick="$('#formModal').classList.add('hidden')" class="border px-4 py-2 rounded-xl text-sm">Cancel</button></div></div></div>`;
 window._fields=fields; window._editingId=null;
}
function showAdd(){ window._editingId=null; $("#formTitle").innerText="Add Record"; renderForm({}); $("#formModal").classList.remove("hidden"); }
function editRec(id){ fetch(`/api/tcode/${TCODE}/records`,{credentials:"include"}).then(r=>r.json()).then(list=>{ const rec=list.find(x=>x.id==id); if(!rec) return; window._editingId=id; $("#formTitle").innerText="Edit Record"; renderForm(rec.data); $("#formModal").classList.remove("hidden"); }); }
function renderForm(data){ $("#formFields").innerHTML=window._fields.map(f=>`<div class="col-span-${f.width||1}"><label class="text-[11px] text-gray-500">${f.label}${f.required?'*':''}</label><${f.type=='textarea'?'textarea':'input'} type="${f.type=='number'?'number':'text'}" id="f_${f.key}" class="w-full border rounded-xl p-2.5 text-sm mt-1" ${f.type=='textarea'?'rows=3':''}>${data[f.key]||''}</${f.type=='textarea'?'textarea':'input'}></div>`).join(""); }
async function saveRec(){ const obj={}; window._fields.forEach(f=>{ const el=$("#f_"+f.key); obj[f.key]=el.value; }); const method=window._editingId?"PUT":"POST"; const url=window._editingId?`/api/tcode/${TCODE}/records/${window._editingId}`:`/api/tcode/${TCODE}/records`; await fetch(url,{method,credentials:"include",headers:{"Content-Type":"application/json"},body:JSON.stringify(obj)}); $("#formModal").classList.add("hidden"); openLayout(CUR_LAYOUT.layout_code); }
async function delRec(id){ if(!confirm("Delete?")) return; await fetch(`/api/tcode/${TCODE}/records/${id}`,{method:"DELETE",credentials:"include"}); openLayout(CUR_LAYOUT.layout_code); }
async function importFile(e){ const fd=new FormData(); fd.append("file",e.target.files[0]); await fetch(`/api/tcode/${TCODE}/import`,{method:"POST",body:fd,credentials:"include"}); openLayout(CUR_LAYOUT.layout_code); }
async function newLayout(){ const name=prompt("Layout name?"); if(!name) return; const title=prompt("Title?")||name; await fetch(`/api/tcode/${TCODE}/layouts`,{method:"POST",credentials:"include",headers:{"Content-Type":"application/json"},body:JSON.stringify({name,title,fields:window._fields||[],layout_config:{columns:3}})}); openTcode(TCODE); }
async function deleteLayout(id){ if(!confirm("Delete layout?")) return; await fetch(`/api/tcode/${TCODE}/layouts/${id}`,{method:"DELETE",credentials:"include"}); openTcode(TCODE); }
async function setDefault(id){ await fetch(`/api/tcode/${TCODE}/layouts/${id}/set-default`,{method:"POST",credentials:"include"}); openTcode(TCODE); }

async function showAdmin(){
 const users=await api("/api/admin/users"); const depts=await api("/api/departments"); const plants=await api("/api/plants"); const allT=await api("/api/tcodes/all");
 $("#main").innerHTML=`<div class="bg-white rounded-2xl border p-6"><h2 class="font-bold text-lg">Admin Panel - MASTER - COMPLETE</h2>
 <div class="mt-6 grid grid-cols-1 lg:grid-cols-2 gap-6">
 <div><h3 class="font-bold text-sm mb-3">Create User</h3><div class="space-y-2"><input id="cu_u" placeholder="Username" class="w-full border p-2 rounded-xl text-sm"><input id="cu_p" placeholder="Password" class="w-full border p-2 rounded-xl text-sm"><input id="cu_fn" placeholder="Full Name" class="w-full border p-2 rounded-xl text-sm"><select id="cu_d" class="w-full border p-2 rounded-xl text-sm">${depts.map(d=>`<option>${d.name}</option>`).join("")}</select><select id="cu_pl" class="w-full border p-2 rounded-xl text-sm">${plants.map(p=>`<option>${p.name}</option>`).join("")}</select><select id="cu_r" class="w-full border p-2 rounded-xl text-sm"><option value="user">User</option><option value="admin">Admin</option>${ME.username=='admin'?'<option value="master">Master</option>':''}</select><button onclick="createUser()" class="w-full bg-black text-white py-2 rounded-xl text-sm">Create User</button></div></div>
 <div><h3 class="font-bold text-sm mb-3">Plants & Departments (MASTER ONLY) - Add/Delete</h3><div class="flex gap-2 mb-2"><input id="new_plant" placeholder="New Plant" class="border p-2 rounded-xl text-sm flex-1"><button onclick="addPlant()" class="bg-black text-white px-3 py-2 rounded-xl text-xs">+ Add Plant</button></div><div class="flex flex-wrap gap-2 mb-4">${plants.map(p=>`<span class="text-xs border px-2 py-1 rounded-full">${p.name} <button onclick="delPlant('${p.name}')" class="text-red-500 ml-1">x</button></span>`).join("")}</div><div class="flex gap-2 mb-2"><input id="new_dept" placeholder="New Dept" class="border p-2 rounded-xl text-sm flex-1"><button onclick="addDept()" class="bg-black text-white px-3 py-2 rounded-xl text-xs">+ Add Dept</button></div><div class="flex flex-wrap gap-2">${depts.map(d=>`<span class="text-xs border px-2 py-1 rounded-full">${d.name} <button onclick="delDept('${d.name}')" class="text-red-500 ml-1">x</button></span>`).join("")}</div></div>
 </div>
 <div class="mt-6"><h3 class="font-bold text-sm mb-3">Users (${users.length}) - Edit / Delete / Modify</h3><div class="overflow-x-auto"><table class="w-full text-xs"><thead><tr class="bg-gray-50"><th class="p-2 text-left">User</th><th class="p-2 text-left">Name</th><th class="p-2">Dept</th><th class="p-2">Plant</th><th class="p-2">Role</th><th class="p-2">Status</th><th class="p-2">Actions</th></tr></thead><tbody>${users.map(u=>`<tr class="border-t"><td class="p-2">${u.username}</td><td class="p-2">${u.full_name}</td><td class="p-2">${u.department}</td><td class="p-2">${u.plant}</td><td class="p-2">${u.role}</td><td class="p-2">${u.status}</td><td class="p-2 flex gap-1"><button onclick="editUser('${u.username}')" class="bg-blue-600 text-white px-2 py-1 rounded-full">Edit</button><button onclick="resetPass('${u.username}')" class="bg-yellow-600 text-white px-2 py-1 rounded-full">Reset PW</button><button onclick="deleteUser('${u.username}')" class="bg-red-600 text-white px-2 py-1 rounded-full">Del</button></td></tr>`).join("")}</tbody></table></div></div>
 <div class="mt-6"><h3 class="font-bold text-sm mb-3">T-Codes - Add / Edit / Delete (MASTER can Delete) - 6 T-Codes</h3><div class="grid grid-cols-6 gap-2 mb-3"><input id="tc_code" placeholder="CODE" class="border p-2 rounded-xl text-xs"><input id="tc_title" placeholder="Title" class="border p-2 rounded-xl text-xs col-span-2"><input id="tc_desc" placeholder="Desc" class="border p-2 rounded-xl text-xs col-span-2"><button onclick="addTcode()" class="bg-black text-white rounded-xl text-xs">+ New T-Code</button></div><div class="overflow-x-auto"><table class="w-full text-xs"><thead><tr class="bg-gray-50"><th class="p-2 text-left">Code</th><th class="p-2 text-left">Title</th><th class="p-2">Depts</th><th class="p-2">Color</th><th class="p-2">Actions</th></tr></thead><tbody>${allT.map(t=>`<tr class="border-t"><td class="p-2 font-bold">${t.code}</td><td class="p-2">${t.title}</td><td class="p-2">${t.allowed_depts}</td><td class="p-2"><div class="w-4 h-4 rounded" style="background:${t.color}"></div></td><td class="p-2"><button onclick="editTcode('${t.code}')" class="text-blue-600">Edit</button> <button onclick="delTcode('${t.code}')" class="text-red-600 ml-2">Del</button></td></tr>`).join("")}</tbody></table></div></div>
 </div>`;
}
async function createUser(){ const fd=new FormData(); fd.append("username",$("#cu_u").value); fd.append("password",$("#cu_p").value); fd.append("full_name",$("#cu_fn").value); fd.append("department",$("#cu_d").value); fd.append("plant",$("#cu_pl").value); fd.append("role",$("#cu_r").value); const r=await fetch("/api/admin/create-user",{method:"POST",body:fd,credentials:"include"}); if(!r.ok) alert(await r.text()); else showAdmin(); }
async function deleteUser(u){ if(!confirm("Delete "+u+"?")) return; const fd=new FormData(); fd.append("username",u); const r=await fetch("/api/admin/delete-user",{method:"POST",body:fd,credentials:"include"}); if(!r.ok) alert(await r.text()); else showAdmin(); }
async function editUser(u){ const fn=prompt("New Full Name?"); if(fn===null) return; const dept=prompt("Department?"); if(dept===null) return; const plant=prompt("Plant?"); if(plant===null) return; const role=prompt("Role (user/admin/master)?"); if(role===null) return; const status=prompt("Status (approved/pending)?")||"approved"; const fd=new FormData(); fd.append("username",u); fd.append("full_name",fn||u); fd.append("department",dept||""); fd.append("plant",plant||""); fd.append("role",role||"user"); fd.append("status",status); await fetch("/api/admin/update-user",{method:"POST",body:fd,credentials:"include"}); showAdmin(); }
async function resetPass(u){ const np=prompt("New password for "+u+"?"); if(!np) return; const fd=new FormData(); fd.append("username",u); fd.append("new_password",np); await fetch("/api/admin/reset-password",{method:"POST",body:fd,credentials:"include"}); alert("Password reset"); }
async function addPlant(){ const n=$("#new_plant").value.trim(); if(!n) return; const fd=new FormData(); fd.append("name",n); await fetch("/api/admin/plants/create",{method:"POST",body:fd,credentials:"include"}); showAdmin(); }
async function delPlant(n){ if(!confirm("Delete plant "+n+"?")) return; const fd=new FormData(); fd.append("name",n); await fetch("/api/admin/plants/delete",{method:"POST",body:fd,credentials:"include"}); showAdmin(); }
async function addDept(){ const n=$("#new_dept").value.trim(); if(!n) return; const fd=new FormData(); fd.append("name",n); await fetch("/api/admin/departments/create",{method:"POST",body:fd,credentials:"include"}); showAdmin(); }
async function delDept(n){ if(!confirm("Delete dept "+n+"?")) return; const fd=new FormData(); fd.append("name",n); await fetch("/api/admin/departments/delete",{method:"POST",body:fd,credentials:"include"}); showAdmin(); }
async function addTcode(){ const fd=new FormData(); fd.append("code",$("#tc_code").value); fd.append("title",$("#tc_title").value); fd.append("description",$("#tc_desc").value); fd.append("color","#3b82f6"); fd.append("icon","ri-apps-line"); fd.append("allowed_depts","Production,Quality Control"); const r=await fetch("/api/admin/tcodes/create",{method:"POST",body:fd,credentials:"include"}); if(!r.ok) alert(await r.text()); else showAdmin(); }
async function editTcode(code){ const title=prompt("New title?"); if(!title) return; const desc=prompt("Description?")||""; const depts=prompt("Allowed departments comma separated?")||"Production"; const fd=new FormData(); fd.append("code",code); fd.append("title",title); fd.append("description",desc); fd.append("color","#3b82f6"); fd.append("icon","ri-apps-line"); fd.append("allowed_depts",depts); fd.append("active",1); await fetch("/api/admin/tcodes/update",{method:"POST",body:fd,credentials:"include"}); showAdmin(); }
async function delTcode(code){ if(!confirm("MASTER DELETE T-Code "+code+"? All records will be deleted!")) return; const fd=new FormData(); fd.append("code",code); const r=await fetch("/api/admin/tcodes/delete",{method:"POST",body:fd,credentials:"include"}); if(!r.ok) alert(await r.text()); else showAdmin(); }
async function logout(){ await fetch("/api/logout",{method:"POST",credentials:"include"}); location.reload(); }
load();
</script></body></html>
"""

@app.get("/", response_class=HTMLResponse)
def home():
    return HTMLResponse(HTML)

@app.get("/health")
def health():
    return {"ok": True, "version": "V15 COMPLETE SECURE"}
