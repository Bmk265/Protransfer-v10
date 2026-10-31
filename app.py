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
def q(sql): return sql.replace("?", "%s") if USE_PG else sql
def fetch_all(sql, params=()):
    conn=get_conn()
    try:
        if USE_PG:
            cur=conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            cur.execute(q(sql), params); return cur.fetchall()
        else: return [dict(r) for r in conn.execute(sql, params).fetchall()]
    finally: conn.close()
def fetch_one(sql, params=()):
    conn=get_conn()
    try:
        if USE_PG:
            cur=conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
            cur.execute(q(sql), params); return cur.fetchone()
        else:
            r=conn.execute(sql, params).fetchone()
            return dict(r) if r else None
    finally: conn.close()
def exec_write(sql, params=()):
    conn=get_conn()
    try:
        if USE_PG: cur=conn.cursor(); cur.execute(q(sql), params)
        else: conn.execute(sql, params)
        conn.commit()
    finally: conn.close()
def init_db():
    conn=get_conn(); cur=conn.cursor()
    if USE_PG: cur.execute("CREATE TABLE IF NOT EXISTS users(username TEXT PRIMARY KEY,password_hash TEXT,full_name TEXT,role TEXT,department TEXT,sub_department TEXT,plant TEXT,status TEXT DEFAULT 'pending',created_at TEXT,updated_at TEXT);CREATE TABLE IF NOT EXISTS sessions(token TEXT PRIMARY KEY,username TEXT,expires TEXT);CREATE TABLE IF NOT EXISTS departments(name TEXT PRIMARY KEY,active INT DEFAULT 1,created_at TEXT);CREATE TABLE IF NOT EXISTS plants(name TEXT PRIMARY KEY,active INT DEFAULT 1,created_at TEXT);CREATE TABLE IF NOT EXISTS org(id INT PRIMARY KEY,name TEXT);CREATE TABLE IF NOT EXISTS tcodes(code TEXT PRIMARY KEY,title TEXT,description TEXT,color TEXT,icon TEXT,allowed_depts TEXT,active INT DEFAULT 1,created_at TEXT,updated_at TEXT,created_by TEXT);CREATE TABLE IF NOT EXISTS tcode_schemas(code TEXT PRIMARY KEY,fields TEXT,layout TEXT,updated_at TEXT,updated_by TEXT);CREATE TABLE IF NOT EXISTS tcode_records(id SERIAL PRIMARY KEY,tcode_code TEXT,data TEXT,status TEXT DEFAULT 'active',created_by TEXT,created_by_name TEXT,created_at TEXT,updated_at TEXT);CREATE TABLE IF NOT EXISTS tcode_layouts(id SERIAL PRIMARY KEY,tcode_code TEXT,layout_code TEXT UNIQUE,name TEXT,title TEXT,is_default INT DEFAULT 0,fields TEXT,layout_config TEXT,description TEXT,created_by TEXT,created_by_name TEXT,created_at TEXT,updated_at TEXT);CREATE TABLE IF NOT EXISTS access_grants(id SERIAL PRIMARY KEY,username TEXT,tcode_code TEXT,layout_code TEXT,granted_by TEXT,granted_at TEXT,UNIQUE(username,tcode_code,layout_code));")
    else: cur.executescript("CREATE TABLE IF NOT EXISTS users(username TEXT PRIMARY KEY,password_hash TEXT,full_name TEXT,role TEXT,department TEXT,sub_department TEXT,plant TEXT,status TEXT DEFAULT 'pending',created_at TEXT,updated_at TEXT);CREATE TABLE IF NOT EXISTS sessions(token TEXT PRIMARY KEY,username TEXT,expires TEXT);CREATE TABLE IF NOT EXISTS departments(name TEXT PRIMARY KEY,active INT DEFAULT 1,created_at TEXT);CREATE TABLE IF NOT EXISTS plants(name TEXT PRIMARY KEY,active INT DEFAULT 1,created_at TEXT);CREATE TABLE IF NOT EXISTS org(id INT PRIMARY KEY,name TEXT DEFAULT 'TCL Pvt Ltd');CREATE TABLE IF NOT EXISTS tcodes(code TEXT PRIMARY KEY,title TEXT,description TEXT,color TEXT,icon TEXT,allowed_depts TEXT,active INT DEFAULT 1,created_at TEXT,updated_at TEXT,created_by TEXT);CREATE TABLE IF NOT EXISTS tcode_schemas(code TEXT PRIMARY KEY,fields TEXT,layout TEXT,updated_at TEXT,updated_by TEXT);CREATE TABLE IF NOT EXISTS tcode_records(id INTEGER PRIMARY KEY AUTOINCREMENT,tcode_code TEXT,data TEXT,status TEXT DEFAULT 'active',created_by TEXT,created_by_name TEXT,created_at TEXT,updated_at TEXT);CREATE TABLE IF NOT EXISTS tcode_layouts(id INTEGER PRIMARY KEY AUTOINCREMENT,tcode_code TEXT,layout_code TEXT UNIQUE,name TEXT,title TEXT,is_default INT DEFAULT 0,fields TEXT,layout_config TEXT,description TEXT,created_by TEXT,created_by_name TEXT,created_at TEXT,updated_at TEXT);CREATE TABLE IF NOT EXISTS access_grants(id INTEGER PRIMARY KEY AUTOINCREMENT,username TEXT,tcode_code TEXT,layout_code TEXT,granted_by TEXT,granted_at TEXT,UNIQUE(username,tcode_code,layout_code));")
    conn.commit(); now=datetime.datetime.now().isoformat()
    def safe(sql, vals):
        try:
            if USE_PG:
                sql2=sql.replace("INSERT OR IGNORE INTO","INSERT INTO").replace("?","%s")
                if "ON CONFLICT" not in sql2:
                    if "departments" in sql2: sql2+=" ON CONFLICT (name) DO NOTHING"
                    elif "plants" in sql2: sql2+=" ON CONFLICT (name) DO NOTHING"
                    elif "org" in sql2: sql2+=" ON CONFLICT (id) DO NOTHING"
                    elif "tcodes" in sql2: sql2+=" ON CONFLICT (code) DO NOTHING"
                    elif "users" in sql2: sql2+=" ON CONFLICT (username) DO NOTHING"
                cur.execute(sql2, vals)
            else: cur.execute(sql, vals)
            conn.commit()
        except: pass
    safe("INSERT OR IGNORE INTO departments VALUES(?,?,?)",("Production",1,now))
    safe("INSERT OR IGNORE INTO departments VALUES(?,?,?)",("Quality Control",1,now))
    safe("INSERT OR IGNORE INTO departments VALUES(?,?,?)",("Final QC",1,now))
    safe("INSERT OR IGNORE INTO departments VALUES(?,?,?)",("Commercial",1,now))
    safe("INSERT OR IGNORE INTO departments VALUES(?,?,?)",("Planning",1,now))
    safe("INSERT OR IGNORE INTO departments VALUES(?,?,?)",("Store",1,now))
    safe("INSERT OR IGNORE INTO plants VALUES(?,?,?)",("Gooty",1,now))
    safe("INSERT OR IGNORE INTO plants VALUES(?,?,?)",("Plant 1400",1,now))
    safe("INSERT OR IGNORE INTO plants VALUES(?,?,?)",("HYD",1,now))
    safe("INSERT OR IGNORE INTO plants VALUES(?,?,?)",("TCL Corporate",1,now))
    safe("INSERT OR IGNORE INTO org VALUES(?,?)",(1,"Thermocables Limited"))
    safe("INSERT OR IGNORE INTO tcodes VALUES(?,?,?,?,?,?,?,?,?,?)",("GCH0","Green Card Records","Green card","#22c55e","ri-checkbox-circle-line","Quality Control,Production",1,now,now,"system"))
    safe("INSERT OR IGNORE INTO tcodes VALUES(?,?,?,?,?,?,?,?,?,?)",("NCR0","Non Conformance Records","Quality issues","#ef4444","ri-error-warning-line","Quality Control,Final QC,Production",1,now,now,"system"))
    safe("INSERT OR IGNORE INTO tcodes VALUES(?,?,?,?,?,?,?,?,?,?)",("PLN0","Planning Details","Master planning","#06b6d4","ri-calendar-check-line","Planning,Production,Commercial,Quality Control,Final QC,Store",1,now,now,"system"))
    safe("INSERT OR IGNORE INTO tcodes VALUES(?,?,?,?,?,?,?,?,?,?)",("PROD","Production Handover","Outer Sheathing","#3b82f6","ri-send-plane-line","Production,Quality Control,Final QC",1,now,now,"system"))
    ph=bcrypt.hashpw("Admin@123".encode(),bcrypt.gensalt()).decode()
    safe("INSERT OR IGNORE INTO users VALUES(?,?,?,?,?,?,?,?,?,?)",("admin",ph,"System Admin","master","Planning","","Gooty","approved",now,now))
    conn.close()
init_db()
def now(): return datetime.datetime.now().isoformat()
def get_user(req:Request):
    t=req.cookies.get("pt")
    if not t: return None
    s=fetch_one("SELECT * FROM sessions WHERE token=?",(t,))
    if not s: return None
    u=fetch_one("SELECT * FROM users WHERE username=?",(s["username"] if isinstance(s, dict) else s["username"],))
    if not u or u["status"]!="approved": return None
    return u
def need(req:Request):
    u=get_user(req)
    if not u: raise HTTPException(401,"login")
    return u
def is_master(u): return u["username"]=="admin" or u["role"]=="master"
def need_admin(req:Request):
    u=need(req)
    if u["role"] not in ["admin","master"]: raise HTTPException(403,"admin")
    return u
def need_master(req:Request):
    u=need(req)
    if not is_master(u): raise HTTPException(403,"MASTER ADMIN ONLY")
    return u
app=FastAPI()
app.add_middleware(CORSMiddleware,allow_origins=["*"],allow_methods=["*"],allow_headers=["*"])
@app.post("/api/login")
def login(res:Response,username:str=Form(...),password:str=Form(...)):
    username=username.lower().strip()
    u=fetch_one("SELECT * FROM users WHERE username=?",(username,))
    if not u or not bcrypt.checkpw(password.encode(),u["password_hash"].encode()): raise HTTPException(401,"invalid")
    if u["status"]!="approved": raise HTTPException(403,"not approved")
    tok=secrets.token_urlsafe(32); exp=(datetime.datetime.now()+datetime.timedelta(days=30)).isoformat()
    exec_write("INSERT INTO sessions VALUES(?,?,?)",(tok,username,exp))
    res.set_cookie("pt",tok,httponly=True,samesite="lax",max_age=2592000)
    return {"ok":True}
@app.post("/api/logout")
def logout(res:Response,req:Request):
    exec_write("DELETE FROM sessions WHERE token=?",(req.cookies.get("pt"),))
    res.delete_cookie("pt"); return {"ok":True}
@app.get("/api/me")
def me(req:Request): return need(req)
@app.get("/api/org")
def get_org(req:Request): need(req); return fetch_one("SELECT * FROM org WHERE id=?",(1,))
@app.post("/api/org")
def set_org(req:Request,name:str=Form(...)): need_master(req); exec_write("UPDATE org SET name=? WHERE id=1",(name,)); return {"ok":True}
@app.get("/api/admin/users")
def admin_list_users(req:Request): need_admin(req); return fetch_all("SELECT * FROM users ORDER BY created_at DESC")
@app.post("/api/admin/create-user")
def admin_create_user(req:Request,username:str=Form(...),password:str=Form(...),full_name:str=Form(...),department:str=Form(...),plant:str=Form(...),role:str=Form(...)):
    need_admin(req); caller=need(req)
    if role=="master" and not is_master(caller): raise HTTPException(403,"Only MASTER can create master")
    username=username.lower().strip()
    if fetch_one("SELECT 1 FROM users WHERE username=?",(username,)): raise HTTPException(400,"User exists")
    ph=bcrypt.hashpw(password.encode(),bcrypt.gensalt()).decode()
    exec_write("INSERT INTO users VALUES(?,?,?,?,?,?,?,?,?,?)",(username,ph,full_name,role,department,"",plant,"approved",now(),now())); return {"ok":True}
@app.post("/api/admin/update-user")
def admin_update_user(req:Request,username:str=Form(...),full_name:str=Form(...),department:str=Form(...),plant:str=Form(...),role:str=Form(...),status:str=Form(...)):
    caller=need_admin(req); target=fetch_one("SELECT * FROM users WHERE username=?",(username.lower().strip(),))
    if not target: raise HTTPException(404,"not found")
    if is_master(target) and not is_master(caller): raise HTTPException(403,"Cannot edit MASTER")
    if role=="master" and not is_master(caller): raise HTTPException(403,"Only MASTER can set master role")
    exec_write("UPDATE users SET full_name=?,department=?,plant=?,role=?,status=?,updated_at=? WHERE username=?",(full_name,department,plant,role,status,now(),username.lower().strip())); return {"ok":True}
@app.post("/api/admin/delete-user")
def admin_delete_user(req:Request,username:str=Form(...)):
    caller=need_admin(req); username=username.lower().strip()
    target=fetch_one("SELECT * FROM users WHERE username=?",(username,))
    if not target: raise HTTPException(404,"not found")
    if is_master(target): raise HTTPException(400,"Cannot delete MASTER ADMIN - Protected")
    if target["role"]=="admin" and not is_master(caller): raise HTTPException(403,"Only MASTER can delete admins")
    if username==caller["username"]: raise HTTPException(400,"Cannot delete self")
    exec_write("DELETE FROM users WHERE username=?",(username,)); exec_write("DELETE FROM sessions WHERE username=?",(username,)); exec_write("DELETE FROM access_grants WHERE username=?",(username,)); return {"ok":True}
@app.post("/api/admin/reset-password")
def admin_reset_password(req:Request,username:str=Form(...),new_password:str=Form(...)):
    need_admin(req); ph=bcrypt.hashpw(new_password.encode(),bcrypt.gensalt()).decode(); exec_write("UPDATE users SET password_hash=? WHERE username=?",(ph,username.lower().strip())); return {"ok":True}
@app.get("/api/departments")
def depts(req:Request): need(req); return fetch_all("SELECT * FROM departments ORDER BY name")
@app.get("/api/admin/departments")
def admin_depts(req:Request): need_admin(req); return fetch_all("SELECT * FROM departments ORDER BY name")
@app.post("/api/admin/departments/create")
def create_dept(req:Request,name:str=Form(...)): need_master(req); exec_write("INSERT INTO departments(name,active,created_at) VALUES(?,?,?) ON CONFLICT (name) DO NOTHING" if USE_PG else "INSERT OR REPLACE INTO departments VALUES(?,?,?)",(name.strip(),1,now())); return {"ok":True}
@app.post("/api/admin/departments/delete")
def delete_dept(req:Request,name:str=Form(...)): need_master(req); exec_write("DELETE FROM departments WHERE name=?",(name,)); return {"ok":True}
@app.post("/api/admin/departments/toggle")
def toggle_dept(req:Request,name:str=Form(...)): need_master(req); r=fetch_one("SELECT active FROM departments WHERE name=?",(name,)); exec_write("UPDATE departments SET active=? WHERE name=?",(0 if r["active"]==1 else 1,name)); return {"ok":True}
@app.get("/api/plants")
def plants(req:Request): need(req); return fetch_all("SELECT * FROM plants ORDER BY name")
@app.get("/api/admin/plants")
def admin_plants(req:Request): need_admin(req); return fetch_all("SELECT * FROM plants ORDER BY name")
@app.post("/api/admin/plants/create")
def create_plant(req:Request,name:str=Form(...)): need_master(req); exec_write("INSERT INTO plants(name,active,created_at) VALUES(?,?,?) ON CONFLICT (name) DO NOTHING" if USE_PG else "INSERT OR REPLACE INTO plants VALUES(?,?,?)",(name.strip(),1,now())); return {"ok":True}
@app.post("/api/admin/plants/delete")
def delete_plant(req:Request,name:str=Form(...)): need_master(req); exec_write("DELETE FROM plants WHERE name=?",(name,)); return {"ok":True}
@app.post("/api/admin/plants/toggle")
def toggle_plant(req:Request,name:str=Form(...)): need_master(req); r=fetch_one("SELECT active FROM plants WHERE name=?",(name,)); exec_write("UPDATE plants SET active=? WHERE name=?",(0 if r["active"]==1 else 1,name)); return {"ok":True}
@app.get("/api/tcodes")
def list_tcodes(req:Request):
    u=need(req); all_c=fetch_all("SELECT * FROM tcodes WHERE active=1 ORDER BY code")
    for tc in all_c: 
        try: tc["layout_count"]=fetch_one("SELECT COUNT(*) as c FROM tcode_layouts WHERE tcode_code=?",(tc["code"],))["c"]
        except: tc["layout_count"]=0
    if is_master(u) or u["role"]=="admin": return all_c
    grants=fetch_all("SELECT tcode_code FROM access_grants WHERE username=?",(u["username"],)); granted=set(g["tcode_code"] for g in grants)
    return [tc for tc in all_c if u["department"] in [x.strip() for x in tc["allowed_depts"].split(",")] or tc["code"] in granted]
@app.get("/api/tcodes/all")
def list_tcodes_all(req:Request): need_admin(req); all_c=fetch_all("SELECT * FROM tcodes ORDER BY code")
    for tc in all_c: 
        try: tc["layout_count"]=fetch_one("SELECT COUNT(*) as c FROM tcode_layouts WHERE tcode_code=?",(tc["code"],))["c"]
        except: tc["layout_count"]=0
    return all_c
@app.post("/api/admin/tcodes/create")
def create_tcode(req:Request,code:str=Form(...),title:str=Form(...),description:str=Form(...),color:str=Form(...),icon:str=Form(...),allowed_depts:str=Form(...)):
    need_admin(req); code=code.upper().strip()
    if fetch_one("SELECT 1 FROM tcodes WHERE code=?",(code,)): raise HTTPException(400,"T-Code exists")
    exec_write("INSERT INTO tcodes VALUES(?,?,?,?,?,?,?,?,?,?)",(code,title,description,color or "#3b82f6",icon or "ri-apps-line",allowed_depts,1,now(),now(),need(req)["username"]))
    default_fields=json.dumps([{"key":"job_no","label":"Job No / SAP No","type":"text","required":True,"width":1},{"key":"customer","label":"Customer","type":"text","width":1},{"key":"remarks","label":"Remarks","type":"textarea","width":3}])
    exec_write("INSERT INTO tcode_schemas VALUES(?,?,?,?,?)",(code,default_fields,json.dumps({"columns":3}),now(),need(req)["username"]))
    exec_write("INSERT INTO tcode_layouts(tcode_code,layout_code,name,title,is_default,fields,layout_config,description,created_by,created_by_name,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",(code,f"{code}-001","Default",f"Default {title} View",1,default_fields,json.dumps({"columns":3}),f"Default for {code}",need(req)["username"],need(req)["full_name"],now(),now()))
    return {"ok":True}
@app.post("/api/admin/tcodes/update")
def update_tcode(req:Request,code:str=Form(...),title:str=Form(...),description:str=Form(...),color:str=Form(...),icon:str=Form(...),allowed_depts:str=Form(...),active:int=Form(...)):
    need_admin(req); exec_write("UPDATE tcodes SET title=?,description=?,color=?,icon=?,allowed_depts=?,active=?,updated_at=? WHERE code=?",(title,description,color,icon,allowed_depts,active,now(),code.upper().strip())); return {"ok":True}
@app.post("/api/admin/tcodes/delete")
def delete_tcode(req:Request,code:str=Form(...)):
    u=need(req); code=code.upper().strip()
    if not is_master(u): raise HTTPException(403,"Only MASTER ADMIN can delete T-Codes")
    exec_write("DELETE FROM tcodes WHERE code=?",(code,)); exec_write("DELETE FROM tcode_schemas WHERE code=?",(code,)); exec_write("DELETE FROM tcode_layouts WHERE tcode_code=?",(code,)); exec_write("DELETE FROM tcode_records WHERE tcode_code=?",(code,)); return {"ok":True}
@app.get("/api/tcode/{code}/layouts")
def list_layouts(code:str,req:Request):
    need(req); rows=fetch_all("SELECT * FROM tcode_layouts WHERE tcode_code=? ORDER BY is_default DESC, layout_code ASC",(code,))
    for r in rows: 
        try: r["fields"]=json.loads(r["fields"])
        except: r["fields"]=[]
        r["layout_config"]=json.loads(r["layout_config"]) if isinstance(r["layout_config"], str) else (r["layout_config"] or {})
    return rows
@app.post("/api/tcode/{code}/layouts")
async def create_layout(code:str,req:Request):
    u=need(req); b=await req.json(); layout_code=b.get("layout_code","").upper().strip()
    if not layout_code:
        cnt=fetch_one("SELECT COUNT(*) as c FROM tcode_layouts WHERE tcode_code=?",(code,))["c"]
        layout_code=f"{code}-{cnt+1:03d}"
    exec_write("INSERT INTO tcode_layouts(tcode_code,layout_code,name,title,is_default,fields,layout_config,description,created_by,created_by_name,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",(code,layout_code,b.get("name",""),b.get("title",""),0,json.dumps(b["fields"]),json.dumps(b.get("layout_config",{"columns":3})),b.get("description",""),u["username"],u["full_name"],now(),now()))
    return {"ok":True,"layout_code":layout_code}
@app.delete("/api/tcode/{code}/layouts/{lid}")
def del_layout(code:str,lid:int,req:Request):
    u=need(req); r=fetch_one("SELECT * FROM tcode_layouts WHERE id=?",(lid,))
    if not r: raise HTTPException(404,"not found")
    if r["is_default"]==1: raise HTTPException(400,"Cannot delete default")
    if r["created_by"]!=u["username"] and not is_master(u): raise HTTPException(403,"not owner")
    exec_write("DELETE FROM tcode_layouts WHERE id=?",(lid,)); return {"ok":True}
@app.post("/api/tcode/{code}/layouts/{lid}/set-default")
def set_default_layout(code:str,lid:int,req:Request):
    need_admin(req); exec_write("UPDATE tcode_layouts SET is_default=0 WHERE tcode_code=?",(code,)); exec_write("UPDATE tcode_layouts SET is_default=1,updated_at=? WHERE id=?",(now(),lid)); return {"ok":True}
@app.get("/api/tcode/{code}/schema")
def get_schema(code:str,req:Request):
    need(req); r=fetch_one("SELECT * FROM tcode_schemas WHERE code=?",(code,))
    if not r: return {"code":code,"fields":[],"layout":{"columns":3}}
    return {"code":r["code"],"fields":json.loads(r["fields"]),"layout":json.loads(r["layout"])}
@app.post("/api/tcode/{code}/schema")
async def save_schema(code:str,req:Request):
    need_admin(req); b=await req.json()
    if USE_PG:
        exec_write("INSERT INTO tcode_schemas(code,fields,layout,updated_at,updated_by) VALUES(?,?,?,?,?) ON CONFLICT (code) DO UPDATE SET fields=EXCLUDED.fields, layout=EXCLUDED.layout, updated_at=EXCLUDED.updated_at",(code,json.dumps(b["fields"]),json.dumps(b.get("layout",{"columns":3})),now(),need(req)["username"]))
        exec_write("UPDATE tcode_layouts SET fields=%s,layout_config=%s,updated_at=%s WHERE tcode_code=%s AND is_default=1",(json.dumps(b["fields"]),json.dumps(b.get("layout",{"columns":3})),now(),code))
    else:
        exec_write("INSERT OR REPLACE INTO tcode_schemas VALUES(?,?,?,?,?)",(code,json.dumps(b["fields"]),json.dumps(b.get("layout",{"columns":3})),now(),need(req)["username"]))
        exec_write("UPDATE tcode_layouts SET fields=?,layout_config=?,updated_at=? WHERE tcode_code=? AND is_default=1",(json.dumps(b["fields"]),json.dumps(b.get("layout",{"columns":3})),now(),code))
    return {"ok":True}
@app.get("/api/tcode/{code}/records")
def list_records(code:str,req:Request):
    need(req); rows=fetch_all("SELECT * FROM tcode_records WHERE tcode_code=? AND status='active' ORDER BY id DESC",(code,))
    for r in rows: r["data"]=json.loads(r["data"]) if isinstance(r["data"], str) else r["data"]
    return rows
@app.post("/api/tcode/{code}/records")
async def create_record(code:str,req:Request):
    u=need(req); b=await req.json(); exec_write("INSERT INTO tcode_records(tcode_code,data,status,created_by,created_by_name,created_at,updated_at) VALUES(?,?,?,?,?,?,?)",(code,json.dumps(b),"active",u["username"],u["full_name"],now(),now())); return {"ok":True}
@app.put("/api/tcode/{code}/records/{rid}")
async def update_record(code:str,rid:int,req:Request):
    u=need(req); b=await req.json(); r=fetch_one("SELECT * FROM tcode_records WHERE id=?",(rid,))
    if r["created_by"]!=u["username"] and not is_master(u) and u["role"]!="admin": raise HTTPException(403,"not owner")
    exec_write("UPDATE tcode_records SET data=?,updated_at=? WHERE id=?",(json.dumps(b),now(),rid)); return {"ok":True}
@app.delete("/api/tcode/{code}/records/{rid}")
def delete_record(code:str,rid:int,req:Request): need_admin(req); exec_write("UPDATE tcode_records SET status='deleted' WHERE id=?",(rid,)); return {"ok":True}
@app.get("/api/admin/access-grants")
def list_access_grants(req:Request): need_admin(req); return fetch_all("SELECT ag.*,u.full_name,u.department FROM access_grants ag LEFT JOIN users u ON ag.username=u.username ORDER BY ag.username,ag.tcode_code")
@app.post("/api/admin/grant-access")
def grant_access(req:Request,username:str=Form(...),tcode_code:str=Form(...),layout_code:str=Form(...)):
    u=need_admin(req); 
    if USE_PG:
        exec_write("INSERT INTO access_grants(username,tcode_code,layout_code,granted_by,granted_at) VALUES(?,?,?,?,?) ON CONFLICT (username,tcode_code,layout_code) DO NOTHING",(username.lower().strip(),tcode_code.upper().strip(),layout_code.upper().strip(),u["username"],now()))
    else:
        exec_write("INSERT OR REPLACE INTO access_grants(username,tcode_code,layout_code,granted_by,granted_at) VALUES(?,?,?,?,?)",(username.lower().strip(),tcode_code.upper().strip(),layout_code.upper().strip(),u["username"],now()))
    return {"ok":True}
@app.post("/api/admin/revoke-access")
def revoke_access(req:Request,grant_id:int=Form(...)): need_admin(req); exec_write("DELETE FROM access_grants WHERE id=?",(grant_id,)); return {"ok":True}
@app.post("/api/admin/upload-excel-preview")
def upload_excel_preview(req:Request,file:UploadFile=File(...)):
    need_admin(req); data=file.file.read(); wb=load_workbook(io.BytesIO(data),data_only=True); ws=wb.active; rows=list(ws.iter_rows(values_only=True))
    if not rows: raise HTTPException(400,"Empty excel")
    headers=[str(h or "").strip() for h in rows[0]]; preview=[[str(c or "") for c in r] for r in rows[1:11] if any(r)]
    return {"headers":headers,"preview":preview,"total_rows":len([r for r in rows[1:] if any(r)])}
@app.post("/api/admin/upload-excel-save-file")
def upload_excel_save_file(req:Request,file:UploadFile=File(...),tcode_code:str=Form(...),mapping:str=Form(...)):
    u=need_admin(req); mapping_dict=json.loads(mapping); col_to_field={}
    for k,v in mapping_dict.items():
        if str(k).isdigit(): col_to_field[int(k)]=v
    data=file.file.read(); wb=load_workbook(io.BytesIO(data),data_only=True); ws=wb.active; rows=list(ws.iter_rows(values_only=True)); count=0
    for r in rows[1:]:
        if not any(r): continue
        data_dict={}
        for col_idx, field_key in col_to_field.items():
            if col_idx < len(r): data_dict[field_key]=str(r[col_idx] or "").strip()
        if not data_dict: continue
        exec_write("INSERT INTO tcode_records(tcode_code,data,status,created_by,created_by_name,created_at,updated_at) VALUES(?,?,?,?,?,?,?)",(tcode_code,json.dumps(data_dict),"active",u["username"],u["full_name"],now(),now())); count+=1
    return {"ok":True,"saved":count}
@app.get("/api/tcode/{code}/export-excel")
def export_excel(code:str,req:Request):
    need(req); rows=fetch_all("SELECT * FROM tcode_records WHERE tcode_code=? AND status='active' ORDER BY id DESC",(code,)); schema=fetch_one("SELECT fields FROM tcode_schemas WHERE code=?",(code,)); fields=json.loads(schema["fields"]) if schema and schema["fields"] else []
    wb=Workbook(); ws=wb.active; ws.title=code; ws.append([f["label"] for f in fields]+["Created By","Created At"])
    for r in rows:
        data=json.loads(r["data"]) if isinstance(r["data"], str) else r["data"]
        ws.append([data.get(f["key"],"") for f in fields]+[r["created_by_name"],r["created_at"]])
    bio=io.BytesIO(); wb.save(bio); bio.seek(0)
    return StreamingResponse(bio, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f"attachment; filename={code}_export.xlsx"})
@app.get("/api/planning/lookup")
def planning_lookup(req:Request,job_no:str="",item_no:str="",sap_no:str=""):
    need(req); q=(job_no or item_no or sap_no or "").strip().lower()
    if not q: return {"found":False}
    rows=fetch_all("SELECT data FROM tcode_records WHERE tcode_code='PLN0' AND status='active'")
    for r in rows:
        data=json.loads(r["data"]) if isinstance(r["data"], str) else r["data"]
        for k,v in data.items():
            if not v: continue
            if q in str(v).strip().lower(): return {"found":True,"data":data,"matched_field":k}
    return {"found":False}
@app.get("/api/planning/all-jobs")
def all_jobs(req:Request):
    need(req); rows=fetch_all("SELECT data FROM tcode_records WHERE tcode_code='PLN0' AND status='active' LIMIT 500")
    jobs=[]
    for r in rows:
        data=json.loads(r["data"]) if isinstance(r["data"], str) else r["data"]
        job=data.get("job_no") or data.get("item_no") or data.get("sap_no") or ""
        if job: jobs.append({"job_no":job,"customer":data.get("customer",""),"material":data.get("material_description") or ""})
    return jobs[:200]
@app.get("/api/planning/suggested-fields")
def suggested_fields(req:Request):
    need_admin(req); rows=fetch_all("SELECT data FROM tcode_records WHERE tcode_code='PLN0' AND status='active' LIMIT 200")
    field_counts={}
    for r in rows:
        data=json.loads(r["data"]) if isinstance(r["data"], str) else r["data"]
        for k in data.keys(): field_counts[k]=field_counts.get(k,0)+1
    suggestions=[{"key":k,"label":k.replace("_"," ").title(),"count":c} for k,c in sorted(field_counts.items(), key=lambda x: -x[1])]
    return {"suggestions":suggestions, "total_pln_records": len(rows)}

# --- HTML BELOW - Keep as is ---
HTML_PAGE = """<!DOCTYPE html><html><head><meta charset="UTF-8"><title>V13 Master</title><script src="https://cdn.tailwindcss.com"></script><link href="https://cdnjs.cloudflare.com/ajax/libs/remixicon/4.2.0/remixicon.min.css" rel="stylesheet"></head><body><div id="root"></div><div id="toasts" class="fixed top-3 right-3 z-[9999]"></div><script>
const $=id=>document.getElementById(id);const toast=(m,t="ok")=>{let e=document.createElement("div");e.className=`px-4 py-3 rounded-xl text-white text-sm ${t=="err"?"bg-red-600":"bg-gray-900"}`;e.textContent=m;$("toasts").appendChild(e);setTimeout(()=>e.remove(),3500)};let U=null,ORG=null;async function api(p,o={}){let r=await fetch(p,{credentials:"include",...o});if(!r.ok)throw new Error(await r.text());let ct=r.headers.get("content-type");if(ct&&ct.includes("application/json"))return r.json();return r;}
async function boot(){try{U=await api("/api/me");ORG=await api("/api/org");app()}catch{auth()}}boot();
function auth(){root.innerHTML=`<div class="min-h-screen flex items-center justify-center p-4 bg-[#0f172a]"><div class="bg-white rounded-2xl p-8 w-full max-w-md"><h2 class="text-2xl font-bold">V13 MASTER ADMIN</h2><div class="mt-6"><input id="lu" placeholder="Username" class="w-full border p-3 rounded-xl mb-3"><input id="lp" type="password" placeholder="Password" class="w-full border p-3 rounded-xl mb-4"><button onclick="login()" class="w-full bg-black text-white py-3 rounded-xl">Login</button><div class="mt-3 text-[11px] bg-gray-50 p-2 rounded-xl">Default: admin / Admin@123</div></div></div></div>`}
async function login(){let f=new FormData();f.append("username",$("lu").value);f.append("password",$("lp").value);try{await api("/api/login",{method:"POST",body:f});location.reload()}catch(e){toast(e.message,"err")}}
function isMaster(){return U && (U.username=='admin' || U.role=='master')}
function shell(inner){return `<div class="min-h-screen"><div class="bg-white border-b h-[56px] flex items-center justify-center"><span class="font-bold" id="orgName">${ORG.name}</span></div><div class="bg-[#f9fafb] border-b flex flex-wrap items-center justify-between px-4 py-2 gap-2"><div><div class="font-bold">V13 Master Admin - Postgres Forever</div><div class="text-[11px] text-gray-500">Add/Delete/Modify Plants, Departments, T-Codes, Users</div></div><div class="flex gap-2 flex-wrap"><button onclick="openAdminPanel()" class="px-3 py-2 rounded-full bg-red-600 text-white text-xs font-bold">Admin Panel - Master</button><button onclick="openUserManager()" class="px-3 py-2 rounded-full bg-black text-white text-xs">Users - Add/Delete/Modify</button><button onclick="openAccessManager()" class="px-3 py-2 rounded-full bg-indigo-600 text-white text-xs">Access</button><button onclick="openExcelUpload('PLN0')" class="px-3 py-2 rounded-full bg-green-600 text-white text-xs">Excel</button><button onclick="v('dash')" class="px-3 py-2 rounded-full bg-white border text-xs">Dashboard</button><button onclick="logout()" class="px-3 py-2 rounded-full bg-white border text-xs">Logout</button></div></div><div class="max-w-[1480px] mx-auto p-5">${inner}</div></div>`}
function app(){v('dash')}
async function v(n){if(n=='dash'){let tcodes=await api("/api/tcodes");$("root").innerHTML=shell(`<div class="flex justify-between mb-4"><h2 class="font-bold">T-Codes • ${tcodes.length} - Where Add/Delete/Modify? Click Admin Panel button on top</h2></div><div class="grid grid-cols-4 gap-4">${tcodes.map(t=>`<div onclick="openTcode('${t.code}')" class="bg-white rounded-[20px] border p-5 cursor-pointer"><div class="w-12 h-12 rounded-xl grid place-items-center text-white" style="background:${t.color}"><i class="${t.icon}"></i></div><div class="mt-4"><div class="text-[11px] font-bold text-gray-500">${t.code}</div><div class="text-[14px] font-bold">${t.title}</div></div></div>`).join('')}</div>`)}}
async function openAdminPanel(){let depts=await api("/api/admin/departments");let plants=await api("/api/admin/plants");let tcodes=await api("/api/tcodes/all");let suggested=await api("/api/planning/suggested-fields");$("root").innerHTML=shell(`<div class="bg-white rounded-2xl border p-6"><div class="flex justify-between"><h2 class="font-bold text-lg">Admin Panel - Where to Add/Delete/Modify Everything</h2><button onclick="v('dash')" class="px-4 py-2 rounded-full bg-gray-100 text-sm">Back</button></div><div class="mt-6 grid grid-cols-12 gap-6"><div class="col-span-6"><div class="bg-gray-50 rounded-2xl p-5 border mb-6"><h3 class="font-bold text-sm flex justify-between"><span>Plants • ${plants.length} - Add/Delete/Modify here</span><button onclick="addPlant()" class="px-3 py-1 rounded-full bg-black text-white text-xs">+ Add Plant</button></h3><div class="mt-3 space-y-2">${plants.map(p=>`<div class="flex items-center justify-between p-2 bg-white border rounded-xl text-xs"><span>${p.name}</span><div class="flex gap-1"><button onclick="togglePlant('${p.name}')" class="px-2 py-1 rounded-full bg-gray-100">${p.active?'Disable':'Enable'}</button><button onclick="deletePlant('${p.name}')" class="px-2 py-1 rounded-full bg-red-50 text-red-600">Del</button></div></div>`).join('')}</div></div><div class="bg-gray-50 rounded-2xl p-5 border"><h3 class="font-bold text-sm flex justify-between"><span>Departments • ${depts.length}</span><button onclick="addDept()" class="px-3 py-1 rounded-full bg-black text-white text-xs">+ Add Dept</button></h3><div class="mt-3 space-y-2">${depts.map(d=>`<div class="flex items-center justify-between p-2 bg-white border rounded-xl text-xs"><span>${d.name}</span><div class="flex gap-1"><button onclick="toggleDept('${d.name}')" class="px-2 py-1 rounded-full bg-gray-100">${d.active?'Disable':'Enable'}</button><button onclick="deleteDept('${d.name}')" class="px-2 py-1 rounded-full bg-red-50 text-red-600">Del</button></div></div>`).join('')}</div></div></div><div class="col-span-6"><div class="bg-gray-50 rounded-2xl p-5 border mb-6"><h3 class="font-bold text-sm flex justify-between"><span>T-Codes • ${tcodes.length} - Add/Delete/Modify here</span><button onclick="createNewTcode()" class="px-3 py-1 rounded-full bg-red-600 text-white text-xs">+ New T-Code</button></h3><div class="mt-3 space-y-2">${tcodes.map(t=>`<div class="flex items-center justify-between p-2 bg-white border rounded-xl text-xs"><div><div class="font-bold">${t.code} - ${t.title}</div><div class="text-[10px] text-gray-500">${t.allowed_depts}</div></div><div class="flex gap-1"><button onclick="editTcode('${t.code}')" class="px-2 py-1 rounded-full bg-white border">Edit</button><button onclick="deleteTcode('${t.code}')" class="px-2 py-1 rounded-full bg-red-600 text-white">Del</button></div></div>`).join('')}</div></div><div class="bg-indigo-50 rounded-2xl p-5 border"><h3 class="font-bold text-sm">Suggested Fields from PLN0 • ${suggested.total_pln_records} records</h3><div class="mt-3 flex flex-wrap gap-2">${suggested.suggestions.map(s=>`<span class="px-2 py-1 bg-white border rounded-full text-[11px]">${s.key} (${s.count})</span>`).join('')||'No data'}</div></div></div></div></div>`);}
async function addPlant(){let n=prompt("Plant Name");if(!n)return;let f=new FormData();f.append("name",n);await api("/api/admin/plants/create",{method:"POST",body:f});openAdminPanel()}
async function deletePlant(n){if(!confirm("Delete "+n+"?"))return;let f=new FormData();f.append("name",n);await api("/api/admin/plants/delete",{method:"POST",body:f});openAdminPanel()}
async function togglePlant(n){let f=new FormData();f.append("name",n);await api("/api/admin/plants/toggle",{method:"POST",body:f});openAdminPanel()}
async function addDept(){let n=prompt("Dept Name");if(!n)return;let f=new FormData();f.append("name",n);await api("/api/admin/departments/create",{method:"POST",body:f});openAdminPanel()}
async function deleteDept(n){if(!confirm("Delete "+n+"?"))return;let f=new FormData();f.append("name",n);await api("/api/admin/departments/delete",{method:"POST",body:f});openAdminPanel()}
async function toggleDept(n){let f=new FormData();f.append("name",n);await api("/api/admin/departments/toggle",{method:"POST",body:f});openAdminPanel()}
async function createNewTcode(){let code=prompt("CODE e.g. DSP0");if(!code)return;let title=prompt("Title");if(!title)return;let desc=prompt("Description")||"";let color=prompt("Color","#3b82f6")||"#3b82f6";let icon=prompt("Icon","ri-apps-line")||"ri-apps-line";let allowed=prompt("Allowed Depts","Planning,Production")||"Planning";let f=new FormData();f.append("code",code.toUpperCase());f.append("title",title);f.append("description",desc);f.append("color",color);f.append("icon",icon);f.append("allowed_depts",allowed);await api("/api/admin/tcodes/create",{method:"POST",body:f});openAdminPanel()}
async function editTcode(code){let tcodes=await api("/api/tcodes/all");let t=tcodes.find(x=>x.code==code);let title=prompt("Title",t.title);if(title===null)return;let desc=prompt("Desc",t.description)||t.description;let color=prompt("Color",t.color)||t.color;let icon=prompt("Icon",t.icon)||t.icon;let allowed=prompt("Allowed Depts",t.allowed_depts)||t.allowed_depts;let active=prompt("Active 1/0",t.active)||"1";let f=new FormData();f.append("code",code);f.append("title",title);f.append("description",desc);f.append("color",color);f.append("icon",icon);f.append("allowed_depts",allowed);f.append("active",active);await api("/api/admin/tcodes/update",{method:"POST",body:f});openAdminPanel()}
async function deleteTcode(code){if(!confirm("Delete "+code+"?"))return;let f=new FormData();f.append("code",code);await api("/api/admin/tcodes/delete",{method:"POST",body:f});openAdminPanel()}
async function openUserManager(){let users=await api("/api/admin/users");let depts=await api("/api/departments");let plants=await api("/api/plants");$("root").innerHTML=shell(`<div class="bg-white rounded-2xl border p-6"><div class="flex justify-between"><h2 class="font-bold text-lg">Users - Add/Delete/Modify</h2><button onclick="v('dash')" class="px-4 py-2 rounded-full bg-gray-100 text-sm">Back</button></div><div class="mt-6 grid grid-cols-12 gap-6"><div class="col-span-5 bg-gray-50 rounded-2xl p-5 border"><h3 class="font-bold text-sm">Create User</h3><div class="mt-4 space-y-3"><input id="new_username" placeholder="Username" class="w-full border p-2.5 rounded-xl text-sm"><input id="new_password" placeholder="Password" class="w-full border p-2.5 rounded-xl text-sm"><input id="new_fullname" placeholder="Full Name" class="w-full border p-2.5 rounded-xl text-sm"><select id="new_dept" class="w-full border p-2.5 rounded-xl text-sm">${depts.map(d=>`<option value="${d.name}">${d.name}</option>`).join('')}</select><select id="new_plant" class="w-full border p-2.5 rounded-xl text-sm">${plants.map(p=>`<option value="${p.name}">${p.name}</option>`).join('')}</select><select id="new_role" class="w-full border p-2.5 rounded-xl text-sm"><option value="user">User</option><option value="admin">Admin</option>${isMaster()?'<option value="master">Master</option>':''}</select><button onclick="doCreateUser()" class="w-full bg-black text-white py-3 rounded-xl">+ Create</button></div><div id="editBox" class="mt-6 hidden bg-white rounded-2xl p-4 border"><h3 class="font-bold text-sm">Edit User - <span id="editUName"></span></h3><div class="mt-3 space-y-2"><input id="edit_fullname" class="w-full border p-2 rounded-xl text-sm"><select id="edit_dept" class="w-full border p-2 rounded-xl text-sm">${depts.map(d=>`<option value="${d.name}">${d.name}</option>`).join('')}</select><select id="edit_plant" class="w-full border p-2 rounded-xl text-sm">${plants.map(p=>`<option value="${p.name}">${p.name}</option>`).join('')}</select><select id="edit_role" class="w-full border p-2 rounded-xl text-sm"><option value="user">User</option><option value="admin">Admin</option>${isMaster()?'<option value="master">Master</option>':''}</select><select id="edit_status" class="w-full border p-2 rounded-xl text-sm"><option value="approved">Approved</option><option value="pending">Pending</option></select><div class="flex gap-2"><button onclick="doUpdateUser()" class="flex-1 bg-black text-white py-2 rounded-xl text-sm">Update</button><button onclick="cancelEdit()" class="px-4 py-2 rounded-xl bg-gray-100 text-sm">Cancel</button></div></div></div></div><div class="col-span-7"><h3 class="font-bold text-sm">All Users • ${users.length}</h3><div class="mt-3 border rounded-xl overflow-auto max-h-[80vh]"><table class="w-full text-xs"><thead class="bg-gray-50"><tr><th class="p-2 text-left">User</th><th class="p-2">Role</th><th class="p-2">Action</th></tr></thead><tbody>${users.map(u=>`<tr class="border-b"><td class="p-2 font-mono"><b>@${u.username}</b><div class="text-[10px]">${u.full_name}</div></td><td class="p-2"><span class="px-2 py-1 rounded-full ${u.role=='master'?'bg-red-600 text-white':u.role=='admin'?'bg-black text-white':'bg-gray-100'} text-[10px]">${u.role}</span></td><td class="p-2 flex gap-1"><button onclick="startEditUser('${u.username}')" class="px-2 py-1 rounded-full bg-indigo-50 text-indigo-600 text-[11px]">Edit/Modify</button><button onclick="deleteUser('${u.username}')" class="px-2 py-1 rounded-full bg-red-50 text-red-600 text-[11px]">Delete</button></td></tr>`).join('')}</tbody></table></div></div></div></div>`);}
let _editingUsername=null;function startEditUser(username){fetch("/api/admin/users",{credentials:"include"}).then(r=>r.json()).then(users=>{let u=users.find(x=>x.username==username);_editingUsername=username;$("editBox").classList.remove("hidden");$("editUName").innerText=username;$("edit_fullname").value=u.full_name;$("edit_dept").value=u.department;$("edit_plant").value=u.plant;$("edit_role").value=u.role;$("edit_status").value=u.status;});}
function cancelEdit(){_editingUsername=null;$("editBox").classList.add("hidden")}
async function doUpdateUser(){let f=new FormData();f.append("username",_editingUsername);f.append("full_name",$("edit_fullname").value);f.append("department",$("edit_dept").value);f.append("plant",$("edit_plant").value);f.append("role",$("edit_role").value);f.append("status",$("edit_status").value);await api("/api/admin/update-user",{method:"POST",body:f});openUserManager();}
async function doCreateUser(){let f=new FormData();f.append("username",$("new_username").value);f.append("password",$("new_password").value);f.append("full_name",$("new_fullname").value);f.append("department",$("new_dept").value);f.append("plant",$("new_plant").value);f.append("role",$("new_role").value);await api("/api/admin/create-user",{method:"POST",body:f});openUserManager();}
async function deleteUser(username){if(!confirm("Delete @"+username+"?"))return;let f=new FormData();f.append("username",username);await api("/api/admin/delete-user",{method:"POST",body:f});openUserManager();}
async function openAccessManager(){let users=await api("/api/admin/users");let tcodes=await api("/api/tcodes/all");let grants=await api("/api/admin/access-grants");$("root").innerHTML=shell(`<div class="bg-white rounded-2xl border p-6"><div class="flex justify-between"><h2 class="font-bold text-lg">Access</h2><button onclick="v('dash')" class="px-4 py-2 rounded-full bg-gray-100 text-sm">Back</button></div><div class="mt-6 grid grid-cols-12 gap-6"><div class="col-span-5 bg-gray-50 rounded-2xl p-5 border"><h3 class="font-bold text-sm">Grant</h3><div class="mt-4 space-y-3"><select id="acc_user" class="w-full border p-2.5 rounded-xl text-sm">${users.map(u=>`<option value="${u.username}">${u.username}</option>`).join('')}</select><select id="acc_tcode" class="w-full border p-2.5 rounded-xl text-sm">${tcodes.map(t=>`<option value="${t.code}">${t.code}</option>`).join('')}</select><select id="acc_layout" class="w-full border p-2.5 rounded-xl text-sm"><option value="">ALL</option></select><button onclick="doGrantAccess()" class="w-full bg-black text-white py-3 rounded-xl">+ Grant</button></div></div><div class="col-span-7"><h3 class="font-bold text-sm">Grants • ${grants.length}</h3><div class="mt-3 border rounded-xl overflow-auto max-h-[70vh]"><table class="w-full text-xs"><thead class="bg-gray-50"><tr><th class="p-2">User</th><th class="p-2">T-Code</th><th class="p-2">Action</th></tr></thead><tbody>${grants.map(g=>`<tr class="border-b"><td class="p-2">@${g.username}</td><td class="p-2">${g.tcode_code}</td><td class="p-2"><button onclick="revokeGrant(${g.id})" class="px-2 py-1 rounded-full bg-red-50 text-red-600 text-[11px]">Revoke</button></td></tr>`).join('')}</tbody></table></div></div></div></div>`);}
async function doGrantAccess(){let user=$("acc_user").value;let tcode=$("acc_tcode").value;let layout=$("acc_layout").value;let f=new FormData();f.append("username",user);f.append("tcode_code",tcode);f.append("layout_code",layout);await api("/api/admin/grant-access",{method:"POST",body:f});openAccessManager();}
async function revokeGrant(id){let f=new FormData();f.append("grant_id",id);await api("/api/admin/revoke-access",{method:"POST",body:f});openAccessManager();}
async function openTcode(code){let layouts=await api("/api/tcode/"+code+"/layouts");let tc=(await api("/api/tcodes/all")).find(x=>x.code==code)||{code:code,title:code,color:"#000"};let saved=localStorage.getItem("layout_"+code);let curLayout=layouts.find(l=>l.id==saved)||layouts.find(l=>l.is_default)||layouts[0];let records=await api("/api/tcode/"+code+"/records");$("root").innerHTML=shell(`<div class="flex justify-between mb-4"><div class="flex items-center gap-3"><div class="w-10 h-10 rounded-xl grid place-items-center text-white" style="background:${tc.color}"><i class="${tc.icon}"></i></div><div><div class="text-[11px] font-bold text-gray-500">${tc.code} • ${layouts.length} layouts</div><div class="text-[18px] font-bold">${tc.title}</div></div></div><button onclick="v('dash')" class="px-4 py-2 rounded-full bg-white border text-sm">Dashboard</button></div><div class="bg-white rounded-2xl border p-4 mb-4 flex items-center gap-4"><label class="text-sm font-bold">Layout:</label><select id="layoutSelect" onchange="switchLayout('${code}',this.value)" class="flex-1 border p-2.5 rounded-xl text-sm">${layouts.map(l=>`<option value="${l.id}" ${l.id==curLayout.id?'selected':''}>${l.layout_code} - ${l.title}</option>`).join('')}</select><a href="/api/tcode/${code}/export-excel" class="px-3 py-2 rounded-full bg-gray-100 text-sm">Export</a></div><div class="grid grid-cols-12 gap-5"><div class="col-span-5 bg-white rounded-2xl border p-5"><h3 class="font-bold text-sm">${curLayout.layout_code}</h3><div id="formArea" class="mt-4 grid grid-cols-3 gap-3">${curLayout.fields.map(f=>`<div class="col-span-${f.width||1}"><label class="text-[11px] font-semibold">${f.label}</label><input id="f_${f.key}" class="w-full border p-2.5 rounded-xl mt-1"></div>`).join('')}</div><div class="mt-4 flex gap-2"><button onclick="saveRecord('${code}')" class="flex-1 bg-black text-white py-3 rounded-xl">+ Save</button><button onclick="clearForm()" class="px-5 py-3 rounded-xl bg-gray-100">Clear</button></div></div><div class="col-span-7 bg-white rounded-2xl border p-5"><h3 class="font-bold text-sm">Records • ${records.length}</h3><div class="mt-3 overflow-auto max-h-[70vh] border rounded-xl"><table class="w-full text-xs"><thead class="bg-gray-50"><tr><th class="p-2">ID</th><th class="p-2">Data</th><th class="p-2">Action</th></tr></thead><tbody>${records.map(r=>`<tr class="border-b"><td class="p-2">#${r.id}</td><td class="p-2">${Object.entries(r.data).slice(0,3).map(([k,v])=>`${k}:${String(v).slice(0,20)}`).join(' ')}</td><td class="p-2"><button onclick="loadEdit(${r.id})" class="px-2 py-1 rounded-full bg-gray-100 text-[11px]">Edit</button><button onclick="delRecord('${code}',${r.id})" class="px-2 py-1 rounded-full bg-red-50 text-red-600 text-[11px]">Del</button></td></tr>`).join('')}</tbody></table></div></div></div>`);window._curLayouts=layouts;window._curLayout=curLayout;window._curRecords=records;}
function switchLayout(code,id){localStorage.setItem("layout_"+code,id);openTcode(code)}
let _editId=null;function clearForm(){_editId=null;document.querySelectorAll("#formArea input").forEach(e=>e.value="");}
function loadEdit(id){let rec=window._curRecords.find(x=>x.id==id);_editId=id;for(let k in rec.data){let el=document.getElementById("f_"+k);if(el)el.value=rec.data[k]}}
async function saveRecord(code){let data={};for(let f of window._curLayout.fields){let el=document.getElementById("f_"+f.key);if(el)data[f.key]=el.value}if(_editId){await api("/api/tcode/"+code+"/records/"+_editId,{method:"PUT",headers:{"Content-Type":"application/json"},body:JSON.stringify(data)});}else{await api("/api/tcode/"+code+"/records",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(data)});}clearForm();openTcode(code);}
async function delRecord(code,id){await api("/api/tcode/"+code+"/records/"+id,{method:"DELETE"});openTcode(code);}
async function openExcelUpload(code){$("root").innerHTML=shell(`<div class="bg-white rounded-2xl border p-6 max-w-4xl mx-auto"><div class="flex justify-between"><h2 class="font-bold text-lg">Upload Excel to ${code}</h2><button onclick="openTcode('${code}')" class="px-4 py-2 rounded-full bg-gray-100 text-sm">Back</button></div><div class="mt-6"><div class="border-2 border-dashed rounded-2xl p-8 text-center bg-gray-50"><input id="excelFile" type="file" accept=".xlsx,.xls" class="hidden" onchange="previewExcel('${code}')"><button onclick="document.getElementById('excelFile').click()" class="px-6 py-3 bg-black text-white rounded-xl">Choose Excel</button><div id="fileName" class="mt-2 text-sm"></div></div><div id="excelPreview" class="hidden"><div class="mt-3 overflow-auto border rounded-xl max-h-[40vh]"><table class="w-full text-xs" id="previewTable"></table></div><div id="mappingArea" class="mt-4 grid grid-cols-3 gap-3"></div><button onclick="saveExcelToPlanning('${code}')" class="mt-4 w-full bg-green-600 text-white py-3 rounded-xl">Load All Rows</button></div></div></div>`);}
let _uploadFile=null;async function previewExcel(code){let fileEl=$("excelFile");let f=new FormData();f.append("file",fileEl.files[0]);let res=await api("/api/admin/upload-excel-preview",{method:"POST",body:f});let schema=await api("/api/tcode/"+code+"/schema");let fields=schema.fields;let mappingHtml=fields.map(ff=>`<div><label class="text-[11px] font-semibold">${ff.label}</label><select id="map_${ff.key}" class="w-full border p-2 rounded-xl text-xs mt-1"><option value="">Ignore</option>${res.headers.map((h,i)=>`<option value="${i}">${h}</option>`).join('')}</select></div>`).join('');$("mappingArea").innerHTML=mappingHtml;let tableHtml=`<thead class="bg-gray-900 text-white"><tr>${res.headers.map(h=>`<th class="p-2 text-left">${h}</th>`).join('')}</tr></thead><tbody>${res.preview.map(r=>`<tr class="border-b">${r.map(c=>`<td class="p-2">${c}</td>`).join('')}</tr>`).join('')}</tbody>`;$("previewTable").innerHTML=tableHtml;$("excelPreview").classList.remove("hidden");_uploadFile=fileEl.files[0];}
async function saveExcelToPlanning(code){let schema=await api("/api/tcode/"+code+"/schema");let fields=schema.fields;let mapping={};fields.forEach(ff=>{let sel=$("map_"+ff.key);if(sel&&sel.value!=="")mapping[sel.value]=ff.key});let fd2=new FormData();fd2.append("file",_uploadFile);fd2.append("tcode_code",code);fd2.append("mapping",JSON.stringify(mapping));let saveRes=await api("/api/admin/upload-excel-save-file",{method:"POST",body:fd2});toast(`Loaded ${saveRes.saved} rows`);openTcode(code);}
async function logout(){await api("/api/logout",{method:"POST"});location.reload()}
</script></body></html>
"""
@app.get("/",response_class=HTMLResponse)
def index(): return HTML_PAGE
if __name__=="__main__":
    import uvicorn
    uvicorn.run(app,host="0.0.0.0",port=int(os.environ.get("PORT",8000)))
